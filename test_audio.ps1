# Tests real VBA WAV merging without a microphone or ASR server.
param([string]$XlamPath = (Join-Path $env:APPDATA 'Microsoft\AddIns\MicShot.xlam'))
$ErrorActionPreference = 'Stop'
$testRoot = Join-Path $env:TEMP ('MicShot_audio_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $testRoot | Out-Null
function Write-TestWav([string]$Path, [byte[]]$Pcm) {
    $stream = [IO.File]::Create($Path)
    $writer = [IO.BinaryWriter]::new($stream)
    try {
        $writer.Write([Text.Encoding]::ASCII.GetBytes('RIFF'))
        $writer.Write([int](36 + $Pcm.Length))
        $writer.Write([Text.Encoding]::ASCII.GetBytes('WAVEfmt '))
        $writer.Write([int]16)
        $writer.Write([int16]1); $writer.Write([int16]1)
        $writer.Write([int]44100); $writer.Write([int]88200)
        $writer.Write([int16]2); $writer.Write([int16]16)
        $writer.Write([Text.Encoding]::ASCII.GetBytes('data'))
        $writer.Write([int]$Pcm.Length)
        $writer.Write($Pcm)
    } finally { $writer.Dispose() }
}
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.DisplayAlerts = $false
    $addin = $excel.Workbooks.Open($XlamPath)
    $macro = "'$($addin.Name)'!SelfTestMerge"
    foreach ($count in @(3, 105)) {
        $dir = Join-Path $testRoot "segments_$count"
        New-Item -ItemType Directory -Path $dir | Out-Null
        $expected = [Collections.Generic.List[byte]]::new()
        for ($i = 1; $i -le $count; $i++) {
            $pcm = [byte[]]@(($i % 256), 0, ($i % 256), 0)   # скобки: PS5.1 не парсит % внутри [byte[]]@( )
            $expected.AddRange($pcm)
            Write-TestWav (Join-Path $dir ('asr_seg_{0:00}.wav' -f $i)) $pcm
        }
        if (-not $excel.Run($macro, $dir, $count)) { throw "Merge failed: $count segments" }
        $merged = [IO.File]::ReadAllBytes((Join-Path $dir 'narration.wav'))
        if ($merged.Length -ne (44 + $expected.Count)) { throw 'Incorrect WAV size' }
        if ([BitConverter]::ToInt32($merged, 40) -ne $expected.Count) { throw 'Incorrect data size' }
        for ($i = 0; $i -lt $expected.Count; $i++) {
            if ($merged[44 + $i] -ne $expected[$i]) { throw "Wrong segment order at byte $i" }
        }
        Write-Host "PASS: $count segments, exact PCM order and WAV sizes"
    }
    $dir = Join-Path $testRoot 'broken'
    New-Item -ItemType Directory -Path $dir | Out-Null
    Write-TestWav (Join-Path $dir 'asr_seg_01.wav') ([byte[]]@(1, 0))
    [IO.File]::WriteAllBytes((Join-Path $dir 'asr_seg_02.wav'), [byte[]]@(0, 1))
    if ($excel.Run($macro, $dir, 2)) { throw 'Truncated WAV was accepted' }
    if (-not (Test-Path (Join-Path $dir 'asr_seg_01.wav'))) { throw 'Original segment deleted' }
    Remove-Item -LiteralPath (Join-Path $dir 'asr_seg_02.wav')
    if ($excel.Run($macro, $dir, 2)) { throw 'Missing WAV was accepted' }
    Write-Host 'PASS: truncated and missing segments rejected; originals retained'
} finally {
    if ($null -ne $excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    Write-Host "Fixtures: $testRoot"
}
