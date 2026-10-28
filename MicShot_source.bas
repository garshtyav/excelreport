Attribute VB_Name = "MicShot"
' =====================================================================
'  MicShot v7.5 - надстройка Excel 365 (классический Excel для
'  Windows). Собирается скриптом build_micshot.ps1.
'
'  Кнопки на вкладке "НАДСТРОЙКИ" (раздел MicShot):
'    "? Отчёт: старт записи"        - одна кнопка-тумблер на старт и
'                                     стоп. Запускает сценарий
'                                     "рассказать и показать проблему":
'                                       * звук с микрофона -> narration.wav
'                                       * ПОЛНЫЙ журнал действий как код VBA
'                                         -> actions_log.bas (выделение, ввод,
'                                         форматирование, листы, книги, фильтры)
'                                       * окно живой расшифровки (MicShotForm):
'                                         статус, распознанный текст, кнопки
'                                       * скриншоты окна: по действиям и
'                                         периодически -> shot_*.png
'                                     Во время записи кнопка нажата и
'                                     читается "¦ Отчёт: стоп + открыть
'                                     папку" - повторный клик собирает
'                                     отчёт в Документы\MicShot_Report_<...>
'                                     и открывает папку.
'    "Скриншот сейчас"              - одиночный снимок окна (PNG)
'    "? Панель"                     - показать/скрыть окно расшифровки
'    "Ключ API…"                    - настройки распознавания
'
'  Живая расшифровка в окне: локальный SAPI (если есть русский движок)
'  плюс "живые тики" - каждые live_sec секунд (asr_config.txt) текущий
'  кусок записи stop+save уходит на сервер, текст появляется в окне;
'  микрофон тут же начинает писать следующий кусок, на "стопе" все
'  куски склеиваются в narration.wav.
'
'  Технически: звук - MCI (winmm.dll), журнал - события Application
'  плюс цикл опроса форматирования (у Excel нет событий на ленту),
'  скриншот - PrintWindow (PW_RENDERFULLCONTENT) + GDI+ PNG,
'  расшифровка речи - профиль из %APPDATA%\MicShot\asr_config.txt
'  (по умолчанию LAN-Whisper 192.168.1.77, запасной - облако Z.ai GLM-ASR).
'  Плавающих окон нет, кроме окна расшифровки (немодальное UserForm).
'  Сеть используется ТОЛЬКО для отправки звука в распознавание.
' =====================================================================

Private Const APP_NAME As String = "MicShot"
Private Const MIC_ALIAS As String = "MicShotRecorder"
' Win32: позиционирование и прозрачность полосы-субтитров
' (объявления уровня модуля должны идти ДО первой процедуры)
Private Const GWL_STYLE As Long = -16
Private Const GWL_EXSTYLE As Long = -20
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_THICKFRAME As Long = &H40000
Private Const WS_EX_LAYERED As Long = &H80000
Private Const LWA_ALPHA As Long = 2
' Скриншоты: по действиям пользователя (клики, правый/двойной клик, ввод,
' смена листа) + редкий периодический снимок (лента и действия вне листа
' событиями не видны - их ловит периодический).
Public Const SHOT_ON_ACTION As Boolean = True        ' снимок по действиям
Public Const ACTION_SHOT_DEBOUNCE_SEC As Long = 1    ' пауза после действия перед снимком
Public Const ACTION_SHOT_MIN_GAP_SEC As Long = 2     ' минимум между снимками по действиям
Public Const SHOT_PERIODIC_SEC As Long = 60          ' периодический снимок, 0 = выключить
' Лист "MicShot Отчёт" в активной книге: по умолчанию выключен - расшифровка
' живёт в transcript.txt (папка отчёта открывается сама после стопа).
Public Const CREATE_REPORT_SHEET As Boolean = False
' Полный журнал действий. Часть действий Excel событиями НЕ отдаёт вообще
' (форматирование, масштаб, автофильтр) - их ловит тихий цикл опроса:
' каждые FMT_POLL_SEC секунд отпечаток формата выделенного диапазона
' сравнивается с предыдущим, отличия уходят в журнал.
Public Const FMT_POLL_SEC As Long = 2               ' период опроса форматирования
Public Const VAL_CACHE_MAX_CELLS As Long = 50       ' кэш старых значений ("было -> стало")
' Слои 2-3 слежения за форматированием: опрос ВИДИМОЙ области (ловит кисть
' форматирования, вставку с форматом и прочее вне выделения) и базовый
' снимок формата на старте с итоговым диффом на стопе (гарантия полноты).
Public Const VIS_FMT_MAX_CELLS As Long = 3000       ' лимит ячеек видимой области на тик
Public Const BASELINE_MAX_CELLS As Long = 20000     ' лимит ячеек листа для базового снимка
' Живая расшифровка в окне панели: каждые LIVE_ASR_SEC секунд хвост звука
' уходит на сервер распознавания, текст появляется во время записи
' (переопределяется строкой live_sec=NN в asr_config.txt, 0 = выключить).
Public Const LIVE_ASR_SEC As Long = 30
' Первый живой тик - раньше периода: быстрая обратная связь в полосе
' (первая волна текста уже через ~10-15 с, дальше - каждые live_sec).
Public Const FIRST_LIVE_TICK_SEC As Long = 10
' Заголовок полосы-субтитров (по нему Win32 находит её окно).
Public Const SUBTITLE_CAPTION As String = "MicShot - субтитры"
' Версия (идёт в кейс при отправке) и boundary multipart-запросов.
Public Const APP_VERSION As String = "7.5"
Public Const HTTP_BOUNDARY As String = "----MicShotFormBoundary7MA4YWxkTrZu0gW"

' ---------------- состояние записи отчёта ----------------
Public gRecording As Boolean
Public gReportDir As String
Public gStartAt As Date
Public gShotCount As Long
Public gShotPending As Boolean
Public gNextActionShot As Date
Public gLastShot As Date
Public gNextShot As Date
Public gMicOk As Boolean
Public gLogFile As Integer
Public gLogLines As Long
Public gLastSelKey As String
Public gEvents As AppEvents
Public gSapiCtx As Object
Public gSapiEvents As CSapiEvents
Public gTranscriptLive As String
' слежение за форматированием и прочими "бесцветными" для событий действиями
Public gFmtRange As Range        ' диапазон, за форматом которого следим (последнее выделение)
Public gFmtFp As String          ' отпечаток формата gFmtRange
Public gNextFmtPoll As Date
Public gLastCalcAt As Date       ' дебаунс сообщений о пересчёте формул
Public gLastZoom As Long
Public gLastFilterKey As String  ' состояние автофильтра/фильтра активного листа
Public gValCache As Object       ' Scripting.Dictionary: адрес -> значение до изменения
Public gVisMap As Object         ' Dictionary "B5" -> токен формата видимой области (слой 2)
Public gVisSheet As String       ' книга|лист, к которому относится gVisMap
Public gBaseSnap As Object       ' Dictionary "книга|лист" -> токены базового снимка (слой 3)
Public gStartBookName As String  ' книга, активная на старте (для копии в отчёт)
Public gWbCopyName As String     ' имя файла-копии книги в папке отчёта
' панель живой расшифровки и живые тики сервера
Public gNextLiveTick As Date
Public gLiveHttp As Object
Public gLiveRequestAt As Date
Public gLiveRequestSeg As Long
Public gSegmentText As Object
Public gAudioMerged As Boolean
Public gLiveBusy As Boolean      ' предыдущий тик ещё не закончился
Public gLiveSegN As Long         ' сколько кусков записи уже сохранено
Public gLiveErrCount As Long     ' подряд неудачных обращений к серверу
Public gLiveGotText As Boolean   ' живая расшифровка уже дала текст
Public gLiveSec As Long          ' фактический период (константа или live_sec из конфига)
Public gLiveFirst As Boolean     ' первый тик ещё не сыграл (идёт с укороченным периодом)

' ---------------- типы Windows API ----------------
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type PICTDESC
    cbSize As Long
    picType As Long
    hBitmap As LongPtr
    hPalette As LongPtr
End Type

Private Type GUID_TYPE
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type GDIPLUS_STARTUP_INPUT
    GdiplusVersion As Long
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

#If VBA7 Then
Private Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" ( _
    ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
    ByVal uReturnLength As Long, ByVal hwndCallback As LongPtr) As Long

Private Declare PtrSafe Function GetWindowRect Lib "user32" ( _
    ByVal hWnd As LongPtr, ByRef lpRect As RECT) As Long
Private Declare PtrSafe Function GetWindowDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" ( _
    ByVal hDC As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
Private Declare PtrSafe Function SelectObject Lib "gdi32" ( _
    ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function BitBlt Lib "gdi32" ( _
    ByVal hDestDC As LongPtr, ByVal x As Long, ByVal y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As LongPtr, _
    ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare PtrSafe Function PrintWindow Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal hDCDest As LongPtr, ByVal nFlags As Long) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    Destination As Any, Source As Any, ByVal Length As Long)
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" ( _
    ByRef pPictDesc As PICTDESC, ByRef riid As GUID_TYPE, _
    ByVal fPictureOwnsHandle As Long, ByRef ipPic As Object) As Long

Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" ( _
    ByRef token As LongPtr, ByRef inputbuf As GDIPLUS_STARTUP_INPUT) As Long
Private Declare PtrSafe Sub GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr)
Private Declare PtrSafe Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" ( _
    ByVal hBmp As LongPtr, ByVal hPal As LongPtr, ByRef gpBitmap As LongPtr) As Long
Private Declare PtrSafe Function GdipSaveImageToFile Lib "gdiplus" ( _
    ByVal image As LongPtr, ByVal fileName As LongPtr, _
    ByRef clsidEncoder As GUID_TYPE, ByRef encoderParams As Any) As Long
Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal image As LongPtr) As Long
Private Declare PtrSafe Function CLSIDFromString Lib "ole32" ( _
    ByVal lpOLECHAR As LongPtr, ByRef pclsid As GUID_TYPE) As Long

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare PtrSafe Function FindWindowA Lib "user32" ( _
    ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function MoveWindow Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal x As Long, ByVal y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare PtrSafe Function GetWindowLongA Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLongA Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal crKey As Long, ByVal bAlpha As Byte, _
    ByVal dwFlags As Long) As Long
#End If

' ------------------------- жизненный цикл -------------------------

Public Sub Auto_Open()
    On Error Resume Next
    Application.CommandBars(APP_NAME).Delete
    Application.CommandBars("MicShotButtons").Delete   ' панель под старым именем
    On Error GoTo 0

    Dim cb As Office.CommandBar
    Set cb = Application.CommandBars.Add(APP_NAME, msoBarTop, False, True)
    ' --- КАСТОМИЗАЦИЯ: подписи кнопок правятся прямо здесь ---
    ' (имя раздела на вкладке "НАДСТРОЙКИ" - константа APP_NAME выше;
    '  первая кнопка - тумблер старт/стоп, её подписи меняются в UpdateToolbar)
    AddButton cb, "Отчёт: старт записи", "MicReportToggle", "Звук + полный журнал действий (VBA) + скриншоты + субтитры", "MicShotToggle"
    AddButton cb, "Скриншот сейчас", "TakeScreenshot", "Одиночный снимок окна Excel (PNG в Документы)", "MicShotShot"
    AddButton cb, "Панель", "MicShot_TogglePanel", "Показать/скрыть полосу расшифровки", "MicShotPanel"
    AddButton cb, "Ключ API…", "OpenAsrConfig", "Настройки распознавания речи (asr_config.txt)", "MicShotKey"
    AddButton cb, "Настройки отправки…", "OpenSendConfig", "API и авторизация отправки отчётов (send_config.txt)", "MicShotSendCfg"
    UpdateToolbar
    cb.Visible = True
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    If gRecording Then MicReportStop True   ' аварийное сохранение без диалогов
    mciSendString "close " & MIC_ALIAS, vbNullString, 0, 0
    Application.CommandBars(APP_NAME).Delete
End Sub

Private Sub AddButton(ByVal cb As Office.CommandBar, ByVal btnCaption As String, _
                      ByVal macroName As String, ByVal tip As String, _
                      Optional ByVal tag As String = "")
    Dim btn As Office.CommandBarButton
    Set btn = cb.Controls.Add(msoControlButton)
    btn.Style = msoButtonCaption
    btn.Caption = btnCaption
    btn.OnAction = "'" & ThisWorkbook.Name & "'!" & macroName
    btn.TooltipText = tip
    If Len(tag) > 0 Then btn.tag = tag
End Sub

' ------------- кнопка-тумблер "старт / стоп" -------------

Public Sub MicReportToggle()
    ' Одна кнопка на старт и стоп: повторный клик останавливает запись,
    ' собирает отчёт и открывает папку.
    If gRecording Then
        MicReportStop False
    Else
        MicReportStart
    End If
End Sub

Public Sub UpdateToolbar()
    ' Перекрашивает кнопку-тумблер под текущее состояние: в простое -
    ' "старт записи", во время записи - нажатая "стоп + открыть папку".
    On Error Resume Next
    Dim cb As Office.CommandBar, btn As Office.CommandBarButton, c As Object
    Set cb = Application.CommandBars(APP_NAME)
    If cb Is Nothing Then Exit Sub
    Set btn = Nothing
    For Each c In cb.Controls
        If TypeOf c Is Office.CommandBarButton Then
            If c.tag = "MicShotToggle" Then Set btn = c
        End If
    Next c
    If btn Is Nothing Then Exit Sub
    btn.Style = msoButtonIconAndCaption
    If gRecording Then
        btn.Caption = "Отчёт: стоп + открыть папку"
        btn.FaceId = 295                                  ' квадратик "стоп"
        btn.TooltipText = "Идёт запись: звук + журнал + скриншоты. Нажмите, чтобы собрать отчёт и открыть папку."
        btn.State = msoButtonDown                         ' визуально нажата
    Else
        btn.Caption = "Отчёт: старт записи"
        btn.FaceId = 294                                  ' красный кружок "запись"
        btn.TooltipText = "Звук + полный журнал действий (VBA) + скриншоты"
        btn.State = msoButtonUp
    End If
End Sub

' -------------------- отчёт о проблеме: старт ---------------------

Public Sub MicReportStart()
    On Error GoTo eh
    If gRecording Then
        MsgBox "Запись отчёта уже идёт: " & gReportDir, vbExclamation, APP_NAME
        Exit Sub
    End If

    gReportDir = MyDocuments() & "\MicShot_Report_" & Format(Now, "yyyy-mm-dd_hhnnss")
    MkDir gReportDir
    gStartAt = Now

    ' 1) микрофон (если недоступен - продолжаем без звука)
    gMicOk = False
    mciSendString "close " & MIC_ALIAS, vbNullString, 0, 0
    If mciSendString("open new type waveaudio alias " & MIC_ALIAS, vbNullString, 0, 0) = 0 Then
        If mciSendString("set " & MIC_ALIAS & " time format ms bitspersample 16 channels 1 " & _
                         "samplespersec 44100 bytespersec 88200 alignment 2", vbNullString, 0, 0) = 0 Then
            gMicOk = (mciSendString("record " & MIC_ALIAS, vbNullString, 0, 0) = 0)
        End If
        If Not gMicOk Then mciSendString "close " & MIC_ALIAS, vbNullString, 0, 0
    End If

    ' 2) журнал действий в стиле VBA
    gLogFile = FreeFile
    Open gReportDir & "\actions_log.bas" For Output As #gLogFile
    Print #gLogFile, "Attribute VB_Name = ""MicShotLog"""
    Print #gLogFile, "' === MicShot: журнал действий пользователя (VBA-стиль) ==="
    Print #gLogFile, "' начало записи: " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #gLogFile, "' файл можно импортировать в VBA и запустить RecordedActions -"
    Print #gLogFile, "' он приблизительно повторит показанные действия."
    Print #gLogFile, "Public Sub RecordedActions()"
    gLogLines = 0
    gLastSelKey = ""
    Set gValCache = CreateObject("Scripting.Dictionary")
    gLastCalcAt = 0
    gLastZoom = 0
    gLastFilterKey = ""
    Set gFmtRange = Nothing
    gFmtFp = ""
    gLiveBusy = False
    Set gLiveHttp = Nothing
    Set gSegmentText = CreateObject("Scripting.Dictionary")
    gAudioMerged = False
    gLiveSegN = 0
    gLiveErrCount = 0
    gLiveGotText = False
    gLiveSec = LiveAsrSec()
    gLiveFirst = True
    Set gVisMap = Nothing
    gVisSheet = ""
    Set gBaseSnap = CreateObject("Scripting.Dictionary")
    gStartBookName = ""
    gWbCopyName = ""

    Set gEvents = New AppEvents
    Set gEvents.App = Application

    ' 3) периодические скриншоты
    gShotCount = 0
    gShotPending = False
    gLastShot = 0
    gRecording = True
    If SHOT_PERIODIC_SEC > 0 Then ScheduleNextShot
    ScheduleFmtPoll
    If gLiveSec > 0 And gMicOk Then ScheduleLiveTick   ' первый тик - через FIRST_LIVE_TICK_SEC
    ' базовый снимок формата активной книги (слой 3); новые книги снимутся
    ' лениво при активации (App_WorkbookActivate)
    On Error Resume Next
    If Not ActiveWorkbook Is Nothing Then
        If ActiveWorkbook.Name <> ThisWorkbook.Name Then gStartBookName = ActiveWorkbook.Name
        SnapshotWorkbookFormat ActiveWorkbook
    End If
    ' следим за форматом сразу с текущего выделения
    On Error Resume Next
    If Not Selection Is Nothing Then
        If TypeOf Selection Is Range Then TrackSelection ActiveSheet, Selection
    End If
    ShowReportPanel

    gTranscriptLive = ""
    TryStartSpeech

    Application.StatusBar = APP_NAME & ": > ИДЁТ ЗАПИСЬ ОТЧЁТА" & _
        IIf(gMicOk, " (звук", " (БЕЗ ЗВУКА") & _
        " + действия + скриншоты по кликам). " & _
        "Говорите и показывайте проблему, в конце нажмите 'Отчёт: стоп'."
    UpdateToolbar
    Exit Sub
eh:
    Dim startError As String
    startError = Err.Description
    On Error Resume Next
    If gRecording Then MicReportStop True
    mciSendString "close " & MIC_ALIAS, vbNullString, 0, 0
    If gLogFile > 0 Then Close #gLogFile
    gLogFile = 0
    Set gEvents = Nothing
    gRecording = False
    MsgBox "MicReportStart: " & startError, vbExclamation, APP_NAME
End Sub

Private Sub ScheduleNextShot()
    If SHOT_PERIODIC_SEC <= 0 Then Exit Sub
    gNextShot = Now + TimeSerial(0, 0, SHOT_PERIODIC_SEC)
    Application.OnTime gNextShot, "'" & ThisWorkbook.Name & "'!MicShot_PeriodicShot"
End Sub

' -------------------- скриншот по таймеру и по действиям -------------------------

Public Sub MicShot_PeriodicShot()
    On Error Resume Next
    If Not gRecording Then Exit Sub
    TakeReportShot "per"
    Application.StatusBar = APP_NAME & ": > ИДЁТ ЗАПИСЬ ОТЧЁТА... скриншотов: " & gShotCount
    If gRecording Then ScheduleNextShot
End Sub

Public Sub MicShot_ActionShot()
    ' Отложенный снимок после действия (дебаунс: серия кликов = один снимок).
    On Error Resume Next
    gShotPending = False
    If Not gRecording Then Exit Sub
    TakeReportShot "act"
    Application.StatusBar = APP_NAME & ": > ИДЁТ ЗАПИСЬ ОТЧЁТА... скриншотов: " & gShotCount
End Sub

Public Sub RequestActionShot()
    ' Вызывается из событий действий; ставит отложенный снимок, если позволяют паузы.
    If Not gRecording Then Exit Sub
    If Not SHOT_ON_ACTION Then Exit Sub
    If gShotPending Then Exit Sub
    If gLastShot > 0 And (Now - gLastShot) < TimeSerial(0, 0, ACTION_SHOT_MIN_GAP_SEC) Then Exit Sub
    gShotPending = True
    gNextActionShot = Now + TimeSerial(0, 0, ACTION_SHOT_DEBOUNCE_SEC)
    Application.OnTime gNextActionShot, "'" & ThisWorkbook.Name & "'!MicShot_ActionShot"
End Sub

Private Sub TakeReportShot(ByVal kind As String)
    On Error Resume Next
    Dim hBmp As LongPtr
    hBmp = CaptureWindowToHbitmap(Application.hWnd)
    If hBmp = 0 Then Exit Sub
    gShotCount = gShotCount + 1
    gLastShot = Now
    Dim p As String
    p = gReportDir & "\shot_" & Format$(gShotCount, "000") & "_" & kind & "_" & Format$(Now, "hhnnss") & ".png"
    If SaveBitmapAsPng(hBmp, p) Then
        DeleteObject hBmp
    Else
        ' GDI+ не сработал - запасной вариант в BMP
        Dim pic As StdPicture
        Set pic = HbitmapToPicture(hBmp)   ' картинка забирает hBmp себе
        SavePicture pic, gReportDir & "\shot_" & Format$(gShotCount, "000") & "_" & kind & ".bmp"
    End If
End Sub

' -------------------- отчёт о проблеме: стоп ----------------------

Public Sub MicReportStop(Optional ByVal silent As Boolean = False)
    On Error GoTo eh
    If Not gRecording Then
        If Not silent Then MsgBox "Запись отчёта сейчас не ведётся.", vbInformation, APP_NAME
        Exit Sub
    End If
    gRecording = False
    UpdateToolbar

    On Error Resume Next
    Application.OnTime gNextShot, "'" & ThisWorkbook.Name & "'!MicShot_PeriodicShot", , False
    Application.OnTime gNextActionShot, "'" & ThisWorkbook.Name & "'!MicShot_ActionShot", , False
    Application.OnTime gNextFmtPoll, "'" & ThisWorkbook.Name & "'!MicShot_FmtPoll", , False
    Application.OnTime gNextLiveTick, "'" & ThisWorkbook.Name & "'!MicShot_LiveTick", , False
    On Error GoTo eh

    If Not gEvents Is Nothing Then
        Set gEvents.App = Nothing
        Set gEvents = Nothing
    End If

    ' итоговый дифф форматирования и копия книги - в журнал и папку отчёта
    ' (до очистки состояний gBaseSnap/gVisMap и до закрытия файла журнала)
    FinalFormatDiff
    SaveWorkbookCopy

    Set gFmtRange = Nothing
    gFmtFp = ""
    Set gValCache = Nothing
    Set gBaseSnap = Nothing
    Set gVisMap = Nothing
    gVisSheet = ""
    gLastFilterKey = ""
    gLastZoom = 0

    On Error Resume Next
    Unload MicShotForm                 ' окно живой расшифровки
    On Error GoTo eh

    If gLogFile > 0 Then
        Print #gLogFile, "End Sub"
        Close #gLogFile
        gLogFile = 0
    End If

    If gMicOk Then
        ' последний кусок записи: stop+save, затем склейка всех кусков
        ' (живые тики сохраняли промежуточные asr_seg_NN.wav)
        mciSendString "stop " & MIC_ALIAS, vbNullString, 0, 0
        gLiveSegN = gLiveSegN + 1
        mciSendString "save " & MIC_ALIAS & " """ & gReportDir & "\asr_seg_" & _
                      Format$(gLiveSegN, "00") & ".wav""", vbNullString, 0, 0
        mciSendString "close " & MIC_ALIAS, vbNullString, 0, 0
        gAudioMerged = MergeWavSegments()
    End If

    If Not gAudioMerged And gLiveSegN > 0 Then gAudioMerged = MergeWavSegments()
    PollLiveRequest
    On Error Resume Next
    If Not gLiveHttp Is Nothing Then gLiveHttp.Abort
    Set gLiveHttp = Nothing
    gLiveBusy = False
    On Error GoTo eh
    StopSpeech
    If Not silent Then RunCloudTranscription   ' расшифровка (профиль из asr_config.txt)
    Dim transcriptSaved As Boolean
    transcriptSaved = SaveTranscript()
    If gAudioMerged And transcriptSaved Then
        If Not gSegmentText Is Nothing Then
            If gSegmentText.Count = gLiveSegN Then KillSegmentFiles
        End If
    End If
    If Not gAudioMerged And gLiveSegN > 0 Then
        Application.StatusBar = APP_NAME & ": исходные аудиофрагменты сохранены отдельно."
    End If
    ' All segments must be merged, transcribed and saved before cleanup.
    If CREATE_REPORT_SHEET Then WriteReportSheet
    WriteReportReadme
    Application.StatusBar = False
    If Not silent Then
        Shell "explorer.exe """ & gReportDir & """", vbNormalFocus
        ShowSendDialog   ' комментарий текстом + список отправки + "Отправить" (заглушка)
    End If
    Exit Sub
eh:
    Application.StatusBar = False
    MsgBox "MicReportStop: " & Err.Description, vbExclamation, APP_NAME
End Sub

Private Sub WriteReportReadme()
    Dim f As Integer
    f = FreeFile
    Open gReportDir & "\README.txt" For Output As #f
    Print #f, "Отчёт MicShot о проблеме"
    Print #f, "========================"
    Print #f, "Начало записи: " & Format$(gStartAt, "yyyy-mm-dd hh:nn:ss")
    Print #f, "Конец записи:  " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Print #f, ""
    Print #f, "Файлы отчёта:"
    If gAudioMerged Then
        Print #f, "  narration.wav   - голосовое описание проблемы"
    ElseIf gLiveSegN > 0 Then
        Print #f, "  ВНИМАНИЕ: полную аудиозапись собрать не удалось."
        Print #f, "  asr_seg_*.wav - сохранённые исходные фрагменты; их нельзя удалять."
    Else
        Print #f, "  (звук не записан: микрофон был недоступен)"
    End If
    If gAudioMerged And Len(Dir$(gReportDir & "\asr_seg_*.wav")) > 0 Then
        Print #f, "  asr_seg_*.wav - исходники сохранены для повторного распознавания."
    End If
    Print #f, "  actions_log.bas - журнал действий как код VBA;"
    Print #f, "                    можно импортировать в VBA и запустить RecordedActions"
    If Len(Dir$(gReportDir & "\transcript.txt")) > 0 Then
        Print #f, "  transcript.txt  - расшифровка голоса в текст"
    End If
    Print #f, "  shot_NNN_<act|per>_*.png - скриншоты окна Excel: " & gShotCount & " шт." & vbCrLf & _
              "     act = по действиям (клики, ввод, смена листа), per = периодические"
    If Len(gWbCopyName) > 0 Then
        Print #f, "  " & gWbCopyName & " - копия книги на момент остановки записи" & vbCrLf & _
                  "     (папка отчёта самодостаточна - её можно целиком отправить на сервер)"
    End If
    Print #f, "  comment.txt     - текстовый комментарий пользователя (появляется,"
    Print #f, "                    если после остановки записи он его написал)"
    Close #f
End Sub

' -------------------- журнал действий (события) --------------------
' Ворота журнала - ОТКРЫТЫЙ ФАЙЛ, а не gRecording: итоговый дифф и копия
' книги пишутся на стопе уже после gRecording=False, пока журнал открыт.

Public Sub LogAction(ByVal codeLine As String, ByVal sheetName As String, _
                     Optional ByVal extraComment As String = "")
    If gLogFile = 0 Then Exit Sub
    If gLogLines >= 5000 Then Exit Sub
    gLogLines = gLogLines + 1
    Dim suffix As String
    suffix = "   ' "
    If Len(extraComment) > 0 Then suffix = suffix & extraComment & "; "
    Print #gLogFile, "    " & codeLine & suffix & Format$(Now, "hh:nn:ss") & " [" & sheetName & "]"
End Sub

Public Sub LogComment(ByVal sText As String, Optional ByVal sheetName As String = "")
    If gLogFile = 0 Then Exit Sub
    If gLogLines >= 5000 Then Exit Sub
    gLogLines = gLogLines + 1
    Dim suffix As String
    suffix = "   ' " & Format$(Now, "hh:nn:ss")
    If Len(sheetName) > 0 Then suffix = suffix & " [" & sheetName & "]"
    Print #gLogFile, "    ' " & sText & suffix
End Sub

Public Function VbaLiteral(ByVal v As Variant) As String
    Dim s As String
    If IsEmpty(v) Or IsNull(v) Then
        VbaLiteral = """"""
    ElseIf IsError(v) Then
        VbaLiteral = """<значение ошибки>"""
    ElseIf IsNumeric(v) Then
        VbaLiteral = CStr(v)
    ElseIf IsDate(v) Then
        VbaLiteral = """" & Format$(v, "yyyy-mm-dd hh:nn:ss") & """"
    Else
        s = CStr(v)
        If Len(s) > 200 Then s = Left$(s, 200) & "..."
        VbaLiteral = """" & Replace(s, """", """""") & """"
    End If
End Function

' --------------- слежение за форматированием и лентой ---------------
' У Excel НЕТ событий на форматирование, масштаб, автофильтр и прочие
' изменения через ленту. Их ловит тихий цикл опроса (MicShot_FmtPoll):
' каждые FMT_POLL_SEC секунд отпечаток формата последнего выделения
' сравнивается с предыдущим; отличия уходят в журнал. Дополнительно
' CheckFmtChange вызывается при каждой смене выделения.

Private Function SafeStr(ByVal v As Variant) As String
    On Error Resume Next
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    SafeStr = CStr(v)
End Function

Private Function HexColor(ByVal c As Long) As String
    On Error Resume Next
    Dim r As Long, g As Long, b As Long
    r = c And 255
    g = (c \ 256) And 255
    b = (c \ 65536) And 255
    HexColor = "#" & Right$("0" & Hex$(r), 2) & Right$("0" & Hex$(g), 2) & Right$("0" & Hex$(b), 2)
End Function

Public Function FmtFingerprint(ByVal rng As Range) As String
    ' Отпечаток формата: свойства первой ячейки (+ последней, если их
    ' несколько) плюс свойства строк/столбцов выделения.
    On Error Resume Next
    If rng Is Nothing Then Exit Function
    Dim f As String, c1 As Range, c2 As Range
    Set c1 = rng.Cells(1, 1)
    f = "numfmt=" & SafeStr(c1.NumberFormat)
    f = f & "|font=" & SafeStr(c1.Font.Name)
    f = f & "|size=" & SafeStr(c1.Font.Size)
    f = f & "|bold=" & SafeStr(c1.Font.Bold)
    f = f & "|italic=" & SafeStr(c1.Font.Italic)
    f = f & "|under=" & SafeStr(c1.Font.Underline)
    f = f & "|strike=" & SafeStr(c1.Font.Strikethrough)
    f = f & "|fcolor=" & HexColor(c1.Font.Color)
    f = f & "|fill=" & HexColor(c1.Interior.Color)
    f = f & "|pat=" & SafeStr(c1.Interior.Pattern)
    f = f & "|alignH=" & SafeStr(c1.HorizontalAlignment)
    f = f & "|alignV=" & SafeStr(c1.VerticalAlignment)
    f = f & "|wrap=" & SafeStr(c1.WrapText)
    f = f & "|orient=" & SafeStr(c1.Orientation)
    f = f & "|indent=" & SafeStr(c1.IndentLevel)
    f = f & "|merge=" & SafeStr(rng.MergeCells)
    f = f & "|bT=" & SafeStr(c1.Borders(8).LineStyle)    ' xlEdgeTop
    f = f & "|bB=" & SafeStr(c1.Borders(9).LineStyle)    ' xlEdgeBottom
    f = f & "|bL=" & SafeStr(c1.Borders(7).LineStyle)    ' xlEdgeLeft
    f = f & "|bR=" & SafeStr(c1.Borders(10).LineStyle)   ' xlEdgeRight
    If rng.Cells.CountLarge > 1 Then
        Set c2 = rng.Cells(rng.Rows.Count, rng.Columns.Count)
        f = f & "|numfmt2=" & SafeStr(c2.NumberFormat)
        f = f & "|bold2=" & SafeStr(c2.Font.Bold)
        f = f & "|fill2=" & HexColor(c2.Interior.Color)
        f = f & "|bT2=" & SafeStr(c2.Borders(8).LineStyle)
        f = f & "|bB2=" & SafeStr(c2.Borders(9).LineStyle)
        f = f & "|bL2=" & SafeStr(c2.Borders(7).LineStyle)
        f = f & "|bR2=" & SafeStr(c2.Borders(10).LineStyle)
    End If
    f = f & "|rowH=" & SafeStr(rng.Rows(1).RowHeight)
    f = f & "|colW=" & SafeStr(rng.Columns(1).ColumnWidth)
    f = f & "|rowHid=" & SafeStr(rng.Rows(1).Hidden)
    f = f & "|colHid=" & SafeStr(rng.Columns(1).Hidden)
    FmtFingerprint = f
End Function

Public Function FmtDiffDesc(ByVal oldFp As String, ByVal newFp As String) As String
    ' Человекочитаемое описание отличий двух отпечатков: "полужирный:
    ' False -> True; заливка: #FFFFFF -> #FFFF00". Повторы (та же правка
    ' у первой и последней ячейки диапазона) не дублируются.
    On Error Resume Next
    Dim a() As String, b() As String, i As Long, out As String, item As String
    a = Split(oldFp, "|")
    b = Split(newFp, "|")
    If UBound(a) <> UBound(b) Then Exit Function
    For i = 0 To UBound(a)
        If StrComp(a(i), b(i), vbBinaryCompare) <> 0 Then
            item = FmtTokenLabel(a(i), b(i))
            If InStr(1, "; " & out & "; ", "; " & item & "; ") = 0 Then
                If Len(out) > 0 Then out = out & "; "
                out = out & item
                If Len(out) > 300 Then
                    out = Left$(out, 300) & "..."
                    Exit For
                End If
            End If
        End If
    Next i
    FmtDiffDesc = out
End Function

Private Function FmtTokenLabel(ByVal oldTok As String, ByVal newTok As String) As String
    On Error Resume Next
    Dim p As Long, prop As String, oldV As String, newV As String
    p = InStr(oldTok, "=")
    prop = Left$(oldTok, p - 1)
    oldV = Mid$(oldTok, p + 1)
    p = InStr(newTok, "=")
    newV = Mid$(newTok, p + 1)
    Select Case prop
        Case "numfmt", "numfmt2": prop = "числовой формат"
        Case "font": prop = "шрифт"
        Case "size": prop = "размер шрифта"
        Case "bold", "bold2": prop = "полужирный"
        Case "italic": prop = "курсив"
        Case "under": prop = "подчёркивание"
        Case "strike": prop = "зачёркнутый"
        Case "fcolor": prop = "цвет текста"
        Case "fill", "fill2": prop = "заливка"
        Case "pat": prop = "узор заливки"
        Case "alignH": prop = "выравнивание по горизонтали"
        Case "alignV": prop = "выравнивание по вертикали"
        Case "wrap": prop = "перенос текста"
        Case "orient": prop = "ориентация"
        Case "indent": prop = "отступ"
        Case "merge": prop = "объединение ячеек"
        Case "bT", "bT2": prop = "граница сверху"
        Case "bB", "bB2": prop = "граница снизу"
        Case "bL", "bL2": prop = "граница слева"
        Case "bR", "bR2": prop = "граница справа"
        Case "rowH": prop = "высота строки"
        Case "colW": prop = "ширина столбца"
        Case "rowHid": prop = "строка скрыта"
        Case "colHid": prop = "столбец скрыт"
    End Select
    FmtTokenLabel = prop & ": " & oldV & " -> " & newV
End Function

Public Sub CheckFmtChange()
    ' Сравнивает текущий формат отслеживаемого диапазона с последним
    ' отпечатком; при изменении - строка в журнал и скриншот.
    On Error Resume Next
    If Not gRecording Then Exit Sub
    If gFmtRange Is Nothing Then Exit Sub
    If Len(gFmtFp) = 0 Then Exit Sub
    Dim nowFp As String
    nowFp = FmtFingerprint(gFmtRange)
    If Len(nowFp) = 0 Then
        ' диапазон перестал существовать (удалены строки/столбцы/лист)
        Set gFmtRange = Nothing
        gFmtFp = ""
        Exit Sub
    End If
    If nowFp = gFmtFp Then Exit Sub
    Dim d As String
    d = FmtDiffDesc(gFmtFp, nowFp)
    gFmtFp = nowFp
    If Len(d) > 0 Then
        LogComment "изменено форматирование " & gFmtRange.Address(False, False) & _
                   " (" & d & ")", gFmtRange.Worksheet.Name
        RequestActionShot
    End If
End Sub

Public Sub TrackSelection(ByVal Sh As Object, ByVal Target As Range)
    ' При каждой смене выделения: 1) проверить формат ПРЕДЫДУЩЕГО диапазона
    ' (лента форматирования событиями не видна), 2) начать следить за новым,
    ' 3) закэшировать его значения для журнала "было -> стало".
    On Error Resume Next
    If Not gRecording Then Exit Sub
    CheckFmtChange
    If Sh Is Nothing Then Exit Sub
    If Target Is Nothing Then Exit Sub
    If Not TypeOf Sh Is Worksheet Then Exit Sub
    Set gFmtRange = Target
    gFmtFp = FmtFingerprint(Target)
    CacheSelectionValues Target
End Sub

Private Sub CacheSelectionValues(ByVal rng As Range)
    On Error Resume Next
    If gValCache Is Nothing Then Set gValCache = CreateObject("Scripting.Dictionary")
    gValCache.RemoveAll
    If rng Is Nothing Then Exit Sub
    If rng.Cells.CountLarge > VAL_CACHE_MAX_CELLS Then Exit Sub
    Dim c As Range
    For Each c In rng.Cells
        gValCache(c.Address(False, False)) = VbaLiteral(c.Value)
    Next c
End Sub

Public Function CachedValue(ByVal addr As String) As String
    ' Старое значение ячейки (из кэша на момент выделения) - для "было -> стало".
    On Error Resume Next
    If gValCache Is Nothing Then Exit Function
    If gValCache.Exists(addr) Then CachedValue = gValCache(addr)
End Function

Public Sub CachedValueUpdate(ByVal addr As String, ByVal v As String)
    On Error Resume Next
    If gValCache Is Nothing Then Exit Sub
    gValCache(addr) = v
End Sub

' ---------- слои 2-3: видимая область и итоговый дифф формата ----------

Private Function HexColorV(ByVal v As Variant) As String
    ' Цвет как Variant (Null для смешанного внутри ячейки) -> "#RRGGBB".
    On Error Resume Next
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    HexColorV = HexColor(v)
End Function

Private Function FlattenArr(ByVal v As Variant, ByVal cnt As Long) As Variant
    ' Свойства диапазона дают СКАЛЯР (одна ячейка), 1-D или 2-D (row, col)
    ' массив. Нормализуем в 1-базный 1-D длины cnt (порядок Cells(): построчно).
    On Error GoTo fail
    Dim out() As Variant
    ReDim out(1 To cnt)
    If Not IsArray(v) Then
        out(1) = v
        FlattenArr = out
        Exit Function
    End If
    Dim r As Long, c As Long, k As Long, nR As Long, nC As Long
    nR = UBound(v, 1)
    nC = 0
    On Error Resume Next
    nC = UBound(v, 2)          ' у 1-D массива второго измерения нет
    On Error GoTo fail
    If nC = 0 Then
        For k = 1 To nR
            If k <= cnt Then out(k) = v(k)
        Next k
    Else
        For r = LBound(v, 1) To nR
            For c = LBound(v, 2) To nC
                k = k + 1
                If k <= cnt Then out(k) = v(r, c)
            Next c
        Next r
    End If
    FlattenArr = out
    Exit Function
fail:
End Function

Private Function BuildAreaTokens(ByVal rng As Range, ByVal maxCells As Long) As Variant
    ' Токены формата по ячейкам диапазона (одна область): свойства читаются
    ' МАССИВАМИ (по одному COM-вызову на свойство), токен ячейки - в том же
    ' "k=v|k=v" формате, что и FmtFingerprint, чтобы работал FmtDiffDesc.
    ' Empty = диапазон пуст/превышен лимит/ошибка.
    On Error GoTo done
    If rng Is Nothing Then Exit Function
    Dim n As Long
    n = rng.Cells.CountLarge
    If n <= 0 Or n > maxCells Then Exit Function
    Dim cnt As Long
    cnt = CLng(n)
    Dim vN As Variant, vS As Variant, vB As Variant, vI As Variant
    Dim vU As Variant, vK As Variant, vF As Variant, vC As Variant
    Dim vP As Variant, vA As Variant, vW As Variant, vM As Variant
    vN = rng.NumberFormat
    vS = rng.Font.Size
    vB = rng.Font.Bold
    vI = rng.Font.Italic
    vU = rng.Font.Underline
    vK = rng.Font.Strikethrough
    vF = rng.Font.Color
    vC = rng.Interior.Color
    vP = rng.Interior.Pattern
    vA = rng.HorizontalAlignment
    vW = rng.WrapText
    vM = rng.MergeCells
    vN = FlattenArr(vN, cnt)
    vS = FlattenArr(vS, cnt)
    vB = FlattenArr(vB, cnt)
    vI = FlattenArr(vI, cnt)
    vU = FlattenArr(vU, cnt)
    vK = FlattenArr(vK, cnt)
    vF = FlattenArr(vF, cnt)
    vC = FlattenArr(vC, cnt)
    vP = FlattenArr(vP, cnt)
    vA = FlattenArr(vA, cnt)
    vW = FlattenArr(vW, cnt)
    vM = FlattenArr(vM, cnt)
    If IsEmpty(vN) Or IsEmpty(vB) Or IsEmpty(vF) Then Exit Function   ' ошибка нормализации
    Dim toks() As String
    ReDim toks(1 To cnt)
    Dim i As Long
    For i = 1 To cnt
        toks(i) = "numfmt=" & SafeStr(vN(i)) & "|size=" & SafeStr(vS(i)) & _
                  "|bold=" & SafeStr(vB(i)) & "|italic=" & SafeStr(vI(i)) & _
                  "|under=" & SafeStr(vU(i)) & "|strike=" & SafeStr(vK(i)) & _
                  "|fcolor=" & HexColorV(vF(i)) & "|fill=" & HexColorV(vC(i)) & _
                  "|pat=" & SafeStr(vP(i)) & "|alignH=" & SafeStr(vA(i)) & _
                  "|wrap=" & SafeStr(vW(i)) & "|merge=" & SafeStr(vM(i))
    Next i
    BuildAreaTokens = toks
done:
End Function

Private Function IsCellInFmtRange(ByVal c As Range) As Boolean
    ' Ячейка внутри отслеживаемого выделения (слой 1 уже пишет такие).
    On Error Resume Next
    If gFmtRange Is Nothing Then Exit Function
    If c Is Nothing Then Exit Function
    If c.Worksheet.Parent.Name & "|" & c.Worksheet.Name <> _
       gFmtRange.Worksheet.Parent.Name & "|" & gFmtRange.Worksheet.Name Then Exit Function
    Dim hit As Range
    Set hit = Intersect(c, gFmtRange)
    IsCellInFmtRange = Not hit Is Nothing
End Function

Private Function ColLetter(ByVal n As Long) As String
    ' Номер столбца -> буквы (A, Z, AA...) без COM-вызовов.
    Dim s As String
    Do While n > 0
        s = Chr$(65 + ((n - 1) Mod 26)) & s
        n = (n - 1) \ 26
    Loop
    ColLetter = s
End Function

Private Function AreaDict(ByVal rng As Range, ByVal toks As Variant) As Object
    ' Dictionary "B5" -> токен. Адреса считаются арифметически по порядку
    ' Cells() (построчно) - без COM-вызова на ячейку.
    On Error Resume Next
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    If rng Is Nothing Or IsEmpty(toks) Then
        Set AreaDict = d
        Exit Function
    End If
    Dim nR As Long, nC As Long, r As Long, c As Long, i As Long
    Dim r0 As Long, c0 As Long
    nR = rng.Rows.Count
    nC = rng.Columns.Count
    r0 = rng.Row
    c0 = rng.Column
    For r = 1 To nR
        For c = 1 To nC
            i = i + 1
            d(ColLetter(c0 + c - 1) & (r0 + r - 1)) = toks(i)
        Next c
    Next r
    Set AreaDict = d
End Function

Public Sub CheckVisibleFmtChange()
    ' Слой 2: опрос видимой области активного листа - ловит изменения
    ' форматирования ВНЕ выделения (кисть, вставка с форматом, целые
    ' строки/столбцы). Сравнение ПО АДРЕСАМ ячеек: устойчиво к прокрутке
    ' и росту UsedRange; сравниваются ячейки, видимые и в прошлый тик;
    ' ячейки внутри выделения пропускает - их пишет слой 1.
    On Error Resume Next
    If Not gRecording Then Exit Sub
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub
    Dim ws As Worksheet, aw As Window, rng As Range
    Set ws = ActiveSheet
    Set aw = ActiveWindow
    If ws Is Nothing Or aw Is Nothing Then Exit Sub
    Set rng = Intersect(ws.UsedRange, aw.VisibleRange)
    If rng Is Nothing Then
        Set gVisMap = Nothing
        Exit Sub
    End If
    Dim toks As Variant, cur As Object, sheetKey As String
    toks = BuildAreaTokens(rng, VIS_FMT_MAX_CELLS)
    Set cur = AreaDict(rng, toks)
    If cur.Count = 0 Then
        Set gVisMap = Nothing
        Exit Sub
    End If
    sheetKey = ws.Parent.Name & "|" & ws.Name
    If gVisMap Is Nothing Or gVisSheet <> sheetKey Then
        Set gVisMap = cur
        gVisSheet = sheetKey
        Exit Sub
    End If
    Dim a As Variant, cell As Range
    Dim outSel As Long, inSel As Long, examined As Long
    Dim firstAddr As String, firstOld As String, firstNew As String
    For Each a In cur.Keys
        If gVisMap.Exists(a) Then
            If cur(a) <> gVisMap(a) Then
                examined = examined + 1
                Set cell = ws.Range(a)
                If IsCellInFmtRange(cell) Then
                    inSel = inSel + 1
                Else
                    outSel = outSel + 1
                    If Len(firstAddr) = 0 Then
                        firstAddr = CStr(a)
                        firstOld = gVisMap(a)
                        firstNew = cur(a)
                    End If
                End If
                If examined >= 200 Then Exit For
            End If
        End If
    Next a
    Set gVisMap = cur
    If outSel = 0 Then Exit Sub
    Dim extra As String, d As String
    If outSel > 1 Then extra = " и ещё " & (outSel - 1) & " ячеек"
    d = FmtDiffDesc(firstOld, firstNew)
    LogComment "изменено форматирование вне выделения: " & firstAddr & extra & _
               IIf(Len(d) > 0, " (" & d & ")", ""), ws.Name
    RequestActionShot
End Sub

Public Sub SnapshotWorkbookFormat(ByVal Wb As Workbook)
    ' Слой 3: базовый снимок формата всех листов книги (для итогового диффа).
    ' Вызывается на старте записи и лениво при активации новой книги.
    On Error Resume Next
    If Not gRecording Then Exit Sub
    If Wb Is Nothing Then Exit Sub
    If Wb.Name = ThisWorkbook.Name Then Exit Sub
    If gBaseSnap Is Nothing Then Set gBaseSnap = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet, key As String, toks As Variant
    For Each ws In Wb.Worksheets
        key = Wb.Name & "|" & ws.Name
        If Not gBaseSnap.Exists(key) Then
            toks = BuildAreaTokens(ws.UsedRange, BASELINE_MAX_CELLS)
            If IsEmpty(toks) Then
                gBaseSnap(key) = "SKIP"       ' пустой или слишком большой лист
            Else
                Set gBaseSnap(key) = AreaDict(ws.UsedRange, toks)
            End If
        End If
    Next ws
End Sub

Private Sub FinalFormatDiff()
    ' Слой 3, на "стопе": сравнение формата каждого снимавшегося листа
    ' с базовым снимком ПО АДРЕСАМ ячеек. Гарантия, что итоговый отчёт
    ' не пропустил изменений, невидимых для слоёв 1-2 (за экраном).
    On Error Resume Next
    If gBaseSnap Is Nothing Then Exit Sub
    If gBaseSnap.Count = 0 Then Exit Sub
    LogComment "=== ИТОГОВЫЙ ДИФФ ФОРМАТИРОВАНИЯ (сравнение со снимком на старте) ==="
    Dim k As Variant, parts() As String, Wb As Workbook, ws As Worksheet
    Dim base As Object, cur As Object, toks As Variant
    Dim a As Variant, nChanged As Long, newCells As Long, examined As Long
    Dim firstAddr As String, firstOld As String, firstNew As String
    For Each k In gBaseSnap.Keys
        parts = Split(CStr(k), "|")
        Set Wb = Workbooks(parts(0))
        If Wb Is Nothing Then
            LogComment "  книга закрыта, сравнение невозможно: " & parts(0)
        Else
            Set ws = Wb.Worksheets(parts(1))
            If ws Is Nothing Then
                LogComment "  лист удалён во время записи: " & parts(0) & " / " & parts(1)
            ElseIf VarType(gBaseSnap(k)) = vbString Then
                LogComment "  лист слишком большой для сравнения: " & parts(0) & " / " & parts(1)
            Else
                Set base = gBaseSnap(k)
                toks = BuildAreaTokens(ws.UsedRange, BASELINE_MAX_CELLS)
                Set cur = AreaDict(ws.UsedRange, toks)
                If cur.Count = 0 Then
                    LogComment "  лист стал слишком большим: " & parts(1)
                Else
                    nChanged = 0
                    newCells = 0
                    firstAddr = ""
                    examined = 0
                    For Each a In cur.Keys
                        If base.Exists(a) Then
                            If cur(a) <> base(a) Then
                                nChanged = nChanged + 1
                                If Len(firstAddr) = 0 Then
                                    firstAddr = CStr(a)
                                    firstOld = base(a)
                                    firstNew = cur(a)
                                End If
                            End If
                        Else
                            newCells = newCells + 1   ' рост данных/новые области
                        End If
                        examined = examined + 1
                        If examined >= 60000 Then Exit For
                    Next a
                    If nChanged = 0 And newCells = 0 Then
                        LogComment "  " & parts(0) & " / " & parts(1) & ": изменений форматирования нет"
                    Else
                        If nChanged > 0 Then
                            LogComment "  " & parts(0) & " / " & parts(1) & ": форматирование изменено в " & _
                                       nChanged & " ячейках, первая " & firstAddr & " (" & _
                                       FmtDiffDesc(firstOld, firstNew) & ")"
                        End If
                        If newCells > 0 Then
                            LogComment "  " & parts(0) & " / " & parts(1) & ": появились новые ячейки (" & _
                                       newCells & ") - рост данных или форматирование новых областей"
                        End If
                    End If
                End If
            End If
        End If
    Next k
    ' листы, созданные после старта (в базовом снимке их нет)
    Dim key2 As String
    For Each k In gBaseSnap.Keys
        parts = Split(CStr(k), "|")
        Set Wb = Workbooks(parts(0))
        If Not Wb Is Nothing Then
            For Each ws In Wb.Worksheets
                key2 = Wb.Name & "|" & ws.Name
                If Not gBaseSnap.Exists(key2) Then
                    LogComment "  лист создан после старта (сравнения нет): " & key2
                    gBaseSnap(key2) = "NEW"    ' чтобы не повторяться
                End If
            Next ws
        End If
    Next k
End Sub

Private Sub SaveWorkbookCopy()
    ' Кладёт копию книги пользователя в папку отчёта - папку можно
    ' целиком отправить на сервер, книга приложена.
    On Error Resume Next
    Dim Wb As Workbook
    Set Wb = ActiveWorkbook
    If Wb Is Nothing Or Wb.Name = ThisWorkbook.Name Then
        If Len(gStartBookName) > 0 Then Set Wb = Workbooks(gStartBookName)
    End If
    If Wb Is Nothing Then Exit Sub
    If Wb.Name = ThisWorkbook.Name Then Exit Sub
    Dim nm As String, dst As String
    nm = Wb.Name
    If InStrRev(nm, ".") = 0 Then nm = nm & ".xlsx"
    dst = gReportDir & "\" & nm
    Wb.SaveCopyAs dst
    If Len(Dir$(dst)) = 0 Then
        dst = gReportDir & "\book_" & Format$(Now, "hhnnss") & "_" & nm
        Wb.SaveCopyAs dst
    End If
    If Len(Dir$(dst)) > 0 Then
        gWbCopyName = nm
        LogComment "к отчёту приложена копия книги: " & nm
    End If
End Sub

Public Sub MicShot_FmtPoll()
    ' Тихий цикл опроса (OnTime каждые FMT_POLL_SEC секунд): ловит то,
    ' что Excel событиями не отдаёт - форматирование (выделение и видимая
    ' область), масштаб, фильтры - и держит полосу-субтитры у низа окна.
    On Error Resume Next
    If Not gRecording Then Exit Sub
    PollLiveRequest
    CheckFmtChange
    CheckVisibleFmtChange
    TrackWindowStuff
    PositionSubtitlesBar
    ScheduleFmtPoll
End Sub

Private Sub ScheduleFmtPoll()
    If Not gRecording Then Exit Sub
    gNextFmtPoll = Now + TimeSerial(0, 0, FMT_POLL_SEC)
    Application.OnTime gNextFmtPoll, "'" & ThisWorkbook.Name & "'!MicShot_FmtPoll"
End Sub

Private Sub TrackWindowStuff()
    On Error Resume Next
    Dim aw As Window
    Set aw = ActiveWindow
    If aw Is Nothing Then Exit Sub
    ' масштаб окна
    If gLastZoom = 0 Then
        gLastZoom = aw.Zoom
    ElseIf aw.Zoom <> gLastZoom Then
        LogComment "масштаб: " & gLastZoom & "% -> " & aw.Zoom & "%"
        gLastZoom = aw.Zoom
        RequestActionShot
    End If
    ' автофильтр и фильтрация активного листа
    If TypeOf ActiveSheet Is Worksheet Then
        Dim k As String, prevAuto As String, prevFlt As String
        k = ActiveSheet.AutoFilterMode & "|" & ActiveSheet.FilterMode
        If Len(gLastFilterKey) = 0 Then
            gLastFilterKey = k
        ElseIf k <> gLastFilterKey Then
            prevAuto = Left$(gLastFilterKey, InStr(gLastFilterKey, "|") - 1)
            prevFlt = Mid$(gLastFilterKey, InStr(gLastFilterKey, "|") + 1)
            If Left$(k, InStr(k, "|") - 1) <> prevAuto Then
                If Left$(k, InStr(k, "|") - 1) = "True" Then
                    LogComment "включён автофильтр", ActiveSheet.Name
                Else
                    LogComment "снят автофильтр", ActiveSheet.Name
                End If
                RequestActionShot
            ElseIf Mid$(k, InStr(k, "|") + 1) <> prevFlt Then
                If Mid$(k, InStr(k, "|") + 1) = "True" Then
                    LogComment "применён фильтр (скрыты строки)", ActiveSheet.Name
                Else
                    LogComment "фильтр снят (показаны все строки)", ActiveSheet.Name
                End If
                RequestActionShot
            End If
            gLastFilterKey = k
        End If
    End If
End Sub

' --------------- комментарий и отправка отчёта (MicShotSendForm) ---------------
' Отправка: папку отчёта пакуем в zip (Shell.Application), POST multipart
' на endpoint из send_config.txt. Ретраи с нарастающей паузой (только для
' устранимых ошибок: сеть/5xx/429), лог в %APPDATA%\MicShot\send_log.txt,
' итог в <папка отчёта>\send_result.txt. endpoint пуст = не настроено.

Public Sub WriteCommentFile(ByVal repDir As String, ByVal txt As String)
    ' Текстовый комментарий пользователя (для тех, кто без микрофона) ->
    ' comment.txt (UTF-8) в папке отчёта.
    On Error Resume Next
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText "Комментарий пользователя (" & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "):" & vbCrLf & txt
    st.SaveToFile repDir & "\comment.txt", 2
    st.Close
End Sub

Public Sub WriteSendConfigTemplate(ByVal cfg As String)
    ' Полный шаблон send_config.txt (синхронно с шаблоном в build_micshot.ps1).
    On Error Resume Next
    MkDir Environ$("APPDATA") & "\MicShot"
    Dim f As Integer
    f = FreeFile
    Open cfg For Output As #f
    Print #f, "# Настройки отправки отчётов в службу поддержки (MicShot)."
    Print #f, "# endpoint пуст = отправка не настроена: диалог покажет подсказку."
    Print #f, "#endpoint=https://support.example.com/api/cases"
    Print #f, "# Авторизация: none | bearer | basic | header"
    Print #f, "auth=none"
    Print #f, "# bearer: Authorization: Bearer <token>"
    Print #f, "# basic:  Authorization: Basic base64(<login>:<password>)"
    Print #f, "# header: заголовок <header_name> со значением <token>"
    Print #f, "token="
    Print #f, "login="
    Print #f, "password="
    Print #f, "header_name=X-API-Key"
    Print #f, "# Метка проекта в кейсе (опционально):"
    Print #f, "project="
    Print #f, "# Поле формы, в которое кладётся zip отчёта:"
    Print #f, "file_field=file"
    Print #f, "# Таймаут ответа и повторы (пауза растёт: delay * backoff^(N-1)):"
    Print #f, "timeout_sec=120"
    Print #f, "retry_count=3"
    Print #f, "retry_delay_sec=5"
    Print #f, "retry_backoff=2"
    Print #f, "# Лог отправок пишется в %APPDATA%\MicShot\send_log.txt,"
    Print #f, "# итог каждой попытки - в <папка отчёта>\send_result.txt"
    Close #f
End Sub

Public Sub OpenSendConfig()
    ' Кнопка "Настройки отправки...": открывает send_config.txt в Блокноте.
    On Error Resume Next
    Dim cfg As String
    cfg = Environ$("APPDATA") & "\MicShot\send_config.txt"
    If Len(Dir$(cfg)) = 0 Then WriteSendConfigTemplate cfg
    Shell "notepad.exe """ & cfg & """", vbNormalFocus
End Sub

Private Sub LoadSendConfig(ByRef endpoint As String, ByRef auth As String, ByRef token As String, _
                           ByRef login As String, ByRef password As String, ByRef headerName As String, _
                           ByRef project As String, ByRef fileField As String, _
                           ByRef timeoutSec As Long, ByRef retryCount As Long, _
                           ByRef retryDelay As Long, ByRef backoff As Long)
    On Error Resume Next
    endpoint = "": auth = "none": token = "": login = "": password = ""
    headerName = "X-API-Key": project = "": fileField = "file"
    timeoutSec = 120: retryCount = 3: retryDelay = 5: backoff = 2
    Dim f As Integer, s As String, cfgPath As String
    cfgPath = Environ$("APPDATA") & "\MicShot\send_config.txt"
    ' файл отсутствует = все значения по умолчанию (иначе Do/EOF на
    ' неоткрытом канале с Resume Next зацикливается навсегда)
    If Len(Dir$(cfgPath)) = 0 Then Exit Sub
    f = FreeFile
    Open cfgPath For Input As #f
    Do While Not EOF(f)
        Line Input #f, s
        s = Trim$(s)
        If Left$(s, 1) = "#" Then GoTo NextLine
        If Left$(s, 9) = "endpoint=" Then endpoint = Trim$(Mid$(s, 10))
        If Left$(s, 5) = "auth=" Then auth = Trim$(Mid$(s, 6))
        If Left$(s, 6) = "token=" Then token = Trim$(Mid$(s, 7))
        If Left$(s, 6) = "login=" Then login = Trim$(Mid$(s, 7))
        If Left$(s, 9) = "password=" Then password = Trim$(Mid$(s, 10))
        If Left$(s, 12) = "header_name=" Then headerName = Trim$(Mid$(s, 13))
        If Left$(s, 8) = "project=" Then project = Trim$(Mid$(s, 9))
        If Left$(s, 11) = "file_field=" Then fileField = Trim$(Mid$(s, 12))
        If Left$(s, 12) = "timeout_sec=" Then timeoutSec = CLng(Val(Mid$(s, 13)))
        If Left$(s, 12) = "retry_count=" Then retryCount = CLng(Val(Mid$(s, 13)))
        If Left$(s, 16) = "retry_delay_sec=" Then retryDelay = CLng(Val(Mid$(s, 17)))
        If Left$(s, 14) = "retry_backoff=" Then backoff = CLng(Val(Mid$(s, 15)))
NextLine:
    Loop
    Close #f
    If timeoutSec < 15 Then timeoutSec = 15
    If retryCount < 1 Then retryCount = 1
    If retryDelay < 1 Then retryDelay = 1
    If backoff < 1 Then backoff = 1
End Sub

Public Sub LogSend(ByVal level As String, ByVal msg As String)
    ' Журнал отправок (ANSI, читается Блокнотом): %APPDATA%\MicShot\send_log.txt
    On Error Resume Next
    MkDir Environ$("APPDATA") & "\MicShot"        ' ignore "already exists"
    Dim f As Integer
    f = FreeFile
    Open Environ$("APPDATA") & "\MicShot\send_log.txt" For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " [" & level & "] " & msg
    Close #f
End Sub

Private Sub WriteSendResult(ByVal repDir As String, ByVal ok As Boolean, _
                            ByVal attemptsMade As Long, ByVal detail As String)
    ' Итог отправки - файлом в папку отчёта: разберущемуся видно состояние.
    On Error Resume Next
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    If ok Then
        st.WriteText "SENT " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " attempts=" & attemptsMade & " " & detail
    Else
        st.WriteText "FAILED " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " attempts=" & attemptsMade & " error=" & detail
    End If
    st.SaveToFile repDir & "\send_result.txt", 2
    st.Close
End Sub

Private Function Base64Encode(ByVal sText As String) As String
    ' ANSI-байты строки -> base64 (для auth=basic).
    On Error GoTo fail
    Dim chars As String, b() As Byte, out As String
    Dim i As Long, n As Long, c0 As Long, c1 As Long, c2 As Long
    chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    b = StrConv(sText, vbFromUnicode)
    n = UBound(b) + 1
    i = 0
    Do While i < n
        c0 = b(i): c1 = 0: c2 = 0
        If i + 1 < n Then c1 = b(i + 1)
        If i + 2 < n Then c2 = b(i + 2)
        out = out & Mid$(chars, (c0 \ 4) + 1, 1)
        out = out & Mid$(chars, ((c0 And 3) * 16 + (c1 \ 16)) + 1, 1)
        If i + 1 < n Then
            out = out & Mid$(chars, ((c1 And 15) * 4 + (c2 \ 64)) + 1, 1)
        Else
            out = out & "="
        End If
        If i + 2 < n Then
            out = out & Mid$(chars, (c2 And 63) + 1, 1)
        Else
            out = out & "="
        End If
        i = i + 3
    Loop
    Base64Encode = out
    Exit Function
fail:
End Function

Private Function SafeFileLen(ByVal p As String) As Long
    On Error Resume Next
    SafeFileLen = FileLen(p)
End Function

Private Function ZipReportFolder(ByVal repDir As String, ByVal zipPath As String) As Boolean
    ' Пакует файлы отчёта в zip через PowerShell Compress-Archive:
    ' синхронно и скрыто (Shell CopyHere в скрытых сессиях молча не срабатывает).
    On Error GoTo failed
    If Len(Dir$(zipPath)) > 0 Then Kill zipPath
    Dim ws As Object, cmd As String
    Set ws = CreateObject("WScript.Shell")
    cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " & _
          """Compress-Archive -Path '" & repDir & "\*' -DestinationPath '" & zipPath & "' -Force"""
    ws.Run cmd, 0, True
    ZipReportFolder = (SafeFileLen(zipPath) > 22)
    Exit Function
failed:
    On Error Resume Next
End Function

Private Function ReadTextFileUtf8(ByVal p As String, ByVal maxChars As Long) As String
    On Error Resume Next
    If Len(Dir$(p)) = 0 Then Exit Function
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.LoadFromFile p
    ReadTextFileUtf8 = Left$(st.ReadText, maxChars)
    st.Close
End Function

Private Function AddCaseField(ByVal nm As String, ByVal v As String) As String
    AddCaseField = vbCrLf & "--" & HTTP_BOUNDARY & vbCrLf & _
                   "Content-Disposition: form-data; name=""" & nm & """" & vbCrLf & vbCrLf & v
End Function

Private Sub BuildCaseBody(ByVal zipPath As String, ByVal fileField As String, _
                          ByVal repDir As String, ByVal project As String, _
                          ByRef outBytes() As Byte)
    ' multipart: zip + поля кейса (id папки, проект, машина, пользователь,
    ' версия, время, комментарий). Тело 0-базное (см. урок в TranscribeChunk).
    On Error GoTo fail
    Dim pre As String, post As String
    pre = "--" & HTTP_BOUNDARY & vbCrLf & _
          "Content-Disposition: form-data; name=""" & fileField & """; filename=""" & _
          Mid$(zipPath, InStrRev(zipPath, "\") + 1) & """" & vbCrLf & _
          "Content-Type: application/zip" & vbCrLf & vbCrLf
    post = AddCaseField("report_id", Mid$(repDir, InStrRev(repDir, "\") + 1))
    If Len(project) > 0 Then post = post & AddCaseField("project", project)
    post = post & AddCaseField("machine", Environ$("COMPUTERNAME"))
    post = post & AddCaseField("user", Environ$("USERNAME"))
    post = post & AddCaseField("app_version", APP_VERSION)
    post = post & AddCaseField("recorded_at", Format$(Now, "yyyy-mm-dd hh:nn:ss"))
    Dim cmt As String
    cmt = ReadTextFileUtf8(repDir & "\comment.txt", 2000)
    If Len(cmt) > 0 Then post = post & AddCaseField("comment", cmt)
    post = post & vbCrLf & "--" & HTTP_BOUNDARY & "--" & vbCrLf

    Dim f As Integer, zipBytes() As Byte
    f = FreeFile
    Open zipPath For Binary Access Read As #f
    ReDim zipBytes(1 To LOF(f))
    Get #f, , zipBytes
    Close #f
    Dim preB() As Byte, postB() As Byte, body() As Byte
    preB = StrConv(pre, vbFromUnicode)
    postB = StrConv(post, vbFromUnicode)
    Dim nPre As Long, nZip As Long, nPost As Long
    nPre = UBound(preB) + 1
    nZip = UBound(zipBytes)
    nPost = UBound(postB) + 1
    ReDim body(0 To nPre + nZip + nPost - 1)
    CopyMemory body(0), preB(0), nPre
    CopyMemory body(nPre), zipBytes(1), nZip
    CopyMemory body(nPre + nZip), postB(0), nPost
    outBytes = body
    Exit Sub
fail:
End Sub



Private Function PostCaseZip(ByVal endpoint As String, ByRef body() As Byte, _
                             ByVal authHeader As String, ByVal authValue As String, _
                             ByVal timeoutSec As Long, ByRef httpStatus As Long, _
                             ByRef errText As String) As String
    On Error GoTo eh
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 30000, 30000, CLng(timeoutSec) * 1000, CLng(timeoutSec) * 1000
    http.Open "POST", endpoint, False
    If Len(authHeader) > 0 And Len(authValue) > 0 Then http.SetRequestHeader authHeader, authValue
    http.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & HTTP_BOUNDARY
    http.SetRequestHeader "User-Agent", "MicShot/" & APP_VERSION
    http.Send body
    httpStatus = http.Status
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write http.ResponseBody
    st.Position = 0
    st.Type = 2
    st.Charset = "utf-8"
    PostCaseZip = st.ReadText
    st.Close
    Exit Function
eh:
    If Err.Number = 12029 Or Err.Number = -2147012867 Or Err.Number = 12002 Or Err.Number = -2147012894 Then
        errText = "server unreachable (" & endpoint & ")"
    Else
        errText = Err.Number & ": " & Err.Description
    End If
End Function

Public Function SendReportToServer(ByVal repDir As String, Optional ByRef resultText As String = "") As Boolean
    ' Отправка кейса: конфиг -> zip -> ретраи -> лог -> send_result.txt.
    ' True = отправлено; False = нет (resultText объясняет: не настроено /
    ' сеть / HTTP). Ретраятся только устранимые ошибки (сеть, 5xx, 429/408).
    On Error GoTo failed
    LogSend "INFO", "SendReportToServer вызван: " & repDir
    If gRecording Then
        resultText = "идёт запись отчёта - сначала остановите"
        Exit Function
    End If
    Dim endpoint As String, auth As String, token As String
    Dim login As String, password As String, headerName As String
    Dim project As String, fileField As String
    Dim timeoutSec As Long, retryCount As Long, retryDelay As Long, backoff As Long
    LoadSendConfig endpoint, auth, token, login, password, headerName, _
                   project, fileField, timeoutSec, retryCount, retryDelay, backoff
    If Len(endpoint) = 0 Then
        resultText = "не настроено: endpoint пуст в %APPDATA%\MicShot\send_config.txt"
        LogSend "WARN", "отправка отменена: endpoint не настроен"
        Exit Function
    End If
    If Len(fileField) = 0 Then fileField = "file"
    LogSend "INFO", "отправка отчёта: " & repDir & " -> " & endpoint & _
                     " (auth=" & auth & ", retries=" & retryCount & ")"

    ' 1) zip
    Dim zipPath As String
    zipPath = Environ$("TEMP") & "\" & Mid$(repDir, InStrRev(repDir, "\") + 1) & ".zip"
    If Not ZipReportFolder(repDir, zipPath) Then
        resultText = "не удалось упаковать папку отчёта в zip"
        LogSend "ERROR", resultText
        WriteSendResult repDir, False, 0, resultText
        Exit Function
    End If
    LogSend "INFO", "zip готов: " & SafeFileLen(zipPath) & " байт"
    ' 2) тело и авторизация
    Dim body() As Byte
    BuildCaseBody zipPath, fileField, repDir, project, body
    If SafeFileLen(zipPath) = 0 Then
        resultText = "zip пропал перед отправкой"
        LogSend "ERROR", resultText
        WriteSendResult repDir, False, 0, resultText
        Exit Function
    End If
    Dim authHeader As String, authValue As String
    Select Case LCase$(auth)
        Case "bearer": authHeader = "Authorization": authValue = "Bearer " & token
        Case "basic":  authHeader = "Authorization": authValue = "Basic " & Base64Encode(login & ":" & password)
        Case "header": authHeader = headerName: authValue = token
    End Select

    Dim attempt As Long, attemptsMade As Long, delaySec As Long
    Dim resp As String, httpStatus As Long, errText As String
    Dim ok As Boolean, retryable As Boolean
    delaySec = retryDelay
    For attempt = 1 To retryCount
        If attempt > 1 Then
            LogSend "INFO", "пауза " & delaySec & " c перед попыткой " & attempt
            Dim untilT As Double
            untilT = Timer + delaySec
            Do While Timer < untilT
                Sleep 200
                DoEvents
            Loop
            delaySec = delaySec * backoff
        End If
        attemptsMade = attempt
        errText = ""
        httpStatus = 0
        resp = PostCaseZip(endpoint, body, authHeader, authValue, timeoutSec, httpStatus, errText)
        If Len(errText) = 0 And httpStatus >= 200 And httpStatus < 300 Then
            ok = True
            Exit For
        End If
        If Len(errText) > 0 Then
            retryable = True                                   ' сеть/таймаут - повтор
        Else
            retryable = (httpStatus = 429 Or httpStatus = 408 Or httpStatus >= 500)
        End If
        LogSend "WARN", "попытка " & attempt & "/" & retryCount & ": " & _
                 IIf(Len(errText) > 0, errText, "HTTP " & httpStatus & ": " & Left$(resp, 150)) & _
                 IIf(retryable, " - будет повтор", " - повтор бессмыслен")
        If Not retryable Then Exit For
    Next attempt

    On Error Resume Next
    Kill zipPath
    On Error GoTo failed

    If ok Then
        resultText = "HTTP " & httpStatus & IIf(Len(resp) > 0, ": " & Left$(resp, 120), "")
        LogSend "INFO", "ОТПРАВЛЕНО (" & resultText & ")"
        WriteSendResult repDir, True, attemptsMade, resultText
        SendReportToServer = True
    Else
        resultText = IIf(Len(errText) > 0, errText, "HTTP " & httpStatus & ": " & Left$(resp, 200))
        LogSend "ERROR", "не отправлено после " & attemptsMade & " попыток: " & resultText
        WriteSendResult repDir, False, attemptsMade, resultText
    End If
    Exit Function
failed:
    resultText = "ошибка: " & Err.Description
    LogSend "ERROR", resultText
End Function

Public Sub ShowSendDialog()
    ' Диалог после стопа: что уйдёт в поддержку + текстовый комментарий
    ' + кнопка "Отправить" / "Не отправлять".
    On Error Resume Next
    MicShotSendForm.Init gReportDir
    MicShotSendForm.Show vbModal
End Sub

Public Sub PrepareSendDialog()
    ' Служебное (автотест): загрузить форму и заполнить список БЕЗ показа.
    On Error Resume Next
    MicShotSendForm.Init gReportDir
End Sub

' --------------- полоса живой расшифровки (MicShotForm) ---------------
' Немодальная форма в виде СУБТИТРОВ: узкая полупрозрачная полоса без
' заголовка и кнопок у нижнего края окна Excel. Только текст расшифровки;
' управление - кнопками тулбара. Позиция/прозрачность поддерживаются
' циклом опроса (PositionSubtitlesBar), чтобы полоса прилипала к низу.

Public Sub ShowReportPanel()
    On Error Resume Next
    MicShotForm.ResetText
    MicShotForm.Show vbModeless
    ' подсказка в полосе, пока не пришёл первый текст/заметка
    MicShotForm.ShowNote "говорите - расшифровка появится здесь (первая волна через " & _
        IIf(gMicOk And gLiveSec > 0, "~" & FIRST_LIVE_TICK_SEC & "-" & gLiveSec & " с", "30 с") & ")"
    PositionSubtitlesBar
End Sub

Public Sub MicShot_TogglePanel()
    ' Кнопка "Панель" на тулбаре: показать/скрыть полосу расшифровки.
    On Error Resume Next
    If Not gRecording Then
        Application.StatusBar = APP_NAME & ": полоса появляется во время записи отчёта."
        Exit Sub
    End If
    If MicShotForm.Visible Then
        MicShotForm.Hide
    Else
        MicShotForm.Show vbModeless
        PositionSubtitlesBar
    End If
End Sub

Public Sub PositionSubtitlesBar()
    ' Прижимает полосу к низу окна Excel, снимает заголовок и делает
    ' полупрозрачной. Вызывается из цикла опроса (раз в 2 с).
    On Error Resume Next
    Dim hWnd As LongPtr
    hWnd = FindWindowA("ThunderDFrame", SUBTITLE_CAPTION)
    If hWnd = 0 Then Exit Sub
    Dim r As RECT
    GetWindowRect Application.hWnd, r
    Dim w As Long, h As Long, x As Long, y As Long
    w = (r.Right - r.Left) * 3 \ 4
    h = 84
    x = r.Left + ((r.Right - r.Left) - w) \ 2
    y = r.Bottom - h - 40                     ' чуть выше строки состояния
    SetWindowLongA hWnd, GWL_STYLE, GetWindowLongA(hWnd, GWL_STYLE) And (Not WS_CAPTION) And (Not WS_THICKFRAME)
    SetWindowLongA hWnd, GWL_EXSTYLE, GetWindowLongA(hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes hWnd, 0, 216, LWA_ALPHA
    MoveWindow hWnd, x, y, w, h, 1      ' после смены стиля - с перерисовкой
End Sub

Public Sub PanelNote(ByVal s As String)
    ' Служебная заметка в полосе (сервер не ответил и т.п.).
    On Error Resume Next
    MicShotForm.ShowNote s
End Sub

Public Function MicShot_PanelShown() As Boolean
    ' Служебное для автотеста: показана ли полоса.
    On Error Resume Next
    MicShot_PanelShown = (gRecording And MicShotForm.Visible)
End Function

Public Function MicShot_LiveText() As String
    ' Служебное для автотеста: накопленный текст расшифровки.
    On Error Resume Next
    MicShot_LiveText = gTranscriptLive
End Function

' --------------- живая расшифровка хвоста (сервер) ---------------

Private Function LiveAsrSec() As Long
    ' Период живых тиков: строка live_sec=NN в asr_config.txt,
    ' по умолчанию LIVE_ASR_SEC; 0 = живая расшифровка выключена.
    On Error Resume Next
    LiveAsrSec = LIVE_ASR_SEC
    Dim f As Integer, s As String
    f = FreeFile
    Open Environ$("APPDATA") & "\MicShot\asr_config.txt" For Input As #f
    Do While Not EOF(f)
        Line Input #f, s
        s = Trim$(s)
        If Left$(s, 9) = "live_sec=" Then LiveAsrSec = CLng(Val(Mid$(s, 10)))
    Loop
    Close #f
End Function

Private Sub ScheduleLiveTick()
    Dim sec As Long
    If gLiveFirst Then
        sec = FIRST_LIVE_TICK_SEC          ' быстрая первая волна текста
    Else
        sec = gLiveSec
    End If
    If sec <= 0 Then sec = gLiveSec
    gNextLiveTick = Now + TimeSerial(0, 0, sec)
    Application.OnTime gNextLiveTick, "'" & ThisWorkbook.Name & "'!MicShot_LiveTick"
End Sub

Public Sub MicShot_LiveTick()
    ' HTTP runs asynchronously; recording resumes BEFORE the request starts.
    On Error GoTo failed
    If Not gRecording Then Exit Sub
    gLiveFirst = False
    If gLiveSec > 0 Then ScheduleLiveTick
    PollLiveRequest
    If Not gMicOk Or gLiveBusy Or gLiveSec <= 0 Then Exit Sub
    If gLiveErrCount >= 3 Then Exit Sub
    Dim segPath As String, errText As String, ignored As String
    gLiveSegN = gLiveSegN + 1
    segPath = gReportDir & "\asr_seg_" & Format$(gLiveSegN, "00") & ".wav"
    mciSendString "stop " & MIC_ALIAS, vbNullString, 0, 0
    Dim saved As Long
    saved = mciSendString("save " & MIC_ALIAS & " """ & segPath & """", vbNullString, 0, 0)
    mciSendString "close " & MIC_ALIAS, vbNullString, 0, 0
    RestartMicRecording
    If saved <> 0 Or Len(Dir$(segPath)) = 0 Then
        gLiveErrCount = gLiveErrCount + 1
        PanelNote "Не удалось сохранить фрагмент аудио; исходные файлы будут сохранены."
        Exit Sub
    End If
    gLiveRequestSeg = gLiveSegN
    gLiveRequestAt = Now
    ignored = TranscribeChunk(segPath, errText, 15, True)
    If Len(errText) > 0 Then
        On Error Resume Next
        If Not gLiveHttp Is Nothing Then gLiveHttp.Abort
        Set gLiveHttp = Nothing
        On Error GoTo failed
        gLiveErrCount = gLiveErrCount + 1
        PanelNote Left$(errText, 100)
    Else
        gLiveBusy = True
    End If
    Exit Sub
failed:
    gLiveBusy = False
    gLiveErrCount = gLiveErrCount + 1
    PanelNote "Ошибка живой расшифровки: " & Err.Description
End Sub

Private Sub PollLiveRequest()
    On Error GoTo failed
    If Not gLiveBusy Or gLiveHttp Is Nothing Then Exit Sub
    ' WaitForResponse(0) checks completion without blocking Excel.
    If Not gLiveHttp.WaitForResponse(0) Then
        If Now - gLiveRequestAt < 30# / 86400# Then Exit Sub
        Err.Raise vbObjectError + 701, , "Таймаут живой расшифровки"
    End If
    Dim t As String
    If gLiveHttp.Status <> 200 Then Err.Raise vbObjectError + 702, , "HTTP " & gLiveHttp.Status
    t = ReadAsrResponse(gLiveHttp)
    gSegmentText(CStr(gLiveRequestSeg)) = t
    gLiveErrCount = 0
    If Len(Trim$(t)) > 0 Then
        gLiveGotText = True
        LiveSpeechCommit t
    End If
    PanelNote ""
    Set gLiveHttp = Nothing
    gLiveBusy = False
    Exit Sub
failed:
    PanelNote "Сбой распознавания; фрагмент будет повторён после стопа: " & Left$(Err.Description, 100)
    On Error Resume Next
    If Not gLiveHttp Is Nothing Then gLiveHttp.Abort
    Set gLiveHttp = Nothing
    gLiveBusy = False
    gLiveErrCount = gLiveErrCount + 1
End Sub

Private Sub RestartMicRecording()
    On Error GoTo failed
    gMicOk = False
    mciSendString "close " & MIC_ALIAS, vbNullString, 0, 0
    If mciSendString("open new type waveaudio alias " & MIC_ALIAS, vbNullString, 0, 0) <> 0 Then GoTo failed
    If mciSendString("set " & MIC_ALIAS & " time format ms bitspersample 16 channels 1 " & _
                    "samplespersec 44100 bytespersec 88200 alignment 2", vbNullString, 0, 0) <> 0 Then GoTo failed
    gMicOk = (mciSendString("record " & MIC_ALIAS, vbNullString, 0, 0) = 0)
    If gMicOk Then Exit Sub
failed:
    PanelNote "Микрофон остановлен. Уже записанные фрагменты сохранены."
End Sub

Private Function WavDataInfo(ByVal wavPath As String, ByRef dataOff As Long, _
                             ByRef dataSize As Long, ByRef headerLen As Long) As Boolean
    ' Walk RIFF chunks, including padding; never search audio bytes for "data".
    On Error GoTo failed
    Dim f As Integer, fileSize As Long, pos As Long, chunkSize As Long
    Dim magic As String * 4, wave As String * 4, chunkId As String * 4
    f = FreeFile
    Open wavPath For Binary Access Read As #f
    fileSize = LOF(f)
    If fileSize < 44 Then GoTo failed
    Get #f, 1, magic
    Get #f, 9, wave
    If magic <> "RIFF" Or wave <> "WAVE" Then GoTo failed
    pos = 13
    Do While pos <= fileSize - 7
        Get #f, pos, chunkId
        Get #f, pos + 4, chunkSize
        If chunkSize < 0 Or chunkSize > fileSize - pos - 7 Then GoTo failed
        If chunkId = "data" Then
            dataOff = pos + 8
            dataSize = chunkSize
            headerLen = pos + 7
            WavDataInfo = True
            Close #f
            Exit Function
        End If
        pos = pos + 8 + chunkSize + (chunkSize Mod 2)
    Loop
failed:
    On Error Resume Next
    If f > 0 Then Close #f
End Function

Private Function MergeWavSegments() As Boolean
    ' Склеивает asr_seg_*.wav (одинаковый формат PCM) в narration.wav.
    ' Вызывается на "стопе", пока куски ещё на диске.
    On Error GoTo failed
    Dim n As Long, i As Long
    n = gLiveSegN
    If n <= 0 Then Exit Function
    Dim paths() As String, offs() As Long, sizes() As Long
    ReDim paths(1 To n): ReDim offs(1 To n): ReDim sizes(1 To n)
    For i = 1 To n
        paths(i) = gReportDir & "\asr_seg_" & Format$(i, "00") & ".wav"
    Next i
    Dim headerLen As Long, h2 As Long, o2 As Long, s2 As Long
    For i = 1 To n
        If Not WavDataInfo(paths(i), o2, s2, h2) Then Exit Function
        offs(i) = o2: sizes(i) = s2
        If i = 1 Then headerLen = h2
    Next i
    Dim totalData As Long, pos As Long
    For i = 1 To n: totalData = totalData + sizes(i): Next i
    If totalData <= 0 Then Exit Function
    Dim out() As Byte, hdr() As Byte, f As Integer
    ReDim out(1 To headerLen + totalData)
    f = FreeFile
    Open paths(1) For Binary Access Read As #f
    ReDim hdr(1 To headerLen)
    Get #f, 1, hdr
    Close #f
    PutLong hdr, 5, headerLen - 8 + totalData      ' размер RIFF
    PutLong hdr, headerLen - 3, totalData          ' размер data
    CopyMemory out(1), hdr(1), headerLen
    pos = headerLen + 1
    Dim db() As Byte
    For i = 1 To n
        If sizes(i) > 0 Then
            ReDim db(1 To sizes(i))
            f = FreeFile
            Open paths(i) For Binary Access Read As #f
            Get #f, offs(i), db
            Close #f
            CopyMemory out(pos), db(1), sizes(i)
            pos = pos + sizes(i)
        End If
    Next i
    f = FreeFile
    Open gReportDir & "\narration.wav" For Binary Access Write As #f
    Put #f, , out
    Close #f
    f = 0
    If Not WavDataInfo(gReportDir & "\narration.wav", o2, s2, h2) Then Exit Function
    MergeWavSegments = (s2 = totalData And FileLen(gReportDir & "\narration.wav") = headerLen + totalData)
    Exit Function
failed:
    On Error Resume Next
    If f > 0 Then Close #f
End Function

Private Sub KillSegmentFiles()
    On Error Resume Next
    Dim sf As String
    sf = Dir$(gReportDir & "\asr_seg_*.wav")
    Do While Len(sf) > 0
        Kill gReportDir & "\" & sf
        sf = Dir$
    Loop
End Sub

' -------------------- панель и распознавание речи --------------------

Public Function SelfTestMerge(ByVal directory As String, ByVal segmentCount As Long) As Boolean
    If gRecording Then Err.Raise vbObjectError + 703, , "Recording is active"
    Dim oldDir As String, oldCount As Long
    oldDir = gReportDir
    oldCount = gLiveSegN
    On Error GoTo cleanup
    gReportDir = directory
    gLiveSegN = segmentCount
    SelfTestMerge = MergeWavSegments()
cleanup:
    gReportDir = oldDir
    gLiveSegN = oldCount
End Function

Public Sub CompileCheck()
    ' Служебный макрос: форсирует компиляцию классов проекта (для автотеста).
    On Error Resume Next
    Dim c As CSapiEvents
    Set c = New CSapiEvents
    Set c = Nothing
End Sub

Public Sub LiveSpeechCommit(ByVal t As String)
    ' Финальный кусок расшифровки (SAPI, живой тик сервера или тест):
    ' в буфер transcript и в окно панели.
    On Error Resume Next
    If Len(t) > 0 Then
        gTranscriptLive = gTranscriptLive & t & " "
        If gRecording Then MicShotForm.AppendLive t, True
    End If
End Sub

Private Sub LiveSpeechInterim(ByVal t As String)
    On Error Resume Next
    If Len(t) > 60 Then t = Left$(t, 60) & "..."
    If gRecording Then
        Application.StatusBar = APP_NAME & ": распознаю: " & t
        MicShotForm.AppendLive t, False
    End If
End Sub

Public Sub OpenAsrConfig()
    ' Открывает конфиг ключа в Блокноте (создаёт при отсутствии).
    On Error Resume Next
    Dim cfg As String
    cfg = Environ$("APPDATA") & "\MicShot\asr_config.txt"
    If Len(Dir$(cfg)) = 0 Then
        Dim f As Integer
        f = FreeFile
        Open cfg For Output As #f
        Print #f, "key="
        Print #f, "endpoint=https://api.z.ai/api/paas/v4/audio/transcriptions"
        Print #f, "model=glm-asr-2512"
        Close #f
    End If
    Shell "notepad.exe """ & cfg & """", vbNormalFocus
End Sub

Private Function AsrKey() As String
    Dim k As String, e As String, m As String, l As String, v As String, c As Long
    LoadAsrConfig k, e, m, l, v, c
    AsrKey = k
End Function

Private Function AsrEndpoint() As String
    Dim k As String, e As String, m As String, l As String, v As String, c As Long
    LoadAsrConfig k, e, m, l, v, c
    AsrEndpoint = e
End Function

Private Sub TryStartSpeech()
    On Error Resume Next
    If Not RussianEngineAvailable() Then
        Application.StatusBar = APP_NAME & ": локальное SAPI-распознавание русского не найдено (не критично)."
        Exit Sub
    End If
    Dim reco As Object
    Set reco = CreateObject("SAPI.SpSharedRecognizer")
    If reco Is Nothing Then Exit Sub
    Set gSapiCtx = reco.CreateRecoContext()
    If gSapiCtx Is Nothing Then Exit Sub
    Set gSapiEvents = New CSapiEvents
    Set gSapiEvents.RC = gSapiCtx
    Dim gram As Object
    Set gram = gSapiCtx.CreateGrammar(1)
    gram.DictationSetState 1
    Application.StatusBar = APP_NAME & ": SAPI-диктовка включена - речь будет попадать в расшифровку."
End Sub

Private Sub StopSpeech()
    On Error Resume Next
    Set gSapiEvents = Nothing
    Set gSapiCtx = Nothing
End Sub

Private Function RussianEngineAvailable() As Boolean
    ' Ищем среди установленных SAPI-распознавателей русский язык (LCID 0x419)
    On Error Resume Next
    Dim cat As Object, t As Object
    Set cat = CreateObject("SAPI.SpObjectTokenCategory")
    If cat Is Nothing Then Exit Function
    cat.SetId "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Recognizers\Tokens", False
    For Each t In cat.EnumerateTokens()
        If InStr(1, LCase$(t.GetAttribute("Language")), "419") > 0 Then
            RussianEngineAvailable = True
            Exit Function
        End If
    Next
End Function

Private Function SaveTranscript() As Boolean
    On Error GoTo failed
    Dim t As String
    t = gTranscriptLive
    If Len(Trim$(t)) = 0 Then Exit Function
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.Charset = "utf-8"
    st.Open
    st.WriteText t
    st.SaveToFile gReportDir & "\transcript.txt", 2
    st.Close
    SaveTranscript = True
    Exit Function
failed:
    On Error Resume Next
    If Not st Is Nothing Then st.Close
End Function

Private Sub WriteReportSheet()
    ' Лист "MicShot Отчёт" в активной книге с текстом расшифровки.
    On Error Resume Next
    Dim t As String
    t = Trim$(gTranscriptLive)
    If Len(t) = 0 Then Exit Sub
    Dim wbR As Workbook
    Set wbR = ActiveWorkbook
    If wbR Is Nothing Then Exit Sub

    Application.DisplayAlerts = False
    wbR.Worksheets("MicShot Отчёт").Delete
    Application.DisplayAlerts = True

    Dim wsT As Worksheet
    Set wsT = wbR.Worksheets.Add(After:=wbR.Worksheets(wbR.Worksheets.Count))
    wsT.Name = "MicShot Отчёт"
    wsT.Range("A1").Value = "Расшифровка речи (MicShot) - " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    wsT.Range("A1").Font.Bold = True
    wsT.Range("A3").Value = t
    wsT.Range("A3").WrapText = True
    wsT.Columns("A").ColumnWidth = 120
    wsT.Activate
End Sub

' -------------------- облачная расшифровка GLM-ASR --------------------

Public Sub TestAsrConfig()
    ' Проверка ключа: записывает 1 секунду с микрофона и отправляет в облако.
    ' Запускать через Alt+F8 -> TestAsrConfig.
    On Error GoTo eh
    Dim key As String, endpoint As String, model As String
    Dim lang As String, vad As String, chunkSec As Long
    LoadAsrConfig key, endpoint, model, lang, vad, chunkSec
    If InStr(1, endpoint, "z.ai", vbTextCompare) > 0 And Len(key) = 0 Then
        MsgBox "Для облака Z.ai нужен ключ API. Откройте файл:" & vbCrLf & _
               Environ$("APPDATA") & "\MicShot\asr_config.txt" & vbCrLf & vbCrLf & _
               "и впишите первой строкой:  key=ВАШ_КЛЮЧ" & vbCrLf & vbCrLf & _
               "Ключ Z.ai создаётся в кабинете z.ai (Manage API keys).", _
               vbInformation, APP_NAME
        Exit Sub
    End If
    Dim wavPath As String
    wavPath = Environ$("TEMP") & "\MicShot_asr_test.wav"
    mciSendString "close " & MIC_ALIAS, vbNullString, 0, 0
    mciSendString "open new type waveaudio alias " & MIC_ALIAS, vbNullString, 0, 0
    mciSendString "set " & MIC_ALIAS & " time format ms bitspersample 16 channels 1 " & _
                  "samplespersec 44100 bytespersec 88200 alignment 2", vbNullString, 0, 0
    mciSendString "record " & MIC_ALIAS, vbNullString, 0, 0
    Application.Wait Now + TimeSerial(0, 0, 1)
    mciSendString "stop " & MIC_ALIAS, vbNullString, 0, 0
    mciSendString "save " & MIC_ALIAS & " """ & wavPath & """", vbNullString, 0, 0
    mciSendString "close " & MIC_ALIAS, vbNullString, 0, 0

    Application.StatusBar = APP_NAME & ": отправляю тестовый файл в GLM-ASR..."
    Dim errText As String, t As String
    t = TranscribeChunk(wavPath, errText)
    Application.StatusBar = False
    Kill wavPath
    If Len(errText) > 0 Then
        MsgBox "Ошибка распознавания: " & errText & vbCrLf & vbCrLf & _
               "Если это HTTP 401/403 - ключ не подходит этому эндпоинту " & _
               "(нужен обычный ключ Z.ai с api.z.ai, не ключ coding-плана).", _
               vbExclamation, APP_NAME
    Else
        MsgBox "Связь с GLM-ASR работает!" & vbCrLf & _
               "Ответ на 1 c тишины: """ & t & """", vbInformation, APP_NAME
    End If
    Exit Sub
eh:
    Application.StatusBar = False
    MsgBox "TestAsrConfig: " & Err.Description, vbExclamation, APP_NAME
End Sub

Public Sub TranscribeFile(ByVal wavPath As String, ByVal outPath As String)
    ' Служебный: расшифровывает WAV текущим профилем (asr_config.txt), без UI.
    ' Результат - текст в outPath, либо "ERROR: ..." при неудаче.
    On Error Resume Next
    Dim errText As String, t As String
    t = TranscribeChunk(wavPath, errText)
    Dim f As Integer
    f = FreeFile
    Open outPath For Output As #f
    If Len(errText) > 0 Then
        Print #f, "ERROR: " & errText
    Else
        Print #f, t
    End If
    Close #f
End Sub

Public Sub RunCloudTranscription()
    If gRecording Then Exit Sub
    ' Reuse successful segments and retry EVERY missing one, in recording order.
    On Error GoTo failed
    If gSegmentText Is Nothing Then Set gSegmentText = CreateObject("Scripting.Dictionary")
    Dim i As Long, segPath As String, t As String, errText As String, Result As String
    For i = 1 To gLiveSegN
        If gSegmentText.Exists(CStr(i)) Then
            t = gSegmentText(CStr(i))
        Else
            segPath = gReportDir & "\asr_seg_" & Format$(i, "00") & ".wav"
            Application.StatusBar = APP_NAME & ": распознавание фрагмента " & i & " из " & gLiveSegN
            errText = ""
            t = TranscribeSavedSegment(segPath, errText)
            If Len(errText) > 0 Then
                t = "[фрагмент " & i & ": не удалось распознать: " & Left$(errText, 150) & "]"
            Else
                gSegmentText(CStr(i)) = t
            End If
        End If
        Result = Result & t & vbCrLf
    Next i
    If gLiveSegN > 0 Then gTranscriptLive = Result
    Exit Sub
failed:
    gTranscriptLive = gTranscriptLive & vbCrLf & "[Ошибка расшифровки: " & Err.Description & "]"
End Sub

Private Function TranscribeSavedSegment(ByVal wavPath As String, ByRef errText As String) As String
    ' Long final segments (including live_sec=0) still respect chunk_sec.
    On Error GoTo failed
    Dim key As String, endpoint As String, model As String, lang As String, vad As String
    Dim chunkSec As Long, n As Long, i As Long, t As String, Result As String
    LoadAsrConfig key, endpoint, model, lang, vad, chunkSec
    n = SplitWavIntoChunks(wavPath, chunkSec)
    If n = 0 Then
        errText = "Не удалось прочитать или разбить WAV"
        Exit Function
    End If
    For i = 1 To n
        t = TranscribeChunk(gReportDir & "\asr_chunk_" & Format$(i, "00") & ".wav", errText)
        If Len(errText) > 0 Then Exit For
        Result = Result & t & " "
    Next i
    If Len(errText) = 0 Then TranscribeSavedSegment = Trim$(Result)
cleanup:
    On Error Resume Next
    For i = 1 To n
        Kill gReportDir & "\asr_chunk_" & Format$(i, "00") & ".wav"
    Next i
    Exit Function
failed:
    errText = Err.Description
    Resume cleanup
End Function

Private Sub LoadAsrConfig(ByRef key As String, ByRef endpoint As String, ByRef model As String, _
                          ByRef lang As String, ByRef vad As String, ByRef chunkSec As Long)
    ' Профиль A (по умолчанию): LAN-Whisper 192.168.1.77 (проект IFRS17, без ключа).
    ' Профиль B: облако Z.ai GLM-ASR (нужен ключ и баланс). Настройки - asr_config.txt.
    On Error Resume Next
    endpoint = "http://192.168.1.77:8000/transcribe"
    model = ""
    lang = "ru"
    vad = "true"
    chunkSec = 60
    key = ""
    Dim f As Integer, s As String
    f = FreeFile
    Open Environ$("APPDATA") & "\MicShot\asr_config.txt" For Input As #f
    Do While Not EOF(f)
        Line Input #f, s
        s = Trim$(s)
        If Left$(s, 1) = "#" Then GoTo NextLine
        If Left$(s, 4) = "key=" Then key = Trim$(Mid$(s, 5))
        If Left$(s, 9) = "endpoint=" Then endpoint = Trim$(Mid$(s, 10))
        If Left$(s, 6) = "model=" Then model = Trim$(Mid$(s, 7))
        If Left$(s, 9) = "language=" Then lang = Trim$(Mid$(s, 10))
        If Left$(s, 11) = "vad_filter=" Then vad = Trim$(Mid$(s, 12))
        If Left$(s, 10) = "chunk_sec=" Then chunkSec = CLng(Val(Mid$(s, 11)))
NextLine:
    Loop
    Close #f
    If chunkSec <= 0 Then chunkSec = 60
End Sub

Private Function SplitWavIntoChunks(ByVal wavPath As String, ByVal maxSec As Long) As Long
    On Error GoTo eh
    Dim f As Integer, fileSize As Long
    f = FreeFile
    Open wavPath For Binary Access Read As #f
    fileSize = LOF(f)
    If fileSize < 44 Then Close #f: Exit Function
    Dim hdr() As Byte
    ReDim hdr(1 To fileSize)
    Get #f, , hdr
    Close #f

    Dim dataOff As Long, dataSize As Long, headerLen As Long, chunkBytes As Long
    If Not WavDataInfo(wavPath, dataOff, dataSize, headerLen) Then Exit Function
    chunkBytes = 88200# * maxSec       ' Recorder always uses 44100 Hz, mono, 16-bit PCM.
    If chunkBytes <= 0 Then Exit Function

    Dim outDir As String
    outDir = Left$(wavPath, InStrRev(wavPath, "\") - 1)

    Dim n As Long, pos As Long, take As Long
    pos = 1
    Do While pos <= dataSize
        n = n + 1
        take = chunkBytes
        If pos + take - 1 > dataSize Then take = dataSize - pos + 1

        Dim Ch() As Byte
        ReDim Ch(1 To headerLen)
        CopyMemory Ch(1), hdr(1), headerLen
        PutLong Ch, 5, headerLen - 8 + take      ' размер RIFF
        PutLong Ch, headerLen - 3, take          ' размер data

        Dim db() As Byte
        ReDim db(1 To take)
        CopyMemory db(1), hdr(dataOff - 1 + pos), take

        Dim g As Integer
        g = FreeFile
        Dim chunkPath As String
        chunkPath = outDir & "\asr_chunk_" & Format$(n, "00") & ".wav"
        If Len(Dir$(chunkPath)) > 0 Then Kill chunkPath
        Open chunkPath For Binary Access Write As #g
        Put #g, , Ch
        Put #g, , db
        Close #g
        pos = pos + take
    Loop
    SplitWavIntoChunks = n
    Exit Function
eh:
    SplitWavIntoChunks = 0
    On Error Resume Next
    If f > 0 Then Close #f
    If g > 0 Then Close #g
End Function

Private Sub PutLong(ByRef b() As Byte, ByVal idx As Long, ByVal v As Long)
    b(idx) = v And 255
    b(idx + 1) = (v \ 256) And 255
    b(idx + 2) = (v \ 65536) And 255
    b(idx + 3) = (v \ 16777216) And 255
End Sub

Private Function TranscribeChunk(ByVal wavPath As String, ByRef errText As String, _
                                 Optional ByVal timeoutSec As Long = 180, _
                                 Optional ByVal asyncRequest As Boolean = False) As String
    ' multipart POST аудиофайла на эндпоинт GLM-ASR. timeoutSec - таймаут
    ' ответа (живые тики шлют короткий хвост и ждут максимум 15 c).
    On Error GoTo eh
    Dim key As String, endpoint As String, model As String
    Dim lang As String, vad As String, chunkSec As Long
    LoadAsrConfig key, endpoint, model, lang, vad, chunkSec

    Dim boundary As String
    boundary = "----MicShotFormBoundary7MA4YWxkTrZu0gW"

    Dim f As Integer, fileBytes() As Byte
    f = FreeFile
    Open wavPath For Binary Access Read As #f
    ReDim fileBytes(1 To LOF(f))
    Get #f, , fileBytes
    Close #f

    Dim pre As String, post As String
    pre = "--" & boundary & vbCrLf & _
          "Content-Disposition: form-data; name=""file""; filename=""audio.wav""" & vbCrLf & _
          "Content-Type: audio/wav" & vbCrLf & vbCrLf
    post = ""
    If Len(model) > 0 Then
        post = post & vbCrLf & "--" & boundary & vbCrLf & _
               "Content-Disposition: form-data; name=""model""" & vbCrLf & vbCrLf & model
    End If
    If Len(lang) > 0 Then
        post = post & vbCrLf & "--" & boundary & vbCrLf & _
               "Content-Disposition: form-data; name=""language""" & vbCrLf & vbCrLf & lang
    End If
    If Len(vad) > 0 Then
        post = post & vbCrLf & "--" & boundary & vbCrLf & _
               "Content-Disposition: form-data; name=""vad_filter""" & vbCrLf & vbCrLf & vad
    End If
    post = post & vbCrLf & "--" & boundary & "--" & vbCrLf

    Dim preB() As Byte, postB() As Byte, body() As Byte
    preB = StrConv(pre, vbFromUnicode)
    postB = StrConv(post, vbFromUnicode)
    Dim nPre As Long, nFile As Long, nPost As Long
    nPre = UBound(preB) + 1
    nFile = UBound(fileBytes)
    nPost = UBound(postB) + 1
    ' Тело обязательно 0-базное: COM-маршалинг 1-базного массива сдвигает
    ' байты, и сервер перестаёт видеть поле "file" (HTTP 422).
    ReDim body(0 To nPre + nFile + nPost - 1)
    CopyMemory body(0), preB(0), nPre
    CopyMemory body(nPre), fileBytes(1), nFile
    CopyMemory body(nPre + nFile), postB(0), nPost

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.SetTimeouts 30000, 30000, CLng(timeoutSec) * 1000, CLng(timeoutSec) * 1000
    http.Open "POST", endpoint, asyncRequest
    If Len(key) > 0 Then http.SetRequestHeader "Authorization", "Bearer " & key
    http.SetRequestHeader "Content-Type", "multipart/form-data; boundary=" & boundary
    If asyncRequest Then Set gLiveHttp = http
    http.Send body
    If asyncRequest Then Exit Function
    ' ResponseText портит UTF-8 (нет charset у ответа) - декодируем байты сами
    Dim respBytes() As Byte, resp As String
    respBytes = http.ResponseBody
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write respBytes
    st.Position = 0
    st.Type = 2
    st.Charset = "utf-8"
    resp = st.ReadText
    st.Close
    If http.Status <> 200 Then
        If InStr(resp, "1113") > 0 Or InStr(resp, "Insufficient balance") > 0 Then
            errText = "НЕТ БАЛАНСА на аккаунте Z.ai (ошибка 1113). Пополните баланс на z.ai - " & _
                      "после этого расшифровка заработает сама, менять ничего не нужно."
        Else
            errText = "HTTP " & http.Status & ": " & Left$(resp, 300)
        End If
        Exit Function
    End If
    TranscribeChunk = ExtractJsonText(resp)
    Exit Function
eh:
    If Err.Number = 12029 Or Err.Number = -2147012867 Or Err.Number = 12002 Or Err.Number = -2147012894 Then
        errText = "сервер распознавания (" & endpoint & ") недоступен - вы в LAN/VPN?"
    Else
        errText = Err.Number & ": " & Err.Description
    End If
End Function

Private Function ReadAsrResponse(ByVal http As Object) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1
    st.Open
    st.Write http.ResponseBody
    st.Position = 0
    st.Type = 2
    st.Charset = "utf-8"
    ReadAsrResponse = ExtractJsonText(st.ReadText)
    st.Close
End Function

Private Function ExtractJsonText(ByVal resp As String) As String
    ' Достаёт значение поля "text" из JSON-ответа (без внешних библиотек).
    On Error Resume Next
    Dim i As Long, j As Long, out As String, Ch As String, n As String
    i = InStr(resp, """text""")
    If i = 0 Then Exit Function
    i = InStr(i, resp, ":")
    If i = 0 Then Exit Function
    i = i + 1
    Do While Mid$(resp, i, 1) = " "
        i = i + 1
    Loop
    If Mid$(resp, i, 1) <> """" Then Exit Function
    i = i + 1
    j = i
    Do While j <= Len(resp)
        Ch = Mid$(resp, j, 1)
        If Ch = "\" Then
            n = Mid$(resp, j + 1, 1)
            If n = """" Then
                out = out & """": j = j + 2
            ElseIf n = "n" Then
                out = out & " ": j = j + 2
            ElseIf n = "t" Then
                out = out & " ": j = j + 2
            ElseIf n = "u" Then
                out = out & ChrW$(CLng("&H" & Mid$(resp, j + 2, 4))): j = j + 6
            Else
                out = out & n: j = j + 2
            End If
        ElseIf Ch = """" Then
            Exit Do
        Else
            out = out & Ch: j = j + 1
        End If
    Loop
    ExtractJsonText = out
End Function

' --------------------------- скриншот -----------------------------

Public Sub TakeScreenshot()
    On Error GoTo eh
    Dim stamp As String
    stamp = Format$(Now, "yyyy-mm-dd_hhnnss")

    Dim hBmp As LongPtr
    hBmp = CaptureWindowToHbitmap(Application.hWnd)
    If hBmp = 0 Then Err.Raise vbObjectError + 513, , "Не удалось получить снимок окна"

    Dim pngPath As String
    pngPath = MyDocuments() & "\Excel_" & stamp & ".png"

    Dim saved As Boolean
    saved = SaveBitmapAsPng(hBmp, pngPath)
    If saved Then
        DeleteObject hBmp
    Else
        Dim pic As StdPicture
        Set pic = HbitmapToPicture(hBmp)
        pngPath = Environ$("TEMP") & "\MicShot_" & stamp & ".bmp"
        SavePicture pic, pngPath
    End If

    Application.StatusBar = APP_NAME & ": скриншот сохранён -> " & pngPath

    Dim insertedNote As String
    On Error Resume Next
    If TypeOf ActiveSheet Is Worksheet Then
        ActiveSheet.Shapes.AddPicture pngPath, msoFalse, msoTrue, _
                                      ActiveCell.Left, ActiveCell.Top, -1, -1
        insertedNote = vbCrLf & "(копия вставлена на активный лист)"
    End If
    On Error GoTo eh

    MsgBox "Скриншот сохранён:" & vbCrLf & pngPath & insertedNote, vbInformation, APP_NAME
    Exit Sub
eh:
    Application.StatusBar = False
    MsgBox "TakeScreenshot: " & Err.Description, vbExclamation, APP_NAME
End Sub

Public Sub SelfTestBmp(ByVal bmpPath As String)
    ' Служебный макрос самопроверки захвата окна (вызывается build-скриптом, без UI).
    On Error GoTo eh
    Dim hBmp As LongPtr
    hBmp = CaptureWindowToHbitmap(Application.hWnd)
    If hBmp = 0 Then Err.Raise vbObjectError + 513, , "capture failed"
    Dim pic As StdPicture
    Set pic = HbitmapToPicture(hBmp)   ' картинка забирает hBmp себе
    SavePicture pic, bmpPath
    Exit Sub
eh:
    Dim f As Integer
    f = FreeFile
    Open Environ$("TEMP") & "\MicShot_selftest.err" For Output As #f
    Print #f, "SelfTestBmp: ошибка " & Err.Number & ": " & Err.Description
    Close #f
End Sub

' ------------------ захват окна и сохранение PNG -------------------

Private Function CaptureWindowToHbitmap(ByVal hWnd As LongPtr) As LongPtr
    Dim r As RECT
    GetWindowRect hWnd, r
    Dim w As Long, h As Long
    w = r.Right - r.Left
    h = r.Bottom - r.Top

    Dim srcWnd As LongPtr, hDCSrc As LongPtr
    If w > 0 And h > 0 Then
        srcWnd = hWnd
        hDCSrc = GetWindowDC(srcWnd)
    Else
        w = 800: h = 600                ' окно скрыто - снимаем весь экран
        srcWnd = 0
        hDCSrc = GetWindowDC(0)
    End If

    Dim hDCMem As LongPtr, hBmp As LongPtr, hOld As LongPtr
    hDCMem = CreateCompatibleDC(hDCSrc)
    hBmp = CreateCompatibleBitmap(hDCSrc, w, h)
    hOld = SelectObject(hDCMem, hBmp)
    ' Excel 365 рисует окно через GPU: обычный BitBlt даёт пустую (белую)
    ' клиентскую область. PrintWindow с PW_RENDERFULLCONTENT (2) получает
    ' реальное содержимое; BitBlt оставлен как запасной вариант.
    If PrintWindow(srcWnd, hDCMem, 2) = 0 Then
        BitBlt hDCMem, 0, 0, w, h, hDCSrc, 0, 0, &HCC0020    ' SRCCOPY
    End If
    SelectObject hDCMem, hOld
    DeleteDC hDCMem
    ReleaseDC srcWnd, hDCSrc

    CaptureWindowToHbitmap = hBmp
End Function

Public Function SaveBitmapAsPng(ByVal hBmp As LongPtr, ByVal pngPath As String) As Boolean
    Dim gsi As GDIPLUS_STARTUP_INPUT
    gsi.GdiplusVersion = 1
    Dim token As LongPtr
    If GdiplusStartup(token, gsi) <> 0 Then Exit Function
    Dim gpBmp As LongPtr, clsid As GUID_TYPE
    If GdipCreateBitmapFromHBITMAP(hBmp, 0, gpBmp) = 0 Then
        CLSIDFromString StrPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), clsid
        SaveBitmapAsPng = (GdipSaveImageToFile(gpBmp, StrPtr(pngPath), clsid, ByVal 0&) = 0)
        GdipDisposeImage gpBmp
    End If
    GdiplusShutdown token
End Function

Private Function HbitmapToPicture(ByVal hBmp As LongPtr) As StdPicture
    Dim pd As PICTDESC
    pd.cbSize = LenB(pd)
    pd.picType = 1                                   ' PICTYPE_BITMAP
    pd.hBitmap = hBmp
    pd.hPalette = 0

    Dim iid As GUID_TYPE
    iid.Data1 = &H7BF80980                           ' IID_IPicture
    iid.Data2 = &HBF32
    iid.Data3 = &H101A
    iid.Data4(0) = &H8B: iid.Data4(1) = &HBB
    iid.Data4(2) = &H0:  iid.Data4(3) = &HAA
    iid.Data4(4) = &H0:  iid.Data4(5) = &H30
    iid.Data4(6) = &HC:  iid.Data4(7) = &HAB

    Dim unk As Object
    If OleCreatePictureIndirect(pd, iid, 1, unk) = 0 Then
        Set HbitmapToPicture = unk
    End If
End Function

Private Function MyDocuments() As String
    MyDocuments = CreateObject("WScript.Shell").SpecialFolders("MyDocuments")
End Function

