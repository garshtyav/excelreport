# VBA-линтер для MicShot: ловит классы ошибок компиляции до сборки.
#   1) дублирующиеся Dim в процедуре          2) дублирующиеся процедуры (в одном модуле)
#   3) модульные Const/Declare после первой процедуры
#   4) вызовы необъявленных процедур ("Sub or Function not defined")
#   5) обращения к процедурам, удалённым при рефакторинге
# Источники: выгрузка из собранного .xlam (%TEMP%\MicShot_build\*_extracted.txt),
# По умолчанию проверяется build_micshot.ps1; --extracted проверяет выгрузку.
# Запуск: python _vba_lint.py   (выход 1 = есть проблемы)
import re, io, sys, os

raw = io.open('build_micshot.ps1', encoding='utf-8-sig').read()

def heredoc(name):
    m = re.search(r"\$" + name + r" = @'\r?\n(.*?)\r?\n'@", raw, re.S)
    return m.group(1) if m else ''

KW = set('''if then else elseif for next do loop while until select case end exit with set let
dim redim const public private sub function property declare option explicit attribute
on error goto resume call print write input line get put open me new stop return erase type enum
true false nothing and or not xor mod like is byval byref optional step each in as
long string integer boolean double variant object byte date single currency any'''.split())

BUILTINS = set('''len trim ltrim rtrim left right mid format dir instr instrrev split join replace
cstr clng cint cdbl cbool cvar csng isnumeric iserror isdate isempty isnull isarray typename
ubound lbound msgbox rgb hex val str chr chrw asc ascw now date time timer int fix abs sgn rnd
randomize err freefile lof eof seek environ iif array createobject getobject kill mkdir
rmdir curdir filelen filedatetime strconv space lcase ucase strreverse shell doevents beep
close unload savepicture deleteobject put get input write print seek lenb
vbinformation vbexclamation vbcritical vbquestion vbcrlf vbnullstring vbtextcompare
vbfromunicode vbunicode
me application activewindow activesheet selection activecell activeworkbook thisworkbook
intersect union workbooks worksheets sheets cells range rows columns
worksheet workbook range cells rows columns chart pivottable hyperlink tableobject window
rectanglelabel commandbar commandbarbutton
lbllive lblinterim lblstatus lblnote tbxlive btnstop btnshot btnhide
lblheader lbxfiles lblcount lblcomment tbxcomment btnsend btnskip'''.split())

REMOVED = ['livedebug', 'micshot_panelshot', 'updatepanelstatus', 'setstatus',
           'tbxlive', 'lblnote', 'lblstatus', 'btnstop', 'btnshot', 'btnhide',
           'extractwavsegment', 'wavdatabytes', 'gliveoffset']

def strip_nc(s):
    s = re.sub(r'"[^"]*"', '""', s)
    return re.sub(r"'.*$", '', s)

def join_continuations(code):
    out, buf = [], ''
    for ln in code.split('\n'):
        s = ln.rstrip()
        if s.endswith(' _'):
            buf += s[:-1] + ' '
            continue
        buf += s
        out.append(buf)
        buf = ''
    if buf:
        out.append(buf)
    return out

problems = []

def collect(src_name, code, procs, globals_out):
    proc = None
    seen_proc = False
    for m in re.finditer(r'Declare PtrSafe (?:Function|Sub)\s+(\w+)', code, re.I):
        globals_out.add(m.group(1).lower())
    for m in re.finditer(r'GoTo\s+(\w+)', code, re.I):
        globals_out.add(m.group(1).lower())
    for idx, line in enumerate(join_continuations(code), 1):
        st = strip_nc(line).strip()
        if not st or re.fullmatch(r'[A-Za-z_]\w*:', st):
            continue
        m = re.match(r'(?:Public |Private |Friend )?(?:Sub|Function)\s+(\w+)', st)
        if m and 'Declare' not in st:
            if proc:
                problems.append(src_name + ": '" + proc['name'] + "' не закрыт до строки " + str(idx))
            sig = st[m.end():]
            args = re.findall(r'(?:ByVal|ByRef|Optional|ParamArray)?\s*([A-Za-z_]\w*)(?:\(\))?\s+(?:As|\()', sig)
            proc = {'name': m.group(1), 'src': src_name, 'line': idx,
                    'args': set(a.lower() for a in args), 'dims': set(), 'calls': [],
                    'body': [], 'dim_line': {}}
            procs.setdefault(proc['name'].lower(), []).append(proc)
            seen_proc = True
            continue
        if re.match(r'End (Sub|Function)\b', st):
            if proc:
                # использование локала ДО его Dim = "Duplicate declaration" при компиляции
                for nm, dl in proc['dim_line'].items():
                    for li, lst in proc['body']:
                        if li < dl and re.search(r'(?<![\w.!"])' + nm + r'\b', lst) \
                           and not lst.lower().startswith('dim '):
                            problems.append(src_name + ": '" + proc['name'] + "': '" + nm +
                                            "' используется до объявления Dim (строки " + str(li) + " и " + str(dl) + ")")
                            break
            proc = None
            continue
        if not proc:
            gm = re.match(r'(?:Public |Private )?(?:Dim |Const )?\s*([A-Za-z_]\w*)\s+As\b', st) or \
                 re.match(r'(?:Public |Private )?Const\s+([A-Za-z_]\w*)\b', st)
            if gm:
                globals_out.add(gm.group(1).lower())
            if seen_proc and re.match(r'(?:Private |Public )?(?:Const|Declare\b)', st):
                problems.append(src_name + ": модульное объявление после первой процедуры (строка " + str(idx) + "): " + st[:60])
            continue
        dm = re.match(r'(?:Static )?Dim\s+(.+)', st, re.I)
        if dm:
            for part in dm.group(1).split(','):
                nm = re.split(r'\s+As\s+', part.strip(), flags=re.I)[0]
                nm = nm.split('(')[0].strip().rstrip('$%&#%!').lower()
                if re.match(r'^\w+$', nm or ''):
                    if nm in proc['dims']:
                        problems.append(src_name + ": '" + nm + "' объявлен дважды в '" + proc['name'] + "' (строка " + str(idx) + ")")
                    if nm in proc['args']:
                        problems.append(src_name + ": '" + nm + "' объявлен Dim поверх параметра в '" + proc['name'] + "' (строка " + str(idx) + ")")
                    proc['dims'].add(nm)
                    proc['dim_line'].setdefault(nm, idx)
            continue
        proc['body'].append((idx, st))
        first = re.match(r'([A-Za-z_]\w*)', st)
        if first and first.group(1).lower() not in KW:
            proc['calls'].append((first.group(1), idx))
        for cm in re.finditer(r'=\s*([A-Za-z_]\w*)\s*\(', st):
            proc['calls'].append((cm.group(1), idx))

xd = os.path.join(os.environ.get('TEMP', ''), 'MicShot_build')
extracted = []
for cn in ('MicShot', 'AppEvents', 'CSapiEvents', 'MicShotForm', 'MicShotSendForm'):
    f = os.path.join(xd, cn + '_extracted.txt')
    if os.path.exists(f):
        extracted.append((cn, io.open(f, encoding='utf-16').read()))
if '--extracted' in sys.argv and len(extracted) != 5:
    sys.exit("Expected all five exported modules for --extracted")
sources = extracted if '--extracted' in sys.argv and extracted else [
    ('MicShot', heredoc('vbaCode')),
    ('AppEvents', heredoc('vbaClass')),
    ('CSapiEvents', heredoc('vbaClassSapi')),
    ('MicShotForm', heredoc('vbaForm')),
    ('MicShotSendForm', heredoc('vbaSendForm'))]

procs = {}
all_globals = set()
for src_name, code in sources:
    collect(src_name, code, procs, all_globals)
all_globals.add('micshotform')
all_globals.add('micshotsendform')

# дубли процедур внутри одного модуля
seen = {}
for nm, places in procs.items():
    by_src = {}
    for pr in places:
        by_src.setdefault(pr['src'], []).append(pr['line'])
    for s, ls in by_src.items():
        if len(ls) > 1:
            problems.append(s + ": ПРОЦЕДУРА '" + nm + "' определена " + str(len(ls)) + " раз (строки " + str(ls) + ")")

# Exit Function в Sub / Exit Sub в Function = compile error
for src_name, code in sources:
    _cur = None
    for idx, line in enumerate(code.split('\n'), 1):
        s = line.strip()
        m2 = re.match(r'(?:Public |Private )?(Sub|Function)\s+(\w+)', s)
        if m2 and 'Declare' not in s:
            _cur = m2.group(1)
            continue
        if re.match(r'End (Sub|Function)', s):
            _cur = None
            continue
        if _cur == 'Sub' and 'Exit Function' in s:
            problems.append(src_name + ": Exit Function внутри Sub (строка " + str(idx) + "): " + s[:60])
        if _cur == 'Function' and re.search(r'Exit Sub', s):
            problems.append(src_name + ": Exit Sub внутри Function (строка " + str(idx) + "): " + s[:60])

# ByVal-массивы запрещены VBA ("Array argument must be ByRef" / ByRef mismatch)
for src_name, code in sources:
    for idx, line in enumerate(code.split('\n'), 1):
        if re.search(r'ByVal\s+\w+\(\)\s+As\s+', line):
            problems.append(src_name + ": ByVal-массив в сигнатуре (строка " + str(idx) + "): " + line.strip()[:70])

# вызовы: известные = все процедуры + глобалы + аргументы + локалы + билтины
defined = set(procs.keys()) | BUILTINS | all_globals
for nm, places in procs.items():
    for pr in places:
        known = defined | pr['args'] | pr['dims']
        for cname, cline in pr['calls']:
            if cname.lower() not in known:
                problems.append(pr['src'] + ": '" + pr['name'] + "': вызов необъявленного '" + cname + "' (строка " + str(cline) + ")")

# обращения к удалённому
for src_name, code in sources:
    for idx, line in enumerate(code.split('\n'), 1):
        for nm in REMOVED:
            if re.search(r'(?<![\w."])' + nm + r'\b', strip_nc(line), re.I):
                problems.append(src_name + ": обращение к удалённому '" + nm + "' (строка " + str(idx) + ")")

if problems:
    print("НАЙДЕНЫ ПРОБЛЕМЫ:")
    for p in problems:
        print("  " + p)
    sys.exit(1)
print("VBA LINT OK: " + str(len(sources)) + " модулей, дубли/вызовы/объявления - чисто")
