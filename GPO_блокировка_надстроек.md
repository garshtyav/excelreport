# Блокировка надстроек Excel через GPO (и сосуществование с MicShot)

Для домена AD с классическим Excel 365 ( desktop). Составлено по проекту MicShot;
точные имена значений реестра создаёт сам ADMX — редактором политик пользоваться
проще, чем править реестр руками.

## 0. Подготовка

1. Скачать с Microsoft «Administrative Template files (ADMX/ADML) for Microsoft 365 Apps/Office 2016/2019».
2. Скопировать `admx`-файлы в Central Store: `\\<домен>\SYSVOL\<домен>\Policies\PolicyDefinitions` (и `ru-RU`/`en-US` ADML).
3. После этого в GPMC появятся ветки: **Конфигурация пользователя → Административные шаблоны → Microsoft Excel 2016/365 → Trust Center** (Центр управления безопасностью).

## 1. Что чем блокируется (важно понимать классы)

| Класс | Пример | Чем управляется |
|---|---|---|
| Книги с макросами и **.xlam/.xla** (как MicShot) | MicShot.xlam | Политики **макросов (VBA)** + **доверенных расположений** + File Block |
| COM/VSTO-надстройки (панель в ленте, DLL) | Power Query-плагины и т.п. | Политики **Trust Center → Надстройки** («Disable all Application Add-ins» и др.) |

«Запретить установку» в чистом виде не существует — .xlam это просто файл, его
запретить копировать нельзя. Блокируется **запуск**: макросы не выполняются, если
файл не в доверенном расположении и/или не подписан.

## 2. Жёсткий вариант: запретить ВСЕ .xlam у всех

**Trust Center → File Block Settings → «Excel Add-in Files»** (.xlam/.xla):
«Blocked — open/save». Реестр (Policies):

```reg
[HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\16.0\Excel\Security\FileBlock]
"xlamfiles"=dword:00000002
```

Минус: блокирует и MicShot. Годится, если MicShot нужен только отдельным людям
(тогда политику вешать не на всех, а на нужный OU, и наоборот).

## 3. Гибкий вариант: макросы запрещены, разрешено только одобренное

Три политики (все: **User Config → Админ. шаблоны → Microsoft Excel → Trust Center**):

1. **Trust Center → Macro Settings → «VBA Macro Notification Settings»** = *«Disable all except digitally signed macros»*.
   Реестр: `HKCU\Software\Policies\Microsoft\Office\16.0\Excel\Security`, `vbawarnings`=4.
2. **Trust Center → Trusted Locations → «Prevent users from adding trusted locations»** = Вкл.
   (то самое место, куда пишет установщик MicShot — станет недоступно пользователю).
3. **Trust Center → Trusted Locations → «Trusted Location #1»** = путь, куда админ
   кладёт одобренные надстройки, например `C:\Program Files\ApprovedAddIns`,
   галка «Allow subfolders». Пользователь этот путь изменить/добавить не может.

Дальше два способа «пропустить» MicShot:

- **Способ А (подпись):** подписать MicShot.xlam корпоративным сертификатом
  (цепочка от CA, которому доверяет домен, — раздаётся через GPO в «Доверенные
  издатели/корневые ЦС»). При vbawarnings=4 подписанный файл с доверенным
  издателем выполняется. Самый чистый корпоративный путь.
- **Способ Б (доверенная папка):** не подписывать, а просто держать MicShot.xlam
  в GPO-доверенном расположении из п.3. Права на папку — только у админов:
  пользователь не сможет подменить/удалить файл (заодно решается защита от удаления).

Раздача файла и регистрации (в оба способа):

- Файл: копировать машинным механизмом (GPP Files, скрипт старта от SYSTEM, Intune/SCCM).
- Регистрация загрузки: OPEN-запись в `HKCU\...\Excel\Options` — через **GPP Registry**
  (User Config → Preferences → Windows Settings → Registry) или логон-скрипт.

## 4. Полный запрет COM/VSTO-надстроек (по желанию, на MicShot не влияет)

**Trust Center → Add-ins**: «Require Application Add-ins to be signed by Trusted
Publisher» и/или «Disable all Application Add-ins». Закрывает класс надстроек,
который ставится через кнопку «Перейти… → COM».

## 5. Ограничения и честные оговорки

- Политики пользователя лежат в `HKCU\...\Policies` — технически пользователь
  владеет своим кустом реестра и «умный» пользователь может переписать значение
  до следующего gpupdate. Политика = барьер от рядовых сотрудников, не от
  целенаправленного обхода. Жёстче — computer-side (HKLM) настройки там, где
  они есть, плюс NTFS-права и AppLocker/SRP на исполнение скриптов.
- Политики применяются при входе/`gpupdate /force`; проверка — `gpresult /h report.html`.
- Если на домене уже действуют такие политики, «пользовательский» установщик
  MicShot (install_micshot.bat) работать перестанет: он пишет доверенное
  расположение в HKCU, а это при «Prevent users from adding…» игнорируется.
  В домене MicShot ставит админ по схеме п. 3.
- Ветка реестра `16.0` одинакова для Office 2016–365 (labeled как Excel 2016 в GPMC).

## 6. Быстрая проверка на одном ПК (без домена, для понимания)

Те же ключи можно прописать в `HKCU\...\Policies` руками (.reg) и запустить Excel —
эффект тот же, что от политики:

```reg
Windows Registry Editor Version 5.00

; макросы: запрещены все, кроме подписанных
[HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\16.0\Excel\Security]
"vbawarnings"=dword:00000004

; запрет пользователю добавлять доверенные расположения
[HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\16.0\Excel\Security\trustedloc]
"NoUserTrustedLocations"=dword:00000001

; центральное доверенное расположение (пример)
[HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\16.0\Excel\Security\trustedloc\Location1]
"Path"="C:\\Program Files\\ApprovedAddIns\\"
"AllowSubFolders"=dword:00000001
"Description"="Одобренные ИТ надстройки"
```

(Имена значений trustedloc-ветки проверяйте по своему office16.admx — в разных
версиях ADMX они слегка отличаются; редактор политик избавляет от этой неопределённости.)
