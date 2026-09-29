# Строит компактный справочник конфигурации (docs\map\*.md) из XML-выгрузки src\cf.
# Запуск из папки системы (C:\1C\<система>): powershell -NoProfile -File ..\tools\build-map.ps1
param(
    [string]$Src = (Join-Path (Get-Location) 'src\cf'),
    [string]$Out = (Join-Path (Get-Location) 'docs\map')
)
$ErrorActionPreference = 'Stop'
$Src = (Resolve-Path $Src).Path
New-Item -ItemType Directory -Force $Out | Out-Null
$utf8 = New-Object System.Text.UTF8Encoding($false)

function ReadText($path) { [System.IO.File]::ReadAllText($path, $utf8) }
function WriteMd($name, $lines) { [System.IO.File]::WriteAllLines((Join-Path $Out $name), [string[]]$lines, $utf8) }
function Esc($s) { if ($null -eq $s) { '' } else { ($s -replace '\|', '\|' -replace '\r?\n', ' ').Trim() } }

# Только блок <Properties> верхнего уровня (до ChildObjects), чтобы не цеплять реквизиты
function Props($text) {
    $i = $text.IndexOf('<ChildObjects')
    if ($i -gt 0) { $text.Substring(0, $i) } else { $text }
}
function Tag($props, $tag) {
    $m = [regex]::Match($props, "<$tag>([^<]*)</$tag>")
    if ($m.Success) { [System.Net.WebUtility]::HtmlDecode($m.Groups[1].Value) } else { $null }
}
function Synonym($props) {
    $m = [regex]::Match($props, '<Synonym>\s*<v8:item>\s*<v8:lang>ru</v8:lang>\s*<v8:content>([^<]*)</v8:content>')
    if ($m.Success) { [System.Net.WebUtility]::HtmlDecode($m.Groups[1].Value) } else { '' }
}
function Items($props, $tag) {
    $m = [regex]::Match($props, "(?s)<$tag>(.*?)</$tag>")
    if (-not $m.Success) { return @() }
    @([regex]::Matches($m.Groups[1].Value, '<xr:Item[^>]*>([^<]*)</xr:Item>') | ForEach-Object { $_.Groups[1].Value })
}

$types = [ordered]@{
    Subsystems = 'Подсистема'; CommonModules = 'ОбщийМодуль'; SessionParameters = 'ПараметрСеанса'; Roles = 'Роль'
    CommonAttributes = 'ОбщийРеквизит'; ExchangePlans = 'ПланОбмена'; FilterCriteria = 'КритерийОтбора'
    EventSubscriptions = 'ПодпискаНаСобытие'; ScheduledJobs = 'РегламентноеЗадание'; FunctionalOptions = 'ФункциональнаяОпция'
    FunctionalOptionsParameters = 'ПараметрФО'; DefinedTypes = 'ОпределяемыйТип'; SettingsStorages = 'ХранилищеНастроек'
    CommonForms = 'ОбщаяФорма'; CommonCommands = 'ОбщаяКоманда'; CommandGroups = 'ГруппаКоманд'; CommonTemplates = 'ОбщийМакет'
    XDTOPackages = 'ПакетXDTO'; WebServices = 'ВебСервис'; HTTPServices = 'HTTPСервис'; WSReferences = 'WSСсылка'
    IntegrationServices = 'СервисИнтеграции'; Bots = 'Бот'; Constants = 'Константа'; Catalogs = 'Справочник'
    Documents = 'Документ'; DocumentNumerators = 'Нумератор'; DocumentJournals = 'ЖурналДокументов'; Enums = 'Перечисление'
    Reports = 'Отчет'; DataProcessors = 'Обработка'; InformationRegisters = 'РегистрСведений'
    AccumulationRegisters = 'РегистрНакопления'; ChartsOfCharacteristicTypes = 'ПВХ'; ChartsOfAccounts = 'ПланСчетов'
    AccountingRegisters = 'РегистрБухгалтерии'; ChartsOfCalculationTypes = 'ПВР'; CalculationRegisters = 'РегистрРасчета'
    BusinessProcesses = 'БизнесПроцесс'; Tasks = 'Задача'
}
# Английское имя типа в ссылках MDObjectRef -> русское
$refTypes = @{
    Catalog = 'Справочник'; Document = 'Документ'; InformationRegister = 'РегистрСведений'; AccumulationRegister = 'РегистрНакопления'
    AccountingRegister = 'РегистрБухгалтерии'; CalculationRegister = 'РегистрРасчета'; Constant = 'Константа'; Enum = 'Перечисление'
    Report = 'Отчет'; DataProcessor = 'Обработка'; CommonModule = 'ОбщийМодуль'; ChartOfCharacteristicTypes = 'ПВХ'
    ChartOfAccounts = 'ПланСчетов'; ChartOfCalculationTypes = 'ПВР'; BusinessProcess = 'БизнесПроцесс'; Task = 'Задача'
    ExchangePlan = 'ПланОбмена'; DocumentJournal = 'ЖурналДокументов'; CommonForm = 'ОбщаяФорма'; CommonCommand = 'ОбщаяКоманда'
    Role = 'Роль'; Subsystem = 'Подсистема'; FilterCriterion = 'КритерийОтбора'; WebService = 'ВебСервис'; HTTPService = 'HTTPСервис'
}
function RuRef($ref) {
    $p = $ref.Split('.', 2)
    if ($p.Count -eq 2 -and $refTypes.ContainsKey($p[0])) { "$($refTypes[$p[0]]).$($p[1])" } else { $ref }
}

# ---------- Индекс всех объектов ----------
$objects = @{}   # folder -> list of @{Name;Syn;Props}
$index = @('# Все объекты: тип.Имя — синоним', '', 'Искать Grep-ом по имени или синониму.', '')
$counts = @()
foreach ($folder in $types.Keys) {
    $dir = Join-Path $Src $folder
    if (-not (Test-Path $dir)) { continue }
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($f in [System.IO.Directory]::GetFiles($dir, '*.xml')) {
        $p = Props (ReadText $f)
        $list.Add(@{ Name = (Tag $p 'Name'); Syn = (Synonym $p); Props = $p })
    }
    $objects[$folder] = $list
    $counts += "| $($types[$folder]) | $($list.Count) |"
    $index += "## $($types[$folder]) ($($list.Count))"
    foreach ($o in ($list | Sort-Object { $_.Name })) { $index += "- $($o.Name) — $(Esc $o.Syn)" }
    $index += ''
}
WriteMd 'objects.md' $index

# ---------- Документы -> регистры и обратный индекс ----------
$docs = @('# Документы: проведение и движения', '', '| Документ | Синоним | Проведение | Регистры (движения) |', '|---|---|---|---|')
$regToDocs = @{}
foreach ($o in ($objects['Documents'] | Sort-Object { $_.Name })) {
    $regs = @(Items $o.Props 'RegisterRecords' | ForEach-Object { RuRef $_ })
    foreach ($r in $regs) {
        if (-not $regToDocs.ContainsKey($r)) { $regToDocs[$r] = New-Object System.Collections.Generic.List[string] }
        $regToDocs[$r].Add($o.Name)
    }
    $docs += "| $($o.Name) | $(Esc $o.Syn) | $(Tag $o.Props 'Posting') | $($regs -join ', ') |"
}
WriteMd 'documents.md' $docs

$regs = @('# Регистры: какие документы пишут движения', '', '| Регистр | Синоним | Документы-регистраторы |', '|---|---|---|')
foreach ($folder in 'AccumulationRegisters', 'InformationRegisters', 'AccountingRegisters', 'CalculationRegisters') {
    foreach ($o in ($objects[$folder] | Sort-Object { $_.Name })) {
        $key = "$($types[$folder]).$($o.Name)"
        $mode = if ($folder -eq 'InformationRegisters') { " ($(Tag $o.Props 'WriteMode'))" } else { '' }
        $d = if ($regToDocs.ContainsKey($key)) { ($regToDocs[$key] | Sort-Object) -join ', ' } else { '—' }
        if ($folder -eq 'InformationRegisters' -and $d -eq '—') { continue }  # независимые РС — только в objects.md
        $regs += "| $key$mode | $(Esc $o.Syn) | $d |"
    }
}
WriteMd 'registers.md' $regs

# ---------- Общие модули ----------
$cm = @('# Общие модули', '', 'Флаги: С — сервер, ВС — вызов сервера, К — клиент (упр.), Г — глобальный, П — привилегированный, ВН — внешнее соединение, ПВИ — повт. исп. значений', '',
        '| Модуль | Флаги | Синоним |', '|---|---|---|')
foreach ($o in ($objects['CommonModules'] | Sort-Object { $_.Name })) {
    $fl = @()
    if ((Tag $o.Props 'Server') -eq 'true') { $fl += 'С' }
    if ((Tag $o.Props 'ServerCall') -eq 'true') { $fl += 'ВС' }
    if ((Tag $o.Props 'ClientManagedApplication') -eq 'true') { $fl += 'К' }
    if ((Tag $o.Props 'Global') -eq 'true') { $fl += 'Г' }
    if ((Tag $o.Props 'Privileged') -eq 'true') { $fl += 'П' }
    if ((Tag $o.Props 'ExternalConnection') -eq 'true') { $fl += 'ВН' }
    $reuse = Tag $o.Props 'ReturnValuesReuse'
    if ($reuse -and $reuse -ne 'DontUse') { $fl += "ПВИ:$reuse" }
    $cm += "| $($o.Name) | $($fl -join ' ') | $(Esc $o.Syn) |"
}
WriteMd 'common-modules.md' $cm

# ---------- Подписки на события ----------
$es = @('# Подписки на события', '', '| Подписка | Событие | Обработчик | Источники |', '|---|---|---|---|')
foreach ($o in ($objects['EventSubscriptions'] | Sort-Object { $_.Name })) {
    $srcTypes = @([regex]::Matches($o.Props, '<v8:Type>cfg:([^<]*)</v8:Type>') | ForEach-Object { $_.Groups[1].Value })
    $srcText = if ($srcTypes.Count -gt 8) { "$($srcTypes.Count) типов: $(($srcTypes | Select-Object -First 5) -join ', ') …" } else { $srcTypes -join ', ' }
    $dt = [regex]::Match($o.Props, '<Source>\s*<v8:TypeSet>cfg:([^<]*)</v8:TypeSet>')
    if ($dt.Success) { $srcText = "набор $($dt.Groups[1].Value)" + $(if ($srcText) { "; $srcText" } else { '' }) }
    $h = (Tag $o.Props 'Handler') -replace '^CommonModule\.', ''
    $es += "| $($o.Name) | $(Tag $o.Props 'Event') | $h | $srcText |"
}
WriteMd 'event-subscriptions.md' $es

# ---------- Регламентные задания ----------
$sj = @('# Регламентные задания', '', '| Задание | Синоним | Метод | Использование | Предопр. |', '|---|---|---|---|---|')
foreach ($o in ($objects['ScheduledJobs'] | Sort-Object { $_.Name })) {
    $m = (Tag $o.Props 'MethodName') -replace '^CommonModule\.', ''
    $sj += "| $($o.Name) | $(Esc $o.Syn) | $m | $(Tag $o.Props 'Use') | $(Tag $o.Props 'Predefined') |"
}
WriteMd 'scheduled-jobs.md' $sj

# ---------- Функциональные опции ----------
$fo = @('# Функциональные опции', '', '| Опция | Синоним | Где хранится | Объектов в составе |', '|---|---|---|---|')
foreach ($o in ($objects['FunctionalOptions'] | Sort-Object { $_.Name })) {
    $loc = RuRef ((Tag $o.Props 'Location') -replace '^(\w+)\.', '$1.')
    $fo += "| $($o.Name) | $(Esc $o.Syn) | $loc | $(@(Items $o.Props 'Content').Count) |"
}
WriteMd 'functional-options.md' $fo

# ---------- Подсистемы (рекурсивно) ----------
$ss = @('# Подсистемы и их состав', '')
function WalkSubsystems($dir, $prefix, $level) {
    if (-not (Test-Path $dir)) { return }
    foreach ($f in ([System.IO.Directory]::GetFiles($dir, '*.xml') | Sort-Object)) {
        $p = Props (ReadText $f)
        $name = Tag $p 'Name'
        $path = if ($prefix) { "$prefix/$name" } else { $name }
        $content = @(Items $p 'Content' | ForEach-Object { RuRef $_ })
        $incl = if ((Tag $p 'IncludeInCommandInterface') -eq 'true') { '' } else { ' _(скрыта в интерфейсе)_' }
        $script:ss += ('#' * [Math]::Min($level + 1, 6)) + " $path — $(Esc (Synonym $p))$incl"
        if ($content.Count) {
            $groups = $content | Group-Object { $_.Split('.')[0] } | Sort-Object Name
            foreach ($g in $groups) { $script:ss += "- **$($g.Name)** ($($g.Count)): $((($g.Group | ForEach-Object { $_.Split('.', 2)[1] }) | Sort-Object) -join ', ')" }
        }
        $script:ss += ''
        WalkSubsystems (Join-Path (Join-Path $dir $name) 'Subsystems') $path ($level + 1)
    }
}
WalkSubsystems (Join-Path $Src 'Subsystems') '' 1
WriteMd 'subsystems.md' $ss

# ---------- Обзор ----------
$cfg = Props (ReadText (Join-Path $Src 'Configuration.xml'))
$bspFile = Join-Path $Src 'CommonModules\ОбновлениеИнформационнойБазыБСП\Ext\Module.bsl'
$bsp = if (Test-Path $bspFile) { ([regex]::Match((ReadText $bspFile), 'Описание\.Версия\s*=\s*"([^"]+)"')).Groups[1].Value } else { '?' }
$libs = @()
foreach ($f in [System.IO.Directory]::GetFiles((Join-Path $Src 'CommonModules'), 'Module.bsl', 'AllDirectories')) {
    if ($f -notmatch '\\ОбновлениеИнформационнойБазы[^\\]*\\Ext\\Module\.bsl$') { continue }
    $t = ReadText $f
    $n = [regex]::Match($t, 'Описание\.Имя\s*=\s*"([^"]+)"'); $v = [regex]::Match($t, 'Описание\.Версия\s*=\s*"([^"]+)"')
    if ($n.Success -and $v.Success) { $libs += "| $($n.Groups[1].Value) | $($v.Groups[1].Value) | $(Split-Path (Split-Path (Split-Path $f)) -Leaf) |" }
}
$locWords = 'Таджик|Казах|Кыргыз|Киргиз|Узбек|Беларус|Белорус|Армени|Украин|Молдав|Азербайдж|Грузия|Грузинск'
$loc = @()
foreach ($folder in $objects.Keys) {
    foreach ($o in $objects[$folder]) {
        if ("$($o.Name) $($o.Syn)" -match $locWords) { $loc += "- $($types[$folder]).$($o.Name) — $(Esc $o.Syn)" }
    }
}
$ov = @(
    '# Обзор конфигурации', '',
    "- Конфигурация: $(Tag $cfg 'Name') — «$(Synonym $cfg)», версия **$(Tag $cfg 'Version')**, поставщик $(Esc (Tag $cfg 'Vendor'))",
    "- Совместимость: $(Tag $cfg 'CompatibilityMode'); БСП: **$bsp**",
    "- Справочник собран: $(Get-Date -Format 'yyyy-MM-dd HH:mm') скриптом C:\1C\tools\build-map.ps1", '',
    '## Файлы справочника',
    '- `objects.md` — все объекты с синонимами (искать Grep-ом)',
    '- `subsystems.md` — дерево подсистем и их состав',
    '- `documents.md` — документы и их движения по регистрам',
    '- `registers.md` — регистры и документы-регистраторы',
    '- `common-modules.md` — общие модули и их флаги',
    '- `event-subscriptions.md`, `scheduled-jobs.md`, `functional-options.md`', '',
    '## Встроенные библиотеки', '', '| Библиотека | Версия | Модуль |', '|---|---|---|') + ($libs | Sort-Object) + @(
    '', '## Количество объектов', '', '| Тип | Кол-во |', '|---|---|') + $counts + @(
    '', "## Объекты с признаками локализации для других стран ($($loc.Count))", '') + $(if ($loc.Count) { $loc | Sort-Object } else { @('- не найдено') })
WriteMd 'README.md' $ov
Write-Host "Готово: $Out"
Get-ChildItem $Out | ForEach-Object { Write-Host ("{0,-26} {1,8:N0} КБ" -f $_.Name, ($_.Length / 1KB)) }
