Import-Module ImportExcel
Import-Module CredentialManager
#------ Init block
Add-Type -AssemblyName System.Windows.Forms
$null = [System.Reflection.Assembly]::Load("System.Xml.ReaderWriter, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
$null = [System.Reflection.Assembly]::Load("System.Drawing.Primitives, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")

try {Add-Type -path FlaUI.Core.dll}catch{$_.Exception.LoaderExceptions;exit 1}

try {
    #$bytes = [System.IO.File]::ReadAllBytes("$($dir)\FlaUI.UIA3.dll")
    #[System.Reflection.Assembly]::Load($bytes)
    Add-Type -path FlaUI.UIA3.dll
}catch
{
    $_.Exception.LoaderExceptions
    exit 1
}

try {Add-Type -Path System.Drawing.Common.dll}catch {$_.Exception.LoaderExceptions;exit 1}
try {Add-Type -Path System.Security.Permissions.dll}catch {$_.Exception.LoaderExceptions;exit 1}
try {Add-Type -Path System.Xml.ReaderWriter.dll}catch {$_.Exception.LoaderExceptions;exit 1}

#----- Vars block
$outFile = [System.IO.Path]::Combine($PSScriptRoot, "out.xlsx")
$resFile = [System.IO.Path]::Combine($PSScriptRoot, "exportTable.xlsx")
$reportFile = ([System.IO.Path]::Combine($PSScriptRoot, "report.xlsx"))

$creds = Get-StoredCredential -AsCredentialObject -Target msk-1c-app
if (![bool]$creds){
    Write-Error "Unable to get credentials from Windows Credential Manager"
    exit 1
}

if ([System.IO.File]::Exists($resFile))
{
    Remove-Item -Path $resFile -Force
}

if ([System.IO.File]::Exists($outFile))
{
    Remove-Item -Path $outFile -Force
}

if ([System.IO.File]::Exists($reportFile))
{
    Remove-Item -Path $reportFile -Force
}

$uia = [FlaUI.UIA3.UIA3Automation]::new()
$uia.TransactionTimeout = [System.TimeSpan]::FromSeconds($uia.TransactionTimeout.TotalSeconds * 2)
$uia.ConnectionTimeout = [System.TimeSpan]::FromSeconds($uia.ConnectionTimeout.TotalSeconds * 2)    
$desktop = $uia.GetDesktop()

#----- functions block
function WaitElement {
    param (
         [string]$xpath,
         [int] $Seconds = 30
    )
    
    try {$el = $desktop.FindFirstByXPath($xpath) }catch{}
    $count = 1 
    while (!$el.IsAvailable){
        if ($count -eq $Seconds){
            Write-Error "Timeout ($($Seconds)s) to found element witn XPATH : $($xpath)"
            exit 1
        }
        $el = $desktop.FindFirstByXPath($xpath) 
        Start-Sleep -Seconds 1
        $count ++
    }
    return $el
}
#----- main block

$app = [FlaUI.Core.Application]::Launch("C:\Program Files\1cv8\common\1cestart.exe") # start 1c

$xpath = "/Window[@ClassName = 'V8TopLevelFrameTaxiStarter' and @Name = 'Запуск 1С:Предприятия']/Pane[5]/Pane/Pane/Button[@Name = '1С:Предприятие']"
$xpathLogin = "/Window[@ClassName = 'V8NewLocalFrameBaseWnd' and @Name = 'Доступ к информационной базе']/Window[@ClassName = 'V8TopLevelFrameTaxiStarter' and @Name = '1С:Предприятие']/Pane[5]/Pane/Pane/ComboBox[@Name = 'Пользователь']"
$xpathPw = "/Window[@ClassName = 'V8NewLocalFrameBaseWnd' and @Name = 'Доступ к информационной базе']/Window[@ClassName = 'V8TopLevelFrameTaxiStarter' and @Name = '1С:Предприятие']/Pane[5]/Pane/Pane/Edit[@Name = 'Пароль']"
$xpathWrongPW = "/Window[@ClassName = 'V8NewLocalFrameBaseWnd' and @Name = 'Доступ к информационной базе']/Window[@ClassName = 'V8NewLocalFrameBaseWnd' and @Name = '1С:Предприятие']/Pane/Pane/Pane/Pane[@Name = 'Идентификация пользователя не выполнена']"
$xpathBtnEnter = "/Window[@ClassName = 'V8NewLocalFrameBaseWnd' and @Name = 'Доступ к информационной базе']/Window[@ClassName = 'V8TopLevelFrameTaxiStarter' and @Name = '1С:Предприятие']/Pane[5]/Pane/Pane/Button[@Name = 'Войти']"


$el = WaitElement -xpath $xpath # click 1C 
$el.Click()

$login = WaitElement -xpath $xpathLogin 
$login.Patterns.Value.Pattern.SetValue($creds.UserName) # type login

$Pw = $desktop.FindFirstByXPath($xpathPw)
$Pw.Patterns.Value.Pattern.SetValue($creds.Password) # type password

$desktop.FindFirstByXPath($xpathBtnEnter).Click() # click Enter

Start-Sleep -Seconds 2

try {
    $wrongCreds = $desktop.FindFirstByXPath($xpathWrongPW)
    if ($wrongCreds.IsAvailable){
        Write-Error "Wrong Login or Password "
        exit 1
    } 
}

catch {}

$xpathDS = "/Window[@ClassName = 'V8TopLevelFrameSDI']/Pane/Pane[13]/Pane/Pane[2]/Pane/Pane[@Name = 'Панель разделов']/TabItem[@Name = 'Денежные средства']"
$ds = WaitElement -xpath $xpathDS -Seconds 120
$ds.Click() # Найти и открыть  меню Денежные средства


$xpathJournal = "/Window[@ClassName = 'V8TopLevelFrameSDI']/Pane/Pane[13]/Pane/Pane[4]/Pane/Menu[@Name = 'Меню функций']/Group[1]/Group[@Name = 'Банк (казначейство)']/MenuItem[@Name = 'Расчетно - платежные документы']"
$desktop.FindFirstByXPath($xpathJournal).Click() # Открыть журнал расчетно-платежных документов


$xpath = "/Window[@ClassName = 'V8TopLevelFrameSDI']/Pane/Pane[13]/Pane/Pane[1]/Pane/Pane[4]/Pane/Pane[1]/Pane/Pane/Pane/Pane[1]/Pane/Tab[@Name = 'Расчетно - платежные документы']/Pane/Pane/ToolBar/Button[@Name = 'Еще']"
$desktop.FindFirstByXPath($xpath).Click() # Клик по кнопке Еще

$xpath = "/Window[@ClassName = 'V8TopLevelFrameSDI']/Window[@ClassName = 'V8CommandPanelPopup']/Menu/MenuItem[@Name = 'Установить период...' and @HelpText = 'Установить период для просмотра']"
$desktop.FindFirstByXPath($xpath).Click()# Клик по кнопке Установить период

$xpath = "/Window[@ClassName = 'V8TopLevelFrameSDI']/Window[@ClassName = 'V8TopLevelFrameSDIsec' and @Name = 'Выберите период']/Pane/Pane[13]/Pane/Pane[1]/Pane/Pane/Pane/Hyperlink[@Name = 'Показать стандартные периоды']"
$desktop.FindFirstByXPath($xpath).Click() # Клик по ссылке Показать стандартные периоды


$xpath = "/Window[@ClassName = 'V8TopLevelFrameSDI']/Window[@ClassName = 'V8TopLevelFrameSDIsec' and @Name = 'Выберите период']/Pane/Pane[13]/Pane/Pane[1]/Pane/Pane/Pane/Pane[1]/Pane/List/ListItem[@Name = 'Сегодня']"
$desktop.FindFirstByXPath($xpath).DoubleClick() # Двойной клик по полю Сегодня

$xpath = "/Window[@ClassName = 'V8TopLevelFrameSDI']/Pane/Pane[13]/Pane/Pane[1]/Pane/Pane[4]/Pane/Pane[1]/Pane/Pane/Pane/Pane[1]/Pane/Tab[@Name = 'Расчетно - платежные документы']/Pane/Pane/ToolBar/Button[@Name = 'Еще']"
$desktop.FindFirstByXPath($xpath).Click() # Клик по кнопке Еще

$xpath = "/Window[@ClassName = 'V8TopLevelFrameSDI' and contains(@Name, 'Бухгалтерия государственного учреждения, редакция 2.0')]/Window[@ClassName = 'V8CommandPanelPopup']/Menu/MenuItem[@Name = 'Вывести список...' and @HelpText = 'Вывести список']"
$desktop.FindFirstByXPath($xpath).Click() # Клик по кнопке Вывести список

$xpath = "/Window[@ClassName = 'V8TopLevelFrameSDI']/Window[@ClassName = 'V8TopLevelFrameSDIsec' and contains(@Name, 'Вывести список')]/Pane/Pane[13]/Pane/Pane[1]/Pane/Pane/Pane/Pane[2]/ToolBar/Button[contains(@Name, 'ОК') and contains(@HelpText, 'Ok')]"
$desktop.FindFirstByXPath($xpath).Click() # Клик по кнопке ОК

$xpath = "/Window[@ClassName = 'V8TopLevelFrameSDI' and contains(@Name, 'Бухгалтерия государственного учреждения, редакция 2.0')]/Pane/Pane[13]/Pane/Pane[1]/Pane/Button[@Name = 'Сохранить']"
$desktop.FindFirstByXPath($xpath).Click() # Клик по кнопке Сохранить

$xpath = "/Window[@ClassName = 'V8TopLevelFrameSDI' and contains(@Name, 'Бухгалтерия государственного учреждения, редакция 2.0')]/Window[@Name = 'Сохранение']/Pane[@ClassName = 'DUIViewWndClassName']/ComboBox[@AutomationId = 'FileNameControlHost' and @ClassName = 'AppControlHost' and @Name = 'Имя файла:']/Edit[@Name = 'Имя файла:']"
$fileName = $desktop.FindFirstByXPath($xpath)
$fileName.Patterns.Value.Pattern.SetValue($outFile) # Заполнить имя файла
 
$xpath = "/Window[@ClassName = 'V8TopLevelFrameSDI' and contains(@Name, 'Бухгалтерия государственного учреждения, редакция 2.0')]/Window[@ClassName = '#32770' and @Name = 'Сохранение']/Pane[@ClassName = 'DUIViewWndClassName']/ComboBox[@AutomationId = 'FileTypeControlHost' and @ClassName = 'AppControlHost' and @Name = 'Тип файла:']"
$desktop.FindFirstByXPath($xpath).Click() # Клик по меню Тип файла


1..7 | ForEach-Object{[System.Windows.Forms.SendKeys]::SendWait("{DOWN}")} # Выбрать тип файла Excel 2007
[System.Windows.Forms.SendKeys]::SendWait("{ENTER}")  # Нажать Enter

$xpath = "/Window[@ClassName = 'V8TopLevelFrameSDI' and contains(@Name, 'Бухгалтерия государственного учреждения, редакция 2.0')]/Window[@ClassName = '#32770' and @Name = 'Сохранение']/Button[@AutomationId = '1' and @ClassName = 'Button' and @Name = 'Сохранить']"
$desktop.FindFirstByXPath($xpath).Click() # Клик по кнопке Сохранить

$xpath = "/Window[@ClassName = 'V8TopLevelFrameSDI']/Pane/Pane[13]/Pane/Pane[1]/Pane/Button[@Name = 'Закрыть']"
$desktop.FindFirstByXPath($xpath).Click() # Клик по кнопке Закрыть

[System.Windows.Forms.SendKeys]::SendWait("%{F4}") # Вызов сочетания клавиш Alt+F4

$xpath = "/Window[@ClassName = 'V8TopLevelFrameSDI']/Window[@ClassName = 'V8TopLevelFrameSDIsec']/Pane/Pane[13]/Pane/Pane[1]/Pane/Pane/Pane/Button[@Name = 'Завершить работу']"
$desktop.FindFirstByXPath($xpath).Click() # Клик по кнопке Завершить работу

$app.Dispose()

$resTable = New-Object System.Data.DataTable
$resTable.Columns.Add( (New-Object System.Data.DataColumn "№ п/п", ([Int]) ))
$resTable.Columns.Add( (New-Object System.Data.DataColumn "Номер", ([String]) ))
$resTable.Columns.Add( (New-Object System.Data.DataColumn "Дата", ([String]) ))
$resTable.Columns.Add( (New-Object System.Data.DataColumn "Сумма", ([String]) ))
$resTable.Columns.Add( (New-Object System.Data.DataColumn "Контрагент", ([String]) ))
$resTable.Columns.Add( (New-Object System.Data.DataColumn "Договор", ([String]) ))
$resTable.Columns.Add( (New-Object System.Data.DataColumn "Комментарий", ([String]) ))

$resTable.Columns[0].AutoIncrement = $true
$resTable.Columns[0].AutoIncrementSeed = 1 
$resTable.Columns[0].AutoIncrementStep = 1

$e = Open-ExcelPackage -Path $outFile 
$ws = $e.Workbook.Worksheets[1]

for($rowIndex = 2; $rowIndex -le $ws.Dimension.Rows; $rowIndex ++)
{
     $newRow = $resTable.NewRow()
     $newRow["Номер"] = $ws.Cells["C$($rowIndex)"].Value
     $newRow["Дата"] = $ws.Cells["D$($rowIndex)"].Value.Substring(0,10)
     $newRow["Сумма"] = $ws.Cells["E$($rowIndex)"].Value
     $newRow["Контрагент"] = $ws.Cells["I$($rowIndex)"].Value
     $newRow["Договор"] = $ws.Cells["J$($rowIndex)"].Value
     $newRow["Комментарий"] = $ws.Cells["N$($rowIndex)"].Value
     $resTable.Rows.Add($newRow)
}


$resTable | Export-Excel -Path $resFile -WorksheetName "Sheet1" -ClearSheet -ExcludeProperty ItemArray, RowError, RowState, Table, HasErrors -StartRow 3 

Remove-Item $outFile

$e = Open-ExcelPackage -Path $resFile
$ws = $e.Workbook.Worksheets[1]

Set-ExcelRange -Address $ws.Cells["A1:G1"] -FontName 'Arial' -Merge -Bold `
    -Value "Реестр платежных поручений на $((Get-Date).ToShortDateString())" -VerticalAlignment Center -HorizontalAlignment Center

Set-ExcelRange -Address $ws.Cells["A3:G3"] -FontName 'Microsoft Sans Serif' -FontSize 8 -VerticalAlignment Center `
    -HorizontalAlignment Center -BackgroundColor ([System.Drawing.Color]::FromArgb(245, 242, 221))

for($i = 4; $i -le $ws.Dimension.Rows; $i++){
    $ws.Cells["A" + $i].Value = [System.Int32]::Parse($ws.Cells["A" + $i].Value)
    $ws.Cells["C" + $i].Value = ($ws.Cells["C" + $i].Value).Substring(0,10)
    try{
        $ws.Cells["D" + $i].Value = [System.Double]::Parse($ws.Cells["D" + $i].Value)
    }
    catch{
        $ws.Cells["D" + $i].Value = $ws.Cells["D" + $i].Value
    }
   
    if ($ws.Cells["G" + $i].Value -ne "#NULL!"){
        $ws.Cells["G" + $i].Hyperlink = $ws.Cells["G" + $i].Value
        $ws.Cells["G" + $i].StyleID = 1 # Стиль изменить на 'гиперссылка'
        $ws.Cells["G" + $i].Style.Font.UnderLine = $true
        $ws.Cells["G" + $i].Style.Font.color.SetColor("Blue")
        $ws.Cells["G" + $i].Value = "Ссылка"
        }
}
    
5,6 | foreach{
    Set-ExcelColumn -ExcelPackage $e -WorksheetName $ws -Column $PSItem -WrapText -Width 28 -VerticalAlignment Center
}
1,2,3,4,7 | foreach {
    Set-ExcelColumn -ExcelPackage $e -WorksheetName $ws -Column $PSItem -AutoSize -HorizontalAlignment Center -VerticalAlignment Center
}
    
Set-ExcelRange -Address $ws.Cells["A3:G$($ws.Dimension.Rows)"] -BorderTop Thin -BorderBottom Thin -BorderLeft thin -BorderRight thin

Close-ExcelPackage $e -SaveAs $reportFile
Remove-Item $resFile
