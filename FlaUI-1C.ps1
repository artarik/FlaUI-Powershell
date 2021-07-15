Import-Module ImportExcel
Import-Module CredentialManager

#------ Init block
Add-Type -AssemblyName System.Windows.Forms
$null = [System.Reflection.Assembly]::Load("System.Xml.ReaderWriter, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
$null = [System.Reflection.Assembly]::Load("System.Drawing.Primitives, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")

$assemblyList = Get-ChildItem -Filter *.dll

foreach ($item in $assemblyList) {
    try {
        Add-Type -Path $item.Name
    }
    catch {
        $_.Exception.LoaderExceptions
        exit 1
    }
}

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

$outFile = [System.IO.Path]::Combine($PSScriptRoot, "out.xlsx")
$resFile = [System.IO.Path]::Combine($PSScriptRoot, "exportTable.xlsx")
$reportFile = ([System.IO.Path]::Combine($PSScriptRoot, "report.xlsx"))

$creds = Get-StoredCredential -AsCredentialObject -Target msk-1c-app
if (![bool]$creds){
    Write-Error "Unable to get Creds from Windows Credential Manager"
    exit 1
}


$xpathWorkFlow = @()
$xpathWorkFlow+= [pscustomobject]@{Action = "Click"; Timeout = 30; Value = "";xpath ="/Window[@ClassName = 'V8TopLevelFrameTaxiStarter' and @Name = 'Запуск 1С:Предприятия']/Pane[5]/Pane/Pane/Button[@Name = '1С:Предприятие']" }
$xpathWorkFlow+= [pscustomobject]@{Action = "Type"; Timeout = 30; Value = $creds.UserName; xpath = "/Window[@ClassName = 'V8NewLocalFrameBaseWnd' and @Name = 'Доступ к информационной базе']/Window[@ClassName = 'V8TopLevelFrameTaxiStarter' and @Name = '1С:Предприятие']/Pane[5]/Pane/Pane/ComboBox[@Name = 'Пользователь']"}
$xpathWorkFlow+= [pscustomobject]@{Action = "Type"; Timeout = 30; Value = $creds.Password; xpath ="/Window[@ClassName = 'V8NewLocalFrameBaseWnd' and @Name = 'Доступ к информационной базе']/Window[@ClassName = 'V8TopLevelFrameTaxiStarter' and @Name = '1С:Предприятие']/Pane[5]/Pane/Pane/Edit[@Name = 'Пароль']"}
$xpathWorkFlow+= [pscustomobject]@{Action = "Click"; Timeout = 30; Value = ""; xpath ="/Window[@ClassName = 'V8NewLocalFrameBaseWnd' and @Name = 'Доступ к информационной базе']/Window[@ClassName = 'V8TopLevelFrameTaxiStarter' and @Name = '1С:Предприятие']/Pane[5]/Pane/Pane/Button[@Name = 'Войти']"}
#$xpathWorkFlow+= [pscustomobject]@{Action = "Click"; Timeout = 5; Value = ""; xpath ="/Window[@ClassName = 'V8NewLocalFrameBaseWnd' and @Name = 'Доступ к информационной базе']/Window[@ClassName = 'V8NewLocalFrameBaseWnd' and @Name = '1С:Предприятие']/Pane/Pane/Pane/Pane[@Name = 'Идентификация пользователя не выполнена']"}
$xpathWorkFlow+= [pscustomobject]@{Action = "Click"; Timeout = 120; Value = ""; xpath ="/Window[@ClassName = 'V8TopLevelFrameSDI']/Pane/Pane[13]/Pane/Pane[2]/Pane/Pane[@Name = 'Панель разделов']/TabItem[@Name = 'Денежные средства']"}
$xpathWorkFlow+= [pscustomobject]@{Action = "Click"; Timeout = 30; Value = ""; xpath ="/Window[@ClassName = 'V8TopLevelFrameSDI']/Pane/Pane[13]/Pane/Pane[4]/Pane/Menu[@Name = 'Меню функций']/Group[1]/Group[@Name = 'Банк (казначейство)']/MenuItem[@Name = 'Расчетно - платежные документы']"}
$xpathWorkFlow+= [pscustomobject]@{Action = "Click"; Timeout = 30; Value = ""; xpath ="/Window[@ClassName = 'V8TopLevelFrameSDI']/Pane/Pane[13]/Pane/Pane[1]/Pane/Pane[4]/Pane/Pane[1]/Pane/Pane/Pane/Pane[1]/Pane/Tab[@Name = 'Расчетно - платежные документы']/Pane/Pane/ToolBar/Button[@Name = 'Еще']"}
$xpathWorkFlow+= [pscustomobject]@{Action = "Click"; Timeout = 30; Value = ""; xpath ="/Window[@ClassName = 'V8TopLevelFrameSDI']/Window[@ClassName = 'V8CommandPanelPopup']/Menu/MenuItem[@Name = 'Установить период...' and @HelpText = 'Установить период для просмотра']"}
$xpathWorkFlow+= [pscustomobject]@{Action = "Click"; Timeout = 30; Value = ""; xpath ="/Window[@ClassName = 'V8TopLevelFrameSDI']/Window[@ClassName = 'V8TopLevelFrameSDIsec' and @Name = 'Выберите период']/Pane/Pane[13]/Pane/Pane[1]/Pane/Pane/Pane/Hyperlink[@Name = 'Показать стандартные периоды']"}
$xpathWorkFlow+= [pscustomobject]@{Action = "DClick"; Timeout = 30; Value = ""; xpath ="/Window[@ClassName = 'V8TopLevelFrameSDI']/Window[@ClassName = 'V8TopLevelFrameSDIsec' and @Name = 'Выберите период']/Pane/Pane[13]/Pane/Pane[1]/Pane/Pane/Pane/Pane[1]/Pane/List/ListItem[@Name = 'Сегодня']"}
$xpathWorkFlow+= [pscustomobject]@{Action = "Click"; Timeout = 30; Value = ""; xpath ="/Window[@ClassName = 'V8TopLevelFrameSDI']/Pane/Pane[13]/Pane/Pane[1]/Pane/Pane[4]/Pane/Pane[1]/Pane/Pane/Pane/Pane[1]/Pane/Tab[@Name = 'Расчетно - платежные документы']/Pane/Pane/ToolBar/Button[@Name = 'Еще']"}
$xpathWorkFlow+= [pscustomobject]@{Action = "Click"; Timeout = 30; Value = ""; xpath ="/Window[@ClassName = 'V8TopLevelFrameSDI']/Window[@ClassName = 'V8CommandPanelPopup']/Menu/MenuItem[@Name = 'Вывести список...' and @HelpText = 'Вывести список']"}
$xpathWorkFlow+= [pscustomobject]@{Action = "Click"; Timeout = 30; Value = ""; xpath ="/Window[@ClassName = 'V8TopLevelFrameSDI']/Window[@ClassName = 'V8TopLevelFrameSDIsec' and contains(@Name, 'Вывести список')]/Pane/Pane[13]/Pane/Pane[1]/Pane/Pane/Pane/Pane[2]/ToolBar/Button[contains(@Name, 'ОК')]"}
$xpathWorkFlow+= [pscustomobject]@{Action = "Click"; Timeout = 30; Value = ""; xpath ="/Window[@ClassName = 'V8TopLevelFrameSDI']/Pane/Pane[13]/Pane/Pane[1]/Pane/Button[@Name = 'Сохранить']"}
$xpathWorkFlow+= [pscustomobject]@{Action = "Type"; Timeout = 30; Value = $outFile; xpath ="/Window[@ClassName = 'V8TopLevelFrameSDI']/Window[@Name = 'Сохранение']/Pane[@ClassName = 'DUIViewWndClassName']/ComboBox[@AutomationId = 'FileNameControlHost' and @ClassName = 'AppControlHost' and @Name = 'Имя файла:']/Edit[@Name = 'Имя файла:']"}
$xpathWorkFlow+= [pscustomobject]@{Action = "Click"; Timeout = 30; Value = ""; xpath ="/Window[@ClassName = 'V8TopLevelFrameSDI']/Window[@ClassName = '#32770' and @Name = 'Сохранение']/Pane[@ClassName = 'DUIViewWndClassName']/ComboBox[@AutomationId = 'FileTypeControlHost' and @ClassName = 'AppControlHost' and @Name = 'Тип файла:']"}
$xpathWorkFlow+= [pscustomobject]@{Action = "Key"; Timeout = 30; Value = "{DOWN}"; xpath =7}
$xpathWorkFlow+= [pscustomobject]@{Action = "Key"; Timeout = 30; Value = "{ENTER}"; xpath =1}
$xpathWorkFlow+= [pscustomobject]@{Action = "Click"; Timeout = 30; Value = ""; xpath ="/Window[@ClassName = 'V8TopLevelFrameSDI']/Window[@ClassName = '#32770' and @Name = 'Сохранение']/Button[@AutomationId = '1' and @ClassName = 'Button' and @Name = 'Сохранить']"}
$xpathWorkFlow+= [pscustomobject]@{Action = "Click"; Timeout = 30; Value = ""; xpath ="/Window[@ClassName = 'V8TopLevelFrameSDI']/Pane/Pane[13]/Pane/Pane[1]/Pane/Button[@Name = 'Закрыть']"}
$xpathWorkFlow+= [pscustomobject]@{Action = "Key"; Timeout = 30; Value = "%{F4}"; xpath =1}
$xpathWorkFlow+= [pscustomobject]@{Action = "Click"; Timeout = 30; Value = ""; xpath ="/Window[@ClassName = 'V8TopLevelFrameSDI']/Window[@ClassName = 'V8TopLevelFrameSDIsec']/Pane/Pane[13]/Pane/Pane[1]/Pane/Pane/Pane/Button[@Name = 'Завершить работу']"}

Remove-Item $outFile -ErrorAction 0
Remove-Item $resFile -ErrorAction 0
Remove-Item $reportFile -ErrorAction 0


$uia = [FlaUI.UIA3.UIA3Automation]::new()
$uia.TransactionTimeout = [System.TimeSpan]::FromSeconds($uia.TransactionTimeout.TotalSeconds * 2)
$uia.ConnectionTimeout = [System.TimeSpan]::FromSeconds($uia.ConnectionTimeout.TotalSeconds * 2)    
$desktop = $uia.GetDesktop()

#----- functions block
function WaitElement {
    param (
         [string]$xpath,
         [int] $Timeout = 30
    )
    
    try {$el = $desktop.FindFirstByXPath($xpath) }catch{}
    $count = 1 
    while (!$el.IsAvailable){
        if ($count -eq $Timeout){
            Write-Error "Timeout ($($Timeout)s) to found element witn XPATH : $($xpath)"
            exit 1
        }
        try {$el = $desktop.FindFirstByXPath($xpath) }catch{}
        Start-Sleep -Seconds 1
        $count ++
    }
    return $el
}


function RunAction {
    param ( $item)

    switch ($item.Action) {
        Click {
            $el = WaitElement -xpath $item.Xpath -Timeout $item.Timeout
            $el.Click()
        }

        DClick {
            $el = WaitElement -xpath $item.Xpath -Timeout $item.Timeout
            $el.DoubleClick()
        }
        Type {
            $el = WaitElement -xpath $item.Xpath -Timeout $item.Timeout
            $el.Patterns.Value.Pattern.SetValue($item.Value)
        }
        Key {
        
            for ($i=1; $i -le $item.xpath; $i++){
                [System.Windows.Forms.SendKeys]::SendWait($item.Value)
            }
        
        }
    }
}


#----- main block

$app = [FlaUI.Core.Application]::Launch("C:\Program Files\1cv8\common\1cestart.exe") # start 1c


foreach ($itemRow in $xpathWorkFlow) {
    RunAction -item $itemRow
}

$app.Dispose()

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
    catch {
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
    
5,6 | ForEach-Object{
    Set-ExcelColumn -ExcelPackage $e -WorksheetName $ws -Column $PSItem -WrapText -Width 28 -VerticalAlignment Center
}
Start-Sleep -Milliseconds 500
for ($col = 1; $col -le 7; $col++) {
    Set-ExcelColumn -ExcelPackage $e -WorksheetName $ws -Column $col -AutoSize -HorizontalAlignment Center -VerticalAlignment Center
}

Set-ExcelRange -Address $ws.Cells["A3:G$($ws.Dimension.Rows)"] -BorderTop Thin -BorderBottom Thin -BorderLeft thin -BorderRight thin

Close-ExcelPackage $e -SaveAs $reportFile
Remove-Item $resFile
Start-Process $reportFile