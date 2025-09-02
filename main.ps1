Add-Type -AssemblyName System.Windows.Forms

$dialog = New-Object System.Windows.Forms.OpenDileDialog
$dialog.InitalDirectory = [System.Environment]::GetFolderPath("MyDesktop")
$dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
$dialog.Title = "Select excel file to append."


$data = Import-Excel "/Users/marcellouscurtis/Documents/Projects/Freelance/Element/Job Cost 7.07.25-7.17.25 Pay Period (7.25.25).xlsx" -WorksheetName "7.07.25-7.17.25 Pay Period"

$data2 = Import-Excel '/Users/marcellouscurtis/Documents/Projects/Freelance/Element/ECD 2025-06-27 Job Cost JE - Complete.xlsx' -WorksheetName '6.09.25-6.20.25 Pay Period'

Write-Host $data

$data3 = $data + $data2

$data3 | Export-Excel "Here.xlsx"