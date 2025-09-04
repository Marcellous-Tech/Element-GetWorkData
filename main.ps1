Add-Type -AssemblyName System.Windows.Forms
# manually time sheet
$dialog = New-Object System.Windows.Forms.FolderBrowserDialog
$dialog.InitialDirectory = [System.Environment]::GetFolderPath("Desktop")
$dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
$dialog.Title = "Select excel file to append."

if($dialog.ShowDialog() -eq "OK"){
    $newFile = $dialog.FileName
}else{
    exit
}

$main = Import-Excel "C:\Users\Marcellous\Desktop\main.xlsx"
$newdata = Import-Excel $newFile
$merged = $main + $newdata

# $desktop = [System.Environment]::GetFolderPath("Desktop")

$merged | Export-Excel $main