Add-Type -AssemblyName System.Windows.Forms

$dialog = New-Object System.Windows.Forms.OpenFileDialog
$dialog.InitialDirectory = [System.Environment]::GetFolderPath("Desktop")
$dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
$dialog.Title = "Select excel file to append."

if($dialog.ShowDialog() -eq "OK"){
    $newFile = $dialog.FileName
}else{
    exit
}

$template = Import-Excel "Template.xlsx"

$merged = $template + $newFile

$desktop = [System.Environment]::GetFolderPath("Deskop")

$retunPath = Join-Path -Path $desktop -ChildPath "main.xlsx"

$merged | Export-Excel $retunPath