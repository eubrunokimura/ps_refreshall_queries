$path = "Seu path aqui"
$files = Get-ChildItem -Path "$($path)\*.xl*"

Write-Host "Iniciando atualizacao de bases consulta"
$app = New-Object -ComObject "Excel.Application"
$app.Visible = $false
$app.DisplayAlerts = $false

$count = 1

foreach ($file in $files) {
    Write-Host "Iniciando o arquivo ($($count)/$($files.count)): $($file.name)"

    Write-Host "Abrindo o arquivo..."
    $wb = $app.workbooks.Open($file)
    $wb.RefreshAll()

    Write-Host "Atualizando..."
    $app.CalculateUntilAsyncQueriesDone()

    Write-Host "Salvando..."
    $wb.Save()

    $app.DisplayAlerts = $true

    ""
    $wb.Close()
    Write-Host "Atualizacao Realizada com sucesso! - $($file.name) "

    $count = $count + 1
}

Write-Host "Finalizando Script..."
$app.Quit()