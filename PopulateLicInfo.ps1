#Объявляем массивы данных
    $WSSerial = @()    
    $WSName = @()
    $Company = @()
    $Comment = @()
    $OfficeActivationKey = @()
    $OfficeLicenseModel = @()
    $OfficeOLPSN = @()
    $OfficeTitle = @()
    $WindowsActivationKey = @()
    $WindowsLicenseModel = @()
    $WindowsOLPSN = @()
    $WindowsTitle = @()

#Подключаемся к источнику данных
Import-Csv "\\zdc5\work\Administrator\licensing\ЛИЦЕНЗИИ МS СВОДНЫЙ.csv" |`
    ForEach-Object {
        $WSSerial += $_."Серийный номер"    
        $WSName += $_."Сетевое имя"
        $Company += $_."Юрлицо"
        $Comment += $_."Примечание"
        $OfficeActivationKey += $_."Ключ активации Office"
        $OfficeLicenseModel += $_."Office License Model"
        $OfficeOLPSN += $_."Номер лицензии Office OLP"
        $OfficeTitle += $_."Версия Офис"
        $WindowsActivationKey += $_."Ключ активации OS"
        $WindowsLicenseModel += $_."Windows License Model"
        $WindowsOLPSN += $_."Номер лицензии Windows OLP"
        $WindowsTitle += $_."Версия Windows"
    }

$TestRequest = read-host -prompt "Тестовый запрос имени ПК"

if ($WSName -contains $TestRequest)
    {
    Write-Host "ПК найден!"
    $Specs = [array]::IndexOf($WSName, $TestRequest)
    Write-Host "Сетевое имя: " $WSName[$Specs]
    Write-Host "Юрлицо: " $Company[$Specs]
    }
else {Write-Host "Что то пошло не так!"}