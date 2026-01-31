#region Variables
$CredFile = "$PSScriptRoot\pass.xml" # Путь к XML файлу с учетными данными
$sources = "\\network\share\example.txt" # Путь к файлу, который нужно переносить
$localcatalogy = "\\network\share\" # Локальный каталог для сохранения копий
$pathto = "/network/share/" #Каталог на удаленном хосте
$year = (Get-Date).Year # Текущий год
$todayDay = Get-Date -Format "dd.MM.yyyy" # Текущий день и год
$timefile = Get-Date -Format "yyyyMMddHHmss" # Формат ггггммдд для локалького файла
$retentionDate = (Get-Date).AddDays(-0) # -14 дней для архивации
$localYearArh = Join-Path -Path $localcatalogy -ChildPath ($year) # Подкаталог с указанием года
$LocalMonthArh = Join-Path -Path $localYearArh -ChildPath $todayDay # Подкаталог с указанием день год
$localfile = Join-Path $LocalMonthArh "course_$timefile.json" # Преобразование локального файла
#$localzip = Join-Path -Path $oldcatalog.FullName "banki.ru-cours_$oldcatalogdate.zip" # Zip архив

$winscpfiles = @( #Обязательные файлы без которых не возможна работа скрипта
    "WinSCP.exe",
    "WinSCPnet.dll"
)
$messageParam = @{
    SmtpServer = "mail.test.local" # SMTP сервер
    From       = "$ENV:COMPUTERNAME@gmail"
    To         = 'ex.ample@gmail' #Почта для отправки уведомлений
    Encoding   = "UTF8"
    Subject    = "Copying File $(get-date -format dd.MM.yyyy)"
    Body       = "Скрипт: $PSCommandPath"
}
$RemoteHostName = "1.1.1.1" #IP/NAME удаленного хоста
$PubkeySSHRemoteHost = "ssh-ed**********************" #Пара SSH, которая регистрируется при входе на хост.
$VerbosePreference = "Continue" 
#Endregion
#region Checks
if (!(Test-Path -Path $CredFile)) {
    $messageParam.Body += "`nНе найден XML файл с учетными данными '$CredFile'"
    Send-MailMessage @messageParam
    throw
}

try {
    $CredObject = Import-Clixml $CredFile -ErrorAction stop
}
catch {
    $messageParam.Body += "`nОшибка при импорте XML файла '$CredFile' с учетными данными."
    Send-MailMessage @messageParam
    throw "Ошибка: $($_.Exception.Message)"
}

foreach ($file in $winscpfiles) {
    if (!(Test-Path -Path (Join-Path -Path $PSScriptRoot -ChildPath $file) )) {
        Write-Verbose "Не найден файл: $file"
        $messageParam.Body += "`nНе найден обязательный файл '$file'"
        Send-MailMessage @messageParam
        throw
    }
}
#Endregion
#region Prepare
$winscpdllpath = Join-Path -Path $PSScriptRoot -ChildPath ($winscpfiles | Where-Object { $_ -like "*.dll" })
try {
    Add-Type -Path $winscpdllpath -ErrorAction Stop 
}
catch {
    $messageParam.Body += "`nНе удается добавить DLL файл '$winscpdllpath'"
    Send-MailMessage @messageParam
    Write-Verbose "Не удалось добавить DLL файл '$winscpdllpath'"
    throw "Ошибка: $($_.Exception.Message)"
}

$sessionOptions = New-Object WinSCP.SessionOptions -Property @{
    Protocol              = [WinSCP.Protocol]::Sftp
    HostName              = $RemoteHostName
    UserName              = $CredObject.UserName 
    Password              = $CredObject.GetNetworkCredential().Password 
    SshHostKeyFingerprint = $PubkeySSHRemoteHost
}

#Endregion
#region Main
Write-Verbose "Поиск файла: $sources"
if (Test-Path -Path $sources) {
    Write-Verbose "Найден файл: '$sources', начинается загрузка на удаленный хост '$RemoteHostName'"
    $session = New-Object WinSCP.Session
    try {
        $session.Open($sessionOptions)
        $session.PutFiles("$sources", "$pathto").Check()
        $session.Dispose()
        Write-Verbose "Успешная загрузка '$sources' на удаленный хост '$RemoteHostName' в каталог '$pathto'"
    }

    Catch {
        $messageParam.Body += "`nОшибка при загрузке файла '$sources' на удаленный хост '$RemoteHostName'`nТекст ошибки: $($_.Exception.Message)"
        Send-MailMessage @messageParam 
        $session.Dispose()
        throw "Ошибка: $($_.Exception.Message)"
    }
    if (!(Test-Path -Path $LocalMonthArh)) {
        try {
            New-Item -ItemType Directory -Path $LocalMonthArh -ErrorAction Stop | Out-Null 
        }
        catch {
            $messageParam.Body += "`nОшибка при создании каталога '$LocalMonthArhs'`nТекст ошибки: $($_.Exception.Message)"
            Send-MailMessage @messageParam
            throw "Ошибка: $($_.Exception.Message)"
        }
    }
    try {
        Move-Item -Path $sources -Destination $localfile -Force -ErrorAction Stop
    }
    catch {
        Write-Verbose "Ошибка при переносе файла '$sources' в локальный каталог '$localYearArh'"
        $messageParam.Body += "`nОшибка при переносе '$sources' в локальный каталог '$localYearArh'`nТекст ошибки: $($_.Exception.Message)"
        Send-MailMessage @messageParam
        throw "Ошибка: $($_.Exception.Message)"
    }
    Write-Verbose "Выполнен перенос файла '$sources' в локальный каталог: '$localcatalogy' с именем '$localfile'"

}
else {
    Write-Verbose "Не найден файл '$sources' для загрузки на удаленный хост"
    exit
}
Write-Verbose "Начинается архивирование"
foreach ($oldcatalog in (Get-ChildItem -Path $localcatalogy -Recurse -Directory | Where-Object {
            $_.Name -match "\d{2}\.\d{2}\.\d{4}" })) {
            
    $oldcatalogdate = [datetime]::ParseExact($oldcatalog, "dd.MM.yyyy", $null)

    if ($oldcatalogdate -lt $retentionDate) {
        try {
            $pathtozip = Split-path $oldcatalog.FullName -Parent
            $zip = Join-Path $pathtozip "banki.ru-cours_$($oldcatalogdate.ToString("MM.yyyy")).zip"
            Compress-Archive -Path $oldcatalog.FullName -DestinationPath $zip -Update -ErrorAction Stop
            Write-Verbose "Папка $oldcatalog с файлом $localfile запакован."
        }
        catch {
            Write-Verbose "Ошибка при создании архива"
            $messageParam.Body += "`nОшибка при cоздании '$zip'`nТекст ошибки: $($_.Exception.Message)"
            Send-MailMessage @messageParam
            throw "Ошибка: $($_.Exception.Message)"
        }
        Remove-Item -Path $oldcatalog.FullName -Force -Recurse
       
    }
}

#Endregion
