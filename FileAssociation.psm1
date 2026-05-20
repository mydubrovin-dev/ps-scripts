<#
.SYNOPSIS
    Модуль для просмотра и изменения файловых ассоциаций
    на уровне системы (HKLM) и отдельных пользователей.
.DESCRIPTION
    Содержит функции:
        Get-FileAssociation      - просмотр ассоциаций
        Set-FileAssociation      - изменение ассоциаций
        Remove-FileAssociation   - удаление пользовательской ассоциации
    Может работать с текущим пользователем, заданным по имени (через NTUSER.DAT)
    и системными ассоциациями (HKLM).
#>

# Вспомогательная функция загрузки куста пользователя
function Load-UserHive {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserName,
        [Parameter(Mandatory=$true)]
        [string]$HiveName = "TempHive_$UserName"
    )
    $ntuserPath = "C:\Users\$UserName\NTUSER.DAT"
    if (-not (Test-Path $ntuserPath)) {
        throw "Не найден файл профиля: $ntuserPath"
    }
    $loadResult = reg load "HKU\$HiveName" $ntuserPath 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "Не удалось загрузить куст: $loadResult"
    }
    return $HiveName
}

# Вспомогательная функция выгрузки куста
function Unload-UserHive {
    param([string]$HiveName)
    reg unload "HKU\$HiveName" 2>$null
}

# Вспомогательная функция получения ProgId из ветки (HKLM или HKU)
function Get-ProgId {
    param(
        [string]$BasePath,   # например "HKLM:\Software\Classes" или "HKU:\TempHive\Software\Classes"
        [string]$Extension   # начинается с точки
    )
    $extPath = "$BasePath\$Extension"
    if (Test-Path $extPath) {
        return (Get-ItemProperty -Path $extPath -Name "(Default)" -ErrorAction SilentlyContinue)."(Default)"
    }
    return $null
}

# Вспомогательная функция получения команды открытия
function Get-OpenCommand {
    param(
        [string]$BasePath,
        [string]$ProgId
    )
    $cmdPath = "$BasePath\$ProgId\shell\open\command"
    if (Test-Path $cmdPath) {
        return (Get-ItemProperty -Path $cmdPath -Name "(Default)" -ErrorAction SilentlyContinue)."(Default)"
    }
    return $null
}

<#
.SYNOPSIS
    Получить текущие файловые ассоциации.
.DESCRIPTION
    Может показать ассоциации для:
    - текущего пользователя (по умолчанию),
    - указанного пользователя (-UserName),
    - всей системы (-System),
    - или всех пользователей (-AllProfiles).
.PARAMETER Extension
    Конкретное расширение (с точкой или без). Если не указано — выводятся все ассоциации.
.PARAMETER UserName
    Имя пользователя, чей профиль нужно прочитать (загружается NTUSER.DAT).
.PARAMETER System
    Показать только системные ассоциации (HKLM).
.PARAMETER AllProfiles
    Показать ассоциации для всех локальных пользователей.
.EXAMPLE
    Get-FileAssociation .txt
    Get-FileAssociation pdf -UserName Ivanov
    Get-FileAssociation -System
    Get-FileAssociation -AllProfiles
#>
function Get-FileAssociation {
    [CmdletBinding(DefaultParameterSetName = 'CurrentUser')]
    param(
        [Parameter(Position=0, ParameterSetName='CurrentUser')]
        [Parameter(Position=0, ParameterSetName='SpecificUser')]
        [Parameter(Position=0, ParameterSetName='System')]
        [Parameter(Position=0, ParameterSetName='AllProfiles')]
        [string]$Extension,

        [Parameter(Mandatory=$true, ParameterSetName='SpecificUser')]
        [string]$UserName,

        [Parameter(Mandatory=$true, ParameterSetName='System')]
        [switch]$System,

        [Parameter(Mandatory=$true, ParameterSetName='AllProfiles')]
        [switch]$AllProfiles
    )

    # Нормализация расширения
    if ($Extension -and $Extension[0] -ne '.') {
        $Extension = ".$Extension"
    }

    $results = @()

    # --- Текущий пользователь (HKCU) ---
    if ($PSCmdlet.ParameterSetName -eq 'CurrentUser') {
        Write-Host "=== Ассоциации текущего пользователя: $env:USERNAME ===" -ForegroundColor Cyan
        $baseHive = "HKCU:\Software\Classes"
        $choiceBase = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts"
        if ($Extension) {
            $progId = Get-ProgId -BasePath $baseHive -Extension $Extension
            $cmd = if ($progId) { Get-OpenCommand -BasePath $baseHive -ProgId $progId } else { $null }
            $userChoice = $null
            $choicePath = "$choiceBase\$Extension\UserChoice"
            if (Test-Path $choicePath) {
                $userChoice = (Get-ItemProperty -Path $choicePath -Name "ProgId" -ErrorAction SilentlyContinue).ProgId
                $userApp = (Get-ItemProperty -Path $choicePath -Name "Application" -ErrorAction SilentlyContinue).Application
                $userChoice = if ($userChoice) { $userChoice } elseif ($userApp) { $userApp }
            }
            [PSCustomObject]@{
                User = $env:USERNAME
                Extension = $Extension
                ProgId = $progId
                Command = $cmd
                UserChoice = $userChoice
            }
        } else {
            $extensions = Get-ChildItem $baseHive | Where-Object { $_.PSChildName -match '^\..+' }
            foreach ($ext in $extensions) {
                $extName = $ext.PSChildName
                $progId = Get-ProgId -BasePath $baseHive -Extension $extName
                if ($progId) {
                    [PSCustomObject]@{
                        User = $env:USERNAME
                        Extension = $extName
                        ProgId = $progId
                    }
                }
            }
        }
    }

    # --- Конкретный пользователь (через куст) ---
    elseif ($PSCmdlet.ParameterSetName -eq 'SpecificUser') {
        Write-Host "=== Ассоциации пользователя: $UserName ===" -ForegroundColor Cyan
        $hiveName = $null
        try {
            $hiveName = Load-UserHive -UserName $UserName
            $baseHive = "HKU:\$hiveName\Software\Classes"
            $choiceBase = "HKU:\$hiveName\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts"
            if ($Extension) {
                $progId = Get-ProgId -BasePath $baseHive -Extension $Extension
                $cmd = if ($progId) { Get-OpenCommand -BasePath $baseHive -ProgId $progId } else { $null }
                $userChoice = $null
                $choicePath = "$choiceBase\$Extension\UserChoice"
                if (Test-Path $choicePath) {
                    $uc = (Get-ItemProperty -Path $choicePath -Name "ProgId" -ErrorAction SilentlyContinue).ProgId
                    $ua = (Get-ItemProperty -Path $choicePath -Name "Application" -ErrorAction SilentlyContinue).Application
                    $userChoice = if ($uc) { $uc } elseif ($ua) { $ua }
                }
                [PSCustomObject]@{
                    User = $UserName
                    Extension = $Extension
                    ProgId = $progId
                    Command = $cmd
                    UserChoice = $userChoice
                }
            } else {
                $extensions = Get-ChildItem $baseHive | Where-Object { $_.PSChildName -match '^\..+' }
                foreach ($ext in $extensions) {
                    $extName = $ext.PSChildName
                    $progId = Get-ProgId -BasePath $baseHive -Extension $extName
                    if ($progId) {
                        [PSCustomObject]@{
                            User = $UserName
                            Extension = $extName
                            ProgId = $progId
                        }
                    }
                }
            }
        }
        finally {
            if ($hiveName) { Unload-UserHive -HiveName $hiveName }
        }
    }

    # --- Системные ассоциации (HKLM) ---
    elseif ($PSCmdlet.ParameterSetName -eq 'System') {
        Write-Host "=== Системные ассоциации (HKLM) ===" -ForegroundColor Cyan
        $baseHive = "HKLM:\Software\Classes"
        if ($Extension) {
            $progId = Get-ProgId -BasePath $baseHive -Extension $Extension
            $cmd = if ($progId) { Get-OpenCommand -BasePath $baseHive -ProgId $progId } else { $null }
            [PSCustomObject]@{
                User = "SYSTEM"
                Extension = $Extension
                ProgId = $progId
                Command = $cmd
            }
        } else {
            $extensions = Get-ChildItem $baseHive | Where-Object { $_.PSChildName -match '^\..+' }
            foreach ($ext in $extensions) {
                $extName = $ext.PSChildName
                $progId = Get-ProgId -BasePath $baseHive -Extension $extName
                if ($progId) {
                    [PSCustomObject]@{
                        User = "SYSTEM"
                        Extension = $extName
                        ProgId = $progId
                    }
                }
            }
        }
    }

    # --- Все профили ---
    elseif ($PSCmdlet.ParameterSetName -eq 'AllProfiles') {
        Write-Host "=== Ассоциации всех локальных пользователей ===" -ForegroundColor Cyan
        $users = Get-ChildItem "C:\Users" -Directory | Where-Object { $_.Name -notin @('Public','Default','All Users','Default User') }
        foreach ($user in $users) {
            $hiveName = $null
            try {
                $hiveName = Load-UserHive -UserName $user.Name -HiveName "Temp_$($user.Name)"
                $baseHive = "HKU:\$hiveName\Software\Classes"
                if ($Extension) {
                    $progId = Get-ProgId -BasePath $baseHive -Extension $Extension
                    if ($progId) {
                        [PSCustomObject]@{
                            User = $user.Name
                            Extension = $Extension
                            ProgId = $progId
                        }
                    }
                } else {
                    $extensions = Get-ChildItem $baseHive -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -match '^\..+' }
                    foreach ($ext in $extensions) {
                        $extName = $ext.PSChildName
                        $progId = Get-ProgId -BasePath $baseHive -Extension $extName
                        if ($progId) {
                            [PSCustomObject]@{
                                User = $user.Name
                                Extension = $extName
                                ProgId = $progId
                            }
                        }
                    }
                }
            }
            catch {
                Write-Warning "Не удалось обработать профиль $($user.Name): $_"
            }
            finally {
                if ($hiveName) { Unload-UserHive -HiveName $hiveName }
            }
        }
    }
}

<#
.SYNOPSIS
    Установить файловую ассоциацию.
.DESCRIPTION
    Может изменять ассоциацию для:
    - текущего пользователя (по умолчанию),
    - конкретного пользователя (-UserName),
    - всей системы (-System, требует прав администратора),
    - всех пользователей (-AllProfiles, требует прав администратора).
.PARAMETER Extension
    Расширение файла (например, ".txt" или "txt").
.PARAMETER ProgId
    Идентификатор программы (ProgId), например "MSEdgeHTM" или полный путь к EXE.
    Можно задать и путь к исполняемому файлу: "C:\Program Files\...\app.exe".
.PARAMETER UserName
    Имя пользователя, для которого меняем ассоциацию (загружается куст NTUSER.DAT).
.PARAMETER System
    Установить системную ассоциацию (HKLM). Требует прав администратора.
.PARAMETER AllProfiles
    Применить ко всем локальным пользователям (изменяет кусты каждого профиля).
    Требует прав администратора.
.EXAMPLE
    Set-FileAssociation .pdf AcroExch.Document.DC -System
    Set-FileAssociation txt "C:\Windows\notepad.exe" -UserName Ivanov
    Set-FileAssociation pdf MyPDFApp -AllProfiles
#>
function Set-FileAssociation {
    [CmdletBinding(DefaultParameterSetName = 'CurrentUser')]
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [string]$Extension,
        [Parameter(Mandatory=$true, Position=1)]
        [string]$ProgId,

        [Parameter(Mandatory=$true, ParameterSetName='SpecificUser')]
        [string]$UserName,

        [Parameter(Mandatory=$true, ParameterSetName='System')]
        [switch]$System,

        [Parameter(Mandatory=$true, ParameterSetName='AllProfiles')]
        [switch]$AllProfiles
    )

    # Нормализация расширения
    if ($Extension[0] -ne '.') {
        $Extension = ".$Extension"
    }

    # Если ProgId содержит "\", значит это путь к исполняемому файлу — создаём временный ProgId или используем как приложение
    $isAppPath = $ProgId.Contains("\")

    # Функция для установки ассоциации в конкретную ветку реестра
    function Set-UserExtension {
        param(
            [string]$ClassesPath,    # например "HKCU:\Software\Classes" или "HKU:\TempHive\Software\Classes"
            [string]$ChoiceBase,     # например "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts" или аналогично в кусте
            [string]$Extension,
            [string]$ProgId,
            [bool]$IsAppPath
        )
        # Прописываем расширение -> ProgId
        $extPath = "$ClassesPath\$Extension"
        New-Item -Path $extPath -Force | Out-Null
        Set-ItemProperty -Path $extPath -Name "(Default)" -Value $ProgId

        # Если был передан путь к приложению, прописываем команду открытия
        if ($IsAppPath) {
            $cmdPath = "$ClassesPath\$ProgId\shell\open\command"
            New-Item -Path $cmdPath -Force | Out-Null
            Set-ItemProperty -Path $cmdPath -Name "(Default)" -Value "`"$ProgId`" `"%1`""
        }

        # Устанавливаем пользовательский выбор (UserChoice)
        $choicePath = "$ChoiceBase\$Extension\UserChoice"
        New-Item -Path $choicePath -Force | Out-Null
        Set-ItemProperty -Path $choicePath -Name "ProgId" -Value $ProgId -Force
        # Если передали путь, добавляем Application
        if ($IsAppPath) {
            Set-ItemProperty -Path $choicePath -Name "Application" -Value $ProgId -Force
        }
    }

    # --- Текущий пользователь ---
    if ($PSCmdlet.ParameterSetName -eq 'CurrentUser') {
        Write-Host "Установка ассоциации '$Extension' -> '$ProgId' для текущего пользователя $env:USERNAME" -ForegroundColor Green
        Set-UserExtension -ClassesPath "HKCU:\Software\Classes" `
                          -ChoiceBase "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts" `
                          -Extension $Extension -ProgId $ProgId -IsAppPath $isAppPath
        Write-Host "Готово." -ForegroundColor Green
    }

    # --- Конкретный пользователь ---
    elseif ($PSCmdlet.ParameterSetName -eq 'SpecificUser') {
        Write-Host "Установка ассоциации '$Extension' -> '$ProgId' для пользователя $UserName" -ForegroundColor Green
        $hiveName = $null
        try {
            $hiveName = Load-UserHive -UserName $UserName
            Set-UserExtension -ClassesPath "HKU:\$hiveName\Software\Classes" `
                              -ChoiceBase "HKU:\$hiveName\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts" `
                              -Extension $Extension -ProgId $ProgId -IsAppPath $isAppPath
            Write-Host "Готово." -ForegroundColor Green
        }
        finally {
            if ($hiveName) { Unload-UserHive -HiveName $hiveName }
        }
    }

    # --- Системная ассоциация (HKLM) ---
    elseif ($PSCmdlet.ParameterSetName -eq 'System') {
        Write-Host "Установка системной ассоциации '$Extension' -> '$ProgId' (HKLM)" -ForegroundColor Green
        # Проверка прав администратора
        $extPath = "HKLM:\Software\Classes\$Extension"
        New-Item -Path $extPath -Force | Out-Null
        Set-ItemProperty -Path $extPath -Name "(Default)" -Value $ProgId
        if ($isAppPath) {
            $cmdPath = "HKLM:\Software\Classes\$ProgId\shell\open\command"
            New-Item -Path $cmdPath -Force | Out-Null
            Set-ItemProperty -Path $cmdPath -Name "(Default)" -Value "`"$ProgId`" `"%1`""
        }
        Write-Host "Готово." -ForegroundColor Green
    }

    # --- Все профили ---
    elseif ($PSCmdlet.ParameterSetName -eq 'AllProfiles') {
        Write-Host "Установка ассоциации '$Extension' -> '$ProgId' для всех локальных пользователей" -ForegroundColor Green
        $users = Get-ChildItem "C:\Users" -Directory | Where-Object { $_.Name -notin @('Public','Default','All Users','Default User') }
        foreach ($user in $users) {
            $hiveName = $null
            try {
                $hiveName = Load-UserHive -UserName $user.Name -HiveName "Temp_$($user.Name)"
                Set-UserExtension -ClassesPath "HKU:\$hiveName\Software\Classes" `
                                  -ChoiceBase "HKU:\$hiveName\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts" `
                                  -Extension $Extension -ProgId $ProgId -IsAppPath $isAppPath
                Write-Host "  OK: $($user.Name)" -ForegroundColor Gray
            }
            catch {
                Write-Warning "  Ошибка для $($user.Name): $_"
            }
            finally {
                if ($hiveName) { Unload-UserHive -HiveName $hiveName }
            }
        }
        Write-Host "Готово." -ForegroundColor Green
    }
}

<#
.SYNOPSIS
    Удалить пользовательскую ассоциацию (сбросить выбор).
.DESCRIPTION
    Удаляет запись UserChoice для заданного расширения у текущего или указанного пользователя.
    После этого будет использоваться системная ассоциация.
.PARAMETER Extension
    Расширение (с точкой или без).
.PARAMETER UserName
    Имя пользователя. Если опущено — текущий пользователь.
.EXAMPLE
    Remove-FileAssociation .pdf
    Remove-FileAssociation pdf -UserName Petrov
#>
function Remove-FileAssociation {
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [string]$Extension,
        [string]$UserName
    )

    if ($Extension[0] -ne '.') { $Extension = ".$Extension" }

    if ($UserName) {
        Write-Host "Сброс ассоциации '$Extension' для пользователя $UserName" -ForegroundColor Yellow
        $hiveName = $null
        try {
            $hiveName = Load-UserHive -UserName $UserName
            $choicePath = "HKU:\$hiveName\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$Extension\UserChoice"
            if (Test-Path $choicePath) {
                Remove-Item -Path $choicePath -Force -Recurse
                Write-Host "Пользовательский выбор удалён." -ForegroundColor Green
            } else {
                Write-Host "Пользовательский выбор отсутствует." -ForegroundColor Yellow
            }
        }
        finally {
            if ($hiveName) { Unload-UserHive -HiveName $hiveName }
        }
    } else {
        Write-Host "Сброс ассоциации '$Extension' для текущего пользователя" -ForegroundColor Yellow
        $choicePath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$Extension\UserChoice"
        if (Test-Path $choicePath) {
            Remove-Item -Path $choicePath -Force -Recurse
            Write-Host "Пользовательский выбор удалён." -ForegroundColor Green
        } else {
            Write-Host "Пользовательский выбор отсутствует." -ForegroundColor Yellow
        }
    }
}

# Экспорт функций
Export-ModuleMember -Function Get-FileAssociation, Set-FileAssociation, Remove-FileAssociation