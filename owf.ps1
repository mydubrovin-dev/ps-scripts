<#
.SYNOPSIS
    Opens a file in Microsoft Word via DCOM running as DOMAIN\USER.
.DESCRIPTION
    Creates COM object Word.Application (DCOM) and opens specified document.
    IMPORTANT: run this script as user DOMAIN\USER.
.PARAMETER FilePath
    Full path to the file to open (e.g., C:\Docs\report.docx).
.PARAMETER ComputerName
    (Optional) Remote computer name where Word runs. If not specified, uses local machine.
.EXAMPLE
    .\OpenWordFile.ps1 -FilePath "D:\Contracts\agreement.docx"
    Opens file in local Word.
.EXAMPLE
    .\OpenWordFile.ps1 -FilePath "\\server\share\doc.docx" -ComputerName "REMOTE-PC"
    Opens file via DCOM on remote PC.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$FilePath,

    [string]$ComputerName = $null
)

$requiredUser = "DOMAIN\USER"
$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

if ($currentUser -ne $requiredUser) {
    Write-Warning "Script is running as $currentUser, not as $requiredUser."
}

try {
    Write-Host "Creating Word object via DCOM..." -ForegroundColor Cyan

    if ($ComputerName) {
        Write-Host "Target computer: $ComputerName"
        $wordType = [Type]::GetTypeFromProgID("Word.Application", $ComputerName)
        if (-not $wordType) {
            throw "Cannot get Word.Application type on $ComputerName. Check DCOM settings."
        }
        $word = [Activator]::CreateInstance($wordType)
    } else {
        $word = New-Object -ComObject Word.Application
    }

    $word.Visible = $true
    $document = $word.Documents.Open($FilePath)
    Write-Host "File '$FilePath' opened successfully in Microsoft Word." -ForegroundColor Green
}
catch {
    Write-Error "Error opening file via DCOM: $_"
    exit 1
}