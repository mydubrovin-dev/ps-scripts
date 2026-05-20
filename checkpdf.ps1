param(
    [string]$PdfPath,
    [switch]$CreateOnly
)

function Write-ColorOutput {
    param([string]$Message, [string]$Color = "White")
    Write-Host $Message -ForegroundColor $Color
}

$dateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Write-Host "[$dateTime] INFO: Starting AcroExch.PDDoc verification script." -ForegroundColor Cyan

# --- 1. Проверка: Убеждаемся, что скрипт запускается в 64-битном PowerShell ---
if ($env:PROCESSOR_ARCHITECTURE -ne "AMD64") {
    Write-ColorOutput -Message "[$dateTime] ERROR: This script must be run in a 64-bit PowerShell console." -Color "Red"
    Write-ColorOutput -Message "[$dateTime] INFO: Exiting." -Color "Cyan"
    exit 1
}
Write-ColorOutput -Message "[$dateTime] INFO: Running in a 64-bit PowerShell environment." -Color "Green"

# --- 2. Создание COM-объекта ---
try {
    Write-ColorOutput -Message "[$dateTime] INFO: Attempting to create COM object 'AcroExch.PDDoc'..." -Color "Cyan"
    $pdDoc = New-Object -ComObject "AcroExch.PDDoc" -ErrorAction Stop
    Write-ColorOutput -Message "[$dateTime] SUCCESS: COM object 'AcroExch.PDDoc' created successfully." -Color "Green"
} catch {
    Write-ColorOutput -Message "[$dateTime] ERROR: Failed to create COM object 'AcroExch.PDDoc'." -Color "Red"
    Write-ColorOutput -Message "[$dateTime] DETAILS: $($_.Exception.Message)" -Color "Red"
    
    if ($_.Exception.Message -match "Access is denied") {
        Write-ColorOutput -Message "[$dateTime] INFO: This is likely a DCOM permission issue. Please check the DCOM configuration for the user 'USR1CV8'." -Color "Yellow"
        Write-ColorOutput -Message "[$dateTime] ACTION: Grant 'Local Launch', 'Local Activation', and 'Local Access' permissions for 'USR1CV8' on the default COM security limits, and possibly on specific application (like Adobe Acrobat)." -Color "Yellow"
    } elseif ($_.Exception.Message -match "Class not registered") {
        Write-ColorOutput -Message "[$dateTime] INFO: 'AcroExch.PDDoc' COM class is not registered." -Color "Yellow"
        Write-ColorOutput -Message "[$dateTime] ACTION: Ensure Adobe Acrobat is properly installed." -Color "Yellow"
    }
    Write-ColorOutput -Message "[$dateTime] INFO: Exiting." -Color "Cyan"
    exit 1
}

# --- 3. Дальнейшие действия с объектом (если требуется) ---
if (-not $CreateOnly) {
    if (-not $PdfPath) {
        Write-ColorOutput -Message "[$dateTime] WARNING: No PDF path provided for testing. Skipping file open operation." -Color "Yellow"
        Write-ColorOutput -Message "[$dateTime] INFO: To test opening a PDF, run the script with: -PdfPath 'C:\path\to\your\file.pdf'" -Color "Cyan"
        # Закрываем объект, если он не нужен
        $pdDoc = $null
    } else {
        if (-not (Test-Path -Path $PdfPath)) {
            Write-ColorOutput -Message "[$dateTime] ERROR: PDF file not found at '$PdfPath'. Exiting." -Color "Red"
            $pdDoc = $null
            exit 1
        }

        try {
            Write-ColorOutput -Message "[$dateTime] INFO: Attempting to open PDF '$PdfPath'..." -Color "Cyan"
            $result = $pdDoc.Open($PdfPath)
            if ($result -eq $true) {
                Write-ColorOutput -Message "[$dateTime] SUCCESS: PDF opened successfully." -Color "Green"
                $pdDoc.Close()
                Write-ColorOutput -Message "[$dateTime] INFO: PDF closed." -Color "Cyan"
            } else {
                Write-ColorOutput -Message "[$dateTime] ERROR: The Open method returned `$false." -Color "Red"
                Write-ColorOutput -Message "[$dateTime] INFO: This might indicate a problem with the PDF file itself or permissions." -Color "Yellow"
            }
        } catch {
            Write-ColorOutput -Message "[$dateTime] ERROR: An unexpected error occurred while opening the PDF." -Color "Red"
            Write-ColorOutput -Message "[$dateTime] DETAILS: $($_.Exception.Message)" -Color "Red"
        } finally {
            # Освобождаем COM-объект
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pdDoc) | Out-Null
        }
    }
} else {
    Write-ColorOutput -Message "[$dateTime] INFO: CreateOnly flag is set. Exiting after object creation." -Color "Cyan"
    # Если объект был создан, но не используется, освобождаем его.
    $pdDoc = $null
}

Write-ColorOutput -Message "[$dateTime] INFO: Script finished." -Color "Cyan"