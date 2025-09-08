function Invoke-SaRAOfficeScrub {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    Write-Host "`n🧹 Starting SaRA Office Scrub..." -ForegroundColor Magenta

    $saraUrl = "https://aka.ms/SaRA_EnterpriseVersionFiles"
    $saraZipPath = "C:\Temp\SaRA.zip"
    $saraExtractPath = "C:\Temp\SaRA"
    $saraExePath = "$saraExtractPath\done\SaRAcmd.exe"
    $maxRetries = 3
    $retryCount = 0
    $downloadSuccess = $false

    try {
        # Ensure C:\Temp exists
        if (-not (Test-Path "C:\Temp")) {
            New-Item -Path "C:\Temp" -ItemType Directory -Force | Out-Null
            Write-Host "📂 Created C:\Temp folder." -ForegroundColor Cyan
        }

        # Download SaRA ZIP with retries
        Write-Host "📥 Downloading SaRA tool..." -ForegroundColor Blue
        while ($retryCount -lt $maxRetries -and -not $downloadSuccess) {
            try {
                Invoke-WebRequest -Uri $saraUrl -OutFile $saraZipPath -ErrorAction Stop
                # Validate file is a ZIP
                $zipTest = [System.IO.Compression.ZipFile]::OpenRead($saraZipPath)
                $zipTest.Dispose()
                $downloadSuccess = $true
                Write-Host "✅ SaRA tool downloaded to $saraZipPath." -ForegroundColor Green
            } catch {
                $retryCount++
                Write-Host "⚠️ Download attempt $retryCount failed: $_" -ForegroundColor Yellow
                if ($retryCount -lt $maxRetries) {
                    Write-Host "Retrying download in 5 seconds..." -ForegroundColor Yellow
                    Start-Sleep -Seconds 5
                } else {
                    throw "Failed to download valid SaRA ZIP after $maxRetries attempts: $_"
                }
            }
        }

        # Extract ZIP
        Write-Host "📦 Extracting SaRA to $saraExtractPath..." -ForegroundColor Blue
        if (Test-Path $saraExtractPath) {
            Remove-Item -Path $saraExtractPath -Recurse -Force -ErrorAction Stop
        }
        Expand-Archive -Path $saraZipPath -DestinationPath $saraExtractPath -Force -ErrorAction Stop
        Write-Host "✅ SaRA extracted successfully." -ForegroundColor Green

        # Verify SaRAcmd.exe exists
        if (-not (Test-Path $saraExePath)) {
            throw "SaRAcmd.exe not found at $saraExePath after extraction."
        }

        # Run SaRA Office Scrub command
        Write-Host "🛠️ Running SaRA Office Scrub command..." -ForegroundColor Blue
        $process = Start-Process -FilePath $saraExePath -ArgumentList "-S OfficeScrubScenario -AcceptEula -OfficeVersion All" -Wait -PassThru -ErrorAction Stop
        if ($process.ExitCode -ne 0) {
            throw "SaRAcmd.exe failed with exit code $($process.ExitCode)."
        }
        Write-Host "✅ SaRA Office Scrub completed successfully." -ForegroundColor Green
    } catch {
        Write-Host "❌ SaRA Office Scrub failed: $_" -ForegroundColor Red
        Write-Host "Please download SaRAcmd.exe from https://aka.ms/SaRA_EnterpriseVersionFiles, extract to C:\Temp\SaRA, and run 'SaRAcmd.exe -S OfficeScrubScenario -AcceptEula -OfficeVersion All' from C:\Temp\SaRA\done manually." -ForegroundColor Yellow
    } finally {
        # Cleanup
        try {
            Write-Host "🧹 Cleaning up SaRA files..." -ForegroundColor DarkGray
            if (Test-Path $saraZipPath) {
                Remove-Item -Path $saraZipPath -Force -ErrorAction Stop
            }
            if (Test-Path $saraExtractPath) {
                Remove-Item -Path $saraExtractPath -Recurse -Force -ErrorAction Stop
            }
            Write-Host "✅ SaRA files cleaned up." -ForegroundColor Green
        } catch {
            Write-Host "⚠️ Failed to clean up SaRA files: $_" -ForegroundColor Yellow
        }
    }
}
Invoke-SaRAOfficeScrub
