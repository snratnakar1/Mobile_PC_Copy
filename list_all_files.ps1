# 1. Connect to the virtual "This PC" folder
$shell = New-Object -ComObject Shell.Application
$thisPC = $shell.NameSpace(0x11)

# 2. Find your connected phone
$phone = $thisPC.Items() | Where-Object { $_.Type -match "Portable|Media|Phone" -or $_.Name -match "Galaxy|Pixel|iPhone|Android" }

if (-not $phone) { 
    Write-Host "Error: No phone found. Ensure it is unlocked and in 'File Transfer' mode." -ForegroundColor Red
} else {
    Write-Host "Connected to phone: $($phone.Name)" -ForegroundColor Green
    
    # 3. Drill down into the phone's internal storage
    $storage = $phone.GetFolder.Items() | Where-Object { $_.Name -match "Internal|Storage|Phone" }
    
    # 4. Target the Download folder (Change "Download" below to your folder name if needed)
    $folder1 = $storage.GetFolder.Items() | Where-Object { $_.Name -eq "DCIM" }

    # 4. Target the Download folder (Change "Download" below to your folder name if needed)
    $folder = $folder1.GetFolder.Items() | Where-Object { $_.Name -eq "Camera" }
    
    if (-not $folder) {
        Write-Host "Could not find target folder. Listing top-level folders instead:" -ForegroundColor Yellow
        $storage.GetFolder.Items() | Select-Object Name
    } else {
        # 5. Output all file names to a text file on your Desktop
        Write-Host "Extracting file names..." -ForegroundColor Cyan
        $files = $folder.GetFolder.Items()
        $files | Select-Object Name | Out-File "$env:USERPROFILE\Documents\phone_files.txt"
        
        Write-Host "Success! Check your PC Docuesnts for 'phone_files.txt'" -ForegroundColor Green
    }
}
