"PowerShell Script to Sort and Move Filess date wise folders"

"`n"

"This is a simple PowerShell script that will sort all of the files and move them to date wise folder according to the file creation date or date taken."

"`n"

"Best Use: Sorting and Moving Images/ Photos"

"`n"

# Collect Input From the User Keyboard

$sourcePath = Read-Host 'Source Path Location.'

$destinationPath = Read-Host 'Destination/ Target Path Location.'


# Initiate the Logic to Check if the Input for Source and Destination Path Locations Are Given.

# Otherwise, It Will Displays an Error

if ([string]::IsNullOrEmpty($sourcePath) -or [string]::IsNullOrEmpty($destinationPath)) {
  Write-Host "Invalid Source & Destination / Target Path Location"
  Pause
}
else {
  Get-ChildItem -Path $sourcePath -Recurse | Move-Item -Destination $destinationPath

  # Create a root folder if it does not exist
  New-Item $destinationPath -ItemType Directory -Force | Out-Null

  # Create a Shell Object
  $shell = New-Object -ComObject Shell.Application

  # Create a Folder Object
  $folder = $shell.Namespace($sourcePath)

  foreach ($file in $folder.Items()) {

    # Get the raw date from file metadata
    $rawDate = ($folder.GetDetailsOf($file,12) -replace [char]8206) -replace [char]8207

    if ($rawDate) {
      try {
        # Parse to Date Object
        $date = [datetime]::ParseExact($rawDate,"g",$null)

        # Provide Date Format for the Folders
        # https://docs.microsoft.com/en-us/dotnet/standard/base-types/custom-date-and-time-format-strings

        $dateString = Get-Date $date -Format "yyyy-MMM-dd"

        # Create a Path
        $newFolderPath = Join-Path $destinationPath $dateString

        # Create Folder if It Does Not Exist
        New-Item $newFolderPath -ItemType Directory -Force | Out-Null

        # Move Files
        Move-Item $file.Path -Destination $newFolderPath -Confirm:$false
      }
      catch {
        # ParseExact failed (would also catch New-Item errors)
      }
    }
    else {
      # Provide Date Format for the Folders
      # https://docs.microsoft.com/en-us/dotnet/standard/base-types/custom-date-and-time-format-strings

      $dateString = $file.LastWriteTime.ToString("yyyy-MMM-dd")

      # Create a Path
      $newFolderPath = Join-Path $destinationPath $dateString

      # Create Folder if It Does Not Exist
      New-Item $newFolderPath -ItemType Directory -Force | Out-Null

      # Move Files
      Move-Item $file.Path -Destination $newFolderPath -Confirm:$false
    }
  }

  # After the end of the loop, the below code will check for empty folders available.

  # If any empty folders available inside the path, it will delete the empty folders.
  Get-ChildItem -Path $destinationPath -Recurse -Force -Directory | Sort-Object -Property FullName -Descending
  | Where-Object { $($_ | Get-ChildItem -Force | Select-Object -First 1).Count -eq 0 } | Remove-Item -Verbose
}

# If you don't want to exit the window, uncomment the below line
# Pause
