do {
    # Prompt the user to enter the folder path
    $folderPath = Read-Host "Enter the full path to the folder containing .pptx files (or type 'exit' to quit)"
    
    # Exit the script if the user types 'exit'
    if ($folderPath -eq "exit") {
        Write-Host "Exiting the script. Goodbye!" -ForegroundColor Cyan
        break
    }
    
    # Check if the folder exists
    if (-Not (Test-Path -Path $folderPath)) {
        Write-Host "The folder path '$folderPath' does not exist. Please check and try again." -ForegroundColor Red
        continue
    }

    try {
        # Load PowerPoint COM Object
        $powerPoint = New-Object -ComObject PowerPoint.Application
        $powerPoint.Visible = $true

        # Get all .pptx files in the folder
        $pptxFiles = Get-ChildItem -Path $folderPath -Filter "*.pptx"

        if ($pptxFiles.Count -eq 0) {
            Write-Host "No .pptx files found in the folder '$folderPath'." -ForegroundColor Yellow
            continue
        }

        # Iterate over each .pptx file
        foreach ($file in $pptxFiles) {
            try {
                # Open the PowerPoint presentation
                $presentation = $powerPoint.Presentations.Open($file.FullName, $false, $false, $true)

                # Define the output .pdf path
                $pdfPath = Join-Path -Path $file.DirectoryName -ChildPath ($file.BaseName + ".pdf")

                # Export to PDF
                $presentation.SaveAs($pdfPath, 32)  # 32 corresponds to ppSaveAsPDF

                # Close the presentation
                $presentation.Close()

                Write-Host "Converted $($file.Name) to PDF successfully." -ForegroundColor Green
            } catch {
                Write-Host "Failed to convert $($file.Name): $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    } catch {
        Write-Host "PowerPoint failed to start: $($_.Exception.Message)" -ForegroundColor Red
        break
    } finally {
        # Quit PowerPoint and release resources
        if ($powerPoint -ne $null) {
            $powerPoint.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($powerPoint) | Out-Null
            Remove-Variable powerPoint
        }
    }

    Write-Host "All files in the folder '$folderPath' have been processed." -ForegroundColor Cyan

} while ($true)
