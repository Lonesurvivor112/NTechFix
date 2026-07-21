# Set the source folder (where the extracted ZIP contents are)
$sourceFolder = "C:\Users\jdoe\contoso.com\IT - Documents\Scripts\Python\AmazonPDF-Tools\SourcePDFs\invoices__05_29_2026__05_11_2026"

# Set the destination folder (where you want all PDFs/images to go)
$destinationFolder = "C:\Users\jdoe\contoso.com\IT - Documents\Scripts\Python\AmazonPDF-Tools\DestinationPDFs"

# Create the destination folder if it doesn't exist
if (!(Test-Path -Path $destinationFolder)) {
    New-Item -ItemType Directory -Path $destinationFolder
}

# Get all PDF and image files recursively
$files = Get-ChildItem -Path $sourceFolder -Recurse -Include *.pdf, *.jpg, *.jpeg, *.png

# Copy each file to the destination folder
foreach ($file in $files) {
    $destinationPath = Join-Path -Path $destinationFolder -ChildPath $file.Name

    # If a file with the same name exists, add a number to avoid overwriting
    $counter = 1
    while (Test-Path $destinationPath) {
        $destinationPath = Join-Path -Path $destinationFolder -ChildPath ("{0}_{1}{2}" -f $file.BaseName, $counter, $file.Extension)
        $counter++
    }

    Copy-Item -Path $file.FullName -Destination $destinationPath
}

Write-Host "✅ All files have been copied to $destinationFolder"
