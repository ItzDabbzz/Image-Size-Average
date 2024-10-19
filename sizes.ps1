##
# This script calculates the average width and height of all images in a folder.
# It then outputs the average width and height, as well as the individual image sizes.
# The output is written to a text file.
#
# Usage: 
#  .\sizes.ps1 -folder "C:\path\to\your\folder" -outputFile "C:\path\to\your\output.txt"
#
# - github.com/ItzDabbzz 2024 GNU GPLv3
##

param(
    [Parameter(Mandatory=$true)]
    [string]$folder,
    
    [Parameter(Mandatory=$true)]
    [string]$outputFile
)

$shell = New-Object -ComObject Shell.Application
$objFolder = $shell.Namespace($folder)

$sizes = @()
$files = Get-ChildItem $folder -File

foreach ($file in $files) {
    $item = $objFolder.ParseName($file.Name)
    $dimensions = $objFolder.GetDetailsOf($item, 31)  # 31 is the index for dimensions
    if ($dimensions -match '(\d+)\s*x\s*(\d+)') {
        $width = [int]$matches[1]
        $height = [int]$matches[2]
        $sizes += @{Width = $width; Height = $height}
    }
}

$averageWidth = ($sizes | Measure-Object -Property Width -Average).Average
$averageHeight = ($sizes | Measure-Object -Property Height -Average).Average

$output = @"
Total images: $($sizes.Count)
Average width: $([math]::Round($averageWidth, 2))
Average height: $([math]::Round($averageHeight, 2))

Individual image sizes:
"@

foreach ($size in $sizes) {
    $output += "`n$($size.Width) x $($size.Height)"
}

$output | Out-File $outputFile

Write-Host "Image size information has been written to $outputFile"
