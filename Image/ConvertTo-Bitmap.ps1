#requires -version 2.0
function ConvertTo-Bitmap {
    <#
    .Synopsis
    Converts an image to a bitmap (.bmp) file.

    .Description
    The ConvertTo-Bitmap function converts image files to .bmp file format.
    You can specify the desired image quality on a scale of 1 to 100.

    ConvertTo-Bitmap takes only COM-based image objects of the type that Get-Image returns.
    To use this function, use the Get-Image function to create image objects for the image files, 
    then submit the image objects to ConvertTo-Bitmap.

    The converted files have the same name and location as the original files but with a .bmp file name extension. 
    If a file with the same name already exists in the location, ConvertTo-Bitmap declares an error. 

    .Parameter Image
    Specifies the image objects to convert.
    The objects must be of the type that the Get-Image function returns.
    Enter a variable that contains the image objects or a command that gets the image objects, such as a Get-Image command.
    This parameter is optional, but if you do not include it, ConvertTo-Bitmap has no effect.

    .Parameter Quality
    A number from 1 to 100 that indicates the desired quality of the .bmp file.
    The default is 100, which represents the best possible quality.

    .Parameter HideProgress
    Hides the progress bar.

    .Parameter Remove
    Deletes the original file. By default, both the original file and new .bmp file are saved. 

    .Notes
    ConvertTo-Bitmap uses the Windows Image Acquisition (WIA) Layer to convert files.

    .Link
    "Image Manipulation in PowerShell": http://blogs.msdn.com/powershell/archive/2009/03/31/image-manipulation-in-powershell.aspx

    .Link
    "ImageProcess object": http://msdn.microsoft.com/en-us/library/ms630507(VS.85).aspx

    .Link 
    Get-Image

    .Link
    ConvertTo-JPEG

    .Example
    Get-Image .\MyPhoto.png | ConvertTo-Bitmap

    .Example
    # Deletes the original BMP files.
    dir .\*.jpg | get-image | ConvertTo-Bitmap –quality 100 –remove -hideProgress

    .Example
    $photos = dir $home\Pictures\Vacation\* -recurse –include *.jpg, *.png, *.gif
    $photos | get-image | ConvertTo-Bitmap
    #>
    param(
    [Parameter(ValueFromPipeline=$true)]    
    $Image,
    
    [ValidateRange(1,100)]
    [int]$Quality = 100
    )
    process {
     if (($image -is [String]) -or ($image -is [System.io.FileInfo])) {Get-Image $image | convertTo-Bitmap -quality $quality ;return}
        if  ($image.count -gt 1) {$image | convertTo-Bitmap -quality $quality ; return}
        if  (-not $image.Loadfile -and -not $image.Fullname) { return }
        write-verbose ("Processing $($image.fullName)")
        $noExtension = $image.Fullname -replace "\.\w*$",""   # "\.\w*$" means dot followed by any number of alpha chars, followed by end of string - i.e file extension
        $process = New-Object -ComObject Wia.ImageProcess
        $convertFilter = $process.FilterInfos.Item("Convert").FilterId
        $process.Filters.Add($convertFilter)
        $process.Filters.Item(1).Properties.Item("Quality") = $quality
        $process.Filters.Item(1).Properties.Item("FormatID") = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
        $newImg = $process.Apply($image.PSObject.BaseObject)
        $newImg.SaveFile("$noExtension.bmp")
    }
}
