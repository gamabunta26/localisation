function ConvertTo-Jpeg {
<#
        .Synopsis
            Converts a file to a JPG of the specified quality in the same folder
        .Description
            Converts a file to a JPG of the specified quality in the same folder. 
            If the file is already a JPG it will be overwritten at the new quality setting
        .Example
            C:\PS>  Dir -recure -include *.tif | Convert-toJPeg .\myImage.bmp
            Creates creates JPG images of quality 100 for all tif files in the current directory and it's sub directories
        .Example
            C:\PS>  Dir -recure -include *.tif | Convert-toJPeg -quality 75
            Creates JPG images of quality 75 for all tif files in the current directory and it's sub directories
        .Parameter Image
            An image object, a path to an image, or a file object representing an image file. It may be passed via the pipeline.
        .Parameter Quality
            Range 1-100, sets image quality (100 highest), lower quality will use higher rates of compression.
            The default is 100. 
    #>
[CmdletBinding()]
    param(
    [Parameter(ValueFromPipeline=$true)]    
    $image,
    
    [ValidateRange(1,100)]
    [int]$quality = 100
    )
    process {
        if (($image -is [String]) -or ($image -is [System.io.FileInfo])) {Get-Image $image | convertTo-Jpeg -quality $quality ;return}
        if  ($image.count -gt 1) {$image | convertTo-Jpeg -quality $quality ; return}
        if  (-not $image.Loadfile -and -not $image.Fullname) { return }
        write-verbose ("Processing $($image.fullName)")
        $noExtension = $image.Fullname -replace "\.\w*$",""   # "\.\w*$" means dot followed by any number of alpha chars, followed by end of string - i.e file extension
        $process = New-Object -ComObject Wia.ImageProcess
        $convertFilter = $process.FilterInfos.Item("Convert").FilterId
        $process.Filters.Add($convertFilter)
        $process.Filters.Item(1).Properties.Item("Quality") = $quality
        $process.Filters.Item(1).Properties.Item("FormatID") = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
        $newImg = $process.Apply($image.PSObject.BaseObject)
        $newImg.SaveFile("$noExtension.jpg")    
    }
}