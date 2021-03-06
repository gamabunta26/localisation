Function Copy-Image {
<#
        .Synopsis
            Copies an image, applying EXIF data from GPS data points   
        .Description
            Copies an image, applying EXIF data from GPS data points   
        .Example
            C:\PS>  Dir E:\dcim –inc IMG*.jpg –rec | Copy-Image -Keywords "Oxfordshire" -rotate -DestPath "$env:userprofile\pictures\oxford" -replace  "IMG","OX-"
            Copies IMG files from folders under E:\DCIM to the user's picture\Oxford folder, replacing IMG in the file name with OX-.
            The Keywords field is set to Oxfordshire, pictures are GeoTagged with the data in $points and rotated. 
        .Parameter Image
            A WIA image object, a path to an image, or a file object representing an image file. It may be passed via the pipeline.
        .Parameter Destination
            The FOLDER to which the file should be saved.
        .Parameter Keywords
            If specified, sets the keywords Exif field.
        .Parameter Title
            If specified, sets the Title Exif field..    
        .Parameter Replace
            If specified, this contains two values seperated by a comma specifying a replacement in the file name
        .Parameter Rotate
            If this switch is specified, the image will be auto-rotated based on its orientation filed
        .Parameter NoClobber
            Unless this switch is specified, a pre-existing image WILL be over-written
        .Parameter ReturnInfo
            If this switch is specified, the path to the saved image will be returned. 
    #>
[CmdletBinding(SupportsShouldProcess=$true)]
Param ( [Parameter(ValueFromPipeline=$true, Mandatory=$true)][Alias("Path","FullName")]$image , 
        [Parameter(Mandatory=$true)][ValidateScript({Test-path $_ })][string]$Destination ,  
        $keywords , $Title, $replace,$filter,[switch]$Rotate,[switch]$NoClobber,[switch]$ReturnInfo, $psc 
)
process {
        if ($psc -eq $null)  {$psc = $pscmdlet} ; if (-not $PSBoundParameters.psc) {$PSBoundParameters.add("psc",$psc)}
        if ($image -is [system.io.fileinfo] ) {$image = $image.FullName }
        if ($image -is [String]             ) {[Void]$PSBoundParameters.Remove("Image") 
                                               Get-Image $image | Copy-Image @PSBoundParameters
                                               return
        }
        if ($Image.count -gt 1              ) {[Void]$PSBoundParameters.Remove("Image") 
                                               $Image | ForEach-object {Copy-Image -image $_ @PSBoundParameters}
                                               return
        }
        if ($image -is [__comObject])  {
           Write-Verbose ("Processing " + $image.fullname)
           if (-not $filter)  {$filter = new-Imagefilter}
           if ($rotate)       {$orient=Get-ExifItem -image  $image       -ExifID $ExifIDOrientation}  # Leave $orient unset if we aren't rotating
           if ($keywords)     {Add-exifFilter       -filter $filter      -ExifID $ExifIDKeywords   -typeid 1101 -string $keywords }
           if ($Title)        {Add-exifFilter       -filter $filter      -ExifID $ExifIDTitle      -typeid 1101 -string $Title    }
           if ($orient -eq 8) {Add-RotateFlipFilter -filter $filter      -angle  270   # Orientation 8=90 degrees, 6=270 degrees, rotate round to 360
                               Add-exifFilter       -filter $filter      -ExifID $ExifIDOrientation -typeid $1003 -value 1      
                               write-verbose "Rotating image counter-clockwise"}
           if ($orient -eq 6) {Add-RotateFlipFilter -filter $filter      -angle  90  
                               Add-exifFilter       -filter $filter      -ExifID $ExifIDOrientation -typeid $1003 -value 1      
                               write-verbose "Rotating image clockwise"}
           if ($replace)      {$SavePath= join-path -Path   (Resolve-Path $Destination) -ChildPath ((Split-Path $image.FullName -Leaf) -Replace $replace)}
           else               {$SavePath= join-path -Path   (Resolve-Path $Destination) -ChildPath  (Split-Path $image.FullName -Leaf)  }
           Set-ImageFilter    -image $image         -filter $filter      -SaveName $savePath -noClobber:$NoClobber -psc $psc
           $orient = $image =  $filter = $null
           if ($returnInfo) {$SavePath}
        }
    }
}