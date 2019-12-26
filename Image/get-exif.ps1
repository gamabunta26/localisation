
#Constants for Strings
$ExifIDDateTimeTaken           = 36867 #  0x9003
$ExifIDImageDescription        = 270   #  0x010E 
$ExifIDMake                    = 271   #  0x010F
$ExifIDModel                   = 272   #  0x0110
$ExifIDSoftware                = 305   #  0x0131
$ExifIDArtist                  = 315   #  0x013B 
$ExifIDCopyright               = 33432 #  0x8298
$ExifIDGPSLatRef               = 1
$ExifIDGPSLongRef              = 3
$ExifIDGPSAltRef               = 5

#Constants For Unicode strings
$ExifIDTitle                   = 40091 #  0x9C9B	XP Title     binary data
$ExifIDAuthor                  = 40093 #  0x9C9D	XP Author
$ExifIDKeywords                = 40094 #  0x9C9E	XP Keywords
$ExifIDSubject                 = 40095 #  0x9C9F	XP Subject

#$ExifIDComment                 = 40092 #  0x9C9C	XP Comment
$ExifIDComment                 = 37510 #  0x9286	UserComment

#Constants For TOKEN numbers
$ExifIDColorSpace              = 40961 #  0xA001   
$ExifIDContrast                = 41992 #  0xA408f
$ExifIDExposureMode            = 41986 #  0xA402
$ExifIDExposureProgram         = 34850 #  0x8822
$ExifIDFlash                   = 37385 #  0x9209
$ExifIDLightSource             = 37384 #  0x9208
$ExifIDMeteringMode            = 37383 #  0x9207
$ExifIDSaturation              = 41993 #  0xA409
$ExifIDSceneCaptutreMode       = 41990 #  0xA406
$ExifIDSharpness               = 41994 #  0xA40A
$ExifIDSubjectRange            = 41996 #  0xA40C
$ExifIDWhiteBalance            = 41987 #  0xA403

#Constants for special case ratios
$ExifIDExposuretime            = 33434 #  0x829A
$ExifIDGPSLattitude            = 2
$ExifIDGPSLongitude            = 4

#Constants for normal numbers
$ExifIDDigitalZoomRatio        = 41988 #  0xA404
$ExifIDExpbias                 = 37380 #  0x9204
$ExifIDFNumber                 = 33437 #  0x829D
$ExifIDFocalLength             = 37386 #  0x920A
$ExifIDFocalLengthIn35mmFormat = 41989 #  0xa405
$ExifIDHeight                  = 40963 #  0xA003
$ExifIDISO                     = 34855 #  0x8827
$ExifIDMaxApperture            = 37381 #  0x9205
$ExifIDWidth                   = 40962 #  0xA002
$ExifIDGPSAltitude             = 6
$ExifIDOrientation             = 274   #  0x0112 
$ExifIDRating                  = 18246 #  0x4746

#Constants For Byte Arrays
$ExifIDFileSource             = 41728 #  0xA300	
$ExifIDMakerNote              = 37500 #  0x927C
$ExifIDGPSVer                 = 0

Function Get-ExifItem {
[CmdletBinding()]
<#
        .Synopsis
            Returns an single item of Generic Exif data
        .Description
            Returns the data part of a single EXIF property item -not the type. 
        .Example
            C:\PS> Get-ExifItem -image $image -ExifID $ExifIDModel
            Returns the Camera model string
        .Parameter image
            The image from which the data will be read 
        .Parameter ExifID
            The ID of the required data field. 
            The module defines constants with names beginning $ExifID for the most used ones. 
    #>
  Param  (  [Parameter(ValueFromPipeline=$true)]
            [__ComObject]$image, 
            $ExifID 
          )
  Process {
    foreach ($id in $exifID) {
        try   {$item = $image.Properties.Item("$ID")  }
        Catch { Write-Verbose "Error getting exif item $ID - probably doesn't exist"  ; continue   }
        if ($item) {
            Write-Verbose "Type is $($item.type)"
            if (($item.Type -eq 1007) -or ($item.Type -eq 1006) )  {                                     # "Rational"=1006;"URational"=1007
                if (($ExifID -eq $ExifIDExposuretime) -and ($item.Value.Numerator -eq 1) ) {"1/$($item.Value.Denominator)"} else {$item.value.value} }
                elseif (($item.type -eq 1101) -or ($item.type -eq 1100)) {$item.value.string() }                                         # "VectorOfByte"=1101
            else {$item.value}
        }
    }
  }
}

Function Get-Exif{
<#
        .Synopsis
            Returns an object containing Exif data
        .Description
            Returns an object containing Exif data
        .Parameter image
            The image from which the data will be read 
        .Example 
             C:\ps> get-image IMG_1234.JPG | get-exif
             Returns the exif summary for the specified file
        .Example
             C:\ps> Get-exif .\*.JPG | format-table -auto Path, iso,Fnumber,exposuretime
             Gets the exif data for all the files in the current folder 
             and returns a table of exposure information.
#>
Param   ( [Parameter(ValueFromPipeline=$true,Mandatory=$true)]$image)
Process {
    if ($image -is [system.io.fileinfo] ) {$image = $image.FullName }
    if ($image -is [String]             ) {$image = Get-Image $image}
    if ($image.count -gt 1              ) {$image | ForEach-Object {Get-Exif $_} ; Return}
    $myvar =[int](Get-ExifItem -image $image -ExifID $ExifIDFlash)
    $flash=""
    If ($myvar –band 1)     {$Flash += "Flash fired"}
    If ($myvar -bAnd 4)     {$Flash += ", return"
                              If ($myvar -bAnd 2) {$flash += " detected"} else {$Flash += "not detected"}
    }
    If ($myvar -bAnd 8)   { If (($myvar-bAnd 16) -eq 16)  {$Flash = "Flash Auto, " + $flash} else {$flash = "Flash on, " + $flash} }
    ElseIf                      ($myvar -bAnd 16)         {$flash = "Flash off. " }
    If ($myvar -bAnd 32)    {$flash =  "No Flash function"    }
    If ($myvar -bAnd 64)    {$flash += ", Red Eye reduction"  }

    $Keywords         = Get-ExifItem -image $image -ExifID $ExifIDKeywords
    if ($keywords) {$keywords = $keywords.Split(";") }
 
    $gps = ""
    $l=(get-ExifItem -image $image -ExifID $ExifIDGPSLattitude)
    if ($l.count -eq 3) {$gps += "$($l[0].value)°$($l[1].value)'$($l[2].value)"""}
    if ($l.count -eq 2) {$gps += "$($l[0].value)°$($l[1].value)'" }
    $gps = $GPS +=  (get-ExifItem -image $image -ExifID $ExifIDGPSLatRef)
    
    $l=(get-ExifItem -image $image -ExifID $ExifIDGPSLongitude)
    if ($l.count -eq 3) {$gps += "  $($l[0].value)°$($l[1].value)'$($l[2].value)"""}
    if ($l.count -eq 2) {$gps += "  $($l[0].value)°$($l[1].value)'" }
    $gps = $GPS +=  (get-ExifItem -image $image -ExifID $ExifIDGPSLongRef)

    $altref = get-ExifItem -image $image -ExifID $ExifIDGPSAltRef
    if ($altRef -EQ 0) {$gps += ", $(get-ExifItem -image $image -ExifID $ExifIDGPSAltitude)M above Sea Level"}
    if ($altRef -EQ 1) {$gps += ", $(get-ExifItem -image $image -ExifID $ExifIDGPSAltitude)M below Sea Level"}

    $dt = (get-ExifItem -image $image -ExifID $ExifIDDateTimeTaken) 
    if ($dt) {$dt = [DateTime]::ParseExact($dt,"yyyy:MM:dd HH:mm:ss",[System.Globalization.CultureInfo]::InvariantCulture) }

    New-Object PSObject -Property @{Flash            =  $flash
                                    Keywords         =  $Keywords  
                                    GPS              =  $GPS
                                    DateTaken        =  $dt
                                    Path             =  $Image.FullName
                                    Manufacturer     =  (Get-ExifItem -image $image -ExifID $ExifIDMake)
                                    Model            =  (Get-ExifItem -image $image -ExifID $ExifIDModel)
                                    Software         =  (Get-ExifItem -image $image -ExifID $ExifIDSoftware)
                                    Subject          =  (Get-ExifItem -image $image -ExifID $ExifIDSubject)
                                    Title            =  (Get-ExifItem -image $image -ExifID $ExifIDTitle)
                                    Comment          =  (Get-ExifItem -image $image -ExifID $ExifIDComment)
                                    Author           =  (Get-ExifItem -image $image -ExifID $ExifIDAuthor)
                                    Copyright        =  (Get-ExifItem -image $image -ExifID $ExifIDCopyright)
                                    Artist           =  (Get-ExifItem -image $image -ExifID $ExifIDArtist)
                                    StarRating       =  (Get-ExifItem -image $image -ExifID $ExifIDRating)
                                    ISO              =  (Get-ExifItem -image $image -ExifID $ExifIDISO)
                                    ExposureBias     =  (Get-ExifItem -image $image -ExifID $ExifIDExpbias)
                                    Exposuretime     =  (Get-ExifItem -image $image -ExifID $ExifIDExposuretime)
                                    FNumber          =  (Get-ExifItem -image $image -ExifID $ExifIDFNumber)
                                    MaxApperture     =  (Get-ExifItem -image $image -ExifID $ExifIDMaxApperture)
                                    FocalLength      =  (Get-ExifItem -image $image -ExifID $ExifIDFocalLength)
                                    FocalLength35mm  =  (Get-ExifItem -image $image -ExifID $ExifIDFocalLengthIn35mmFormat)
                                    DigitalZoomRatio =  (Get-ExifItem -image $image -ExifID $ExifIDDigitalZoomRatio)
                                    Height           =  (Get-ExifItem -image $image -ExifID $ExifIDHeight )
                                    Width            =  (Get-ExifItem -image $image -ExifID $ExifIDWidth)
                                    SubjectRange     =  @{1="Macro"; 2="Close"; 3="Distant"}[[int](      Get-ExifItem -image $image -ExifID $ExifIDSubjectRange)]
                                    ExposureMode     =  @{0="Auto"; 1="Manual"; 2="Auto Bracket"}[[int]( Get-ExifItem -image $image -ExifID $ExifIDExposureMode)]
                                    WhiteBalance     =  @{0="Auto"; 1="Manual" }[[int](                  Get-ExifItem -image $image -ExifID $ExifIDWhiteBalance)] 
                                    Contrast         =  @{0="Normal"; 1="Soft"; 2="Hard" }[[int](        Get-ExifItem -image $image -ExifID $ExifIDContrast)]
                                    Sharpness        =  @{0="Normal"; 1="Soft"; 2="Hard" }[[int](        Get-ExifItem -image $image -ExifID $ExifIDSharpness)]
                                    Saturation       =  @{0="Normal"; 1="Low"; 2 ="High"}[[int](         Get-ExifItem -image $image -ExifID $ExifIDSaturation)]   
                                    Orientation      =  @{1="0"; 3="180"; 6="270"; 8="90"}[[int](        Get-ExifItem -image $image -ExifID $ExifIDOrientation) ]  # (0 -=Row 0 is Top, Col 0 is left / 180=Inverted, Row 0 Bottom and col 0 is Right / 270=90 Degrees CounterClockWise, Row 0 is right and col 0 is top / 90=90 Degrees ClockWise, row 0 is left, Col 0 is bottom
                                    ColorSpace       =  @{1="sRGB"; 2="Adobe RGB" }[[int](               Get-ExifItem -image $image -ExifID $ExifIDColorSpace) ] #'(the value of 2 is not standard EXIF. Instead, an Adobe RGB image is indicated by "Uncalibrated" with an InteropIndex of "R03"') 
                                    FileSource       =  @{1="Film scanner"; 2="Print scanner"; 3="Digital still camera"}[[int](                Get-ExifItem -image $image -ExifID $ExifIDFileSource)]
                                    CaptureMode      =  @{0="Standard"; 1="Landscape"; 2="Portrait"; 3="NightScene"}[[int](                    Get-ExifItem -image $image -ExifID $ExifIDSceneCaptutreMode)]
                                    MeteringMode     =  @{1="Av"; 2="Centre"; 3="Spot"; 4="Multi-Spot"; 5="Multi-Segment"; 6="Partial"}[[int]( Get-ExifItem -image $image -ExifID $ExifIDMeteringMode)]
                                    ExposureProgram  =  @{1="Manual"; 2="Program: Normal"; 3="Aperture Priority"; 4="Shutter Priority"; 5="Program: Creative"; 6="Program: Action";  7="Portrait Mode"; 8="Landscape Mode"}[[int]( Get-ExifItem -image $image -ExifID $ExifIDExposureProgram)] # Manual includes  Bulb, X-Sync Mode  Pentax Sv Mode and TAv mode report 0 - unknown.  Creative is  Depth of field Biased and Action is shutter biased
                                    LightSource      =  @{0="Auto";  1="Daylight"; 2="Fluorescent";  3="Tungsten"; 4="Flash"; 9="Fine Weather"; 10="Cloudy Weather"; 11="Shade"; 12="Daylight Fluorescent"; 13="Day White Fluorescent"; 14="Cool White Fluorescent"; 15="White Fluorescent"; 17="Standard Light A"; 18="Standard Light B"; 19="Standard Light C"; 20="D55"; 21="D65"; 22="D75"; 23="D50"; 24="ISO Studio Tungsten"}[[int]( Get-ExifItem -image $image -ExifID $ExifIDLightSource)]
                                    }
  }
}
