Function ConvertTo-DateTime { 
<#
        .Synopsis
            Takes a date and time as text, parses it and returns a [DateTime]
        .Description
            Takes a date and time as text, parses it and returns a [DateTime]
        .Example
            C:\PS> ConvertTo-DateTime "2010-01-31 12:23:34"  "yyyy-MM-dd HH:mm:ss"
            Returns 31 January 2010 12:23:34 
        .Parameter date
            A text string containing the date
        .Parameter format
            A text string containing the formatting information
#>
param([string]$Date, [string]$Format) 
    [DateTime]::ParseExact($Date,$Format,[System.Globalization.CultureInfo]::InvariantCulture) 
}


Function Set-Offset {
<#
        .Synopsis
            Sets a global variable $offset from a picture.
        .Description
            Calculates the offset between the time shown on a data logger in a picture,
            and the time taken recorded by the camera
        .Example
            C:\PS> Set-Offset
            Will prompt for the path to the picture and the date and time in it
        .Example
            C:\PS> Set-Offset -referenceImagePath "D:\DCIM\100PENTX\IMG43210.jpg" -refDate "31 January 2010 12:23:34 +1"
            Will calculate the offset from image , given the date in the picture
        .Parameter ReferenceImagePath
            (Alias Path) The path to the reference image
        .Parameter RefDate
            The date in the picture
    #>
[CmdletBinding()]
Param ( [Parameter(ValueFromPipelineByPropertyName=$true)][Alias('FullName','Path')]$ReferenceImagePath ,
        [DateTime]$RefDate
      ) 
        if (-not $refdate) {$RefDate = ([datetime]( Read-Host ("Please enter the Date & time in the reference picture, formatted as" + [char]13 + [Char]10 +
                                                               "Either MM/DD/yyyy HH:MM:SS ±Z or dd MMMM yyyy HH:mm:ss ±Z"))).touniversalTime()
        } 
        if (-not $ReferenceImagePath) {$ReferenceImagePath  = Read-Host "Please enter the path to the picture"}
        if ($ReferenceImagePath -and (test-path $ReferenceImagePath) -and $RefDate) { 
            $picTime = (get-image $ReferenceImagePath | get-Exif -verbose:$false).dateTaken
            $Global:offset  = ($picTime - $refdate).totalSeconds
            write-verbose "OffSet = $Global:offset"
        }
}


Function Get-NearestPoint {
<#
        .Synopsis
            From a set of timestamped data points, returns the one nearest to a given time
        .Description
            From a set of timestamped data points, returns the one nearest to a given time
        .Example
            C:\PS> $point = get-nearestPoint -DataPoints $points -ColumnName "DateTime" -MatchingTime $dt 
            Returns the point in $pointswhere the "DateTime" column is nearest to $dt
        .Parameter DataPoints
            An array containing the data points
        .Parameter ColumnName
            The name of the column in the points array that holds the dateTime to match against
        .Parameter MatchingTime
            The time of the item being sought
    #>
[CmdletBinding()]
Param ( $DataPoints , $ColumnName , $MatchingTime)
        write-verbose "Checking $($dataPoints.count) points for one where $columnName is closest to $matchingtime"
        $variance = [math]::Abs(($dataPoints[0].$columnName - $MatchingTime).totalseconds)
        $i = 1 
        do {
           # write-progress -Activity "looking" -Status "looking" -CurrentOperation $i 
            $v = [math]::Abs(($dataPoints[$i].$columnName - $MatchingTime).totalseconds)
            if ($v -le $variance) {$i ++ ; $variance = $v } 
        } while (($v -eq $variance) -and ($i -lt $datapoints.count)) 
        write-verbose "Point $I matched with variance of $variance seconds"
        $datapoints[($i -1)]
        # write-progress -Activity "looking" -Status "looking" -Completed
}


Function Get-CSVGPSData {
<#
        .Synopsis
            Gets GPS Data from a CSV file
        .Description
            Gets GPS Data from a CSV file
        .Example
            C:\PS> $points = Get-CSVGPSData .\20100420161012.log -offset $offset 
            Reads the GPS data from the Comma seperated log file, 
            applying the offset in $offset - storing the result in $points.
        .Parameter Path
            The path to the file
        .Parameter Offset
            The offset to apply to the logged data
    #>
Param  ([Parameter(Mandatory=$true )][Alias("Filename","FullName")]$Path , $offset) 
        ## Check your date format different GPS devices return different numbers of decimals for the .f (fraction part). 
        $Dateformat = "yyyy-MM-dd HH:mm:ss"
        Import-Csv -Path $path -Header "Date","Lat","Lon","altitude","bearing","MetersPerSec","H_acc","V_acc","blank","Network"|
            select  -property  @{Name="DateTime"; Expression = {(   ConvertTo-DateTime $_.Date $DateFormat ).addSeconds($offset) }}  ,
                               @{Name="MPH";      Expression = {    [system.math]::Round((2.237  * $_.MetersPerSec),1) }} , @{Name="KPH"; Expression = { [system.math]::Round((3.6 * $_.Knots),1) }} ,
                               knots, bearing ,          
                               @{Name="LatDMS";   Expression = {  @([math]::truncate([math]::Abs( [double]$_.lat                               )   )  , 
                                                                    [math]::truncate([math]::Abs( [double]$_.lat      - [math]::truncate( [double]$_.lat) )*60)  , 
                                                                    [math]::round(   [math]::Abs(([double]$_.lat *60) - [math]::truncate(([double]$_.lat *60)) )*60 ,2) )  }}   ,
                               lat , @{Name="NS"; Expression = {if ([double]$_.lat -gt 0) {"N"} Else {"S"}  }} ,      
                               @{Name="LonDMS";   Expression = {  @([math]::truncate([math]::Abs( [double]$_.lon                               )   )  , 
                                                                    [math]::truncate([math]::Abs( [double]$_.lon      - [math]::truncate( [double]$_.lon) )*60)  , 
                                                                    [math]::round(   [math]::Abs(([double]$_.lon *60) - [math]::truncate(([double]$_.lon *60)) )*60 ,2) )  }}  , 
                               lon , @{Name="EW"; Expression = {if ([double]$_.lon -gt 0) {"E"} Else {"W"}  }} ,
                               @{Name="AltM";     Expression = {    [math]::round(  ([double]$_.Altitude       ), 1) }} ,
                               @{Name="AltFT";    Expression = {    [math]::round(  ([double]$_.Altitude * 3.28), 1) }}  |
                Sort-object -property datetime               
}


Function Get-GPXData {
<#
        .Synopsis
            Gets GPS Data from a GPX format XML file
        .Description
            Gets GPS Data from a GPX format XML file
        .Example
            C:\PS> $points = Get-GPXData .\20100420161012.GPX -offset $offset 
            Reads the GPS data from the GPX XML log file, 
            applying the offset in $offset - storing the result in $points.
        .Parameter Path
            The path to the file
        .Parameter Offset
            The offset to apply to the logged data
    #>
Param  ([Parameter(Mandatory=$true)][Alias("FileName","FullName")]$Path , $offset) 
        ([xml](Get-Content $path)).gpx.trk.trkseg.trkpt |
             select -property  @{Name="DateTime"; Expression = { ([dateTime]$_.time).toUniversalTime().addSeconds($offset) }}  ,
                               @{Name="MPH";      Expression = {  [system.math]::Round((0.621  * $_.Speed),1) }} , 
                               @{Name="KPH";      Expression = {  $_.speed }} ,
                               @{Name="knots";    Expression = {  [system.math]::Round(($_.Speed / 1.852 ),1) }},
                               @{Name="Bearing";  Expression = {  $_.course}}  ,          
                               @{Name="LatDMS";   Expression = {  @([math]::truncate([math]::Abs( [double]$_.lat                               )   )  , 
                                                                    [math]::truncate([math]::Abs( [double]$_.lat      - [math]::truncate( [double]$_.lat) )*60)  , 
                                                                    [math]::round(   [math]::Abs(([double]$_.lat *60) - [math]::truncate(([double]$_.lat *60)) )*60 ,2) )  }}   ,
                               lat , @{Name="NS"; Expression = {if ([double]$_.lat -gt 0) {"N"} Else {"S"}  }} ,      
                               @{Name="LonDMS";   Expression = {  @([math]::truncate([math]::Abs( [double]$_.lon                               )   )  , 
                                                                    [math]::truncate([math]::Abs( [double]$_.lon      - [math]::truncate( [double]$_.lon) )*60)  , 
                                                                    [math]::round(   [math]::Abs(([double]$_.lon *60) - [math]::truncate(([double]$_.lon *60)) )*60 ,2) )  }}  , 
                               lon , @{Name="EW"; Expression = {if ([double]$_.lon -gt 0) {"E"} Else {"W"}  }} ,
                               @{Name="AltM";     Expression = {    [math]::round(  ([double]$_.ele       ), 1) }} ,
                               @{Name="AltFT";    Expression = {    [math]::round(  ([double]$_.ele * 3.28), 1) }}  |
                Sort-object -property datetime                
}


Function Get-NMEAData {
<#
        .Synopsis
            Gets GPS Data from a text file of NMEA sentences 
        .Description
            Gets GPS Data from a text file of NMEA sentences
        .Example
            C:\PS> $points = Get-NMEAData .\20100420161012.LOG -offset $offset 
            Reads the GPS data from the NMEA file, 
            applying the offset in $offset - storing the result in $points.
        .Parameter Path
            The path to the file
        .Parameter Offset
            The offset to apply to the logged data
    #>
Param  ([Parameter(Mandatory=$true)][Alias("FileName","FullName")] $Path , $offset, [Switch]$NoAltitude) 
        ## Check your date format different GPS devices return different numbers of decimals for the .f (fraction part). 
        $Dateformat = "ddMMyyHHmmss.f"
        $TimeFormat = "HHmmss.f" 
        if (-not $NoAltitude) {$altPoints = (Import-Csv -path $path -Header "type","time","lat","ns","lon","ew","quality","sattelites","HDofP","Altitude","Units","age","ref" | 
                    where {$_.type -eq '$GPGGA'} )  | 
                       select-object -property "lat","ns","lon","ew","altitude",
                                               @{Name="DateTime"; Expression = {(ConvertTo-DateTime $_.time $TimeFormat).timeofday }} | sort datetime
        }
        Import-Csv -Path $path -Header "Type","Time","status","lat","NS","lon","EW","Knots","bearing","Date","blank","checksum" | 
            where {$_.type -eq '$GPRMC' -and $_.Time } |
                select  -property  @{Name="DateTime"; Expression = { (ConvertTo-DateTime ($_.Date+$_.Time) $DateFormat).addSeconds($offset) }}  ,
                                   knots, 
                                   @{Name="MPH"; Expression = { [system.math]::Round((1.15  * $_.Knots),1) }} ,
                                   @{Name="KPH"; Expression = { [system.math]::Round((1.852 * $_.Knots),1) }} ,
                                   bearing, 
                                   lat, NS, @{Name="LatDMS"; Expression = {@([math]::truncate($_.lat / 100) , 
                                                                             [math]::truncate($_.lat % 100) , 
                                                                             [math]::round((60 * ($_.lat - [math]::truncate($_.lat ))),2))}},
                                   lon, EW ,@{Name="LonDMS"; Expression = {@([math]::truncate($_.lon / 100) , 
                                                                             [math]::truncate($_.lon % 100) , 
                                                                             [math]::round((60 * ($_.lon - [math]::truncate($_.lon ))),2))}},  
                                   @{Name="AltM"; expression={if ($alt) {(Get-NearestPoint -datapoints $altPoints -columnname "DateTime" `
                                                                                -matchingtime (ConvertTo-DateTime $_.time $TimeFormat).timeofday).altitude }}}
}


Function Get-SuuntoData {
<#
        .Synopsis
            Gets Scuba dive data from Suunto CSV export files 
        .Description
            Gets Scuba dive data from Suunto CSV export files 
        .Example
            C:\PS> $points = Get-SuuntoData . -offset $offset -minDive 90 
            Reads the data from the SDM and SDM$PRO files in the current folder, 
            discarding the data for dives below 90, and applying the offset in $offset
        .Parameter Path
            The path to the FOLDER containing the CSV files 
        .Parameter Offset
            The offset to apply to the logged data
        .Parameter MinDive
            The number of the first dive to include in the returned data
         .Parameter SelectMinDive
            If this switch is present the user will be prompted to select the first dive.
    #>
Param  ($Path="." , $offset , $minDive=0, [switch]$SelectMinDive ) 
        $SDM    = join-path -Path $path -ChildPath 'SDM.CSV'
        $SDMpro = join-path -Path $path -ChildPath 'SDM$PRO.CSV'
        if (-not ((Test-path $sdm) -and (test-path $SDMPro)) ) {Write-Host  'You need to specify a path where SDM.CSV and SDM$Pro.csv can be found' ; return}
        $Dives=(import-csv $sdm     -header "UniqueDiveID","DiveNumber","DiveDate","TimeOfDay","Series","DCDiveNumber","DiveTime","SurfaceInterval","MaxDepth",
                                            "MeanDepth","DCType", "DCSerialNumber","DCPersonalData","DCSampleRate","DCAltitudeMode","DCPersonalMode",
                                            "SolutionTimeadjustment","Modified","Location", "Site","Weather","WaterVisibility","AirTemp","WaterTemp",
                                            "WaterTempAtEnd","Partner","DiveMaster","BoatName","CylinderDesc","CylinderSize","CylinderUnitsCode",
                                            "CylinderWorkPressure","CylinderStartPreessure","CylinderEndPressure","SACRate","SACUnits","UserField1",
                                            "UserField2", "UserField3","UserField4","UserField5","Weight","OxygenPercent","OLFPercent","OTUFlag"  |
                    select-object -property  UniqueDiveID , DCSampleRate ,
                                             @{Name="DateTime"; expression={(ConvertTo-DateTime ($_.DiveDate + $_.TimeOfDay) "dd/MM/yyyyHH:mm").addSeconds($offset)}} , 
                                             @{Name="lat"     ; expression={$_.UserField4 -split ","}},
                                             @{Name="Lon"     ; expression={$_.UserField5 -split ","}},
                                             @{Name="Description"; expression={($_.Site +", " + $_.Location + ": "+ $_.WaterTemp + "°C") -replace "^,\s",""}} )
        if ( $selectMinDive ) { $minDive = (Select-List -Property Description -InputObject $dives).uniqueDiveId} 
        write-verbose "Looking at dives with ID greater than or equal to $minDive" 
        import-csv $sdmPro          -header "UniqueDiveID","SegmentNumber","SegmentDepth","ASCFlag","SLOWFlag", "CeilingFlag","SURFFlag","AttentionFlag","UserFlag","SafeStopFlag" | 
            where-object {[int]$_.uniqueDiveID -ge $minDive} | 
                select-object     -property @{name="DateTime";    expression={$global:DiveID=$_.UniqueDiveID
                                                                              if ($global:dive.uniqueDiveID -ne $global:DiveID) {$global:dive = ($dives | where-object {$_.UniqueDiveID -eq $global:DiveID})
                                                                                                                                 Write-verbose "Processing dive $($global:dive.description)"
                                                                              }
                                                                              $global:dive.DateTime.AddSeconds([int]$_.SegmentNumber * [int]$global:dive.DCSampleRate )}},
                                            @{name="Description"; expression={$_.SegmentDepth +"M - " +$global:Dive.description}}, SegmentDepth ,
                                            @{name="Lat"        ; expression={                         $global:Dive.Lat}},
                                            @{name="Lon"        ; expression={                         $global:Dive.Lon}}                                             
}


Function Convert-GPStoEXIFFilter {
<#
        .Synopsis
            Builds a collection of EXIF WIA filters to set GPS data from a point 
        .Description
            Builds a collection of EXIF WIA filters to set GPS data from a point 
        .Example
            C:\PS> $filter = Convert-GPStoEXIFFilter 51,36,7 "N" 1,33,54 "W"
            Creates a new filter chain and adds the Exif Filters to add the GPS data to images
        .Parameter LATDMS
           The lattitude as an array of 3 numbers for Degrees, Minutes and Seconds
        .Parameter NS
           N for lattitude north of the equator, S for lattitude South of the equator   
        .Parameter LONDMS
           The longitude as an array of 3 numbers for Degrees, Minutes and Seconds
        .Parameter EW
           E for longitude East of Greenwich W for longitude West of GreenWich
        .Parameter AltM
           Altitude in Meters above mean Sea level    
    #>
param ( [Parameter(Mandatory=$true)]$LatDMS,
        [Parameter(Mandatory=$true)]$NS,
        [Parameter(Mandatory=$true)]$LONDMS,
        [Parameter(Mandatory=$true)]$EW,
        $AltM          
)    
process {
        $filter = new-Imagefilter  
        if (-not $filter.Apply)  { return }
        $ExifVervalue = New-Object -ComObject "WIA.Vector"
        $ExifVervalue.Add([byte]2)
        $ExifVervalue.Add([byte]2)
        $ExifVervalue.Add([byte]0)
        $ExifVervalue.Add([byte]0)

        $LongDMSValue = New-Object -ComObject "WIA.Vector"
        $LonDMS | Foreach {$v = New-Object -ComObject wia.rational ; $v.numerator = [int32]($_ * 1000000) ; $v.denominator = 1000000 ; $longDmsValue.add($v) }
        $LatDMSValue = New-Object -ComObject "WIA.Vector"
        $LatDMS | Foreach {$v = New-Object -ComObject wia.rational ; $v.numerator = [int32]($_ * 1000000) ; $v.denominator = 1000000 ; $latDmsValue.add($v) }

        Add-exifFilter -filter $filter -ExifID "$ExifidGPSVer"       -typeid "$ExifVectorOfBytes"     -value $ExifVervalue
        Add-exifFilter -filter $filter -ExifID "$ExifIDGPSLatRef"    -typeid "$ExifString"            -value $NS
        Add-exifFilter -filter $filter -ExifID "$ExifIDGPSLongRef"   -typeid "$ExifString"            -value $EW
        Add-exifFilter -filter $filter -ExifID "$ExifIDGPSLongitude" -typeid "$ExifVectorOfRationals" -value $longDmsValue
        Add-exifFilter -filter $filter -ExifID "$ExifIDGPSLattitude" -typeid "$ExifVectorOfRationals" -value $LatDMSValue
        if ($AltM)  {
            if ($altM -ge 0) { Add-exifFilter -filter $filter -ExifID "$ExifIDGPSAltRef" -typeid "$ExifByte" -value ([byte]0) }
            else             { Add-exifFilter -filter $filter -ExifID "$ExifIDGPSAltRef" -typeid "$ExifByte " -value ([byte]1) }
            Add-exifFilter -filter $filter -ExifID "$ExifIDGPSAltitude" -typeid "$ExifUnsignedRational" -Numerator ([uint32]([math]::Abs($AltM)  * 100) ) -denominator 100
           }
           return $filter 
    }
}


Function Copy-GPSImage {
<#
        .Synopsis
            Copies an image, applying EXIF data from GPS data points   
        .Description
            Copies an image, applying EXIF data from GPS data points   
        .Example
            C:\PS>  Dir E:\dcim –inc IMG*.jpg –rec | Copy-GpsImage -Points $Points -Keywords "Oxfordshire" -rotate -DestPath "$env:userprofile\pictures\oxford" -replace  "IMG","OX-"
            Copies IMG files from folders under E:\DCIM to the user's picture\Oxford folder, replacing IMG in the file name with OX-.
            The Keywords field is set to Oxfordshire, pictures are GeoTagged with the data in $points and rotated. 
        .Parameter Image
            A WIA image object, a path to an image, or a file object representing an image file. It may be passed via the pipeline.
        .Parameter Points
            An array of GPS data points 
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
           If this switch  is specified the GPS point associated with each image is returned. 
           Note that the point in the points collection is updated as a side effect, so this can be combined with Out-Null
           The points returned - or isolated from the collection with a a where command can be used 
           to plot picture locations on a map 
    #>
[CmdletBinding(SupportsShouldProcess=$true)]
Param ( [Parameter(ValueFromPipeline=$true, Mandatory=$true)][Alias("Path","FullName")]$image , 
        [Parameter(Mandatory=$true)]$points, 
        [Parameter(Mandatory=$true)]$Destination ,  
        $keywords , $Title, $replace,[switch]$Rotate,[switch]$NoClobber, [switch]$ReturnInfo, $psc 
)
process {
        if ($psc -eq $null)  {$psc = $pscmdlet} ; if (-not $PSBoundParameters.psc) {$PSBoundParameters.add("psc",$psc)}
        if ($image -is [system.io.fileinfo] ) {$image = $image.FullName }
        if ($image -is [String]             ) {$image = Get-Image $image}
        if ($Image.count -gt 1              ) {[Void]$PSBoundParameters.Remove("Image") 
                                               $Image | ForEach-object {Copy-GPSImage -image $_ @PSBoundParameters}
                                               return
        }
        if ($image -is [__comObject])  {
            $dt = get-ExifItem -image $image -ExifID $ExifIDDateTimeTaken
            if ($dt) {$dt = [DateTime]::ParseExact($dt,"yyyy:MM:dd HH:mm:ss",[System.Globalization.CultureInfo]::InvariantCulture) 
                      # write-verbose $dt
                      $point  = get-nearestPoint  -DataPoints $points      -ColumnName "DateTime" -MatchingTime $dt 
                      $filter = Convert-GPStoEXIFFilter -LatDMS $point.Latdms -NS $point.NS -LONDMS $point.londms -EW $point.ew -AltM $point.altM 
                      [Void]$PSBoundParameters.Remove("Points") 
                      $path = Copy-Image -filter $filter @PSBoundParameters
                      If ($ReturnInfo) {if ($point.paths ) {if ($point.paths -notcontains $path ) {$point.paths +=  $path ; $point}}
                                        else {$point | Add-Member -MemberType Noteproperty -Name "paths" -Value @($path) -PassThru -force}
                      }        
            }
        }
    }
}


Function Convert-SuuntotoExifFilter {
<#
        .Synopsis
            Adds a collection of EXIF WIA filters to set Scuba diving data from a point 
        .Description
            Adds a collection of EXIF WIA filters to set Scuba diving data from a point 
        .Example
            C:\PS> $filter Convert-SuuntotoExifFilter $Point
            Creates a new filter chain and adds the ExifFileters to add the data in $point to images
        .Parameter Point
            A GPS data point 
        .Parameter title
            Pre-existing title
    #>
param( [Parameter(Mandatory=$true)]$Point,
       [String]$title       
)    
process {
        IF (($Point.LAT.Count -eq 4 ) -and ($Point.Lon.Count -eq 4 )) {
               $filter = Convert-GPStoEXIFFilter -LatDMS ([single[]]($point.lat[0..2])) -NS ($point.lat[3] -replace "\s","") -LONDMS ([single[]]($point.lon[0..2])) -EW ($point.lon[3] -replace "\s","")
        }
        Else { $filter = New-Object -ComObject Wia.ImageProcess  }
        if (-not $filter.Apply)  { return }
        if ($title)              { $title = $title +": " + $Point.Description }
        else                     { $title =                $Point.Description }
        
        
      # Add-exifFilter -filter $filter -ExifID $ExifIDImageDescription        -typeid $ExifString -string $title 
       # $filter.Filters.Item($filter.Filters.Count).Properties.Item("Remove")    = $true
        Add-exifFilter -filter $filter -ExifID $ExifIDTitle        -typeid $ExifVectorOfBytes -string $title 
        Add-exifFilter -filter $filter -ExifID $ExifIDGPSAltRef    -typeid $ExifByte          -value ([byte]1)
                 
        $altitudeValue             =  New-Object -ComObject wia.rational 
        $altitudeValue.numerator   = [uint32]([double]$point.SegmentDepth  * 100) 
        $altitudeValue.denominator = 100 
        $filter.Filters.Add( $filter.FilterInfos.Item("Exif").FilterId)     
        $filter.Filters.Item($filter.Filters.Count).Properties.Item("ID")    = "$ExifIDGPSAltitude"
        $filter.Filters.Item($filter.Filters.Count).Properties.Item("Type")  = 1007 # URational
        $filter.Filters.Item($filter.Filters.Count).Properties.Item("Value") = $altitudeValue        
        return $filter 
    }
}


Function Copy-SuutoImage {
<#
        .Synopsis
            Copies an image, applying EXIF data from Scuba diving log data points   
        .Description
            Copies an image, applying EXIF data from Scuba diving log data points   
        .Example
            C:\PS>  Dir E:\dcim –inc IMG*.jpg –rec | Copy-SuuntoImage -Points $Points -Keywords "Ocean; Bahamas"  -DestPath "$env:userprofile\pictures\Diving" -replace  "IMG_","DIVE"
            Copies IMG files from folders under E:\DCIM to the user's picture\Diving folder, replacing IMG in the file name with DIVE.
            The Keywords field is set to Ocean; and pictures are Tagged with the data in $points 
        .Parameter Image
            A WIA image object, a path to an image, or a file object representing an image file. It may be passed via the pipeline.
        .Parameter Points
            An array of GPS data points 
        .Parameter Destination
            The FOLDER to which the file should be saved.
        .Parameter Keywords
            If specified, sets the keywords Exif field.
        .Parameter Title
            If specified, sets the Title Exif field..    
        .Parameter Replace
            If specified, this contains two values seperated by a comma specifying a replacement in the file name
        .Parameter NoClobbber
            Unless this switch is specified, a pre-existing image WILL be over-written
    #>
[CmdletBinding(SupportsShouldProcess=$true)]
Param ( [Parameter(ValueFromPipeline=$true, Mandatory=$true)][Alias("Path","FullName")]$image , 
        [Parameter(Mandatory=$true)]$points, 
        [Parameter(Mandatory=$true)]$Destination ,  
        $keywords , $replace,[switch]$NoClobber, $psc )
process {
        if ($psc -eq $null)  {$psc = $pscmdlet} ; if (-not $PSBoundParameters.psc) {$PSBoundParameters.add("psc",$psc)}
        if ($image -is [system.io.fileinfo] ) {$image = $image.FullName }
        if ($image -is [String]             ) {$image =(Resolve-Path $image -errorAction "SilentlyContinue") | ForEach-Object {$_.path }}
        if ($Image.count -gt 1              ) {[Void]$PSBoundParameters.Remove("Image") 
                                               $Image | ForEach-object {Copy-GPSImage -image $_ @PSBoundParameters}
                                               return
        }
        if ($image -is [String]             ) {$image = Get-Image $image}
        if ($image -is [__comObject])  {
            $title  = $null # ( get-ExifItem -image $image -ExifID $ExifIDTitle         )
            $dt     = ( get-ExifItem -image $image -ExifID $ExifIDDateTimeTaken )
            if ($dt) {$dt = [DateTime]::ParseExact($dt,"yyyy:MM:dd HH:mm:ss",[System.Globalization.CultureInfo]::InvariantCulture) 
                      write-verbose $dt
                      $point  = get-nearestPoint -DataPoints $points -ColumnName "DateTime" -MatchingTime $dt
                      $filter = Convert-SuuntotoExifFilter -Point  $point  -title  $title              
                      [Void]$PSBoundParameters.Remove("Points") 
                      Copy-Image -filter $filter @PSBoundParameters
#                      if ($keywords)  {Add-exifFilter -filter $filter -ExifID $ExifIDKeywords -typeid 1101 -string $keywords }
 #                     if ($replace)   {$SavePath= join-path -Path $Destination -ChildPath ((Split-Path $image.FullName -Leaf) -Replace $replace)}
  #                    else            {$SavePath= join-path -Path $Destination -ChildPath  (Split-Path $image.FullName -Leaf)  }
   #                   Set-ImageFilter -image $image -passThru -filter $filter | Save-image -fileName $savePath -NoClobber:$NoClobber -psc $psc
                      $point = $dt = $image =  $filter = $null
           }
       }
   }
}


Function Merge-GPSPoints {
<#
        .Synopsis
            Merges a set of GPS points, producing 1 point per minute (or longer)
        .Description
            Averages the points logged each minute, or longer interval specifed by -interval
        .Example
            C:\PS>  merge-gpsPoints -points $points
            Returns the points in $points combined to 1 average point per minute
        .Parameter Points
            An array of GPS data points 
        .Parameter Interval
            The interval in minutes over which points should be averaged (default 1)
    #>
Param ([Parameter(ValueFromPipeline=$true, Mandatory=$true)]$points, $interval=1)
begin   {$PointsToMerge = @()     }
Process {$PointsToMerge += $points}
End     {
         $pointsToMerge | Select-Object -Property knots, @{Name="Minute"       ; Expression={$_.dateTime.date.addhours($_.datetime.hour).addMinutes($interval*[math]::Truncate($_.dateTime.minute/$interval)).tostring("dd MMMM yyyy HH:mm")  }} ,  
                                                         @{name="formattedLon" ; expression={$lon = $_.londms[0] + ($_.londms[1] /60) + ($_.londms[2] /3600)  ;  if ($_.EW -eq "W") {$lon *= -1};  $lon}} ,
                                                         @{name="formattedLat" ; expression={$lat = $_.latdms[0] + ($_.latdms[1] /60) + ($_.latdms[2] /3600)  ;  if ($_.NS -eq "S") {$Lat *= -1};  $lat}} | 
                               Group-Object -Property minute  | 
                                 Select-Object -Property @{name="DateTime"   ;expression={[DateTime]$_.name}},@{name="aveKnots"; expression={( $_.GROUP | MEASURE-OBJECT -Property knots -Average).AVERAGE}},
                                                         @{name="lat"        ;expression={( $_.GROUP | MEASURE-OBJECT -Property formattedLat -Average).AVERAGE}},
                                                         @{name="lon"        ;expression={( $_.GROUP | MEASURE-OBJECT -Property formattedLon -Average).AVERAGE}}
        }
}

function get-gpsDistance {
param ($point1 ,$point2 , $Units="KM")

$conv =[system.math]::pi  /180
$lat1 = $conv * $point1.lat
$lon1 = $conv * $point1.lon

$lat2 = $conv * $point2.lat
$lon2 = $conv * $point2.lon

$distanceInKm = [Math]::Acos([math]::Sin($lat1)*[math]::Sin($lat2) +  [math]::Cos($lat1)*[math]::Cos($lat2)*[math]::Cos($Lon2 - $lon1) ) *6371

switch ($units) {
     "KM"     {$distanceInKm}
     "Miles"  {$distanceInKm * 0.6213 }
     "NM"     {$distanceInKm * 0.54   }
     "Meters" {$distanceInKm * 1000   }
}
}

function get-gpsBearing {
param ($point1 ,$point2 )

$conv =[system.math]::pi  /180
$lat1 = $conv * $point1.lat
$lon1 = $conv * $point1.lon

$lat2 = $conv * $point2.lat
$lon2 = $conv * $point2.lon

#   tc1=mod(atan2(sin(lon2-lon1)*cos(lat2), cos(lat1)*sin(lat2)-sin(lat1)*cos(lat2)*cos(lon2-lon1)),2*pi)

[math]::Round(((([math]::Atan2( ([math]::Sin($lon2- $lon1) * [math]::Cos($lat2))   ,
                                [math]::Cos($Lat1)        * [math]::Sin($Lat2) - [math]::Sin($Lat1)*[math]::Cos($lat2)*[math]::Cos($Lon2 - $lon1)   ) + 2*[math]::pi ) % (2*[math]::pi )) / $conv),0)
  
}


function Select-SeperateGPSPoints {
param ([Parameter(ValueFromPipeline=$true, Mandatory=$true)]$points , $KM=0.04)
begin   {$m= @() }
Process {$m+= $points}
End     {0..($m.Count-1) | foreach -Begin {$previous=$m[0]}`
                                -Process {$d=(get-gpsDistance  $m[$_] $previous -Units "KM") 
                                          if  ( $d -gt $KM) {Add-Member -force -InputObject $Previous -MemberType Noteproperty -Name "aveKnots" -Value ($d/(($m[$_].dateTime - $previous.DateTime ).totalhours) )
                                                             Add-Member -force -InputObject $previous -MemberType Noteproperty -Name "Bearing" -Value ([single](get-gpsBearing  $Previous $m[$_] ))
                                                              $previous
                                                              $Previous = $m[$_] }
                                         }`
                                -end  {$m[-1]  }
}}


Function ConvertTo-GPX{
<#
        .Synopsis
            Converts a set of GPS points to a GPX file to be imported by other programs
        .Description
            Converts a set of GPS points to a GPX file to be imported by other programs
        .Example
            C:\PS> merge-gpsPoints -points $points | convertto-GPX | Set-Content 2010-04-04.gpx -Encoding utf8
            Takes the result of merging the Points in $points and writes it as a GPX file.
            Note that the file will not read properly if output as unicode so -encoding UTF8 is required
        .Example
            C:\PS> $points | where {$_.paths} ) | convertto-GPX -name {Split-Path $_.paths[0] -Leaf} | out-file -Encoding utf8 -FilePath temp.gpx
            Takes the points which have a path set after using Copy-GPS with the -ReturnInfo switch, 
            and outputs a file with the short file name of the the picture as the label for the data point.
            Note that the file will not read properly if output as unicode so -encoding UTF8 is required.
        .Parameter Points
            An array of GPS data points - may be passed via the the pipeline
        .Parameter name
            A label for the data point written as a code block. The default is the dateTime specified as {$_.dateTime}
    #>
Param   ( [Parameter(ValueFromPipeline=$true,Mandatory=$true)]$points,[ScriptBlock]$name={$_.DateTime} )
Begin   { '<?xml version="1.0" encoding="UTF-8"?><gpx xmlns="http://www.topografix.com/GPX/1/1" ' +  
                           ' creator="PowerShell" version="1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '+ 
                           ' xsi:schemaLocation="http://www.topografix.com/GPX/1/1 http://www.topografix.com/GPX/1/1/gpx.xsd"> '+
                           ' <metadata><name>PowerShellExport</name> </metadata>' 
}
process { $points | foreach {'<wpt lat="{0}" lon="{1}"><name>{2}</name> </wpt>' -f $_.lat,$_.lon,( Invoke-Expression $(".{$name}") )} }
End     { '</gpx>'} 
}
                           

Function Out-MapPoint {
<#
        .Synopsis
            Creates a set of PushPins in Map point from a set of GPS points
        .Description
            Creates a set of PushPins in Map point from a set of GPS points
        .Example
            C:\ps> merge-gpsPoints $points | Out-MapPoint -name  {"{0} - {1:00} MPH " -f $_.DateTime,$_.Aveknots*1.1508 } -symbol {if ($_.aveknots -lt 10) {5}  elseif ($_.aveknots -gt 40) {7} else {6}}
            Takes the result of merging the Points in $points and creates a Mappoint map
            The label is based on the the date and speed converted to MPH for example "14:00 - 30 MPH", and the symbol is colour coded based on speed
        .Example    
            C:\ps> (merge-gpsPoints $points | select lat,lon,paths,datetime  ) +
                    ($points | where {$_.paths} | select lat,lon,paths,datetime) | Out-MapPoint -symbol {if ($_.paths) {79} else {1}}
            Merges the GPS points, and combines them with those which have a path to a picture 
            set by using copy-gps with the -return info switch. 
            The "walking" points are give a red push pin and the photo sites are given a camera symbol.
            The label defaults to the date and time.
        .Example
             C:\ps> Get-NMEAData F:\copilot\gpstracks\Jul2110.gps | Merge-GPSPoints  | Select-SeperateGPSPoints -KM 0.03  
             | Out-MapPoint -linkPoints -symbol {Switch ([math]::truncate((([single]$_.bearing)+22.5)/45)) { 0 {128} ; 1 {136} ; 2 {131} ; 3 {138} ; 
                 4 {129} ; 5 {139} ; 6{130} ; 7 {137 }; 8 {128} }}  -linecol {if     ($_.dateTime -lt [datetime]"07/21/2010 12:36:00") {16711680} 
                                                                              elseif ($_.dateTime -lt [datetime]"07/21/2010 13:06:00") {65280} else {255} }
             reads , merges and selects distinct GPS points. then sets pins on the map based on the bearing and draws lines coloured for different stages of the journey
        .Parameter Points
            An array of GPS data points - may be passed via the the pipeline
        .Parameter Name
            The Pushpin Label, written as a code block. The default is the dateTime specified as {$_.dateTime}
        .Parameter Symbol
            The Pushpin type ID, written as a code block - a pin type maybe specified as  {79}    
        .Parameter LinkPoints
            If Specified, lines will be drawn between the points
        .Parameter LineCol
            line colour as a written as a code block {1} = black, 255=Red, 65280=Green, 16711680=Blue, 65535=Yellow, 16711935=Magenta, 16776960=cyan
    #>
Param ([Parameter(valueFromPipeLine=$true)]$points , [ScriptBlock]$name={$_.DateTime} , [ScriptBlock]$symbol, [switch]$linkPoints, [ScriptBlock]$LineCol={1})
Begin {
    if (-not $Global:mpapp) {$Global:MPApp = New-Object -ComObject "Mappoint.Application"}
    $Global:MPApp.Visible = $true
    $map = $Global:mpapp.ActiveMap
    $global:prevLoc = $null
    }
Process {  $points | foreach-object { $location=$map.GetLocation($_.Lat, $_.lon)
               $nameText = Invoke-Expression $(".{$name}") 
               if ($symbol -is [scriptblock]) {$symbolID = Invoke-Expression $(".{$Symbol}") }
               $Pin = $map.AddPushpin($location, $nameText)
               if ($symbol) {$pin.symbol = $symbolID}
               if ($linkPoints) {
                    if ($Global:prevLoc) { $line = $map.shapes.addLine($global:prevloc,$location)
                                           $line.zorder(5)
                                           $line.line.forecolor = Invoke-Expression $(".{$LineCol}") 
                                           $line.line.EndArrowhead = $true 
                                           }
                    $global:prevLoc = $location
               }
           }
        }
}


Function Resolve-ImagePlace {
<#
        .Synopsis
            Queries the GeoNames Web service to translate EXIF Lat/Long information to a place name
        .Description
            Queries the GeoNames Web service to translate EXIF Lat/Long information to a place name
        .Example
            C:\ps> resolve-ImagePlace ".\IMG_1234.jpg" 
            Returns the place information for the image
        .Parameter Image
            The image object or file to test - may be passed via the pipeline
    #>
Param ( [Parameter(ValueFromPipeline=$true, Mandatory=$true)][Alias("Path","FullName")]$image )
       if ($image -is [system.io.fileinfo] ) {$image = $image.FullName }
       if ($image -is [String]             ) {$image =(Resolve-Path $image -errorAction "SilentlyContinue") | ForEach-Object {$_.path }}
       if ($Image.count -gt 1              ) {$Image | ForEach-object {resolve-ImagePlace -image }
                                              return
       }
       if ($image -is [String]             ) {$image = Get-Image $image}
       if ($image -is [__comObject]        ) {
            $l=(get-ExifItem -image $image -ExifID $ExifIDGPSLattitude)
            if ($l.count -eq 3) {$lat = $($l[0].value) + ($l[1].value /60 ) + ($l[2].value / 3600)}
            if ($l.count -eq 2) {$lat = $($l[0].value) + ($l[1].value)/60 }
            if ((get-ExifItem -image $image -ExifID $ExifIDGPSLatRef) -eq "S") {$lat= $lat * -1}
            $lat = [math]::Round($lat,5)                
            $l=(get-ExifItem -image $image -ExifID $ExifIDGPSLongitude)
            if ($l.count -eq 3) {$lon =  ($l[0].value) + ($l[1].value / 60) + ($l[2].value/3600) }
            if ($l.count -eq 2) {$lon =  ($l[0].value) + ($l[1].value / 60) }
            if ((get-ExifItem -image $image -ExifID $ExifIDGPSLongRef) -eq "W") {$lon= $lon * -1} 
            $lon = [math]::Round($lon,5)
            write-verbose ("Lat {0}, Lon {1}" -f $lat,$lon) 
            $url = "http://ws.geonames.org/extendedFindNearby?lat=$lat&lng=$lon"
            write-debug $url
            if ($Global:WebClient -eq $null) {$Global:WebClient=new-object System.Net.WebClient  }
            $x = (([xml]($Global:WebClient.DownloadString($url)))  )
            $x.geonames.geoname[-1] | % {write-verbose ("Lat {0}, Lon {1}, Name {2}, ID {3}" -f $_.lat,$_.lng,$_.name,$_.geoNameID) }
            $x.geonames.geoname | % -begin {$n=""} -Process {$n = $_.name +", " +$n} -end {$n -replace '\,\s$',""}
       }
}
