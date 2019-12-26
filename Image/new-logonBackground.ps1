Function New-LogonBackground{
<#
       .SYNOPSIS
            Changes the logon wallpaper, using desktop search to find a suitable image.
       .PARAMETER Keyword
            Filters the pictures to those tagged with a specific keyword.
       .PARAMETER Path
            Allows the location to be searched for pictures to be controlled. 
            An empty string will cause the whole index to be searched
            The default is the logged on users 'My Pictures' folder. 
            Subdirectories are searched. 
       .PARAMETER JPEGQuality
            The Saved file must be less than 256KB, this parameter allows the quality
            to be altered reduce the size if required. The Default is 75
       .PARAMETER Width
            If not specified width will be obtained by checking the properties of the screen
            specifying width allows this to be over-writter
       .PARAMETER Height
            If not specified width will be obtained by checking the properties of the screen
            specifying Height allows this to be over-written
            
        .EXAMPLE
           New-LogonBackground -Keyword "Portfolio" -path ""
           Sets a new logon background from a picture stored anywhere on 
             
#>
[CmdletBinding()]

Param (
    [String]$Keyword    = "Portfolio" ,
    [String]$Path       = [system.environment]::GetFolderPath( [system.environment+specialFolder]::MyPictures ) ,
       [Int]$JPEGQuality= 70,
       [Int]$width      = 0,
       [Int]$height     = 0,
    [Switch]$SetRegistry     
)
    if ($SetRegistry ) {
    #If called with -SetRegistry :
    # 1. Anticipate failure if not run as admin
    # 2. Try to set the flag. 
    # 3. Quit without setting the background 
       try{
         Set-ItemProperty -PATH "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\Background" -Name oembackground -Value 1 -ErrorAction Stop
       }
       catch [System.Security.SecurityException]{
            Write-Warning "Permission Denied - you  need to run as administrator"
       }
       return  
    }
    #Check we can actually change the logon background. 
    if ( (Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\Background).oembackground -ne 1) {
                  Write-Warning "Registry Key OEMBackground under HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI\Background needs to be set to 1"
                  Write-Warning "Run AS ADMINISTRATOR with -SetRegistry to set the key and try again."
    }
    Set-content -ErrorAction "Silentlycontinue" -Path "$env:windir\System32\oobe\Info\Backgrounds\testFile.txt" -Value "This file was to create test for write access. It is safe to remove"
    if (-not $?) {write-warning "Can't create files in  $env:windir\System32\oobe\Info\Backgrounds please set permissions and try again"; return}
    else         {Remove-Item -Path "$env:windir\System32\oobe\Info\Backgrounds\testFile.txt"} 

    if (-not($width -and $height)) {
        $mymonitor      = Get-WmiObject Win32_DesktopMonitor -Filter "availability = '3'" | select -First 1
        $width, $height = $mymonitor.ScreenWidth, $mymonitor.ScreenHeight
        if  ($width -eq 1366)                          {$width         = 1360}
        if (($width -eq 1920) -and ($height -eq 1080)) {$width,$height = 1360,768}
    } 
    if (@("768x1280" ,"900x1440"  ,"960x1280" ,"1024x1280" ,"1280x1024" ,"1024x768" , "1280x960" ,"1600x1200",
         "1440x900" ,"1920x1200" ,"1280x768" ,"1360x768") -notcontains "$($width)x$($height)" ) {
         write-warning "Screen resolution is not one of the defaults. You may need to specify width and height"
    }
       
    $MonitorAspect      = $Width / $height
    $SaveName           = "$env:windir\System32\oobe\Info\Backgrounds\Background$($width)x$height.jpg"
            
    # Run the search ditch any results with the wrong orientation then select a random one and open it 
    write-verbose "Seaching $path,for Pictures at least $width Pixels by $height tagged $Keyword "
    $myimage            = Get-IndexedItem -path $path -recurse -Filter "Kind=Picture","keywords='$keyword'",
                                           "store=File","width >= $width ","height >= $height " | 
                               where-object {($_.width -gt $_.height) -eq ($width -gt $height)} | get-random | get-image
    If ($?) {
        # to make event logging work, use an elevated instance of Powershell to run New-EventLog -Source PSLogonBackground -LogName application 
        write-eventlog -logname Application -source PSLogonBackground -eventID 31365 -message "Loaded $($myImage.FullName) [ $($myImage.Width) x $($myImage.Height) ]" -ErrorAction silentlycontinue
        write-Verbose "Loaded $($myImage.FullName) [ $($myImage.Width) x $($myImage.Height) ]" 
 
        # Now create a chain of image filters to make the image the right shape and size and suitably compressed JPG
        $myfilter       = New-Imagefilter     
        #We might be lucky: the image might be in the right proportions. If not, crop it to shape. 
        $imageAspect    = $myImage.Width / $myimage.Height
        if ($imageAspect -gt $MonitorAspect)  #image is too wide - crop left and right 
            {$margin    = [int](($myimage.width - ($myImage.height * $MonitorAspect)) /2)
            write-verbose "Cropping $margin pixels from left and right"
            Add-CropFilter -filter $myfilter -left $margin -right $margin -top 0 -bottom 0
        }
        if ($imageAspect -lt $MonitorAspect)  #image is too tall - crop top and bottom. 
            {$margin    = [int](($myimage.height - ($myImage.width / $MonitorAspect)) /2)
            write-verbose "Cropping $margin pixels from top and bottom"
            Add-CropFilter -filter $myfilter -left 0 -right 0 -top $margin -bottom $margin 
        }
        #It might be the right size too. If not resize it. The crop might be imperfect, so allow a small change in aspect ratio
        if (($myimage.Height -ne $height) -or ($myimage.Width -ne $width)) 
            {write-verbose "Scaling image to $width x $height" 
             Add-ScaleFilter -filter $myfilter -width    $width   -height  $height -DoNotPreserveAspectRatio
        }
        if ($myfilter.filters.count )  {$myImage = Set-ImageFilter -image $myimage -filter $myfilter}
        #Finally set JPG format with a suitable degree of compression, apply the filters and save. 
        write-verbose "Image wil be saved as a JPG file with quality of $JPEGQuality / 100" 
        
        Set-ImageFilter -filter (Add-ConversionFilter -typeName "JPG" -quality $JPEGQuality -pass) -image $myimage -Save $saveName   

        # Did it work ?
        $item = get-item $saveName 
        while ($item.length -ge 250kb  -and ($JPEGQuality -ge 15) ) {
              $JPEGQuality -= 5
              Write-warning "File too big - Setting Quality to $Jpegquality and trying again"
              Set-ImageFilter -filter (Add-ConversionFilter -typeName "JPG" -quality $JPEGQuality -pass) -image $myimage -Save $saveName   
              $item = get-item $saveName 
        }      
        if ($item.length -le 250KB) {write-verbose "Successfully made the new background." ; $item 
                                     write-eventlog -logname Application -source PSLogonBackground -eventID 31366 -message "Saved $($Item.FullName) : size $($Item.length)"  -ErrorAction silentlycontinue }
    }
}