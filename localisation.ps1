cls
Import-Module -Name D:\powershell\Image


$a = Get-Exif -image $args | select GPS
if ($a.GPS)
{
    Write-Host $a.GPS
    
    $tab_a = $a.GPS.Split( "," )
    $a = $tab_a[0]
    $tab_b = $a.Split( " " )
    $var_n = $tab_b[0]
    $var_e = $tab_b[2]

        
    try {
    $chrome = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
    $url = "https://www.google.com/maps/place/$var_n+$var_e"
    Start-Process "$chrome" $url
    }
    catch {
        Start-Process $url
    }

}
