Function Get-BingPhotos
{
 [CmdletBinding()]
Param()
 if (Get-Command -Name "New-Imagefilter" -ErrorAction SilentlyContinue) {$filter = new-Imagefilter |  Add-exifFilter -passThru -ExifID $ExifIDRating -typeid  1003 -value 5 }

if ($WebClient -eq $null) {$Global:WebClient=new-object System.Net.WebClient  }
 $url = "http://themeserver.microsoft.com/default.aspx?p=Bing&c=Desktop&m=en-US"
 ([xml]($webClient.DownloadString($url).replace("item","RssItem"))).rss.channel.RssItem | 
    Foreach{ $p = join-path "C:\Users\Public\Pictures\Sample Pictures" ($_.enclosure.url -split "/")[-1] 
             write-verbose $_.enclosure.url 
             if (-not (test-path $p)) {$WebClient.DownloadFile($_.enclosure.url,$p) 
                                       if ($filter)  { $image = Get-Image $p 
                                                       Set-ImageFilter -image $image -PassThru -filter $filter | Save-image 
                                                       Get-ExifItem    -image $image -ExifID $ExifIDTitle    
                                                     } 
                                       else  { get-item $p }
                                       }
           }   

 
}