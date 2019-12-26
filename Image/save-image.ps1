function Save-image {
   <#
        .Synopsis
            Saves a Windows Image Acquisition image
        .Description
            Saves a Windows Image Acquisition image
        .Example   
            C:\ps> $image | Save-image -NoClobber -fileName {$_.FullName -replace ".jpg$","-small.jpg$"}
            Saves the JPG image(s) in $image, in the same folder as the source JPG(s), appending
            -small to the file name(s), so that "MyImage.JPG" becomes "MyImage-Small.JPG"
            Existing images will not be overwritten.
        .Parameter image
            The image or images the filter will be applied to; images may be passed via the pipeline.
            If multiple images are passed either no filename must be included 
            (so the image will be saved under its original name), or the fileName must be a code block,
            otherwise the images will all be written over the same file. 
        .Parameter passThru
            If set, the image or images will be emitted onto the pipeline       
        .Parameter filename
            If not set the existing file will be overwritten. The filename may be a string, 
            or be a script block - as in the example
        .Parameter NoClobber
            specifies the target file should not be over written if it already exists
    #>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
      [Parameter(ValueFromPipeline=$true,Mandatory=$true)]
      $image,
      [parameter(ValueFromPipelineByPropertyName=$true)][Alias("Path","FullName")][ValidateNotNullOrEmpty()]
      $fileName ,
      [switch]$passThru,
      [switch]$NoClobber, $psc )

process {
      if ( $psc -eq $null )            { $psc = $pscmdlet }   ; if (-not $PSBoundParameters.psc) {$PSBoundParameters.add("psc",$psc)}
      if ( $image.count -gt 1       )  { [Void]$PSBoundParameters.Remove("Image") ;  $image | ForEach-object {Save-Image -Image $_ @PSBoundParameters }  ; return}
      if ($filename -is [scriptblock]) {$fname = Invoke-Expression $(".{$filename}") }
      else                             {$fname = $filename } 
      if (test-path $fname)            {if     ($noclobber) {write-warning "$fName exists and WILL NOT be overwritten"; if ($passthru) {$image} ; Return }
                                        elseIF ($psc.shouldProcess($FName,"Delete file")) {Remove-Item  -Path $fname -Force -Confirm:$false }
                                        else   {Return}
      }  
      if ((Test-Path -Path $Fname -IsValid) -and ($pscmdlet.shouldProcess($FName,"Write image")))  { $image.SaveFile($FName) }
      if ($passthru) {$image} 
   }
}
