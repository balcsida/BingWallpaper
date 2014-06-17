# Parameters
param(
  [String]$Path = "C:\BingWallpapers\",
  [Boolean]$FlushDir = $TRUE,
  [String]$Mkt = "en-ww",
  [Int]$NumberOfImages = 3
)

# Code origin:
# http://stackoverflow.com/questions/22696981/get-http-request-and-tolerate-500-server-error-in-powershell
function Get-WebSiteStatusCode {
    param (
        [String] $testUri,
        $maximumRedirection = 10
    )
    $request = $null
    try {
        $request = Invoke-WebRequest -Uri $testUri -MaximumRedirection $maximumRedirection -ErrorAction SilentlyContinue
    } catch [System.Net.WebException] {
        $request = $_.Exception.Response
    } catch {
        Write-Error $_.Exception
        return $null
    }
    $request.StatusCode
}

# Code origin:
# http://techibee.com/powershell/powershell-script-to-get-desktop-screen-resolution/1615
# Slightly modifyed to handle multiple screens
function Get-ScreenResolution {
	[void] [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
	#[void] [Reflection.Assembly]::LoadWithPartialName("System.Drawing")
	$Screens = [System.Windows.Forms.Screen]::AllScreens

    $i = 0
    $OutputObj = New-Object -TypeName PSobject

	foreach ($Screen in $Screens) {
        $OutputObj | Add-Member -MemberType NoteProperty -Name Id -Value $i
		$OutputObj[$i] | Add-Member -MemberType NoteProperty -Name DeviceName -Value $Screen.DeviceName
		$OutputObj[$i] | Add-Member -MemberType NoteProperty -Name Width -Value $Screen.Bounds.Width
		$OutputObj[$i] | Add-Member -MemberType NoteProperty -Name Height -Value $Screen.Bounds.Height
		$OutputObj[$i] | Add-Member -MemberType NoteProperty -Name IsPrimaryMonitor -Value $Screen.Primary
        $i++
	}
    return $OutputObj
}

"Flush Directory: "+$FlushDir
# Testing directory
if(!(Test-Path $Path)){
    New-Item $Path -ItemType directory
} else {
    if($FlushDir){
        Get-ChildItem $Path -Include *.jpg -Recurse | Remove-Item
    }
}
# Check: 1 <= NumberOfImages <= 8
if($NumberOfImages -lt 1){
    Write-Error -Message "You must download at least one image!" -Category InvalidData
    Break
} elseif ($NumberOfImages -eq 1){
    $testUri = "http://www.bing.com" + $bingxml.images.image.urlBase + "_" + $Width + "x" + $Height + ".jpg"
} elseif ($NumberOfImages -gt 1 -and $NumberOfImages -le 8){
    $testUri = "http://www.bing.com" + $bingxml.images.image.urlBase[0] + "_" + $Width + "x" + $Height + ".jpg"
} else {
    Write-Error -Message "Max downloadable image is 8!" -Category InvalidData
    Break
}
"═" * 50
$bingxmlurl = "http://www.bing.com/HPImageArchive.aspx?format=xml&mbl=1&n="+$NumberOfImages+"&mkt="+$mkt
switch (Get-WebSiteStatusCode -testUri $bingxmlurl) { 
    200 {"OK!"}
    default {
        Write-Error -Message "Network Error" -Category InvalidData
        Break
    }
}
[xml]$bingxml = (Invoke-WebRequest $bingxmlurl).Content
$screen = Get-ScreenResolution
"Screen Resolution: "+$screen.Width+"x"+$screen.Height
$Width = $screen.Width
$Height = $screen.Height

Write-Host $testUri
"═" * 50
"Response:"+(Get-WebSiteStatusCode -testUri $testUri)
switch (Get-WebSiteStatusCode -testUri $testUri) { 
    "NotFound" {
        "The image is not available in this resoluion"
        "Will use the default 1366 x 768"
        $Width = 1366
        $Height = 768
    }
    200 {"OK!"}
    default {
        Write-Error -Message "Network Error" -Category InvalidData
        Break
    }
}
"═" * 50
[Int]$i=1
foreach ($bingurl in $bingxml.images.image) {
    $url = "http://www.bing.com" + $bingurl.urlBase + "_" + $Width + "x" + $Height + ".jpg"
    Write-Host $i": "$url
    $name = Split-Path $url -Leaf
    Invoke-WebRequest $url -OutFile ($Path+$name)
    Write-Host "Saved:"($Path+$name)
    Write-Host
    $i++
}