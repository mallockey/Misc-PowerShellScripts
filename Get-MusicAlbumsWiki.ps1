param(
	[Switch]$AllYears,
	[String]$ArtistToSearch,
	[Switch]$ExportCache,
	[Switch]$UseCache,
	[Array]$Year =(Get-Date).Year,
	[String]$Month
)

function Test-URL{
  param(
	$CurrentURL
  )

  try{
    $HTTPRequest = [System.Net.WebRequest]::Create($CurrentURL)
    $HTTPResponse = $HTTPRequest.GetResponse()
    $HTTPStatus = [Int]$HTTPResponse.StatusCode

    if($HTTPStatus -ne 200) {
      return $False
    }

    $HTTPResponse.Close()

  }catch{
		return $False
  }	
  return $True
}

if($AllYears){
	$Year = 2005..(Get-Date).Year
}

if($ExportCache){
	if(!(Test-Path "$PSScriptRoot\AlbumCache\")){
		New-Item "$PSScriptRoot\AlbumCache\" -ItemType Directory
	}
}

$ListOfMonths = "January","February","March","April","May","June","July","August","September","October","November","December"
$ArrayOfAlbums = [System.Collections.ArrayList]@()

$AlbumProps = [Ordered]@{
	Album = $Null
	Artist = $Null
	ReleaseDate = $Null
	ReleaseYear = $Null
}

foreach($CurrentYear in $Year){

	$Counter = 0
	if($UseCache){
		try{
			$ImportedAlbum = Import-Csv "$PSScriptRoot\AlbumCache\ListOf$($CurrentYear)Albums.csv"
		}catch{
			Write-Warning "Cache not found for $($CurrentYear) albums."
			Write-Warning "Please confirm you have exported $($CurrentYear) by typing in the below:"
			Write-Warning "Get-AlbumsFromWiki -Year $($CurrentYear) -ExportCache"
			Write-Warning "Then rerun the script using the -UseCache Parameter"
			continue
		}
		
		$ArrayOfAlbums += $ImportedAlbum
	}else{
		if((Test-URL -CurrentURL "https://en.wikipedia.org/wiki/List_of_$($CurrentYear)_albums")-eq $True){
		
			$Wiki2020AlbumsWebPage = Invoke-WebRequest "https://en.wikipedia.org/wiki/List_of_$($CurrentYear)_albums"
			$Content = $Wiki2020AlbumsWebPage | ForEach-Object {$_.Content}
			$Content = $Content -split '\r?\n'
			
			foreach($Line in $Content){
				$ListOfMonths | ForEach-Object {
					if($Line -like "*$_<br />*"){
						$ReleaseDate = $line
					}
				}
				
				if($Line -like "*<i>*</i>"){
					$AlbumObj = New-Object -TypeName PSObject -Prop $AlbumProps
					$CurrentAlbum = $Line
					$CurrentArtistIndex = $Counter - 2
					$CurrentArtist = $Content[$CurrentArtistIndex]
					$CurrentArtist = $CurrentArtist -replace '<[^>]+>',''
					$CurrentAlbum = $CurrentAlbum -replace '<[^>]+>',''
					$CurrentAlbum = $CurrentAlbum -replace '&amp;',''
					$ReleaseDate = $ReleaseDate -replace '<[^>]+>',''
					$AlbumObj.Album = $CurrentAlbum
					$AlbumObj.Artist = $CurrentArtist
					$AlbumObj.ReleaseDate = $ReleaseDate
					$AlbumObj.ReleaseYear = $CurrentYear
					$ArrayOfAlbums.Add($AlbumObj) | Out-Null
				}
				
				$Counter++
			}
			if($ExportCache){
				$ArrayOfAlbums | Where-Object {$_.ReleaseYear -eq "$CurrentYear"} | Export-Csv "$PSScriptRoot\AlbumCache\ListOf$($CurrentYear)Albums.csv" -NoTypeInformation
			}
		}
		else{
			Write-Warning "https://en.wikipedia.org/wiki/List_of_$($CurrentYear)_albums is not a valid link"
		}
	}
}

$ArrayOfAlbums | Where-Object {$_.Artist -like "*$ArtistToSearch*" -and $_.ReleaseDate -like "*$Month*"}

