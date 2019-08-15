Param(
[Int]$NumberOfArticles = 10
)
Add-Type -AssemblyName System.Web
$Header = @"
<title>News</title>
<style>

h1, h5, th { 
	text-align: center;
	font-family: arial;
} 
h2 {
	text-align: center;
  Font-size: 15px;
}
table { 
	margin: auto; font-family: Segoe UI; box-shadow: 10px 10px 5px #888; border: thin ridge grey; 
} 
th { 
	background: #0046c3; 
	color: #fff; 
	max-width: 400px; 
	padding: 5px 10px; 
} 
td { 
	font-size: 13px; 
	padding: 5px 20px; 
	color: #000; 
} 
tr { 
	background: #b8d1f3; 
}
tr:nth-child(even) { 
	background: #dae5f4; 
} 
tr:nth-child(odd) { 
	background: #b8d1f3; 
}
tr:hover {
	background: yellow;
}
</style>
"@

$date = get-date -Format d
$time = get-date -Format t
$todaysDateAndTime = $date + " " + $time
function getHackerNews {
	$hNewsWP = "https://news.ycombinator.com/"
	$hNewsWP = Invoke-WebRequest -uri $hNewsWP

	$hackerNewsLinks = $hNewsWP.links | Where-Object {$_.Class -eq "storylink"} |
	Select-Object -Property  @{ Label ='<tr><th>Hacker News</th></tr>'; 
	expression = {"<a href=`'$($_.href)`'target=`"_blank`">$($_.InnerHTML)</a>"};}|Select-Object -First $NumberOfArticles | ConvertTo-Html 
	ConvertTo-Html -Body "<h2>Hacker News</h2>" | Out-File "C:\Kits\News.html" -Append

	$var = [System.Web.HttpUtility]::HtmlDecode($hackerNewsLinks)
	$var  | Out-File "C:\Kits\News.html" -Append
}
function getScienceDaily{
	$url = "https://www.sciencedaily.com/news/computers_math/computer_programming"
	$webpage = Invoke-WebRequest -Uri $url
	$date = (Get-Date).Year
	$scienceDailyLinks = $webpage.links | Where-Object {$_.href -like "*releases/$($date)*" -and $_.InnerText -notlike "*read more*" } | Select-Object -First $NumberOfArticles

	foreach($link in $scienceDailyLinks){
		$link.href = "<a href=`'https://www.sciencedaily.com$($link.href)`' target=`"_blank`">$($link.InnerText)</a>" 
	}

	$scienceDailyLinks = $scienceDailyLinks | Select-Object  href
	$scienceDailyLinks = $scienceDailyLinks | ConvertTo-Html
	ConvertTo-Html -Body "<h2>Science Daily</h2>" | Out-File "C:\Kits\News.html" -Append

	$secondVar = [System.Web.HttpUtility]::HtmlDecode($scienceDailyLinks)
	$secondVar | Out-File "C:\Kits\News.html" -Append
}
ConvertTo-Html -head $header -Body  "<h1>$todaysDateAndTime's news</h1>" | Out-File "C:\kits\News.html" 

getHackerNews
getScienceDaily
