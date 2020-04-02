param(
    [String]$ExportFolder
)

$HeaderCSS = @"
<title>My News Sites</title>
<style>

body{
  background: linear-gradient(to right, #003366 0%, #000009 100%);
}

h1{ 
  font-family: arial;
} 

a {
  color: white;
  text-decoration: none;
}

#header {
  display:flex;
  justify-content:space-between;
  padding: 20px;
  text-align: center;
  color: white;
  font-size: 12px;
  border-bottom-left-radius: 20px;
  border-bottom-right-radius: 20px;
  width: 90%;
  margin-left: 4%;
}

.headerLinks {
  align-self:flex-end;
  display:flex;
}

table { 
  border: 1px solid #ddd;
  font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
  border-collapse: collapse;
  background-color:rgb(192,168,100,0.2);
  margin: 10px;
  width: 400px;
} 

th { 
  padding-top: 12px;
  padding-bottom: 12px;
  text-align: center;
  border-bottom: white 1px solid;
  color: white;
} 

td { 
  font-size: 13px; 
  padding: 15px; 
  color: #000; 
} 

tr:hover {
  background-color: rgb( 241, 196, 15,50%);
}

#container{
  display:flex;
  flex-wrap: wrap;
  padding:10px;
  margin-left: 8%
}

.linkItem{
  justify-content: space-between;
  padding: 10px;
}

#footer{
  justify-content:space-between;
  padding: 20px;
  background:   rgba(0, 0, 0, 0.75);
  color: white;
  font-size: 12px;
  width: 100%;
  text-align: center;
}

</style>
"@

function Get-TimeStamp{
  return ((Get-Date -format d)) + " " + (Get-Date -Format t)
}
function Get-HackerNews {

    $HackerNewsWebResponse = Invoke-WebRequest -Uri "https://news.ycombinator.com/"
    $Links = $HackerNewsWebResponse.links | Where-Object {$_.Class -eq "storylink"} 

    $HTMLString = @()
    $HTMLString += "<table>"
    $HTMLString += "<tr><th>Hacker News</th></tr>"

    foreach($Link in $Links){
        $HTMLString += "<tr>"
        $HTMLString += "$("<td><a href=`'$($Link.href)`'target=`'_blank`'>$($Link.InnerHTML)</a></td>")"
        $HTMLString += "</tr>"
    }

    $HTMLString += "</table>"
  
  return $HTMLString

}
function Get-ScienceDaily{

    $webpage = Invoke-WebRequest -Uri "https://www.sciencedaily.com/news/computers_math/computer_programming"
    $date = (Get-Date).Year
    $scienceDailyLinks = $webpage.links | Where-Object {$_.href -like "*releases/$($date)*" -and $_.InnerText -notlike "*read more*" }  | Select-Object -First 25
    $HTMLString = @()
    $HTMLString += "<table>"
    $HTMLString += "<tr><th>Science Daily</th></tr>"
    foreach($link in $scienceDailyLinks){
        $HTMLString += "<tr>"

        if($Link.InnerHTML -like "*<img*"){
            $Link.InnerHTML = "Title N/A"
        }

        $HTMLString += "<td><a href=`'https://www.sciencedaily.com$($link.href)`' target=`"_blank`">$($link.InnerHTML)</a></td>" 
        $HTMLString += "</tr>"
    }

  $HTMLString += "</table>"
  return $HTMLString

}
function Get-RedditSubReddit{

    param(
        [String]$SubReddit
    )

    $ProgrammingLinks = (Invoke-RestMethod -Uri https://www.reddit.com/r/$SubReddit/hot/.json).data.children.data | Sort-Object score -Descending | Select-Object Title, url
    $HTMLString = @()
    $HTMLString += "<table>"
    $HTMLString += "<tr><th>/r/$SubReddit</th></tr>"

    foreach($Link in $ProgrammingLinks){
        $HTMLString += "<tr>"
        $HTMLString += "<td><a href=`'$($Link.Url)`' target=`"_blank`">$($link.title)</a></td>" 
        $HTMLString += "</tr>"
    }

    $HTMLString += "</table>"
    return $HTMLString

}

############################################################Start here!###########################################################

if(!($ExportFolder)){
    $ExportFolder = $PSScriptRoot
}
try{
    $HeaderCSS | Out-File $ExportFolder\News.html -ErrorAction "Stop"
}catch{
    Write-Warning "Unable to output file to $ExportFolder"
    Write-Warning "Please confirm you have permission to the path above or try another path"
    exit
}

$HTMLBody = @"
<body>
    <div id=`"header`">
        <h1>My News Sites</h1>
        <div class=`"headerLinks`">
          <div class=`"linkItem`">
            <h2><a href=`"https://lithub.com/`" target=`"_blank`">Lithub</a></h2> 
          </div>
          <div class=`"linkItem`">
            <h2><a href= `"https://thereader.mitpress.mit.edu/`" target=`"_blank`">The MIT Press Reader</a></h2>
          </div>
          <div class=`"linkItem`">
            <h2><a href= `"https://docs.microsoft.com/en-us/mem/intune/fundamentals/whats-new`" target=`"_blank`">Intune News</a></h2>
          </div>
        </div>
    </div>
    <div id=`"container`">
        $(Get-HackerNews)
        $(Get-ScienceDaily)
        $(Get-RedditSubReddit -SubReddit programming)
        $(Get-RedditSubReddit -SubReddit powershell)
    </div> 
    <div id=`"footer`">
        <h2><a href= `"https://github.com/mallockey/`" target=`"_blank`">Github</a></h2>
        This page was created using PowerShell <br>
        Created on: $(Get-TimeStamp)
    </div>
</body>
"@

$HTMLBody | Out-File $ExportFolder\News.html -Append

$VerbosePreference = "Continue"
if(Test-Path $ExportFolder\News.html){
    Write-Verbose "News.html file successfully exported to $($ExportFolder)"
}else{
    Write-Warning "An unknown error occured, the News was file was not exported"
}
