function Write-ColorOutput{

	param(
		[ValidateSet("Black", "DarkBlue", "DarkGreen", "DarkCyan", "DarkRed", "DarkMagenta", 
		             "DarkYellow", "Gray", "DarkGray", "Blue", "Green", "Cyan", 
					 "Red", "Magenta", "Yellow", "White")]
		[String]$ForegroundColor,
		[ValidateSet("Black", "DarkBlue", "DarkGreen", "DarkCyan", "DarkRed", "DarkMagenta", 
		             "DarkYellow", "Gray", "DarkGray", "Blue", "Green", "Cyan", 
					 "Red", "Magenta", "Yellow", "White")]
		[String]$BackgroundColor,
		[String]$Message
	)

    # save the current color
    $OriginalFGC = $host.UI.RawUI.ForegroundColor
    $OriginalBGC = $host.UI.RawUI.BackgroundColor
	
    # set the new color
    $host.UI.RawUI.ForegroundColor = $ForegroundColor
    $host.UI.RawUI.BackgroundColor = $BackgroundColor

    # output
    Write-Output $Message

    # restore the original color
    $host.UI.RawUI.ForegroundColor = $OriginalFGC
    $host.UI.RawUI.BackgroundColor = $OriginalBGC
}
