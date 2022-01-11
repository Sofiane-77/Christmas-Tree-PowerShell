function Center() {
    #Returns the input value with the amount of space needed to be centered on each line
    Param ([string]$Text)

    $CenteredString = [System.Collections.ArrayList]::new() # Array will contains each line centered

    foreach($line in $Text -split [System.Environment]::NewLine) {
        $line = $line.Trim()
        [void]$CenteredString.Add(("{0}{1}" -f (' ' * (([Math]::Max(0, $Host.UI.RawUI.BufferSize.Width / 2) - [Math]::Floor($line.Length / 2)))), $line))
    }

    return $CenteredString -Join [System.Environment]::NewLine # Join each line to return a string
}

function AddDecorationTags() {
    #Replace caracter with #color#caracter# pattern
    Param([string]$Text, [Hashtable]$Decorations)

    foreach($decoration in $Decorations.keys) {
        $Text = $Text.Replace("$decoration", "#$($Decorations.$decoration)#$decoration#")
    }

    return $Text
}

function RemoveDecorationTags() {
    #Remove the pattern
    Param([string]$Text, [Hashtable]$Decorations)

    foreach($decoration in $Decorations.keys) {
        $Text = $Text.Replace("#$($Decorations.$decoration)#$decoration#", "$decoration")
    }

    return $Text
}

function Write-Host-Color() {
    Param([string]$DecoratedText, [array]$Colors, [string]$DefaultFGColor)

    $currentColor = $DefaultFGColor

    $allColors = $Colors + 'Random' # Add random to the list of possible colors

    # Iterate through splitted Messages
	foreach( $string in $DecoratedText.split('#') ){
		# If a string between #-Tags is equal to any predefined color, and is equal to the defaultcolor: set current color
		if ( $allColors -contains $string -and $currentColor -eq $DefaultFGColor ){
            # if random chosen, we need to set a real color
            if ($string -ieq 'random') {
                $string = ($Colors | Get-Random)
            }
			$currentColor = $string
		}else{
			# If string is a output message, than write string with current color (with no line break)
			Write-Host -nonewline -f $currentColor "$string"
			# Reset current color
			$currentColor = $DefaultFGColor
		}
	}

	# Single write-host for the final line break
	Write-Host

}

$christmasTree = @"
         |
        -+-
         A
        /=\
      i/ O \i
      /=====\
      /  i  \
    i/ O * O \i
    /=========\
    /  *   *  \
  i/ O   i   O \i
  /=============\
  /  O   i   O  \
i/ *   O   O   * \i
/=================\
       |___|
"@

$animationLoopNumber = 50 # Number of times to loop animation

$animationSpeed = 300 # Time in milliseconds to show each frame

$decorations = @{
'O' = 'random';
'|___|' = 'red'; # Trunk of christmas tree
} # Hash Table where key => caracter we want to animate / value => color we want to display this caracter

$foregroundColors = @('Black', 'DarkBlue', 'DarkGreen', 'DarkCyan', 'DarkRed', 'DarkMagenta', 'DarkYellow', 'Gray', 'DarkGray', 'Blue', 'Green', 'Cyan', 'Red', 'Magenta', 'Yellow', 'White');

$currentColor = [Console]::ForegroundColor # Current color of terminal

$currentBufferSize = $Host.UI.RawUI.BufferSize.Width # Default Width of terminal

$currentCursorSize = $Host.UI.RawUI.CursorSize

$christmasTree = Center -Text $christmasTree # Center in terminal

$christmasTree = AddDecorationTags -Text $christmasTree -Decorations $decorations

if (![string]::IsNullOrWhiteSpace($Host.UI.RawUI.CursorSize) -or $Host.UI.RawUI.CursorSize -gt 0) {
    $Host.UI.RawUI.CursorSize = 0 # Hide the cursor before display the tree
}

# Messages displayed

[int]$currentYear = Get-Date -UFormat "%Y"

# The following year is displayed only if the date is greater than December 23
if ((Get-Date) -gt (Get-Date -Year $currentYear -Month 12 -Day 23)) {
    $newYear = $currentYear + 1
}else{
    $newYear = $currentYear
}

$messages = @{
    'MerryChristmas' = @{
        'Text' = "Joyeux Noël !!".ToUpper();
        'Colors' = ("Red", "Green")
    };
    'DevMessage' = @{
        'FormattedText' = "Et beaucoup de {0} en $newYear";
        'FormattedValue' = "CODE"
        'FormattedTextColor' = 'White'
    };
    'HappyNewYear' = @{
        'Text' = "Bonne année !".ToUpper();
        'Colors' = ("Yellow", "Cyan")
    };
}

function CustomDecoratedMessage() {
    Param([string]$FormattedText, [string]$FormattedValue, [boolean]$Centered)

    $text = ""

    if ($Centered) {
        $FormattedText = (Center -Text ($FormattedText -f $FormattedValue)).Split($FormattedText[0])[0] + $FormattedText
    }


    $decorations = @{
        "$FormattedValue" = 'random'
    }

    $text = AddDecorationTags -Text $FormattedValue -Decorations $decorations

    return $FormattedText -f $text
}

# End Messages displayed

# Merry Christmas Song

$playSong = $true

if ($playSong) {
    # We created another thread to be able to play the song AND display the tree
    $Runspace = [runspacefactory]::CreateRunspace()
    $PowerShell = [powershell]::Create()
    $PowerShell.Runspace = $Runspace
    $Runspace.Open()
    $PowerShell.AddScript({

        $Duration = @{
            WHOLE     = 1600;
        }
        $Duration.HALF = $Duration.WHOLE/2;
        $Duration.QUARTER = $Duration.HALF/2;
        $Duration.EIGHTH = $Duration.QUARTER/2;
        $Duration.SIXTEENTH = $Duration.EIGHTH/2;

        # Hashtable where key = note and value = frequency
        $Notes = @{
            A4 = 440;
            B4 = 493.883301256124;
            C5 = 523.251130601197;
            D5 = 587.329535834815;
            E5 = 659.25511382574;
            F5 = 698.456462866008;
            G4 = 391.995435981749;
        }


        [console]::beep($Notes.G4,$Duration.EIGHTH)

        Start-Sleep -m $Duration.SIXTEENTH

        [console]::beep($Notes.C5,$Duration.EIGHTH)

        Start-Sleep -m $Duration.SIXTEENTH

        [console]::beep($Notes.C5,$Duration.EIGHTH)

        [console]::beep($Notes.D5,$Duration.EIGHTH)

        [console]::beep($Notes.C5,$Duration.EIGHTH)

        [console]::beep($Notes.B4,$Duration.EIGHTH)

        [console]::beep($Notes.A4,$Duration.EIGHTH)

        Start-Sleep -m $Duration.SIXTEENTH

        [console]::beep($Notes.A4,$Duration.EIGHTH)

        Start-Sleep -m $Duration.SIXTEENTH

        [console]::beep($Notes.A4,$Duration.EIGHTH)

        Start-Sleep -m $Duration.SIXTEENTH

        [console]::beep($Notes.D5,$Duration.EIGHTH)

        Start-Sleep -m $Duration.SIXTEENTH

        [console]::beep($Notes.D5,$Duration.EIGHTH)

        [console]::beep($Notes.E5,$Duration.EIGHTH)

        [console]::beep($Notes.D5,$Duration.EIGHTH)

        [console]::beep($Notes.C5,$Duration.EIGHTH)

        [console]::beep($Notes.B4,$Duration.EIGHTH)

        Start-Sleep -m $Duration.SIXTEENTH

        [console]::beep($Notes.G4,$Duration.EIGHTH)

        Start-Sleep -m $Duration.SIXTEENTH

        [console]::beep($Notes.G4,$Duration.EIGHTH)

        Start-Sleep -m $Duration.SIXTEENTH

        [console]::beep($Notes.E5,$Duration.EIGHTH)

        Start-Sleep -m $Duration.SIXTEENTH

        [console]::beep($Notes.E5,$Duration.EIGHTH)

        [console]::beep($Notes.F5,$Duration.EIGHTH)

        [console]::beep($Notes.E5,$Duration.EIGHTH)

        [console]::beep($Notes.D5,$Duration.EIGHTH)

        [console]::beep($Notes.C5,$Duration.EIGHTH)

        Start-Sleep -m $Duration.SIXTEENTH

        [console]::beep($Notes.A4,$Duration.EIGHTH)

        Start-Sleep -m $Duration.SIXTEENTH

        [console]::beep($Notes.G4,$Duration.EIGHTH)

        Start-Sleep -m $Duration.SIXTEENTH

        [console]::beep($Notes.A4,$Duration.EIGHTH)

        Start-Sleep -m $Duration.SIXTEENTH

        [console]::beep($Notes.D5,$Duration.EIGHTH)

        Start-Sleep -m $Duration.SIXTEENTH

        [console]::beep($Notes.B4,$Duration.EIGHTH)

        Start-Sleep -m $Duration.SIXTEENTH

        [console]::beep($Notes.C5,$Duration.EIGHTH)

    })

    $PowerShell.BeginInvoke()
}
# End Merry Christmas Song

$i = 0

do {

    if ($currentBufferSize -ine $Host.UI.RawUI.BufferSize.Width) {

        $christmasTree = RemoveDecorationTags -Text $christmasTree -Decorations $decorations # Remove Tags to be able to center correctly

        $christmasTree = Center -Text $christmasTree # Adjust to the new BufferSize

        $christmasTree = AddDecorationTags -Text $christmasTree -Decorations $decorations

        $currentBufferSize = $Host.UI.RawUI.BufferSize.Width

        if (![string]::IsNullOrWhiteSpace($Host.UI.RawUI.CursorSize) -or $Host.UI.RawUI.CursorSize -gt 0) {
            $Host.UI.RawUI.CursorSize = 0 # Hide the cursor again because it is displayed when we resize the window
        }
    }

    Clear-Host

    Write-Host-Color -DecoratedText $christmasTree -Colors $foregroundColors -DefaultFGColor 'Green' # Color of the chrismas tree

    Write-Host (Center -Text $messages.MerryChristmas.Text) -ForegroundColor ($messages.MerryChristmas.Colors | Get-Random)

    Write-Host-Color (CustomDecoratedMessage -FormattedText $messages.DevMessage.FormattedText -FormattedValue $messages.DevMessage.FormattedValue -Centered $true) -Colors $foregroundColors -DefaultFGColor $messages.DevMessage.FormattedTextColor

    Write-Host (Center -Text $messages.HappyNewYear.Text) -ForegroundColor ($messages.HappyNewYear.Colors | Get-Random)

    Start-Sleep -Milliseconds $animationSpeed

    $i++

} until ($i -eq $animationLoopNumber)

# We dont need to set CursorSize if the value is null
if (![string]::IsNullOrWhitespace($currentCursorSize)) {
    $Host.UI.RawUI.CursorSize = $currentCursorSize
}