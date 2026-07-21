$art = @"
       ,   ,
         ,-`{-`/
      ,-~ , \ {-~~-,
    ,~  ,   ,`,-~~-,`,
  ,`   ,   { {      } }                                             }/
 ;     ,--/`\ \    / /                                     }/      /,/
;  ,-./      \ \  { {  (                                  /,;    ,/ ,/
; /   `       } } `, `-`-.___                            / `,  ,/  `,/
 \|         ,`,`    `~.___,---}                         / ,`,,/  ,`,;
  `        { {                                     __  /  ,`/   ,`,;
        /   \ \                                 _,`, `{  `,{   `,`;`
       {     } }       /~\         .-:::-.     (--,   ;\ `,}  `,`;
       \\._./ /      /` , \      ,:::::::::,     `~;   \},/  `,`;     ,-=-
        `-..-`      /. `  .\_   ;:::::::::::;  __,{     `/  `,`;     {
                   / , ~ . ^ `~`\:::::::::::<<~>-,,`,    `-,  ``,_    }
                /~~ . `  . ~  , .`~~\:::::::;    _-~  ;__,        `,-`
       /`\    /~,  . ~ , '  `  ,  .` \::::;`   <<<~```   ``-,,__   ;
      /` .`\ /` .  ^  ,  ~  ,  . ` . ~\~                       \\, `,__
     / ` , ,`\.  ` ~  ,  ^ ,  `  ~ . . ``~~~`,                   `-`--, \
    / , ~ . ~ \ , ` .  ^  `  , . ^   .   , ` .`-,___,---,__            ``
  /` ` . ~ . ` `\ `  ~  ,  .  ,  `  ,  . ~  ^  ,  .  ~  , .`~---,___
/` . `  ,  . ~ , \  `  ~  ,  .  ^  ,  ~  .  `  ,  ~  .  ^  ,  ~  .  `-,
"@

$artLines = $art -split "`n"
$colors = @("Red", "DarkRed", "Yellow", "DarkYellow", "Green", "DarkGreen", "Cyan", "DarkCyan", "Blue", "DarkBlue", "Magenta", "DarkMagenta")
$startTime = Get-Date
$duration = New-TimeSpan -Seconds 25

# Clear the console first
Clear-Host

# Save starting cursor position
$startPos = $host.UI.RawUI.CursorPosition

while ((Get-Date) -lt ($startTime + $duration)) {
    # Shift colors for animation effect
    $firstColor = $colors[0]
    $colors = $colors[1..($colors.Length - 1)] + $firstColor
    
    # Return cursor to starting position
    $host.UI.RawUI.CursorPosition = $startPos
    
    # Display the art with current color configuration
    for ($lineIndex = 0; $lineIndex -lt $artLines.Length; $lineIndex++) {
        $line = $artLines[$lineIndex]
        $cursorLeft = $startPos.X
        
        for ($charIndex = 0; $charIndex -lt $line.Length; $charIndex++) {
            $char = $line[$charIndex]
            # Skip coloring spaces
            if ($char -ne ' ') {
                $colorIndex = ($lineIndex + $charIndex) % $colors.Count
                Write-Host $char -ForegroundColor $colors[$colorIndex] -NoNewline
            } else {
                Write-Host $char -NoNewline
            }
            $cursorLeft++
        }
        
        # Move to the next line
        $host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates $startPos.X, ($startPos.Y + $lineIndex + 1)
    }
    
    # Small delay to control animation speed
    Start-Sleep -Milliseconds 100
}

# Move cursor below the art when done
$host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates 0, ($startPos.Y + $artLines.Length + 1)