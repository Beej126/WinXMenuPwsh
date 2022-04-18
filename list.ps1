cls

function IsElevated([string]$lnkFile) {
  $fs = New-Object System.IO.FileStream($lnkFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read)
  $fs.Seek(21, [System.IO.SeekOrigin]::Begin) | Out-Null
  $result = ($fs.ReadByte() -eq 0x22)
  $fs.Close()
  return $result
}

$wsh = New-Object -ComObject WScript.Shell

$shortcutFileNames = [System.IO.Directory]::GetFiles("$env:LOCALAPPDATA\Microsoft\Windows\WinX\Group3", "*.lnk")

$table = [System.Collections.ArrayList]@()

for ($i = $shortcutFileNames.Length-1; $i -gt -1; $i--) { # Win+X menu is structured in reverse order
  $fn = $shortcutFileNames[$i];
  $titleMatch = [System.Text.RegularExpressions.Regex]::Match($fn, ".*?- (.*?).lnk");
  $title = $titleMatch.Success ? $titleMatch.Groups[1].Value : [System.IO.Path]::GetFileNameWithoutExtension($fn);
  $sc = $wsh.CreateShortcut($fn); # http://stackoverflow.com/a/4909475/813599
  $table.Add([PSCustomObject]@{
    title = $title 
    target = $sc.TargetPath
    dir = $sc.WorkingDirectory
    elev = $(IsElevated $fn).ToString()
    args = $sc.Arguments
  }) | Out-Null
}

$table | format-table