cls

if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "Must be run as Administrator"
    pause
    exit
}
if (-not (Test-Path -Path 'faves.md')) { Write-Error "faves.md doesn't exist"; break; }
$faveLinks = gc faves.md
if (
    $faveLinks.Length -lt 3 -or
    $faveLinks[0].Split("|").Length -lt 5
) {
    Write-Error "faves.md must be formatted as:"
    Write-Error "    Display Name | exe path | Working Dir | Elevated (True/False) | Arguments"
    Write-Error "    ------------ | -------- | ----------- | --------------------- | ---------"
    Write-Error "    Terminal     | wt.exe   |             | True                  |          "
    pause
    exit
}

$grp3Path = "$env:LOCALAPPDATA\Microsoft\Windows\WinX\Group3"
rmdir -Recurse -ErrorAction SilentlyContinue $grp3Path
rmdir -Recurse -ErrorAction SilentlyContinue "C:\Users\Default\AppData\Local\Microsoft\Windows\WinX\Group3" # this new one came along with Win11
mkdir $grp3Path | Out-Null

# this mapping is crucial for Windows to display the resulting shortcuts - fascinating
$systemFolderGuids = @{
    [Environment]::ExpandEnvironmentVariables("%programfiles%")    = "{905E63B6-C1BF-494E-B29C-65B732D3D21A}"
    [Environment]::ExpandEnvironmentVariables("%windir%\system32") = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}"
    [Environment]::ExpandEnvironmentVariables("%windir%")          = "{F38BF404-1D43-42F2-9305-67DE0B28FC23}"
}

function MapSystemFolders([string]$source) {
    $systemFolderGuids.GetEnumerator() | % {
        # echo "source: $source, key: $($_.Key), val: $($_.Value)"
        $source = $source -ireplace [regex]::Escape($_.Key), $_.Value #ireplace is case-INsensitive
    }
    return $source;
}

# from: https://stackoverflow.com/questions/40915420/how-to-expand-variable-in-powershell/56627626#56627626
# example: %bin%\blah\blah => c:\bin\blah\blah
function ExpandVars([string]$source) {
    $varMatches = [regex]::Matches($source, '%(.*?)%')
    if (-not $varMatches.Success) { return $source }

    $varMatches | % {
        if (-not (test-path "env:$($_.Groups[1])")) {
            write-error "env var not defined: $($_.Groups[1])"
            pause
            exit
        }
    }

    # $subst = $source -ireplace '%(.*?)%', '$($env:$1)'
    # # write-host "source: $source, expanded: $($ExecutionContext.InvokeCommand.ExpandString($subst))"
    # return $ExecutionContext.InvokeCommand.ExpandString($subst)
    return [Environment]::ExpandEnvironmentVariables($source)
}

# from: http://blog.coretech.dk/hra/create-shortcut-with-elevated-rights/
function ElevateLnk([string]$lnkFile) {
    $fs = New-Object System.IO.FileStream($lnkFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite)
    $fs.Seek(21, [System.IO.SeekOrigin]::Begin) | Out-Null
    $fs.WriteByte(0x22)
    $fs.Close()
}

#from: https://devblogs.microsoft.com/scripting/use-powershell-to-interact-with-the-windows-api-part-1/#using-add-type-to-call-the-copyitem-function
$MethodDefinition = @'
    [DllImport("Shlwapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError=true)]
    public static extern int HashData([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1, SizeParamIndex = 1)] [In] byte[] pbData, int cbData, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1, SizeParamIndex = 3)] [Out] byte[] piet, int outputLen);
'@
$shlwapi = Add-Type -MemberDefinition $MethodDefinition -Name 'Shlwapi' -Namespace 'Win32' -PassThru
Add-Type -Path "Microsoft.WindowsAPICodePack.dll"
Add-Type -Path "Microsoft.WindowsAPICodePack.Shell.dll" #ShellFile

# lifted following hash algorithm from: http://winaero.com/download.php?view.21
# this is the magic kicker that makes the shortcuts extraordinary and must be done for them to show up in the start menu
function HashLnk($lnkFile) {
    $text = MapSystemFolders $lnkFile.TargetPath.ToLower()
    
    if ($lnkFile.Arguments.Length -gt 0) { $text += $lnkFile.Arguments }
    # this magic string crucial for the hash to be accepted by Windows
    $text += "do not prehash links.  this should only be done by the user."
    $text = $text.ToLower();

    # echo "args.length: $($lnkFile.Arguments.Length), text: $text"
    $inBytes = [System.Text.Encoding]::GetEncoding(1200).GetBytes($text)
    $byteCount = $inBytes.Length
    $outBytes = [byte[]]::new($byteCount) 
    $hashResult = $shlwapi::HashData($inBytes, $byteCount, $outBytes, $byteCount)
    if ($hashResult -ne 0) { throw("Shlwapi::HashData failed: {Marshal.GetLastWin32Error()}") }
    $propertyWriter = [Microsoft.WindowsAPICodePack.Shell.ShellFile]::FromFilePath($lnkFile.FullName).Properties.GetPropertyWriter()
    $propertyWriter.WriteProperty("System.Winx.Hash", [System.BitConverter]::ToUInt32($outBytes, 0))
    $propertyWriter.Close()
}


$lineNum = 0
$linkCount = $faveLinks.Length
$faveLinks | select -skip 2 | % {
    $name, $exe, $dir, $elev, $cargs = ($_.Split("|"))
    if ($name.StartsWith(";")) { $linkCount--; return; } # commented out, skip it
    $lineNum++

    [bool]::TryParse($elev, [ref]$elev) | Out-Null

    $wsh = New-Object -ComObject WScript.Shell

    # for whatever reason Windows displays the WinX menu in reverse numeric order???
    # so we create the names counting down
    $lnkName = ($linkCount - $lineNum - 1).ToString().PadLeft(2, '0') + " - " + $name.Trim() + ".lnk"
    $lnkPath = [System.IO.Path]::Combine($grp3Path, $lnkName)
            
    $lnk = $wsh.CreateShortcut($lnkPath)
    $lnk.Description = $name.Trim()
    $lnk.TargetPath = [string](ExpandVars $exe.Trim()) # had to force cast to string or assignment would be blank???
    $lnk.WorkingDirectory = $dir.Trim()
    $lnk.Arguments = [string](ExpandVars $cargs.Trim())
    
    $lnk.Save();

    if ($elev) { ElevateLnk $lnkPath }

    HashLnk $lnk
    echo "num: $lineNum, name: $($lnk.Description), target: $($lnk.TargetPath), dir: $($lnk.WorkingDirectory), args: $($lnk.Arguments)"
}

echo ""
echo "will now recycle Explorer.exe to reload Win+X menu entries"
pause

kill -Name explorer | wait-process
explorer $PSScriptRoot