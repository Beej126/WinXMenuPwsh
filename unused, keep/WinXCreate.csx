#r "Microsoft.WindowsAPICodePack.Shell.dll"
#r "Microsoft.WindowsAPICodePack.dll"
#r "Interop.IWshRuntimeLibrary.dll"

using IWshRuntimeLibrary;
using Microsoft.WindowsAPICodePack.Shell;
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

private static string ExpandEnvVar(string withVars) {
  var matches = Regex.Matches(withVars, "%.*%");
  foreach(Match match in matches) {
    if (Environment.GetEnvironmentVariable(match.Value.Trim('%')) == null)
      Console.WriteLine($"[91mNot Defined: {match.Value}[0m");
  }
  return Environment.ExpandEnvironmentVariables(withVars);
}

internal class ReverseSorter : System.Collections.IComparer
{
  int System.Collections.IComparer.Compare(object x, object y)
  {
    return new System.Collections.CaseInsensitiveComparer().Compare(y, x);
  }
}

private static readonly KeyValuePair<string, string>[] SystemFolderMapping = new[]
{
  new KeyValuePair<string, string>(ExpandEnvVar("%programfiles%").ToLower(), "{905E63B6-C1BF-494E-B29C-65B732D3D21A}"),
  new KeyValuePair<string, string>(ExpandEnvVar("%windir%\\system32").ToLower(), "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}"),
  new KeyValuePair<string, string>(ExpandEnvVar("%windir%").ToLower(), "{F38BF404-1D43-42F2-9305-67DE0B28FC23}"),
  new KeyValuePair<string, string>(ExpandEnvVar("%systemroot%").ToLower(), "{F38BF404-1D43-42F2-9305-67DE0B28FC23}")
};

private static readonly WshShell wsh = new WshShell();

private static readonly string WinXFolder = ExpandEnvVar(@"%LOCALAPPDATA%\Microsoft\Windows\WinX");
private static readonly System.Collections.IComparer ReverseSort = new ReverseSorter();

public static class WinXHasher
{
  [DllImport("Shlwapi.dll", CharSet = CharSet.Unicode, ExactSpelling = true, SetLastError=true)]
  internal static extern int HashData([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1, SizeParamIndex = 1)] [In] byte[] pbData, int cbData, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.U1, SizeParamIndex = 3)] [Out] byte[] piet, int outputLen);

  private static bool _isFileList = false;
  private static string _groupFolderPath;
  public static void HashLnk(string lnkFile)
  {
    if (Path.GetExtension(lnkFile)?.ToLower() != ".lnk")
    {
      if (_isFileList) return;
      _isFileList = true;
      if (_groupFolderPath == null) {
        _groupFolderPath = Path.Combine(WinXFolder, "Group"+GetNextMaxGroup());
        Directory.CreateDirectory(_groupFolderPath);
      }
      var lines = System.IO.File.ReadAllLines(lnkFile);
      for (var lineNum=2; lineNum < lines.Length; lineNum++) //skip first 2 header lines
      {
        var line = lines[lineNum];
        var chunks = line.Split('|');
        if (chunks.Length != 5) {
          Console.WriteLine("Listing file needs to be lines in following bar separated format (no quotes):");
          Console.WriteLine("Display Name | exe path | Elevated (True/False) | Arguments");
          return;
        }
        var title = chunks[0].Trim();
        if (title.StartsWith("#")) { continue; } //commented out
        var lnkPath = Path.Combine(_groupFolderPath, 
            //the menu structure works in reverse of the top to bottom order in the shortcuts listing file
            (lines.Length - lineNum - 1).ToString().PadLeft(2, '0') + " - " + title + ".lnk");
        var wshShortcut = (IWshShortcut)wsh.CreateShortcut(lnkPath);
        wshShortcut.Description = title;
        wshShortcut.TargetPath = ExpandEnvVar(chunks[1].Trim());
        wshShortcut.WorkingDirectory = chunks[2].Trim();
        bool isElevated; bool.TryParse(chunks[3].Trim(), out isElevated);
        wshShortcut.Arguments = ExpandEnvVar(chunks[4].Trim()); //doing pre eval on envVar for commands that don't do themselves
        
        Console.WriteLine($"lnk: {lnkPath}, target: {wshShortcut.TargetPath}, args: {wshShortcut.Arguments}");
        wshShortcut.Save();
        
        if (isElevated) ElevateLnk(lnkPath);
        
        HashLnk(lnkPath); 
      }
      return;
    }
  
    var lnk = (IWshShortcut)wsh.CreateShortcut(lnkFile);
    
    var text = lnk.TargetPath;
    
    // this mapping is also crucial for windows to display the resulting shortcuts - fascinating
    foreach(var kv in SystemFolderMapping) text = text.ToLower().Replace(kv.Key, kv.Value);

    //this is the magic kicker that makes the shortcuts special and show up in the start menu
    //lifted following hash algorithm from: http://winaero.com/download.php?view.21
    if (lnk.Arguments.Length > 0) text += lnk.Arguments;
    //this magic string appears to be necessary for the hash to be accepted by Windows
    text += "do not prehash links.  this should only be done by the user.";
    text = text.ToLower();
    var inBytes = Encoding.GetEncoding(1200).GetBytes(text);
    var byteCount = inBytes.Length;
    var outBytes = new byte[byteCount];
    var hashResult = HashData(inBytes, byteCount, outBytes, byteCount);
    if (hashResult != 0) throw new Exception("Shlwapi::HashData failed: {Marshal.GetLastWin32Error()}");
    using (var propertyWriter = ShellFile.FromFilePath(lnkFile).Properties.GetPropertyWriter())
    {
      propertyWriter.WriteProperty("System.Winx.Hash", BitConverter.ToUInt32(outBytes, 0));
    }
  }
  
  private static int GetNextMaxGroup()
  {
    var directories = Directory.GetDirectories(WinXFolder, "Group*", SearchOption.TopDirectoryOnly);
    Array.Sort(directories, ReverseSort);
    foreach (var path in directories)
    {
      var s = Path.GetFileName(path)?.Substring(5).Trim();
      if (s == null) continue;
      int result;
      if (int.TryParse(s, out result)) return result+1;
    }
    return 0;
  }

  //from: http://blog.coretech.dk/hra/create-shortcut-with-elevated-rights/
  private static void ElevateLnk(string lnkFile)
  {
    using (FileStream fs = new FileStream(lnkFile, FileMode.Open, FileAccess.ReadWrite))
    {
      fs.Seek(21, SeekOrigin.Begin);
      fs.WriteByte(0x22);
    }
  }
  
}

WinXHasher.HashLnk(Env.ScriptArgs[0]);