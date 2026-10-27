Option Explicit
Dim fso, sh, root, py, app, url, cmd, i, ok
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh = CreateObject("WScript.Shell")
root = fso.GetParentFolderName(WScript.ScriptFullName)
py = root & "\.venv\Scripts\python.exe"
app = root & "\app\app.py"
url = "http://127.0.0.1:8502"
sh.CurrentDirectory = root

If Not fso.FileExists(py) Then
  MsgBox "Tsy hita ny Python ao amin'ny app. Alefaso indray ny INSTALLER_WINDOWS.bat.", 16, "MEMOIRE CORRECTEUR PRO gloss"
  WScript.Quit 1
End If

ok = IsServerReady(url)
If Not ok Then
  cmd = Chr(34) & py & Chr(34) & " -m streamlit run " & Chr(34) & app & Chr(34) & " --server.port 8502 --server.address 127.0.0.1 --server.headless true"
  sh.Run cmd, 0, False
  For i = 1 To 60
    WScript.Sleep 1000
    If IsServerReady(url) Then
      ok = True
      Exit For
    End If
  Next
End If

If ok Then
  sh.Run url, 1, False
Else
  MsgBox "Tsy afaka nanomboka ny serveur. Andramo sokafana LANCER_WINDOWS.bat hijerena ny hafatra erreur.", 16, "MEMOIRE CORRECTEUR PRO gloss"
End If

Function IsServerReady(checkUrl)
  On Error Resume Next
  Dim req
  Set req = CreateObject("MSXML2.XMLHTTP")
  req.Open "GET", checkUrl & "/_stcore/health", False
  req.Send
  IsServerReady = (Err.Number = 0 And req.Status = 200)
  Err.Clear
  Set req = Nothing
  On Error GoTo 0
End Function
