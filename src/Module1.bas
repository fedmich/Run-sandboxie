Attribute VB_Name = "Module1"
Option Explicit

Public Sub Sandbox_app(box$, exe$, Optional args$)
    Dim cmd$
    cmd = """C:\Program Files\Sandboxie\Start.exe"" /box:" & box & " """ & _
        exe & """  " & args
    Shell cmd, vbNormalFocus
End Sub




Public Sub Main()
    
    Form1.Show
End Sub
