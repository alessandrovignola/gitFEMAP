Sub Main()
    Dim RetVal
    Dim pythonExe, scriptPath
    pythonExe = "C:\Python 310\venv\Scripts\pythonw.exe"
    scriptPath = "C:\Users\alessandro.vignola\gitFEMAP\main.py"

    ' Run Python script
    RetVal = Shell(pythonExe & " " & """" & scriptPath & """", vbHide)
End Sub