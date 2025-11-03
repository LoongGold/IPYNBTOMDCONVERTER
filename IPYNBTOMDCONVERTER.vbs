' IPYNB to MD Converter
Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' Check if file was dragged
If WScript.Arguments.Count = 0 Then
    MsgBox "Please drag .ipynb files onto this script to convert them to .md", vbInformation, "IPYNB to MD Converter"
    WScript.Quit
End If

inputFile = WScript.Arguments(0)

' Check file extension
If LCase(fso.GetExtensionName(inputFile)) <> "ipynb" Then
    MsgBox "Error: Only .ipynb files are supported!", vbCritical, "Error"
    WScript.Quit
End If

' Check if file exists
If Not fso.FileExists(inputFile) Then
    MsgBox "Error: File does not exist!", vbCritical, "Error"
    WScript.Quit
End If

' Build output file path
outputFile = fso.GetParentFolderName(inputFile) & "\" & fso.GetBaseName(inputFile) & ".md"

On Error Resume Next

' Convert using jupyter nbconvert
WshShell.Run "cmd /c jupyter nbconvert --to markdown """ & inputFile & """", 0, True

If Err.Number <> 0 Then
    MsgBox "Conversion failed! Please make sure Jupyter Notebook is installed." & vbCrLf & "Install command: pip install jupyter", vbCritical, "Conversion Failed"
    WScript.Quit
End If

' Check if output file was created
If fso.FileExists(outputFile) Then
    MsgBox "Conversion successful!" & vbCrLf & "Original: " & fso.GetFileName(inputFile) & vbCrLf & "New: " & fso.GetFileName(outputFile) & vbCrLf & "Location: " & fso.GetParentFolderName(inputFile), vbInformation, "Conversion Complete"
Else
    MsgBox "Conversion failed! Could not find the generated Markdown file.", vbCritical, "Conversion Failed"
End If