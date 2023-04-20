Attribute VB_Name = "modDotEnv"
'Activate the library "Microsoft Scripting Runtime"

'Import necessary libraries
Option Explicit
Dim FSO As FileSystemObject
Dim envDict As Scripting.Dictionary

Public Sub LoadEnv()
    'Initialize FileSystemobject and Dictionary objects
    Set FSO = New FileSystemObject
    Set envDict = New Scripting.Dictionary
    
    'Define path to .env file
    Dim envPath As String
    envPath = ThisWorkbook.Path & "\.env"
    
    'Check if .env file exists
    If Not FSO.FileExists(envPath) Then
        MsgBox "Could not find .env file"
        Exit Sub
    End If
    
    'Read .env file and add variables to dictionary
    Dim envFile As TextStream
    Set envFile = FSO.OpenTextFile(envPath, ForReading)
    Do Until envFile.AtEndOfStream
        Dim line As String
        line = envFile.ReadLine
        If InStr(line, "=") > 0 Then
            Dim parts() As String
            parts = Split(line, "=")
            'Delete single and doubles quotes
            envDict(parts(0)) = Replace(Replace(parts(1), "'", ""), """", "")
        End If
    Loop
    envFile.Close
End Sub

'Get value of environment variable
Public Function GetEnv(key As String) As Variant
    If envDict.Exists(key) Then
        GetEnv = envDict(key)
    Else
        GetEnv = Null
    End If
End Function

Public Sub Example()
    LoadEnv
    
    ' Get environment variable
    Debug.Print GetEnv("USERNAME")
    Debug.Print GetEnv("PASSWORD")
End Sub

