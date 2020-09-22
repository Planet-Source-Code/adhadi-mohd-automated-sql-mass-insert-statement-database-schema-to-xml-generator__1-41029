Attribute VB_Name = "PSErrors"
'-------------------------
' Write To File
' Implementation of Error Logging Function
'-------------------------
' Updated by hadi, 2 November 2002
' Single File for all problem
' Record Application (Component Name, Major, Minor and revision)
' Record Date And Time
' record Passed Error String


Dim lLogFile As String
Dim lLogPath As String


Public Property Let LogPath(ByVal vdata As String)
    lLogPath = vdata
End Property


Public Property Let LogFile(ByVal vdata As String)
    lLogFile = vdata
End Property


Public Sub WriteToFile(strError As String, Optional lAddStatement As String, Optional ModuleName As String)
On Error Resume Next

    Dim nFile As Integer
    Dim strFile As String
    Dim tdate
    Dim ttime
    Dim theapp
    Dim appfile
    
    '--------------------------
    ' If there is an error message given , record into the LOG file
    '--------------------------
    If Len(strError) > 0 Then
    
        theapp = "Application: " & App.Title & "(" & App.Major & "." & App.Minor & "." & App.Revision & ")"
        appfile = "File: " & App.Path & "\" & App.EXEName
        
        '--------------------------
        ' If given, record the module/method/function name
        '--------------------------
        If Not IsMissing(ModuleName) Then
            appfile = appfile & " (" & ModuleName & ")"
        End If
        
        nFile = FreeFile
        
        '--------------------------
        ' File Name (by default is [application path]\Log\[application title]_ErrorLog.Log)
        '--------------------------
        If Len(lLogFile) = 0 Then lLogFile = App.Title & "_ErrorLog.Log"
        If Len(lLogPath) = 0 Then lLogPath = App.Path & "\Log\"
        
        strFile = lLogPath & lLogFile
        
        '--------------------------
        ' Format the Date
        '--------------------------
        tdate = Format(Date, "DD/MM/YYYY")
        ttime = Format(Time, "HH:MM:SS:MS")
6
        Open strFile For Append As nFile
        Print #nFile, "------------------------Recorded at *" & tdate & " " & ttime & "*"
        Print #nFile, theapp
        Print #nFile, appfile
        
        Print #nFile, "Message:" & strError
        '--------------------------
        ' If given, record the additional statement (eg: sql statement, login statement)
        '--------------------------
        If Not IsMissing(lAddStatement) Then
            If lAddStatement <> "" Then
                Print #nFile, "Statement:" & lAddStatement
            End If
        End If
        Print #nFile, " "
        Close nFile
    End If

End Sub

Public Function ADOErrors(objConn As Object) As String
Dim objError        ' error object
Dim errString As String

If objConn.Errors.Count > 0 Then

'------------------------------
' Collect All Connection errors
'------------------------------
For Each objError In objConn.Errors
    errString = errString & "Error No: " & objError.Number & ";" & _
                            "NativeError: " & objError.NativeError & ";" & _
                            "SQLState:" & objError.NativeError & ";" & _
                            "Source:" & objError.Source & vbCrLf & _
                            "Description:" & objError.Description & vbCrLf

Next

'------------------------------
' Return the Errors
'------------------------------
ADOErrors = errString
    
End If

End Function


Public Function theErrors() As String

Dim errString As String

'------------------------------
' Collect All errors information
'------------------------------

    errString = errString & "Error No: " & Err.Number & ";" & _
                            "Source:" & Err.Source & vbCrLf & _
                            "Description:" & Err.Description & vbCrLf



'------------------------------
' Return the Errors
'------------------------------
theErrors = errString
    


End Function


Public Function FixSQL(ByVal pstr As String) As String
    FixSQL = Replace(pstr, "'", "''")
End Function

