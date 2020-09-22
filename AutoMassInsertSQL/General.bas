Attribute VB_Name = "General"





Public ServerName As String
Public Login As String
Public Password As String
Public Database As String

Public txtFile As String


Public Function ConnStr() As String
    ConnStr = "Provider=SQLOLEDB.1;Persist Security Info=False;" & _
              "User ID=" & Login & _
              ";Initial Catalog=" & Database & _
              ";Data Source=" & ServerName & _
              ";Password=" & Password
End Function



Public Function FixSQL(ByVal pstr As String) As String
    FixSQL = Replace(pstr, "'", "''")
End Function

Public Sub WriteText(ByVal fname As String, ByVal fContent As String)
    Dim nFile As Integer
    Dim strFile As String
    
    nFile = FreeFile
    strFile = fname
    Open strFile For Append As nFile
    Print #nFile, fContent
    Close nFile
End Sub
