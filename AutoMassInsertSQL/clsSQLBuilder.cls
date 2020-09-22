VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSQLBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------
' Automated SQL MASS Insert Statement Generator
' Implementation of Automated SQL Insert Statement Generator (with VB6)
' By Adhadi Mohd, Kuala Lumpur.
' Adhadi@hotpop.com
' 22 Nov 2002
'
'
' The reason of doing is because I need to install my application together
' with database (schema & lookup data) at the client site.
' Of course we can do this using different approach like BACKUP/RESTORE,
' or access upsizing. In my case, I need to script all the schema and
' the data to have a better control to the setup file
' Previously, i wrote the lookup data manually.
' Luckilly, Josh Carderonello's article gave me the basic idea on how to perform the automation.
' MS Project Server 2002 installation also doing the database installation
' using the same approach
'
' Original Idea:
'       From article:
'       How to Write SQL to Dynamically Script Mass INSERT Statement Scripts
'       http://www.sql-server-performance.com/jc_write_sql_script.asp
'       by Josh Calderonello
'
'
' Requirements
' a. VB6X.DLL (if it not included here, you can find this component at planetsourcecode)
' b. SQLDMO components (installed with SQL7, SQL2000 or MSDE)
' c. MSXML 3.0
'
' Issuess
' a. Can't extract IMAGE data (Image)
' b. Identity data field was skipped. (It's my requirement)


Public Event Percentage1(ByVal Percent As Long)
Public Event Percentage2(ByVal Percent As Long, ByVal strText As String)

Dim sqlsvr As SQLDMO.SQLServer
Private Enum SQLInstructor
    IntructSelect = 0
    InstructInsert = 1
    instructUpdate = 2
    intructWhere = 3
    instructGetDatatype = 4
End Enum



Private Sub ConnectSvr()
'------------
' Open SQL Server Connection
'------------
Set sqlsvr = New SQLDMO.SQLServer
sqlsvr.Connect ServerName, Login, Password
End Sub



Private Function TableList(ByVal prmDbName As String) As String
Dim tblXML As String
Dim tbls As New SQLDMO.Table
Dim cols As New SQLDMO.Column
Dim tblCount As Integer
Dim totalTable As Integer
Dim blnCancel As Boolean




    
    
    tblXML = "<DATABASE DBNAME=""" & prmDbName & """ DATE=""" & Now & """>"
    
    '------------
    ' Loop for each tables
    '------------
    totalTable = sqlsvr.Databases(prmDbName).Tables.Count
    
    
    tt = 0
    tblCount = 0
    For Each tbls In sqlsvr.Databases(prmDbName).Tables
    tt = tt + 1
    RaiseEvent Percentage1((tt / totalTable) * 100)
    
    If Not tbls.SystemObject Then
    tblCount = tblCount + 1

    tblXML = tblXML & "<TABLE "
        
        '------------
        ' Get Table Information
        '------------
        tblXML = tblXML & " ID=""" & tbls.ID & """"
        tblXML = tblXML & " NAME=""" & tbls.Name & """>"
            
        '------------
        ' Get Columns Information
        ' Loop for each columns
        '------------
        For Each cols In tbls.Columns
            tblXML = tblXML & "<COLUMN  "
            tblXML = tblXML & " ID=""" & cols.ID & """"
            tblXML = tblXML & " NAME=""" & cols.Name & """"
            tblXML = tblXML & " DATATYPE=""" & cols.Datatype & """"
            tblXML = tblXML & " LENGTH=""" & cols.Length & """"
            tblXML = tblXML & " IDENTITY=""" & cols.Identity & """/>"
        Next
        
    tblXML = tblXML & "</TABLE>"
    End If
    
    
    Next
    tblXML = tblXML & "<COUNT TABLE=""" & tblCount & """/>"
    tblXML = tblXML & "</DATABASE>"
    
    TableList = tblXML

Set cols = Nothing
Set tbls = Nothing
End Function


Private Sub disconnectsvr()
'------------
' Close the SQL server connection
'------------
sqlsvr.Close
Set sqlsvr = Nothing
End Sub


Public Sub GenerateDBSchema(ByVal strSchema As String)
Dim ltblList As String
Dim ftxt As Object

Set ftxt = CreateObject("vb6x.filetool")

ConnectSvr
    ltblList = TableList(Database)
disconnectsvr

ftxt.WriteFile ltblList, strSchema, True

Set ftxt = Nothing

End Sub

Public Function GenerateSQLInsert(ByVal strSchema As String, ByVal targetFile As String) As String


Dim xmlDoc As New MSXML2.DOMDocument30
Dim objNodeList As IXMLDOMNodeList
Dim tblNode As IXMLDOMNode
Dim tblChild As IXMLDOMNode

Dim arrCols As Variant
Dim arrConnector
Dim ado As New ADODB.Recordset
Dim conn As New ADODB.Connection
  
Dim connector As String
Dim col As Variant

xmlDoc.async = False
xmlDoc.Load (strSchema)

'--------------
' Get the list of all tables in the XML data
'--------------
Set objNodeList = xmlDoc.getElementsByTagName("TABLE")
'MsgBox objNodeList.Length
allout = 0
WriteText targetFile, "SET PARSEONLY ON"  ' Use for Parsing the file"
WriteText targetFile, "USE " & Database
WriteText targetFile, " "
For M = 0 To (objNodeList.Length - 1)
  Set tblNode = objNodeList.nextNode
  tbl = tblNode.Attributes(1).Text
  
  '''''''''''''''''
  'If UCase(tbl) = "CATEGORIES" Then
  If 1 = 1 Then
  
  col = ""
  connector = tbl
  If tblNode.hasChildNodes Then
  k = 0
  
  For Each child In tblNode.childNodes
      ' skip the Identity data
      If UCase(child.Attributes(4).Text) <> "TRUE" Then
        k = k + 1
        If k = 1 Then jjoin = "" Else jjoin = ","
        ljoin = ", "
         ' Get the Connector
        col = col & jjoin & "[" & FormatColumn(child.Attributes(1).Text, child.Attributes(2).Text, , child.Attributes(4).Text, InstructInsert) & "]"
        connector = connector & ljoin & FormatColumn(child.Attributes(1).Text, child.Attributes(2).Text, , child.Attributes(4).Text, instructGetDatatype)
    End If
  Next
  End If
  

  ' i want to get all data insert statement dynamically



  arrConnector = Split(connector, ",")
  
  sql = "SELECT " & col & " from [" & arrConnector(0) & "]"
  txtwork = "Reading table " & arrConnector(0) & "...."
  If conn.State = adStateClosed Then
    conn.Open ConnStr
  End If
  conn.CursorLocation = adUseClient
  Set ado = conn.Execute(sql)
  
  totalrecord = ado.RecordCount
  
  If Not ado.EOF Then
      tt = 0
      WriteText targetFile, " "
      WriteText targetFile, "-- Inserting into Table " & arrConnector(0)
      WriteText targetFile, "SELECT 'Now Inserting Table " & arrConnector(0) & "'"
      While Not ado.EOF
      tt = tt + 1
      allout = allout + 1
      
      i = 0
      datstr = ""
      
        RaiseEvent Percentage2((tt / totalrecord) * 100, txtwork)
        
        For Each fld In ado.Fields
          i = i + 1
          If i = 1 Then ljoin = "" Else ljoin = ","
          tdat = LTrim(RTrim(ado(i - 1)))
          
          
          ' Fix integer null value
          If IsNull(tdat) And Len(Trim(arrConnector(i))) = 0 Then tdat = "NULL"
          
          If Trim(arrConnector(i)) = "-" Then
            tdat = ""    ' Handle Image DataType
            arrConnector(i) = "'"
          End If
          
          If Not IsNull(tdat) Then
            tdat = FixSQL(tdat)
          End If
          
          datstr = datstr & ljoin & Trim(arrConnector(i)) & tdat & Trim(arrConnector(i))
        Next
    
      
      WriteText targetFile, "INSERT INTO [" & arrConnector(0) & "] (" & col & ") VALUES (" & datstr & ") "
      
      ado.MoveNext
      Wend
  End If
  Set ado = Nothing
  
  '''''''''''''
  End If
  
  
  Next
        RaiseEvent Percentage2(0, allout & " Insert code generated!")
If conn.State = adStateOpen Then
    conn.Close
End If

Set conn = Nothing

Set tblNode = Nothing
Set objNodeList = Nothing


End Function

Private Function FormatColumn(colName As String, _
                              Optional colDataType As String, _
                              Optional colLength As String, _
                              Optional colIdentity As String, _
                              Optional Instructor As SQLInstructor) As String
                              

If Instructor = IntructSelect Then
    FormatColumn = colName
End If

If Instructor = InstructInsert Then
    FormatColumn = colName
End If

If Instructor = instructGetDatatype Then
    TD = UCase(colDataType)
    
    If UCase(colIdentity) = "TRUE" Then
        TD = "*"
    End If
    
        ' Detect the Datatype and return a suitable quote
        If InStr(1, TD, "INT") Or InStr(1, TD, "NUMERIC") Or InStr(1, TD, "MONEY") Then
            'numeric based datatye dos not use any quote
            TD = ""
        ElseIf InStr(1, TD, "CHAR") Or InStr(1, TD, "BIT") Or InStr(1, TD, "TEXT") Or InStr(1, TD, "DATETIME") Or InStr(1, TD, "UNIQUEIDENTIFIER") Or InStr(1, TD, "SYSNAME") Then
            ' Quote for character based data type
            TD = "'"
        ElseIf InStr(1, TD, "IMAGE") Then
            TD = "-"
        Else
            TD = "'"
        End If
    

    FormatColumn = TD

End If

End Function


Private Function FixSQL(thisstring As Variant)
    If Not IsNull(thisstring) Then
    FixSQL = Replace(thisstring, "'", "''")
    End If
End Function



