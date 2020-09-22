VERSION 5.00
Begin VB.Form Viewer 
   Caption         =   "SQL Tool"
   ClientHeight    =   6975
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8565
   LinkTopic       =   "Form1"
   ScaleHeight     =   6975
   ScaleWidth      =   8565
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   7200
      Top             =   5640
   End
   Begin VB.CommandButton btnParse 
      Caption         =   "Parse Statement"
      Height          =   375
      Left            =   6360
      TabIndex        =   24
      Top             =   4920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton btnViewData 
      Caption         =   "View Populated Data"
      Height          =   375
      Left            =   6360
      TabIndex        =   23
      Top             =   4440
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton btnViewSchema 
      Caption         =   "View Schema"
      Height          =   375
      Left            =   6360
      TabIndex        =   22
      Top             =   2400
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton btnExit 
      Caption         =   "Close && Exit"
      Height          =   375
      Left            =   6360
      TabIndex        =   10
      Top             =   6480
      Width           =   2055
   End
   Begin VB.CommandButton btnReadData 
      Caption         =   "Populate Data"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000A&
      Caption         =   "Generate Statement"
      Height          =   2415
      Left            =   240
      TabIndex        =   5
      Top             =   3840
      Width           =   5895
      Begin VB.TextBox txtDataresults 
         Appearance      =   0  'Flat
         Height          =   1095
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   1080
         Width           =   5535
      End
      Begin VB.Label lblWork 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label lblStatement 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   7
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "Data Population:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Connection Properties"
      Height          =   1455
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   8175
      Begin VB.CommandButton btnListDatabases 
         Caption         =   "List Databases"
         Height          =   375
         Left            =   5520
         TabIndex        =   21
         Top             =   360
         Width           =   2535
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         Left            =   1560
         TabIndex        =   19
         Top             =   1080
         Width           =   1695
      End
      Begin VB.TextBox txtLogin 
         Height          =   285
         Left            =   1560
         TabIndex        =   17
         Text            =   "sa"
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtServer 
         Height          =   285
         Left            =   1560
         TabIndex        =   15
         Text            =   "(local)"
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox cboDatabases 
         Height          =   315
         Left            =   5520
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label6 
         Caption         =   "Password"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Login"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "SQL Server"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Select Database:"
         Height          =   255
         Left            =   3840
         TabIndex        =   14
         Top             =   960
         Width           =   1455
      End
   End
   Begin VB.CommandButton btnReadSchema 
      Caption         =   "Read Schema"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6360
      TabIndex        =   0
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Database Schema Reader"
      Height          =   1935
      Left            =   240
      TabIndex        =   1
      Top             =   1800
      Width           =   5895
      Begin VB.TextBox txtSchemaresults 
         Appearance      =   0  'Flat
         Height          =   1095
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   720
         Width           =   5535
      End
      Begin VB.Label Label1 
         Caption         =   "Reading database schema:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblDbSchema 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   2
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Label lblWaiting 
      Caption         =   "Now reading....."
      Height          =   255
      Left            =   6360
      TabIndex        =   25
      Top             =   5400
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      Index           =   1
      X1              =   240
      X2              =   8400
      Y1              =   6375
      Y2              =   6375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      Index           =   0
      X1              =   240
      X2              =   8400
      Y1              =   6360
      Y2              =   6360
   End
End
Attribute VB_Name = "Viewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mySQL As clsSQLBuilder
Attribute mySQL.VB_VarHelpID = -1
Dim schemafile As String
Dim datafile As String
Dim hcounter As Integer


Private Sub btnExit_Click()
End
End Sub

Private Sub btnListDatabases_Click()
On Error Resume Next
Screen.MousePointer = vbHourglass
Dim sqlsvr As SQLDMO.SQLServer
Set sqlsvr = New SQLDMO.SQLServer
Dim db As New SQLDMO.Database
sqlsvr.Connect ServerName, Login, Password

h = sqlsvr.ConnectionID

If h > 0 Then

For Each db In sqlsvr.Databases
    Me.cboDatabases.AddItem db.Name
Next

Me.cboDatabases.ListIndex = 0
Database = cboDatabases.Text

sqlsvr.Close
Set SQLServer = Nothing

If Database <> "Combo1" Then
    btnReadSchema.Enabled = True
End If

Else
    MsgBox "SQL Server not found"
End If

Screen.MousePointer = vbNormal

End Sub


' ------------------
' Transact SQL Parser using OSQL
' ------------------

' SQL parser is hard to find. If you know any parser component available..., email me.
' Ive tried so-called mssqlparser component but I couldnt figured out the results.
' I don't have time to invent one, so I decided to use OSQL utilities instead.
' At least, it get the jobs done.
' I just want to know if my generated script contains error.

Private Sub btnParse_Click()
Dim sqlfile As Object

Set sqlfile = CreateObject("vb6x.filetool")

Dim osql As String
Dim results As String
Dim z As Boolean
Dim l As Long


lblWaiting.Visible = True


osql = "OSQL -S [servername] -U [loginname] -P [password] -i [inputfile] -o [outputfile]"

osql = Replace(osql, "[servername]", ServerName)
osql = Replace(osql, "[loginname]", Login)
osql = Replace(osql, "[password]", Password)
osql = Replace(osql, "[inputfile]", datafile)
osql = Replace(osql, "[outputfile]", "Parsed.txt")

sqlfile.WriteFile osql, App.Path & "\executioner.bat", True
k = Val(lblWork.Caption) * 500

For l = 1 To k
Next

z = sqlfile.FileExist("Parsed.txt")


If z Then
Kill "Parsed.txt"
End If

z = False
 
Shell App.Path & "\executioner.bat", vbHide


Set sqlfile = Nothing
Timer1.Enabled = True

End Sub

Private Sub btnReadData_Click()

If Database = "Combo1" Then
    MsgBox "Please select a valid database!"
Exit Sub
End If


    time1 = Time
    Set mySQL = New clsSQLBuilder
    
    datafile = "Scripts\" & Format(Time, "HHMMSSMS") & ".txt"
    
    mySQL.GenerateSQLInsert schemafile, datafile
    
    
    
    time2 = Time
    
    Time3 = DateDiff("s", time1, time2)
    
    lblStatement.Caption = "Data populated in " & Time3 & " seconds"
    txtDataresults.Text = "Data was saved in " & datafile
    btnViewData.Visible = True
    btnParse.Visible = True
    
    Set mySQL = Nothing

End Sub

Private Sub btnReadSchema_Click()
If Database = "Combo1" Then
    MsgBox "Please select a valid database!"
Exit Sub
End If



    time1 = Time
    schemafile = App.Path & "\dbschema.xml"
    
    Set mySQL = New clsSQLBuilder
    
    ' Generate The Database Schema
    mySQL.GenerateDBSchema schemafile
    
    time2 = Time
    Time3 = DateDiff("s", time1, time2)
    
    lblDbSchema.Caption = "Data populated in " & Time3 & " seconds"
    txtSchemaresults.Text = "Schema was saved in " & schemafile
    
    ' Enable view button
    btnViewSchema.Visible = True
    btnReadData.Enabled = True
    Set mySQL = Nothing
    

End Sub

Private Sub btnViewData_Click()
Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE " & App.Path & "\" & datafile, vbMaximizedFocus
End Sub

Private Sub btnViewSchema_Click()
Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE " & schemafile, vbMaximizedFocus
End Sub

Private Sub cboDatabases_Click()
Database = cboDatabases.Text
btnViewSchema.Visible = False
btnReadData.Enabled = False
btnViewData.Visible = False
btnParse.Visible = False

lblDbSchema = ""
txtSchemaresults = ""
lblWork = ""
lblStatement = ""
txtDataresults = ""



End Sub

Private Sub Form_Load()
    Login = txtLogin
    Password = txtPassword
    Database = cboDatabases.Text
    ServerName = txtServer
    


              
              
End Sub

Private Sub Form_Unload(Cancel As Integer)
End

End Sub

Private Sub mySQL_Percentage1(ByVal Percent As Long)
lblDbSchema.Caption = Percent & "%"
DoEvents
End Sub

Private Sub mySQL_Percentage2(ByVal Percent As Long, ByVal strText As String)
lblStatement.Caption = Percent & "%"
lblWork.Caption = strText
DoEvents
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim sqlfile As Object
Set sqlfile = CreateObject("VB6x.FIletool")

    results = sqlfile.ReadFile("Parsed.txt")
    
    If results = txtDataresults.Text Then
    hcounter = hcounter + 1
        If hcounter > 10 Then
            Timer1.Enabled = False
            lblWaiting.Visible = False
        
            Exit Sub
        End If
    End If
    
    
    txtDataresults.Text = results
    DoEvents

Set sqlfile = Nothing

End Sub

Private Sub txtLogin_Change()
Login = txtLogin.Text

End Sub

Private Sub txtPassword_Change()
Password = txtPassword.Text
End Sub

Private Sub txtServer_Change()
ServerName = txtServer.Text
End Sub
