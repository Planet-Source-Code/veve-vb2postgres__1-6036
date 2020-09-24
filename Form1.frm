VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Visual Basic and PG-SQL Interface"
   ClientHeight    =   1725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   ScaleHeight     =   1725
   ScaleWidth      =   3840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Modify Data"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create Data"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.Label Label1 
         Caption         =   "Waiting"
         Height          =   255
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Coded by Vasoo Veerapen
'and revised on the 13th of February 2000
'
'  ( Ever heard of the Island of Mauritius eh??? )
'  http://www.geocities.com/TheTropics/Shores/4216
'
'Many people have asked me how to access POSTGRES from
'VB so here goes...
'
'This project demonstrates the basics of using RDO connections
'with a POSTGRES ( PG-SQL ) server running on a LINUX box.

'I am NO LINUX expert, so I can't give much help if something
'goes wrong with LINUX. I have tested the code, though, and it
'worked fine for me ;-)

'This project has some pre-requisites although you can modify
'many of the connection parameters to suit your own...
'
'(1) You need a POSTGRES server installed on the LINUX box
'
'    I used POSTGRES version 6.5.3 available from
'    http://www.postgres.org
'
'(2) Create a database called 'thilo' on the LINUX box
'(3) Get a PG-SQL ODBC driver for this darned Windows machine.
'
'    You can get the latest driver from
'    http://www.insightdist.com/psqlodbc
'
'    The driver should work on 95/98 and NT
'
'(4) I think ??? you need VB Professional/Enterprise editions
'(5) You need to create a DSN for the ODBC connection !!!
'
' --- > OK lets roll em sleeves up and get to work...

Option Explicit

Dim cn As New rdoConnection

Private Sub Command1_Click()

  Dim SQL As String
  
  'Lets connect to the database
  Label1.Caption = "Connecting"
  Label1.Refresh
  With cn
    .Connect = "DSN=PostgreSQL;" _
     & "UID=;PWD=;DATABASE=thilo;"
    .LoginTimeout = 20
    .CursorDriver = rdUseClientBatch
    .EstablishConnection rdDriverNoPrompt
  End With
  While cn.StillConnecting
    DoEvents
  Wend
  
  'Wipe "testtable" if it exists
  Label1.Caption = "Wiping table"
  Label1.Refresh
  On Error GoTo DropErr
  SQL = "DROP TABLE testtable;"
  cn.Execute SQL
  On Error GoTo 0
  While cn.StillExecuting
    DoEvents
  Wend
  
  'Now create a table
  Label1.Caption = "Creating table"
  Label1.Refresh
  SQL = "CREATE TABLE testtable " _
    & "(FirstName varchar(25), LastName varchar(25));"
  cn.Execute SQL
  While cn.StillExecuting
    DoEvents
  Wend
  
  'Insert some data
  Label1.Caption = "Inserting data"
  Label1.Refresh
  SQL = "Insert INTO testtable " _
    & " Values ('Vasoo', 'Veerapen');"
  cn.Execute SQL, rdExecDirect
  While cn.StillExecuting
    DoEvents
  Wend
  
  Label1.Caption = "Data inserted"
  Label1.Refresh
  
  Exit Sub

DropErr:
  Debug.Print "Table not wiped!"
  Resume Next

End Sub

Private Sub Command2_Click()
 
  Dim qy As New rdoQuery
  Dim rs As rdoResultset
  
  'Read some data and lock it from other users...
  Label1.Caption = "Selecting data"
  Label1.Refresh
  
  With qy
     Set .ActiveConnection = cn
     .SQL = "SELECT Lastname, Firstname FROM testtable WHERE LastName='Veerapen';"
  End With
  Set rs = qy.OpenResultset(rdOpenKeyset, rdConcurLock)
  While cn.StillExecuting
    DoEvents
  Wend
  
  Label1.Caption = "Modifying data"
  Label1.Refresh
  
  'Modify the data
  rs("Lastname").SourceTable = "testtable"
  rs("Firstname").SourceTable = "testtable"
  rs("Lastname").KeyColumn = True
  
  rs.Edit
  rs!Firstname = "Alex"
  rs.Update
  Debug.Print rs!Lastname
  rs.Close
  qy.Close

  Label1.Caption = "Data modified"
  Label1.Refresh


End Sub

