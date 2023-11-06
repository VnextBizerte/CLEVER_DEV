VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1095
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1215
      Left            =   1320
      TabIndex        =   0
      Top             =   840
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'9999'
'Badii
Private Sub Command1_Click()
Dim db2 As New Connection
Dim Sql As String
Dim rs As New Recordset
Dim mstream As ADODB.Stream

With db2
    .CursorLocation = adUseClient
    .ConnectionTimeout = 20
    .CommandTimeout = 200
    .Open "Provider=sqloledb;User ID=sa ;password=$$vnext2020 ;Data Source=VS026 ;initial catalog =IQURA "
End With

Sql = "select ID_ETUDIANT,[image] from ETUDIANT where [image] is not null"
rs.Open Sql, db2, adOpenStatic, adLockOptimistic
While Not rs.EOF
    Set mstream = New Stream
    mstream.Type = adTypeBinary
    mstream.Open
    mstream.Write rs("image")
    mstream.SaveToFile "E:\ETUDIANTS\IQURA\" & rs("ID_ETUDIANT") & ".PNG", adSaveCreateOverWrite
    rs.MoveNext
Wend

rs.Close
db2.Close
End Sub

Public Sub Load_image(F As Form, code_stagiaire As String)
Dim Sql As String
Dim rs As New Recordset
Dim mstream As ADODB.Stream
    
Sql = "select image from stagiaire where code=" & code_stagiaire
rs.Open Sql, db, adOpenStatic, adLockOptimistic
If Not rs.EOF Then
    If Not IsNull(rs("image")) Then
        Set mstream = New Stream
        mstream.Type = adTypeBinary
        mstream.Open
        mstream.Write rs("image")
        mstream.SaveToFile App.Path & "\images\tmp.bmp", adSaveCreateOverWrite
        F.Img.Picture = LoadPicture(App.Path & "\images\tmp.bmp")
    Else
        F.Img.Picture = LoadPicture(App.Path & "\Images\personne.bmp")
    End If
Else
    F.Img.Picture = LoadPicture(App.Path & "\Images\personne.bmp")
End If
rs.Close
End Sub

Private Sub Command2_Click()
Dim db2 As New Connection

Dim Sql As String
Dim rs As New Recordset


With db2
    .CursorLocation = adUseClient
    .ConnectionTimeout = 20
    .CommandTimeout = 200
    '.Open "Provider=sqloledb;User ID=sa ;password=NeverForget ;Data Source=VW002\DEV ;initial catalog =CLEVER "
    '.Open "Provider=sqloledb;User ID=sa ;password=$$vnext2020 ;Data Source=VS026 ;initial catalog =CLEVER_EVAL "
    .Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=e:\;Extended Properties=DBASE IV"
End With

Sql = "select * from ARTICLE.DBF "
rs.Open Sql, db2, adOpenStatic, adLockOptimistic
rs.Close
End Sub
