VERSION 5.00
Begin VB.Form Frm_grille_maj 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MAJ Grille"
   ClientHeight    =   6885
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9645
   Icon            =   "Frm_grille_maj.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   9645
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   9495
      Begin VB.CommandButton CMDQUERY 
         Caption         =   "QUERY"
         Height          =   735
         Left            =   240
         TabIndex        =   18
         Top             =   3840
         Width           =   2295
      End
      Begin VB.TextBox TXTSQL 
         Height          =   1815
         Left            =   240
         TabIndex        =   17
         Top             =   4800
         Width           =   9015
      End
      Begin VB.CheckBox ChkFields 
         DataField       =   "CONDITION_SOCIETE_GRID"
         Height          =   255
         Index           =   6
         Left            =   2520
         TabIndex        =   14
         Top             =   2640
         Width           =   375
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00FFFFFF&
         DataField       =   "CONDITION_PARTICULIERE_GRID"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0,000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   2520
         TabIndex        =   13
         Top             =   3000
         Width           =   4815
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00FFFFFF&
         DataField       =   "NOM_GRID"
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0,000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1036
            SubFormatType   =   0
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   2520
         TabIndex        =   0
         Top             =   720
         Width           =   4815
      End
      Begin VB.ComboBox CboFields 
         DataField       =   "CHAMPS_ID_GRID"
         Height          =   315
         Index           =   3
         Left            =   2520
         TabIndex        =   3
         Text            =   "CboFields"
         Top             =   2160
         Width           =   4815
      End
      Begin VB.ComboBox CboFields 
         DataField       =   "KEY_ID_GRID"
         Height          =   315
         Index           =   2
         Left            =   2520
         TabIndex        =   2
         Text            =   "CboFields"
         Top             =   1680
         Width           =   4815
      End
      Begin VB.ComboBox CboFields 
         DataField       =   "TABLE_GRID"
         Height          =   315
         Index           =   1
         Left            =   2520
         TabIndex        =   1
         Text            =   "CboFields"
         Top             =   1200
         Width           =   4815
      End
      Begin VB.CommandButton CmdCancel 
         BackColor       =   &H00C0E0FF&
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   3960
         Width           =   975
      End
      Begin VB.CommandButton CmdOK 
         BackColor       =   &H00C0FFC0&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7200
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3960
         Width           =   855
      End
      Begin VB.Label lblLabels 
         Caption         =   "Condition particulière:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   16
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Condition société:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nom:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "ID CHAMPS:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "ID KEY:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Table:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label LblFields 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   0
         Left            =   2520
         TabIndex        =   8
         Top             =   240
         Width           =   4815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Code:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Frm_grille_maj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Table = "GRID"

Private Sub Cbofields_Click(Index As Integer)

If Index = 1 Then
    Call remplir_combo1
End If

End Sub

Private Sub CmdCancel_Click()
Unload Me
End Sub

Private Sub CmdOK_Click()
Dim Sql As String

''on error goto erreur

If baseG = "" Then
    If Insert_TB(Me, Table, 1, 1, 1, 0, 0, 0, "", 1) Then
        Unload Me
    End If
Else
    If Update_TB(Me, Table, 1, 1, 1, 0, 0, " ID_GRID='" & baseG & "' ", baseG, 1) Then
        Unload Me
    End If
End If

Exit Sub

erreur:
MsgBox Err.Description

End Sub


Private Sub CMDQUERY_Click()
Dim Sql As String

''on error goto erreur

If baseG = "" Then
    Call Insert_TB(Me, Table, 1, 1, 1, 0, 0, 0, "", 0)
Else
    Call Update_TB(Me, Table, 1, 1, 1, 0, 0, " ID_GRID='" & baseG & "' ", baseG, 0)
End If

Exit Sub

erreur:
MsgBox Err.Description

End Sub

Private Sub Form_Load()
Dim Sql As String
Dim rs As New Recordset

Call remplir_combo

If baseG <> "" Then
    Sql = "select * from GRID where ID_GRID='" & baseG & "' "
    rs.Open Sql, db, adOpenStatic, adLockOptimistic
    LblFields(0).Caption = baseG
    CboFields(1).Text = rs("TABLE_GRID")
    CboFields(1).Enabled = False
    CboFields(2).Text = rs("KEY_ID_GRID")
    CboFields(3).Text = IfNull(rs("CHAMPS_ID_GRID"), "")
    txtFields(4).Text = rs("NOM_GRID")
    If rs("CONDITION_SOCIETE_GRID") = True Then
        ChkFields(6).Value = 1
    Else
        ChkFields(6).Value = 0
    End If
    txtFields(5).Text = IfNull(rs("CONDITION_PARTICULIERE_GRID"), "")
    rs.Close
End If

End Sub

Private Sub Form_Activate()
txtFields(4).SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
        Unload Me
End Select
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If

End Sub

Private Sub remplir_combo()
Dim Sql As String
Dim rs As New Recordset

CboFields(1).Clear

Sql = "select name from sys.sysobjects where xtype='U' or xtype='V' order by name"
rs.Open Sql, db, adOpenStatic, adLockOptimistic
While Not rs.EOF
  CboFields(1).AddItem rs(0)
  rs.MoveNext
Wend
rs.Close

End Sub

Private Sub remplir_combo1()
Dim Sql As String
Dim rs As New Recordset

CboFields(2).Clear
CboFields(3).Clear
If CboFields(1).Text <> "" Then
    CboFields(3).AddItem ""
    Sql = "select CHAMPS from VUE_TABLES where [TABLE]='" & CboFields(1).Text & "' "
    rs.Open Sql, db, adOpenStatic, adLockOptimistic
    While Not rs.EOF
        CboFields(2).AddItem rs(0)
        CboFields(3).AddItem rs(0)
        rs.MoveNext
    Wend
    rs.Close
End If

End Sub

