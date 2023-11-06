VERSION 5.00
Begin VB.Form Frm_transfert_document_maj2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transfert entête"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7275
   Icon            =   "Frm_transfert_document_maj2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7275
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7215
      Begin VB.TextBox txtFields 
         BackColor       =   &H00FFFFFF&
         DataField       =   "NOM_TRANSFERT"
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
         Index           =   0
         Left            =   2520
         TabIndex        =   0
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox txtID 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         DataField       =   "GRID_ID"
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   4215
      End
      Begin VB.ComboBox CboFields 
         DataField       =   "CHAMPS_SOURCE"
         Height          =   315
         Index           =   1
         Left            =   2520
         TabIndex        =   1
         Text            =   "CboFields"
         Top             =   1080
         Width           =   4215
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
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2640
         Width           =   855
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
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2640
         Width           =   975
      End
      Begin VB.ComboBox CboFields 
         DataField       =   "CHAMPS_DESTINATION"
         Height          =   315
         Index           =   2
         Left            =   2520
         TabIndex        =   2
         Text            =   "CboFields"
         Top             =   1560
         Width           =   4215
      End
      Begin VB.Label lblLabels 
         Caption         =   "Champs source:"
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
         TabIndex        =   9
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nom transfert:"
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
         TabIndex        =   8
         Top             =   600
         Width           =   1815
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
      Begin VB.Label lblLabels 
         Caption         =   "Champs destination:"
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
         TabIndex        =   6
         Top             =   1560
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Frm_transfert_document_maj2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const Table = "TRANSFERT_DOCUMENT_CONFIG_CHAMPS"

Private Sub CmdCancel_Click()
Unload Me
End Sub

Private Sub CmdOK_Click()
Dim Sql As String

''on error goto erreur

If ID_transfert_ChampsG = "" Then
    If Insert_TB(Me, Table, 1, 0, 1, 0, 0, 0, "", 1) Then
        Unload Me
    End If
Else
    If Update_TB(Me, Table, 1, 0, 1, 0, 0, " ID_TRANSFERT_CHAMPS='" & ID_transfert_ChampsG & "' ", ID_transfert_ChampsG, 1) Then
        Unload Me
    End If
End If

Exit Sub

erreur:
MsgBox Err.Description

End Sub


Private Sub Form_Load()
Dim Sql As String
Dim rs As New Recordset

Call remplir_combo

If ID_transfert_ChampsG <> "" Then
    Sql = "select CHAMPS_SOURCE,CHAMPS_DESTINATION from TRANSFERT_DOCUMENT_CONFIG_CHAMPS where ID_TRANSFERT_CHAMPS='" & ID_transfert_ChampsG & "' "
    rs.Open Sql, db, adOpenStatic, adLockOptimistic
    txtID.Text = ID_transfert_ChampsG
    txtFields(0).Text = NumeroG
    CboFields(1).Text = rs("CHAMPS_SOURCE")
    CboFields(2).Text = rs("CHAMPS_DESTINATION")
    rs.Close
Else
    txtID.Text = ""
    txtFields(0).Text = NumeroG
End If

End Sub

Private Sub Form_Activate()
CboFields(1).SetFocus
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
Dim i As Integer
Dim Table1 As String
Dim Table2 As String

If Typ_tableG = "E" Then
    Sql = "select TABLE_TD from TYPE_DOCUMENT where ID_TD='" & Type_documentG1 & "' "
    rs.Open Sql, db, adOpenStatic, adLockOptimistic
    Table1 = rs(0)
    rs.Close
    Sql = "select TABLE_TD from TYPE_DOCUMENT where ID_TD='" & Type_documentG2 & "' "
    rs.Open Sql, db, adOpenStatic, adLockOptimistic
    Table2 = rs(0)
    rs.Close
Else
    Sql = "select TABLE_LIGNE_TD from TYPE_DOCUMENT where ID_TD='" & Type_documentG1 & "' "
    rs.Open Sql, db, adOpenStatic, adLockOptimistic
    Table1 = rs(0)
    rs.Close
    Sql = "select TABLE_LIGNE_TD from TYPE_DOCUMENT where ID_TD='" & Type_documentG2 & "' "
    rs.Open Sql, db, adOpenStatic, adLockOptimistic
    Table2 = rs(0)
    rs.Close
End If

Sql = "select CHAMPS from VUE_TABLES V " _
& " where [TABLE]='" & Table1 & "' " _
& " and not exists(select 1 from TRANSFERT_DOCUMENT_CONFIG_CHAMPS T where T.CHAMPS_SOURCE=V.Champs  and T.TRANSFERT_ID=" & NumeroG & ")"

rs.Open Sql, db, adOpenStatic, adLockOptimistic
CboFields(1).Clear
While Not rs.EOF
    CboFields(1).AddItem Trim$(rs(0))
    rs.MoveNext
Wend
rs.Close


Sql = "select CHAMPS from VUE_TABLES V " _
& " where [TABLE]='" & Table2 & "' " _
& " and not exists(select 1 from TRANSFERT_DOCUMENT_CONFIG_CHAMPS T where T.CHAMPS_DESTINATION=V.Champs  and T.TRANSFERT_ID=" & NumeroG & ")"

rs.Open Sql, db, adOpenStatic, adLockOptimistic
CboFields(2).Clear
While Not rs.EOF
    CboFields(2).AddItem Trim$(rs(0))
    rs.MoveNext
Wend
rs.Close


End Sub


