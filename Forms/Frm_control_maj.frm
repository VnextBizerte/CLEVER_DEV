VERSION 5.00
Begin VB.Form Frm_control_maj 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controls"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9585
   Icon            =   "Frm_control_maj.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   9585
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   7095
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   9495
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "TRONQUER_ZERO_CTRL"
         Height          =   195
         Index           =   11
         Left            =   2520
         TabIndex        =   29
         Top             =   4485
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "MONNAIE_CTRL"
         Height          =   195
         Index           =   10
         Left            =   2520
         TabIndex        =   28
         Top             =   4125
         Width           =   255
      End
      Begin VB.CommandButton CMDQUERY 
         Caption         =   "QUERY"
         Height          =   735
         Left            =   4080
         TabIndex        =   27
         Top             =   4800
         Width           =   2535
      End
      Begin VB.TextBox TXTSQL 
         Height          =   1215
         Left            =   240
         TabIndex        =   26
         Text            =   "Text1"
         Top             =   5760
         Width           =   9015
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00FFFFFF&
         DataField       =   "VALEUR_VIDE_CTRL"
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
         Index           =   9
         Left            =   2520
         TabIndex        =   25
         Top             =   3720
         Width           =   6615
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "CHAMPS_CACHE_CTRL"
         Height          =   195
         Index           =   9
         Left            =   2520
         TabIndex        =   2
         Top             =   1800
         Width           =   255
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00FFFFFF&
         DataField       =   "LISTE_DDL"
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
         Index           =   8
         Left            =   2520
         TabIndex        =   7
         Top             =   3360
         Width           =   6615
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00FFFFFF&
         DataField       =   "VALEUR_DEFAUT_CTRL"
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
         Index           =   7
         Left            =   2520
         TabIndex        =   6
         Top             =   3000
         Width           =   6615
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00FFFFFF&
         DataField       =   "ID_DESIGN_CTRL"
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
         Index           =   2
         Left            =   2520
         TabIndex        =   0
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00FFFFFF&
         DataField       =   "PAGE_ID"
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
         Index           =   1
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   2415
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "CHAMPS_CALCULE_CTRL"
         Height          =   195
         Index           =   4
         Left            =   2520
         TabIndex        =   3
         Top             =   2040
         Width           =   255
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
         Left            =   7080
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4800
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
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   4800
         Width           =   975
      End
      Begin VB.ComboBox CboFields 
         DataField       =   "CHAMPS_CTRL"
         Height          =   315
         Index           =   3
         Left            =   2520
         TabIndex        =   1
         Text            =   "CboFields"
         Top             =   1440
         Width           =   4815
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "MODIFIABLE_INSERT_CTRL"
         Height          =   195
         Index           =   5
         Left            =   2520
         TabIndex        =   4
         Top             =   2445
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "MODIFIABLE_UPDATE_CTRL"
         Height          =   195
         Index           =   6
         Left            =   2520
         TabIndex        =   5
         Top             =   2685
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tronquer zéro:"
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
         Index           =   12
         Left            =   240
         TabIndex        =   31
         Top             =   4440
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Monnaie:"
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
         Index           =   11
         Left            =   240
         TabIndex        =   30
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Valeur VIDE:"
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
         Index           =   10
         Left            =   240
         TabIndex        =   24
         Top             =   3720
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Caché:"
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
         Index           =   9
         Left            =   240
         TabIndex        =   23
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label LblFields 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   9
         Left            =   5280
         TabIndex        =   22
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Liste DDL:"
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
         Index           =   8
         Left            =   240
         TabIndex        =   21
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Défaut:"
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
         TabIndex        =   20
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "ID design:"
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
         TabIndex        =   19
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Calculé:"
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
         Index           =   7
         Left            =   240
         TabIndex        =   18
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Page:"
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
         TabIndex        =   17
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
         TabIndex        =   16
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Champs:"
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
         TabIndex        =   15
         Top             =   1440
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Actif insert:"
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
         TabIndex        =   14
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Actif update:"
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
         TabIndex        =   13
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label LblFields 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   0
         Left            =   2520
         TabIndex        =   12
         Top             =   240
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Frm_control_maj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Table = "CONTROLE"

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
    If Update_TB(Me, Table, 1, 1, 1, 0, 0, " ID_CTRL='" & baseG & "' ", baseG, 1) Then
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
    Call Update_TB(Me, Table, 1, 1, 1, 0, 0, " ID_CTRL='" & baseG & "' ", baseG, 0)
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
    LblFields(0).Caption = baseG
    txtFields(1).Text = pageG
    LblFields(9).Caption = tableG
    
    Sql = "select * from CONTROLE where ID_CTRL='" & baseG & "' "
    rs.Open Sql, db, adOpenStatic, adLockOptimistic
    txtFields(2).Text = IfNull(rs("ID_DESIGN_CTRL"), "")
    CboFields(3).Text = rs("CHAMPS_CTRL")
    If rs("CHAMPS_CALCULE_CTRL") Then
        ChkFields(4).Value = 1
    Else
        ChkFields(4).Value = 0
    End If
    If rs("CHAMPS_CACHE_CTRL") Then
        ChkFields(9).Value = 1
    Else
        ChkFields(9).Value = 0
    End If
    If rs("MODIFIABLE_INSERT_CTRL") Then
        ChkFields(5).Value = 1
    Else
        ChkFields(5).Value = 0
    End If
    If rs("MODIFIABLE_UPDATE_CTRL") Then
        ChkFields(6).Value = 1
    Else
        ChkFields(6).Value = 0
    End If
    txtFields(7).Text = IfNull(rs("VALEUR_DEFAUT_CTRL"), "")
    txtFields(8).Text = IfNull(rs("LISTE_DDL"), "")
    txtFields(9).Text = IfNull(rs("VALEUR_VIDE_CTRL"), "")
    If rs("MONNAIE_CTRL") Then
        ChkFields(10).Value = 1
    Else
        ChkFields(10).Value = 0
    End If
    If rs("TRONQUER_ZERO_CTRL") Then
        ChkFields(11).Value = 1
    Else
        ChkFields(11).Value = 0
    End If

Else
    LblFields(0).Caption = baseG
    txtFields(1).Text = pageG
    LblFields(9).Caption = tableG
End If

End Sub

Private Sub Form_Activate()
txtFields(1).SetFocus
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

CboFields(3).Clear
If champsG <> "" Then CboFields(3).AddItem champsG

Sql = "select Champs from VUE_TABLES V where [TABLE]='" & tableG & "' and not exists (select 1 from CONTROLE C where C.CHAMPS_CTRL=V.CHAMPS and PAGE_ID=" & pageG & " ) order by position"
rs.Open Sql, db, adOpenStatic, adLockOptimistic
While Not rs.EOF
  CboFields(3).AddItem rs(0)
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




