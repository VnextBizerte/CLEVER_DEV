VERSION 5.00
Begin VB.Form Frm_grille_colonne_maj 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MAJ Colonnes"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9510
   Icon            =   "Frm_grille_colonne_maj.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   9510
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   9255
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9495
      Begin VB.TextBox txtFields 
         BackColor       =   &H00FFFFFF&
         DataField       =   "LONGUEUR_TYPE_COLONNE"
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
         Index           =   23
         Left            =   7680
         TabIndex        =   45
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "VISIBLE_DEFAUT_COLONNE"
         Height          =   195
         Index           =   22
         Left            =   2520
         TabIndex        =   43
         Top             =   7680
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "ACCEPTER_NULL_COLONNE"
         Height          =   195
         Index           =   21
         Left            =   2520
         TabIndex        =   42
         Top             =   7320
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "ZERO_VIDE_COLONNE"
         Height          =   195
         Index           =   20
         Left            =   2520
         TabIndex        =   40
         Top             =   6960
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "TRONQUER_ZERO_COLONNE"
         Height          =   195
         Index           =   19
         Left            =   2520
         TabIndex        =   38
         Top             =   6600
         Width           =   255
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00FFFFFF&
         DataField       =   "CHECK_COLONNE"
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
         Index           =   18
         Left            =   2520
         TabIndex        =   35
         Top             =   6240
         Width           =   1335
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "CALCULE_COLONNE"
         Height          =   195
         Index           =   17
         Left            =   2520
         TabIndex        =   33
         Top             =   5880
         Width           =   255
      End
      Begin VB.TextBox TXTSQL 
         Height          =   735
         Left            =   480
         TabIndex        =   32
         Text            =   "Text1"
         Top             =   8400
         Width           =   8775
      End
      Begin VB.CommandButton CMDQUERY 
         Caption         =   "QUERY"
         Height          =   855
         Left            =   5160
         TabIndex        =   31
         Top             =   5280
         Width           =   1815
      End
      Begin VB.ComboBox CboFields 
         DataField       =   "ALIGN_COLONNE"
         Height          =   315
         Index           =   16
         Left            =   2520
         TabIndex        =   30
         Text            =   "Combo1"
         Top             =   5400
         Width           =   1695
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "GRID_LISTE_COLONNE"
         Height          =   195
         Index           =   15
         Left            =   2520
         TabIndex        =   28
         Top             =   3840
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "MONNAIE_COLONNE"
         Height          =   195
         Index           =   14
         Left            =   2520
         TabIndex        =   27
         Top             =   5040
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "ACTIF_COLONNE"
         Height          =   195
         Index           =   4
         Left            =   8520
         TabIndex        =   9
         Top             =   7920
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "READONLY_COLONNE"
         Height          =   195
         Index           =   13
         Left            =   2520
         TabIndex        =   8
         Top             =   4680
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "VISIBLE_COLONNE"
         Height          =   195
         Index           =   12
         Left            =   2520
         TabIndex        =   7
         Top             =   4320
         Value           =   1  'Checked
         Width           =   255
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00FFFFFF&
         DataField       =   "LARGEUR_COLONNE"
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
         Index           =   11
         Left            =   2520
         TabIndex        =   4
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00FFFFFF&
         DataField       =   "CHAMPS_COLONNE"
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
         TabIndex        =   0
         Top             =   600
         Width           =   4815
      End
      Begin VB.TextBox txtFields 
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
         Index           =   0
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00FFFFFF&
         DataField       =   "VALEUR_DEFAUT_COLONNE"
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
         Top             =   3360
         Width           =   6615
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00FFFFFF&
         DataField       =   "LISTE_COLONNE"
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
         Index           =   6
         Left            =   2520
         TabIndex        =   5
         Top             =   3000
         Width           =   6615
      End
      Begin VB.ComboBox CboFields 
         DataField       =   "ORDRE_COLONNE"
         Height          =   315
         Index           =   3
         Left            =   2520
         TabIndex        =   2
         Text            =   "CboFields"
         Top             =   1680
         Width           =   1575
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00FFFFFF&
         DataField       =   "NOM_COLONNE"
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
         TabIndex        =   1
         Top             =   960
         Width           =   4815
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
         TabIndex        =   10
         Top             =   5400
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
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   5400
         Width           =   975
      End
      Begin VB.ComboBox CboFields 
         DataField       =   "TYPE_COLONNE"
         Height          =   315
         Index           =   5
         Left            =   2520
         TabIndex        =   3
         Text            =   "CboFields"
         Top             =   2160
         Width           =   2775
      End
      Begin VB.Label lblLabels 
         Caption         =   "Longueur type:"
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
         Index           =   20
         Left            =   5880
         TabIndex        =   46
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Visible défaut :"
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
         Index           =   19
         Left            =   240
         TabIndex        =   44
         Top             =   7680
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Accepter NULL :"
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
         Index           =   18
         Left            =   240
         TabIndex        =   41
         Top             =   7320
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Zéro vide :"
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
         Index           =   17
         Left            =   240
         TabIndex        =   39
         Top             =   6960
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tronquer zéro :"
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
         Index           =   16
         Left            =   240
         TabIndex        =   37
         Top             =   6600
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Check :"
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
         Index           =   15
         Left            =   240
         TabIndex        =   36
         Top             =   6240
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Calculer:"
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
         Index           =   14
         Left            =   240
         TabIndex        =   34
         Top             =   5880
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Alignement:"
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
         Index           =   13
         Left            =   240
         TabIndex        =   29
         Top             =   5400
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
         Index           =   9
         Left            =   240
         TabIndex        =   26
         Top             =   5040
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Actif:"
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
         Left            =   6240
         TabIndex        =   25
         Top             =   7965
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Readonly:"
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
         TabIndex        =   24
         Top             =   4680
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Visible:"
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
         TabIndex        =   23
         Top             =   4320
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Largeur:"
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
         TabIndex        =   22
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Grid liste:"
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
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Ordre:"
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
         TabIndex        =   19
         Top             =   1680
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
         TabIndex        =   18
         Top             =   960
         Width           =   1935
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
         Caption         =   "Valeur par défaut:"
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
         TabIndex        =   15
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Liste:"
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
         Top             =   3000
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Type contrôle:"
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
         Top             =   2160
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Frm_grille_colonne_maj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Table = "GRID_COLONNE"

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
    If Update_TB(Me, Table, 1, 1, 1, 0, 0, " ID_GRID_COLONNE='" & baseG & "' ", baseG, 1) Then
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
    Call Update_TB(Me, Table, 1, 1, 1, 0, 0, " ID_GRID_COLONNE='" & baseG & "' ", baseG, 0)
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
    Sql = "select * from GRID_COLONNE where ID_GRID_COLONNE='" & baseG & "' "
    rs.Open Sql, db, adOpenStatic, adLockOptimistic
    txtFields(1).Text = champsG
    txtFields(2).Text = rs("NOM_COLONNE")
    CboFields(3).Text = rs("ORDRE_COLONNE")
    If rs("ACTIF_COLONNE") Then
        ChkFields(4).Value = 1
    Else
        ChkFields(4).Value = 0
    End If
    CboFields(5).Text = IfNull(rs("TYPE_COLONNE"), "")
    txtFields(6).Text = IfNull(rs("LISTE_COLONNE"), "")
    txtFields(7).Text = IfNull(rs("VALEUR_DEFAUT_COLONNE"), "")
    txtFields(0).Text = GridG
    txtFields(11).Text = rs("LARGEUR_COLONNE")
    
    If rs("VISIBLE_COLONNE") Then
        ChkFields(12).Value = 1
    Else
        ChkFields(12).Value = 0
    End If
    If rs("READONLY_COLONNE") Then
        ChkFields(13).Value = 1
    Else
        ChkFields(13).Value = 0
    End If
    
    If rs("MONNAIE_COLONNE") Then
        ChkFields(14).Value = 1
    Else
        ChkFields(14).Value = 0
    End If
    
    If rs("GRID_LISTE_COLONNE") Then
        ChkFields(15).Value = 1
    Else
        ChkFields(15).Value = 0
    End If
    If rs("CALCULE_COLONNE") Then
        ChkFields(17).Value = 1
    Else
        ChkFields(17).Value = 0
    End If
        
    txtFields(18).Text = IfNull(rs("CHECK_COLONNE"), "")
    
    If rs("TRONQUER_ZERO_COLONNE") Then
        ChkFields(19).Value = 1
    Else
        ChkFields(19).Value = 0
    End If
    If rs("ZERO_VIDE_COLONNE") Then
        ChkFields(20).Value = 1
    Else
        ChkFields(20).Value = 0
    End If
    If rs("ACCEPTER_NULL_COLONNE") Then
        ChkFields(21).Value = 1
    Else
        ChkFields(21).Value = 0
    End If
    If rs("VISIBLE_DEFAUT_COLONNE") Then
        ChkFields(22).Value = 1
    Else
        ChkFields(22).Value = 0
    End If
    
    txtFields(23).Text = IfNull(rs("LONGUEUR_TYPE_COLONNE"), "")
    
    CboFields(16).Text = rs("ALIGN_COLONNE")
    rs.Close
Else
    txtFields(1).Text = champsG
    txtFields(2).Text = champsG
    ChkFields(4).Value = 1
    txtFields(0).Text = GridG
End If

End Sub

Private Sub Form_Activate()
txtFields(2).SetFocus
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


CboFields(5).Clear

CboFields(5).AddItem "BIT"
CboFields(5).AddItem "CUSTOM"
CboFields(5).AddItem "DATE"
CboFields(5).AddItem "DATETIME"
CboFields(5).AddItem "NUMERIC"
CboFields(5).AddItem "POURCENTAGE"
CboFields(5).AddItem "QUANTITE"
CboFields(5).AddItem "TEXT"


Sql = "declare @i int " _
& " declare @sql varchar(MAX)" _
& " set @sql=';WITH CTE as (select 1 as Ordre union ' "
Sql = Sql & " set @i=2 " _
& " while @i<100 " _
& " begin " _
& " set @sql=@sql+'select '+cast(@i as varchar(3))+' union ' " _
& " set @i=@i+1 " _
& " End" _
& " set @sql=@sql+' select 100)' " _
& " set @sql=@sql+ ' select * from CTE C where not exists (select 1 from GRID_COLONNE G where G.GRID_ID=" & GridG & " and C.ordre=G.Ordre_colonne)' " _
& " Select @sql "
rs.Open Sql, db, adOpenStatic, adLockOptimistic

If baseG <> "" Then
    Sql = rs(0) & " union select " & ordreG & " order by ordre "
Else
    Sql = rs(0) & " order by ordre "
End If
rs.Close
rs.Open Sql, db, adOpenStatic, adLockOptimistic
CboFields(3).Clear
While Not rs.EOF
    CboFields(3).AddItem Trim$(Str(rs(0)))
    rs.MoveNext
Wend
rs.Close

CboFields(16).Clear
CboFields(16).AddItem "L"
CboFields(16).AddItem "C"
CboFields(16).AddItem "R"
CboFields(16).Text = "L"

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


