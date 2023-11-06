VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form Frm_grille 
   Caption         =   "Grille"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "Frm_grille.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   WindowState     =   2  'Maximized
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   2655
      Left            =   240
      TabIndex        =   12
      Top             =   4800
      Width           =   12975
      _Version        =   393216
      _ExtentX        =   22886
      _ExtentY        =   4683
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "Frm_grille.frx":1272
   End
   Begin FPSpreadADO.fpSpread grid1 
      Height          =   3255
      Left            =   360
      TabIndex        =   11
      Top             =   8040
      Width           =   12735
      _Version        =   393216
      _ExtentX        =   22463
      _ExtentY        =   5741
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "Frm_grille.frx":1446
   End
   Begin FPSpreadADO.fpSpread grid 
      Height          =   3495
      Left            =   240
      TabIndex        =   10
      Top             =   360
      Width           =   12975
      _Version        =   393216
      _ExtentX        =   22886
      _ExtentY        =   6165
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SpreadDesigner  =   "Frm_grille.frx":161A
   End
   Begin VB.CommandButton CmdAnnuler2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Annuler"
      Height          =   780
      Left            =   18960
      Picture         =   "Frm_grille.frx":17EE
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton CmdEdit2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Edition"
      Height          =   780
      Left            =   17760
      Picture         =   "Frm_grille.frx":1DC2
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CheckBox ChkSelection2 
      Height          =   255
      Left            =   3120
      TabIndex        =   6
      Top             =   4320
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox ChkSelection 
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   0
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Edition"
      Height          =   780
      Left            =   16320
      Picture         =   "Frm_grille.frx":23CE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Ajouter"
      Height          =   780
      Left            =   15120
      Picture         =   "Frm_grille.frx":29DA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Fermer"
      Height          =   780
      Left            =   18960
      Picture         =   "Frm_grille.frx":2FA5
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnnuler 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Annuler"
      Height          =   780
      Left            =   17520
      Picture         =   "Frm_grille.frx":35DF
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Selection ligne-------------------"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   4320
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Selection ligne-------------------"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "Frm_grille"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'GRILLE
'-----
'-------------
'----------------'
Option Explicit

Dim Row_Actif As Integer
Dim Row_Actif2 As Integer

'Private Sub remplir_combo()
'Dim sql As String
'Dim rs As New Recordset
'
'CboTable.Clear
'
'sql = "select name from sys.sysobjects where xtype='U' order by name"
'rs.Open sql, db, adOpenStatic, adLockOptimistic
'While Not rs.EOF
'  CboTable.AddItem rs(0)
'  rs.MoveNext
'Wend
'rs.Close
'
'End Sub

Sub init_grid()
Dim i As Integer

grid.MaxCols = 7
grid.MaxRows = 0

i = 1

grid.Row = 0
grid.col = i
grid.Text = "ID"
grid.ColWidth(i) = 10
i = i + 1
grid.col = i
grid.Text = "Nom"
grid.ColWidth(i) = 20
i = i + 1
grid.col = i
grid.Text = "Table"
grid.ColWidth(i) = 20
i = i + 1
grid.col = i
grid.Text = "Key_ID"
grid.ColWidth(i) = 10
i = i + 1
grid.col = i
grid.Text = "Champs_ID"
grid.ColWidth(i) = 10
i = i + 1
grid.col = i
grid.Text = "COND_SOCIETE"
grid.ColWidth(i) = 10
i = i + 1
grid.col = i
grid.Text = "COND_PARTICUL"
grid.ColWidth(i) = 40

grid.OperationMode = 3

grid.UserColAction = UserColActionSort

End Sub

Sub init_grid2()
Dim i As Integer

grid2.MaxCols = 29
grid2.MaxRows = 0

i = 1
grid2.Row = 0

grid2.col = i
grid2.Text = "ID_GRID_COLONNE"
grid2.ColWidth(i) = 10
i = i + 1
grid2.col = i
grid2.Text = "Etat"
grid2.ColWidth(i) = 10
i = i + 1
grid2.col = i
grid2.Text = "Colonne"
grid2.ColWidth(i) = 30
i = i + 1
grid2.col = i
grid2.Text = "Nom"
grid2.ColWidth(i) = 30
i = i + 1
grid2.col = i
grid2.Text = "Actif"
grid2.ColWidth(i) = 10
i = i + 1
grid2.col = i
grid2.Text = "Align"
grid2.ColWidth(i) = 10
i = i + 1
grid2.col = i
grid2.Text = "Ordre"
grid2.ColWidth(i) = 10
i = i + 1
grid2.col = i
grid2.Text = "Visible"
grid2.ColWidth(i) = 10
i = i + 1
grid2.col = i
grid2.Text = "Readonly"
grid2.ColWidth(i) = 10
i = i + 1
grid2.col = i
grid2.Text = "Monnaie"
grid2.ColWidth(i) = 10
i = i + 1
grid2.col = i
grid2.Text = "Largeur"
grid2.ColWidth(i) = 10
i = i + 1
grid2.col = i
grid2.Text = "Type control"
grid2.ColWidth(i) = 15
i = i + 1
grid2.col = i
grid2.Text = "Liste col"
grid2.ColWidth(i) = 10
i = i + 1
grid2.col = i
grid2.Text = "Par défaut"
grid2.ColWidth(i) = 10
i = i + 1
grid2.col = i
grid2.Text = "Grid colonne"
grid2.ColWidth(i) = 10
i = i + 1
grid2.col = i
grid2.Text = "Primary Key"
grid2.ColWidth(i) = 10
i = i + 1
grid2.col = i
grid2.Text = "Foreign Key"
grid2.ColWidth(i) = 10
i = i + 1
grid2.col = i
grid2.Text = "Unique"
grid2.ColWidth(i) = 10
i = i + 1
grid2.col = i
grid2.Text = "CHECK"
grid2.ColWidth(i) = 10
i = i + 1
grid2.col = i
grid2.Text = "Position"
grid2.ColWidth(i) = 10
i = i + 1
grid2.col = i
grid2.Text = "Défaut"
grid2.ColWidth(i) = 10
i = i + 1
grid2.col = i
grid2.Text = "Type"
grid2.ColWidth(i) = 10
i = i + 1
grid2.col = i
grid2.Text = "max"
grid2.ColWidth(i) = 10
i = i + 1
grid2.col = i
grid2.Text = "Precision"
grid2.ColWidth(i) = 10
i = i + 1
grid2.col = i
grid2.Text = "Virgule"
grid2.ColWidth(i) = 10
i = i + 1
grid2.col = i
grid2.Text = "table Key"
grid2.ColWidth(i) = 10
i = i + 1
grid2.col = i
grid2.Text = "Champs Key"
grid2.ColWidth(4) = 10
i = i + 1
grid2.col = i
grid2.Text = "Check_Clause"
grid2.ColWidth(i) = 10
i = i + 1
grid2.col = i
grid2.Text = "Description"
grid2.ColWidth(i) = 10

grid2.OperationMode = 3

grid2.UserColAction = UserColActionSort

End Sub

Private Sub Afficher_grid()
Dim Sql As String
Dim i As Integer
Dim rs As New Recordset

grid.Visible = False
grid.MaxRows = 0

Sql = "select * from GRID order by Table_grid"

rs.Open Sql, db, adOpenKeyset, adLockOptimistic
If rs.RecordCount <> 0 Then
    grid.MaxRows = rs.RecordCount
    For i = 1 To rs.RecordCount
        grid.Row = i
        grid.col = 1
        grid.Lock = True
        grid.BackColor = &H8000000F
        grid.TypeHAlign = TypeHAlignCenter
        grid.Text = Str(rs("ID_GRID"))
        grid.col = 2
        grid.Lock = True
        grid.TypeMaxEditLen = 200
        grid.BackColor = &H8000000F
        grid.Text = rs("NOM_GRID")
        grid.col = 3
        grid.Lock = True
        grid.TypeMaxEditLen = 200
        grid.BackColor = &H8000000F
        grid.Text = rs("TABLE_GRID")
        grid.col = 4
        grid.Lock = True
        grid.BackColor = &H8000000F
        grid.Text = rs("KEY_ID_GRID")
        grid.col = 5
        grid.Lock = True
        grid.BackColor = &H8000000F
        grid.Text = IfNull(rs("CHAMPS_ID_GRID"), "")
        grid.col = 6
        grid.Lock = True
        grid.BackColor = &H8000000F
        grid.CellType = CellTypeCheckBox
        grid.TypeCheckCenter = True
        If IsNull(rs("CONDITION_SOCIETE_GRID")) Then
            grid.Text = 0
        Else
            grid.Text = rs("CONDITION_SOCIETE_GRID")
        End If
        grid.col = 7
        grid.Lock = True
        grid.BackColor = &H8000000F
        grid.Text = IfNull(rs("CONDITION_PARTICULIERE_GRID"), "")
        
        rs.MoveNext
    Next i
End If
grid.Visible = True
rs.Close

End Sub

Private Sub Afficher_grid2()
Dim Sql As String
Dim i As Integer
Dim j As Integer
Dim rs As New Recordset
Dim ID_GRID As Integer
Dim Table As String
Dim couleur

grid2.Visible = False
grid2.MaxRows = 0

grid.Row = grid.ActiveRow
grid.col = 1
ID_GRID = grid.Text
grid.Row = grid.ActiveRow
grid.col = 3
Table = grid.Text
If grid.Text <> "" Then
    Sql = "select * from VUE_GRID_COLONNE where ID_GRID=" & ID_GRID & " and [TABLE]='" & Table & "' order by position"
    
    rs.Open Sql, db, adOpenKeyset, adLockOptimistic
    If rs.RecordCount <> 0 Then
        grid2.MaxRows = rs.RecordCount
        For i = 1 To rs.RecordCount
            grid2.Row = i
            
            If IsNull(rs("GRID_ID")) Then
                couleur = &H8000000F
            Else
                If rs("ACTIF_COLONNE") = 0 Then
                    couleur = &HC0C0FF
                Else
                    couleur = &HFF00&
                End If
            End If
            j = 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.TypeHAlign = TypeHAlignCenter
            If IsNull(rs("ID_GRID_COLONNE")) Then
                grid2.Text = ""
            Else
                grid2.Text = Str(rs("ID_GRID_COLONNE"))
            End If
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.CellType = CellTypeCheckBox
            grid2.TypeCheckCenter = True
            If IsNull(rs("GRID_ID")) Then
                grid2.Text = 0
            Else
                grid2.Text = 1
            End If
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.Text = rs("CHAMPS")
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.Text = IfNull(rs("NOM_COLONNE"), "")
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.CellType = CellTypeCheckBox
            grid2.TypeCheckCenter = True
            If IsNull(rs("ACTIF_COLONNE")) Then
                grid2.Text = 0
            Else
                grid2.Text = rs("ACTIF_COLONNE")
            End If
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.TypeHAlign = TypeHAlignCenter
            If IsNull(rs("ALIGN_COLONNE")) Then
                grid2.Text = ""
            Else
                grid2.Text = rs("ALIGN_COLONNE")
            End If
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.TypeHAlign = TypeHAlignCenter
            If IsNull(rs("ORDRE_COLONNE")) Then
                grid2.Text = ""
            Else
                grid2.Text = Str(rs("ORDRE_COLONNE"))
            End If
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.CellType = CellTypeCheckBox
            grid2.TypeCheckCenter = True
            If IsNull(rs("VISIBLE_COLONNE")) Then
                grid2.Text = 0
            Else
                grid2.Text = rs("VISIBLE_COLONNE")
            End If
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.CellType = CellTypeCheckBox
            grid2.TypeCheckCenter = True
            If IsNull(rs("READONLY_COLONNE")) Then
                grid2.Text = 0
            Else
                grid2.Text = rs("READONLY_COLONNE")
            End If
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.CellType = CellTypeCheckBox
            grid2.TypeCheckCenter = True
            If IsNull(rs("MONNAIE_COLONNE")) Then
                grid2.Text = 0
            Else
                grid2.Text = rs("MONNAIE_COLONNE")
            End If
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.TypeHAlign = TypeHAlignCenter
            If IsNull(rs("LARGEUR_COLONNE")) Then
                grid2.Text = ""
            Else
                grid2.Text = Str(rs("LARGEUR_COLONNE"))
            End If
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.TypeHAlign = TypeHAlignCenter
            grid2.Text = IfNull(rs("TYPE_COLONNE"), "")
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.Text = IfNull(rs("LISTE_COLONNE"), "")
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.Text = IfNull(rs("VALEUR_DEFAUT_COLONNE"), "")
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.CellType = CellTypeCheckBox
            grid2.TypeCheckCenter = True
            If IsNull(rs("GRID_LISTE_COLONNE")) Then
                grid2.Text = 0
            Else
                grid2.Text = Str(rs("GRID_LISTE_COLONNE"))
            End If
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.CellType = CellTypeCheckBox
            grid2.TypeCheckCenter = True
            grid2.Text = rs("Primary Key")
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.CellType = CellTypeCheckBox
            grid2.TypeCheckCenter = True
            grid2.Text = rs("Foreign Key")
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.CellType = CellTypeCheckBox
            grid2.TypeCheckCenter = True
            grid2.Text = rs("Unique")
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.CellType = CellTypeCheckBox
            grid2.TypeCheckCenter = True
            grid2.Text = rs("Check")
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.TypeHAlign = TypeHAlignCenter
            grid2.BackColor = couleur
            grid2.Text = rs("POSITION")
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.Text = rs("défaut")
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.Text = rs("type")
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.Text = rs("max")
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.Text = rs("precision")
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.Text = rs("virgule")
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.Text = rs("table key")
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.Text = rs("champs key")
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.Text = rs("check_clause")
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.Text = rs("description")
            
            rs.MoveNext
        Next i
    End If
    rs.Close
End If

Call grid2.SetSelection(1, Row_Actif2, 1, Row_Actif2)

grid2.Visible = True

End Sub


Private Sub ChkSelection_Click()
If ChkSelection.Value = 1 Then
    grid.OperationMode = 3
    grid2.OperationMode = 3
Else
    grid.OperationMode = 0
    grid2.OperationMode = 0
End If

End Sub

Private Sub ChkSelection2_Click()
If ChkSelection2.Value = 1 Then
    grid2.OperationMode = 3
Else
    grid2.OperationMode = 0
End If

End Sub

Private Sub cmdAdd_Click()
baseG = ""
Frm_grille_maj.Show 1
Call Afficher_grid
Call Afficher_grid2
End Sub


Private Sub CmdAnnuler2_Click()
Dim Sql As String
Dim rep

'on error goto erreur

grid2.col = 1
grid2.Row = grid2.ActiveRow
If Not IsNumeric(grid2.Text) Then
    MsgBox "Cette colonne n'est pas visible !", vbExclamation
Else
    rep = MsgBox("Etes vous sûr d'annuler cette colonne ?", vbYesNo)
    If rep = 6 Then
        Sql = "delete GRID_COLONNE where ID_GRID_COLONNE=" & grid2.Text
        'db.Execute "insert   (SQL) values ('" & Replace(Sql, "'", "''") & "')"

        db.Execute Sql
        Call Afficher_grid2
    End If
End If
Exit Sub

erreur:
MsgBox Err.Description
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdEdit_Click()
grid.Row = grid.ActiveRow
grid.col = 1
If IsNumeric(grid.Text) Then
    baseG = grid.Text
    Frm_grille_maj.Show 1
    Call Afficher_grid
    Call grid.SetSelection(1, Row_Actif, 1, Row_Actif)
    Call Afficher_grid2
    Call grid2.SetSelection(1, Row_Actif2, 1, Row_Actif2)
End If
End Sub

Private Sub cmdEdit2_Click()
grid2.col = 1
grid2.Row = grid2.ActiveRow
If IsNumeric(grid2.Text) Then
    baseG = grid2.Text
    grid2.col = 2
    If grid2.Text = 1 Then
        grid2.col = 3
        champsG = grid2.Text
        grid2.col = 7
        ordreG = grid2.Text
        grid.Row = grid.ActiveRow
        grid.col = 1
        GridG = grid.Text
        Frm_grille_colonne_maj.Show 1
        Call Afficher_grid2
        Call grid2.SetSelection(1, Row_Actif2, 1, Row_Actif2)
    Else
        baseG = ""
        grid.Row = grid.ActiveRow
        grid.col = 1
        GridG = grid.Text
        Frm_grille_colonne_maj.Show 1
        Call Afficher_grid2
        Call grid2.SetSelection(1, Row_Actif2, 1, Row_Actif2)
    End If
Else
    baseG = ""
    grid.Row = grid.ActiveRow
    grid.col = 1
    GridG = grid.Text
    grid2.col = 3
    champsG = grid2.Text
    Frm_grille_colonne_maj.Show 1
    Call Afficher_grid2
End If
End Sub

Private Sub Form_Load()
grid.MaxRows = 0
grid2.MaxRows = 0
Call init_grid
Call init_grid2
Call Afficher_grid
Call Afficher_grid2
End Sub

Private Sub grid_Click(ByVal col As Long, ByVal Row As Long)
If Row_Actif <> Row Then
    Call Afficher_grid2
    Row_Actif = grid.ActiveRow
    Row_Actif2 = 1
End If

End Sub

Private Sub grid_KeyPress(KeyAscii As Integer)
If grid.ActiveRow <> Row_Actif Then
    Call Afficher_grid2
    Row_Actif = grid.ActiveRow
    Row_Actif2 = 1
End If

End Sub

Private Sub grid2_KeyPress(KeyAscii As Integer)
Row_Actif2 = grid2.ActiveRow
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
        Unload Me
End Select
End Sub

Private Sub Grid_dblClick(ByVal col As Long, ByVal Row As Long)
If Row > 0 Then
    Row_Actif = grid.ActiveRow
    Row_Actif2 = grid2.ActiveRow
    cmdEdit_Click
End If
End Sub

Private Sub Grid2_dblClick(ByVal col As Long, ByVal Row As Long)
If Row > 0 Then
    Row_Actif = grid.ActiveRow
    Row_Actif2 = grid2.ActiveRow
    cmdEdit2_Click
End If
End Sub

