VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form Frm_page 
   Caption         =   "Page"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "Frm_page.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   2415
      Left            =   240
      TabIndex        =   13
      Top             =   8160
      Width           =   11895
      _Version        =   393216
      _ExtentX        =   20981
      _ExtentY        =   4260
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
      SpreadDesigner  =   "Frm_page.frx":1272
   End
   Begin FPSpreadADO.fpSpread grid1 
      Height          =   3495
      Left            =   120
      TabIndex        =   12
      Top             =   4440
      Width           =   12015
      _Version        =   393216
      _ExtentX        =   21193
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
      SpreadDesigner  =   "Frm_page.frx":1446
   End
   Begin FPSpreadADO.fpSpread grid 
      Height          =   3135
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   11655
      _Version        =   393216
      _ExtentX        =   20558
      _ExtentY        =   5530
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
      SpreadDesigner  =   "Frm_page.frx":161A
   End
   Begin VB.CommandButton cmdEdit1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Edition"
      Height          =   780
      Left            =   17640
      Picture         =   "Frm_page.frx":17EE
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Ajouter"
      Height          =   780
      Left            =   16440
      Picture         =   "Frm_page.frx":1DFA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdAnnuler1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Annuler"
      Height          =   780
      Left            =   18840
      Picture         =   "Frm_page.frx":23C5
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CheckBox ChkSelection2 
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   3960
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CheckBox ChkSelection 
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CommandButton cmdAnnuler 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Annuler"
      Height          =   780
      Left            =   17640
      Picture         =   "Frm_page.frx":2999
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Fermer"
      Height          =   780
      Left            =   18960
      Picture         =   "Frm_page.frx":2F6D
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Ajouter"
      Height          =   780
      Left            =   15240
      Picture         =   "Frm_page.frx":35A7
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Edition"
      Height          =   780
      Left            =   16440
      Picture         =   "Frm_page.frx":3B72
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
      Left            =   240
      TabIndex        =   7
      Top             =   3960
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
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Frm_page"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Row_Actif As Integer
Dim Row_Actif1 As Integer
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

grid.MaxCols = 12
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
grid.Text = "Table Update"
grid.ColWidth(i) = 20
i = i + 1
grid.col = i
grid.Text = "Grid"
grid.ColWidth(i) = 10
i = i + 1
grid.col = i
grid.Text = "Nom Grid"
grid.ColWidth(i) = 10
i = i + 1
grid.col = i
grid.Text = "Table Grid"
grid.ColWidth(i) = 10
i = i + 1
grid.col = i
grid.Text = "Key Grid"
grid.ColWidth(i) = 10
i = i + 1
grid.col = i
grid.Text = "Champs Grid"
grid.ColWidth(i) = 10
i = i + 1
grid.col = i
grid.Text = "Grid Rechercher"
grid.ColWidth(i) = 10
i = i + 1
grid.col = i
grid.Text = "Grid Lister"
grid.ColWidth(i) = 10
i = i + 1
grid.col = i
grid.Text = "Actif"
grid.ColWidth(i) = 10
grid.OperationMode = 3

grid.UserColAction = UserColActionSort

End Sub

Sub init_grid1()
Dim i As Integer

grid1.MaxCols = 11
grid1.MaxRows = 0

i = 1
grid1.Row = 0
grid1.col = i
grid1.Text = "ID"
grid1.ColWidth(i) = 10
i = i + 1
grid1.col = i
grid1.Text = "ID Page"
grid1.ColWidth(i) = 10
i = i + 1
grid1.col = i
grid1.Text = "ID Design"
grid1.ColWidth(i) = 10
i = i + 1
grid1.col = i
grid1.Text = "Champs"
grid1.ColWidth(i) = 20
i = i + 1
grid1.col = i
grid1.Text = "Caché"
grid1.ColWidth(i) = 10
i = i + 1
grid1.col = i
grid1.Text = "Calculé"
grid1.ColWidth(i) = 10
i = i + 1
grid1.col = i
grid1.Text = "Modif Insert"
grid1.ColWidth(i) = 10
i = i + 1
grid1.col = i
grid1.Text = "Modif Update"
grid1.ColWidth(i) = 10
i = i + 1
grid1.col = i
grid1.Text = "Défaut"
grid1.ColWidth(i) = 10
i = i + 1
grid1.col = i
grid1.Text = "Liste"
grid1.ColWidth(i) = 10
i = i + 1
grid1.col = i
grid1.Text = "VIDE"
grid1.ColWidth(i) = 10

grid1.OperationMode = 3

grid1.UserColAction = UserColActionSort

End Sub

Sub init_grid2()
Dim i As Integer

grid2.MaxCols = 24
grid2.MaxRows = 0

i = 1
grid2.Row = 0

grid2.col = i
grid2.Text = "ID_GRID_COLONNE"
grid2.ColWidth(i) = 10
i = i + 1
grid2.col = i
grid2.Text = "Visible"
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
grid2.Text = "Ordre"
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
Dim j As Integer
Dim rs As New Recordset
Dim couleur

grid.Visible = False
grid.MaxRows = 0

Sql = "select * from [VUE_PAGE] order by NOM_PAGE "

rs.Open Sql, db, adOpenKeyset, adLockOptimistic
If rs.RecordCount <> 0 Then
    grid.MaxRows = rs.RecordCount
    For i = 1 To rs.RecordCount
        If rs("ACTIF_PAGE") Then
            couleur = &H8000000F
        Else
            couleur = &HC0E0FF
        End If
        j = 1
        grid.Row = i
        grid.col = j
        grid.Lock = True
        grid.BackColor = couleur
        grid.TypeHAlign = TypeHAlignCenter
        grid.Text = Str(rs("ID_PAGE"))
        j = j + 1
        grid.col = j
        grid.Lock = True
        grid.BackColor = couleur
        grid.Text = rs("NOM_PAGE")
        j = j + 1
        grid.col = j
        grid.Lock = True
        grid.TypeMaxEditLen = 200
        grid.BackColor = couleur
        grid.Text = rs("TABLE_PAGE")
        j = j + 1
        grid.col = j
        grid.Lock = True
        grid.TypeMaxEditLen = 200
        grid.BackColor = couleur
        grid.Text = IfNull(rs("TABLE_PAGE_UPDATE"), "")
        j = j + 1
        grid.col = j
        grid.Lock = True
        grid.BackColor = couleur
        If IsNull(rs("GRID_PAGE_ID")) Then
            grid.Text = ""
        Else
            grid.Text = Str(rs("GRID_PAGE_ID"))
        End If
        j = j + 1
        grid.col = j
        grid.Lock = True
        grid.TypeMaxEditLen = 200
        grid.BackColor = couleur
        grid.Text = IfNull(rs("NOM_GRID"), "")
        j = j + 1
        grid.col = j
        grid.Lock = True
        grid.TypeMaxEditLen = 200
        grid.BackColor = couleur
        grid.Text = IfNull(rs("TABLE_GRID"), "")
        j = j + 1
        grid.col = j
        grid.Lock = True
        grid.TypeMaxEditLen = 200
        grid.BackColor = couleur
        grid.Text = IfNull(rs("KEY_ID_GRID"), "")
        j = j + 1
        grid.col = j
        grid.Lock = True
        grid.TypeMaxEditLen = 200
        grid.BackColor = couleur
        grid.Text = IfNull(rs("CHAMPS_ID_GRID"), "")
        j = j + 1
        grid.col = j
        grid.Lock = True
        grid.BackColor = couleur
        If IsNull(rs("GRID_PAGE_RECHERCHER_ID")) Then
            grid.Text = ""
        Else
            grid.Text = Str(rs("GRID_PAGE_RECHERCHER_ID"))
        End If
        j = j + 1
        grid.col = j
        grid.Lock = True
        grid.BackColor = couleur
        If IsNull(rs("GRID_PAGE_LISTER_ID")) Then
            grid.Text = ""
        Else
            grid.Text = Str(rs("GRID_PAGE_LISTER_ID"))
        End If
        j = j + 1
        grid.col = j
        grid.Lock = True
        grid.BackColor = couleur
        grid.CellType = CellTypeCheckBox
        grid.TypeCheckCenter = True
        grid.Text = rs("ACTIF_PAGE")
        
        rs.MoveNext
    Next i
End If

Call grid.SetSelection(1, Row_Actif, 1, Row_Actif)

grid.Visible = True

rs.Close

End Sub

Private Sub Afficher_grid1()
Dim Sql As String
Dim i As Integer
Dim j As Integer
Dim rs As New Recordset
Dim couleur

grid1.Visible = False
grid1.MaxRows = 0

grid.Row = grid.ActiveRow
grid.col = 1
If Trim$(grid.Text) <> "" Then
    Sql = "select * from CONTROLE where PAGE_ID='" & grid.Text & "' "
    
    rs.Open Sql, db, adOpenKeyset, adLockOptimistic
    If rs.RecordCount <> 0 Then
        grid1.MaxRows = rs.RecordCount
        For i = 1 To rs.RecordCount
            grid1.Row = i
            
            couleur = &HFF00&
            j = 1
            grid1.col = j
            grid1.Lock = True
            grid1.BackColor = couleur
            grid1.TypeHAlign = TypeHAlignCenter
            grid1.Text = Str(rs("ID_CTRL"))
            j = j + 1
            grid1.col = j
            grid1.Lock = True
            grid1.BackColor = couleur
            grid1.Text = Str(rs("PAGE_ID"))
            j = j + 1
            grid1.col = j
            grid1.Lock = True
            grid1.BackColor = couleur
            grid1.Text = IfNull(rs("ID_DESIGN_CTRL"), "")
            j = j + 1
            grid1.col = j
            grid1.Lock = True
            grid1.BackColor = couleur
            grid1.Text = rs("CHAMPS_CTRL")
            j = j + 1
            grid1.col = j
            grid1.Lock = True
            grid1.BackColor = couleur
            grid1.CellType = CellTypeCheckBox
            grid1.TypeCheckCenter = True
            grid1.Text = rs("CHAMPS_CACHE_CTRL")
            j = j + 1
            grid1.col = j
            grid1.Lock = True
            grid1.BackColor = couleur
            grid1.CellType = CellTypeCheckBox
            grid1.TypeCheckCenter = True
            grid1.Text = rs("CHAMPS_CALCULE_CTRL")
            j = j + 1
            grid1.col = j
            grid1.Lock = True
            grid1.BackColor = couleur
            grid1.CellType = CellTypeCheckBox
            grid1.TypeCheckCenter = True
            grid1.Text = rs("MODIFIABLE_INSERT_CTRL")
            j = j + 1
            grid1.col = j
            grid1.Lock = True
            grid1.BackColor = couleur
            grid1.CellType = CellTypeCheckBox
            grid1.TypeCheckCenter = True
            grid1.Text = rs("MODIFIABLE_UPDATE_CTRL")
            j = j + 1
            grid1.col = j
            grid1.Lock = True
            grid1.BackColor = couleur
            grid1.Text = IfNull(rs("VALEUR_DEFAUT_CTRL"), "")
            j = j + 1
            grid1.col = j
            grid1.Lock = True
            grid1.BackColor = couleur
            grid1.TypeHAlign = TypeHAlignCenter
            grid1.Text = IfNull(rs("LISTE_DDL"), "")
            j = j + 1
            grid1.col = j
            grid1.Lock = True
            grid1.BackColor = couleur
            grid1.Text = IfNull(rs("VALEUR_VIDE_CTRL"), "")
            rs.MoveNext
        Next i
    End If
    rs.Close
End If

Call grid1.SetSelection(1, Row_Actif1, 1, Row_Actif1)

grid1.Visible = True

End Sub

Private Sub Afficher_grid2()
Dim Sql As String
Dim i As Integer
Dim j As Integer
Dim rs As New Recordset
Dim couleur

grid2.Visible = False
grid2.MaxRows = 0

grid.Row = grid.ActiveRow
grid.col = 5
If Trim$(grid.Text) <> "" Then
    Sql = "select * from VUE_GRID_COLONNE where ID_GRID='" & grid.Text & "' order by position"
    
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
            If IsNull(rs("ORDRE_COLONNE")) Then
                grid2.Text = ""
            Else
                grid2.Text = Str(rs("ORDRE_COLONNE"))
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
            grid2.TypeHAlign = TypeHAlignCenter
            If IsNull(rs("GRID_LISTE_COLONNE")) Then
                grid2.Text = ""
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
    grid1.OperationMode = 3
    grid2.OperationMode = 3
Else
    grid.OperationMode = 0
    grid1.OperationMode = 0
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
Frm_page_maj.Show 1
Call Afficher_grid
Call Afficher_grid1
Call Afficher_grid2
End Sub

Private Sub CmdAnnuler2_Click()
Dim Sql As String
Dim rep

'on error goto erreur

grid2.col = 1
grid2.Row = grid.ActiveRow
If Not IsNumeric(grid2.Text) Then
    MsgBox "Cette colonne n'est pas visible !", vbExclamation
Else
    rep = MsgBox("Etes vous sûr d'annuler cette colonne ?", vbYesNo)
    If rep = 6 Then
        Sql = "delete GRID_COLONNE where ID_GRID_COLONNE=" & grid2.Text
        db.Execute Sql
        Call Afficher_grid1
        Call Afficher_grid2
    End If
End If
Exit Sub

erreur:
MsgBox Err.Description
End Sub

Private Sub cmdAdd1_Click()
baseG = ""
grid.Row = grid.ActiveRow
grid.col = 1
pageG = grid.Text
grid.col = 3
tableG = grid.Text

Frm_control_maj.Show 1

Call Afficher_grid
Call Afficher_grid1
Call Afficher_grid2
End Sub

Private Sub cmdAnnuler1_Click()
Dim Sql As String
Dim rep

'on error goto erreur

grid1.col = 1
grid1.Row = grid1.ActiveRow
    rep = MsgBox("Etes vous sûr d'annuler ce controle ?", vbYesNo)
    If rep = 6 Then
        Sql = "delete CONTROLE where ID_CTRL=" & grid1.Text
        db.Execute Sql
        db.Execute "insert ZZ_QUERY (SQL) values ('" & Replace(Sql, "'", "''") & "')"
        Call Afficher_grid1
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
    Frm_page_maj.Show 1
    Call Afficher_grid
    Call grid.SetSelection(1, Row_Actif, 1, Row_Actif)
    Call Afficher_grid1
    Call grid1.SetSelection(1, Row_Actif1, 1, Row_Actif1)
    Call Afficher_grid2
    Call grid2.SetSelection(1, Row_Actif2, 1, Row_Actif2)
End If
End Sub

Private Sub cmdEdit1_Click()
grid1.Row = grid1.ActiveRow
grid1.col = 1
baseG = grid1.Text
grid1.col = 4
champsG = grid1.Text

grid.Row = grid.ActiveRow
grid.col = 1
pageG = grid.Text
grid.col = 3
tableG = grid.Text

Frm_control_maj.Show 1
Call Afficher_grid
Call grid.SetSelection(1, Row_Actif, 1, Row_Actif)
Call Afficher_grid1
Call grid1.SetSelection(1, Row_Actif1, 1, Row_Actif1)
Call Afficher_grid2
Call grid2.SetSelection(1, Row_Actif2, 1, Row_Actif2)

End Sub

Private Sub Form_Load()
grid.MaxRows = 0
grid1.MaxRows = 0
grid2.MaxRows = 0
Call init_grid
Call init_grid1
Call init_grid2
Call Afficher_grid
Call Afficher_grid1
Call Afficher_grid2
End Sub

Private Sub grid_Click(ByVal col As Long, ByVal Row As Long)
If Row_Actif <> Row Then
    Call Afficher_grid1
    Call Afficher_grid2
    Row_Actif = grid.ActiveRow
    Row_Actif1 = 1
    Row_Actif2 = 1
End If
End Sub

Private Sub grid_KeyPress(KeyAscii As Integer)
If grid.ActiveRow <> Row_Actif Then
    Call Afficher_grid1
    Call Afficher_grid2
    Row_Actif = grid.ActiveRow
    Row_Actif1 = 1
    Row_Actif2 = 1
End If
End Sub

Private Sub grid1_Click(ByVal col As Long, ByVal Row As Long)
Row_Actif1 = grid1.ActiveRow
End Sub

Private Sub grid1_KeyPress(KeyAscii As Integer)
Row_Actif1 = grid1.ActiveRow
End Sub

Private Sub grid2_Click(ByVal col As Long, ByVal Row As Long)
Row_Actif2 = grid2.ActiveRow
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
    Row_Actif1 = grid1.ActiveRow
    Row_Actif2 = grid2.ActiveRow
    cmdEdit_Click
End If
End Sub

Private Sub Grid1_dblClick(ByVal col As Long, ByVal Row As Long)
If Row > 0 Then
    Row_Actif = grid.ActiveRow
    Row_Actif1 = grid1.ActiveRow
    Row_Actif2 = grid2.ActiveRow
    cmdEdit1_Click
End If
End Sub

