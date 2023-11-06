VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#6.0#0"; "fpSpr60.ocx"
Begin VB.Form Frm_transfert_document 
   Caption         =   "Page"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "Frm_transfert_document.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin FPSpreadADO.fpSpread fpSpread1 
      Height          =   2415
      Left            =   10080
      TabIndex        =   13
      Top             =   4560
      Width           =   4815
      _Version        =   393216
      _ExtentX        =   8493
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
      SpreadDesigner  =   "Frm_transfert_document.frx":1272
   End
   Begin FPSpreadADO.fpSpread grid1 
      Height          =   2895
      Left            =   240
      TabIndex        =   12
      Top             =   4560
      Width           =   8895
      _Version        =   393216
      _ExtentX        =   15690
      _ExtentY        =   5106
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
      SpreadDesigner  =   "Frm_transfert_document.frx":1446
   End
   Begin FPSpreadADO.fpSpread grid 
      Height          =   2775
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   12135
      _Version        =   393216
      _ExtentX        =   21405
      _ExtentY        =   4895
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
      SpreadDesigner  =   "Frm_transfert_document.frx":161A
   End
   Begin VB.CommandButton cmdAjouter2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ajouter"
      Height          =   780
      Left            =   18960
      Picture         =   "Frm_transfert_document.frx":17EE
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdAjouter 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Ajouter"
      Height          =   780
      Left            =   6840
      Picture         =   "Frm_transfert_document.frx":1DFA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton cmdSupprimer 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Supprimer"
      Height          =   780
      Left            =   8040
      Picture         =   "Frm_transfert_document.frx":23C5
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CheckBox ChkSelection 
      Height          =   255
      Left            =   3120
      TabIndex        =   4
      Top             =   120
      Value           =   1  'Checked
      Width           =   255
   End
   Begin VB.CommandButton cmdSupprimer2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Supprimer"
      Height          =   780
      Left            =   18960
      Picture         =   "Frm_transfert_document.frx":2999
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Fermer"
      Height          =   780
      Left            =   18960
      Picture         =   "Frm_transfert_document.frx":2F6D
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Ajouter"
      Height          =   780
      Left            =   15240
      Picture         =   "Frm_transfert_document.frx":35A7
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdEdit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Edition"
      Height          =   780
      Left            =   16440
      Picture         =   "Frm_transfert_document.frx":3B72
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "LIGNE-------------------"
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
      Index           =   1
      Left            =   10320
      TabIndex        =   10
      Top             =   4200
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "ENTETE-------------------"
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
      TabIndex        =   6
      Top             =   4200
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
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "Frm_transfert_document"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Row_Actif As Integer
Dim Row_Actif1 As Integer
Dim Row_Actif2 As Integer

Sub init_grid()
Dim i As Integer

grid.MaxCols = 6
grid.MaxRows = 0

i = 1
grid.Row = 0
grid.col = i
grid.Text = "ID_TRANSFERT"
grid.ColWidth(i) = 20
i = i + 1
grid.col = i
grid.Text = "NOM_TRANSFERT"
grid.ColWidth(i) = 30
i = i + 1
grid.col = i
grid.Text = "TD_SOURCE_ID"
grid.ColWidth(i) = 20
i = i + 1
grid.col = i
grid.Text = "TD_DESTINATION_ID"
grid.ColWidth(i) = 20
i = i + 1
grid.col = i
grid.Text = "NUMERO_SOURCE"
grid.ColWidth(i) = 20
i = i + 1
grid.col = i
grid.Text = "NUMERO_DESTINATION"
grid.ColWidth(i) = 20

grid.OperationMode = 3
grid.UserColAction = UserColActionSort

End Sub

Sub init_grid1()
Dim i As Integer

grid1.MaxCols = 4
grid1.MaxRows = 0

i = 1
grid1.Row = 0
grid1.col = i
grid1.Text = "ID_TRANSFERT_CHAMPS"
grid1.ColWidth(i) = 0
i = i + 1
grid1.col = i
grid1.Text = "NUMERO"
grid1.ColWidth(i) = 0
i = i + 1
grid1.col = i
grid1.Text = "CHAMPS_SOURCE"
grid1.ColWidth(i) = 40
i = i + 1
grid1.col = i
grid1.Text = "CHAMPS_DESTINATION"
grid1.ColWidth(i) = 40

grid1.OperationMode = 3

grid1.UserColAction = UserColActionSort

End Sub

Sub init_grid2()
Dim i As Integer

grid2.MaxCols = 4
grid2.MaxRows = 0

i = 1
grid2.Row = 0
grid2.col = i
grid2.Text = "ID_TRANSFERT_CHAMPS"
grid2.ColWidth(i) = 0
i = i + 1
grid2.col = i
grid2.Text = "NUMERO"
grid2.ColWidth(i) = 0
i = i + 1
grid2.col = i
grid2.Text = "CHAMPS_SOURCE"
grid2.ColWidth(i) = 40
i = i + 1
grid2.col = i
grid2.Text = "CHAMPS_DESTINATION"
grid2.ColWidth(i) = 40

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
couleur = &H8000000F

Sql = "select * from TRANSFERT_DOCUMENT_CONFIG order by ID_TRANSFERT"

rs.Open Sql, db, adOpenKeyset, adLockOptimistic
If rs.RecordCount <> 0 Then
    grid.MaxRows = rs.RecordCount
    For i = 1 To rs.RecordCount
        j = 1
        grid.Row = i
        grid.col = j
        grid.Lock = True
        grid.BackColor = couleur
        grid.TypeHAlign = TypeHAlignCenter
        grid.Text = Str(rs("ID_TRANSFERT"))
        j = j + 1
        grid.col = j
        grid.Lock = True
        grid.BackColor = couleur
        grid.Text = rs("NOM_TRANSFERT")
        j = j + 1
        grid.col = j
        grid.Lock = True
        grid.BackColor = couleur
        grid.TypeHAlign = TypeHAlignCenter
        grid.Text = Str(rs("TD_SOURCE_ID"))
        j = j + 1
        grid.col = j
        grid.Lock = True
        grid.BackColor = couleur
        grid.TypeHAlign = TypeHAlignCenter
        grid.Text = Str(rs("TD_DESTINATION_ID"))
        j = j + 1
        grid.col = j
        grid.Lock = True
        grid.BackColor = couleur
        grid.TypeHAlign = TypeHAlignCenter
        grid.Text = Str(rs("NUMERO_ENTETE"))
        j = j + 1
        grid.col = j
        grid.Lock = True
        grid.BackColor = couleur
        grid.TypeHAlign = TypeHAlignCenter
        grid.Text = Str(rs("NUMERO_LIGNE"))
        
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
grid.col = 5
If Trim$(grid.Text) <> "" Then
    Sql = "select * from TRANSFERT_DOCUMENT_CONFIG_CHAMPS where TRANSFERT_ID='" & grid.Text & "' "
    
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
            grid1.Text = Str(rs("ID_TRANSFERT_CHAMPS"))
            j = j + 1
            grid1.col = j
            grid1.Lock = True
            grid1.BackColor = couleur
            grid1.Text = Str(rs("TRANSFERT_ID"))
            j = j + 1
            grid1.col = j
            grid1.Lock = True
            grid1.BackColor = couleur
            grid1.Text = IfNull(rs("CHAMPS_SOURCE"), "")
            j = j + 1
            grid1.col = j
            grid1.Lock = True
            grid1.BackColor = couleur
            grid1.Text = IfNull(rs("CHAMPS_DESTINATION"), "")
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
grid.col = 6
If Trim$(grid.Text) <> "" Then
    Sql = "select * from TRANSFERT_DOCUMENT_CONFIG_CHAMPS where TRANSFERT_ID='" & grid.Text & "' "
    
    rs.Open Sql, db, adOpenKeyset, adLockOptimistic
    If rs.RecordCount <> 0 Then
        grid2.MaxRows = rs.RecordCount
        For i = 1 To rs.RecordCount
            grid2.Row = i
            
            couleur = &HFF00&
            j = 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.TypeHAlign = TypeHAlignCenter
            grid2.Text = Str(rs("ID_TRANSFERT_CHAMPS"))
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.Text = Str(rs("TRANSFERT_ID"))
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.Text = IfNull(rs("CHAMPS_SOURCE"), "")
            j = j + 1
            grid2.col = j
            grid2.Lock = True
            grid2.BackColor = couleur
            grid2.Text = IfNull(rs("CHAMPS_DESTINATION"), "")
            rs.MoveNext
        Next i
    End If
    rs.Close
End If

Call grid2.SetSelection(1, Row_Actif1, 1, Row_Actif1)

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

Private Sub cmdAjouter_Click()
grid.Row = grid.ActiveRow
grid.col = 3
Type_documentG1 = grid.Text
grid.col = 4
Type_documentG2 = grid.Text
grid.col = 5
NumeroG = grid.Text


ID_transfert_ChampsG = ""
grid1.col = 3
Champs_sourceG = grid1.Text
grid1.col = 4
Champs_destinationG = grid1.Text

Typ_tableG = "E"

Frm_transfert_document_maj1.Show 1

Call Afficher_grid
Call Afficher_grid1
Call Afficher_grid2

End Sub

Private Sub cmdAjouter2_Click()
grid.Row = grid.ActiveRow
grid.col = 3
Type_documentG1 = grid.Text
grid.col = 4
Type_documentG2 = grid.Text
grid.col = 6
NumeroG = grid.Text


ID_transfert_ChampsG = ""
grid2.col = 3
Champs_sourceG = grid2.Text
grid2.col = 4
Champs_destinationG = grid2.Text

Typ_tableG = "L"

Frm_transfert_document_maj1.Show 1

Call Afficher_grid
Call Afficher_grid1
Call Afficher_grid2

End Sub

Private Sub cmdSupprimer_Click()
Dim Sql As String
Dim rep

'on error goto erreur

grid1.col = 1
grid1.Row = grid1.ActiveRow
    rep = MsgBox("Etes vous sûr d'annuler ce champs ?", vbYesNo)
    If rep = 6 Then
        Sql = "delete TRANSFERT_DOCUMENT_CONFIG_CHAMPS where ID_TRANSFERT_CHAMPS=" & grid1.Text
        db.Execute Sql
        Call Afficher_grid1
End If
Exit Sub

erreur:
MsgBox Err.Description
End Sub

Private Sub cmdSupprimer2_Click()
Dim Sql As String
Dim rep

'on error goto erreur

grid2.col = 1
grid2.Row = grid2.ActiveRow
    rep = MsgBox("Etes vous sûr d'annuler ce champs ?", vbYesNo)
    If rep = 6 Then
        Sql = "delete TRANSFERT_DOCUMENT_CONFIG_CHAMPS where ID_TRANSFERT_CHAMPS=" & grid2.Text
        db.Execute Sql
        Call Afficher_grid2
End If
Exit Sub

erreur:
MsgBox Err.Description
End Sub


Private Sub cmdAnnuler1_Click()

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

grid.Row = grid.ActiveRow
grid.col = 3
Type_documentG1 = grid.Text
grid.col = 4
Type_documentG2 = grid.Text
'Grid.col = 4
'NumeroG = Grid.Text
grid.col = 5
NumeroG = grid.Text

Typ_tableG = "E"

grid1.Row = grid1.ActiveRow
grid1.col = 1
ID_transfert_ChampsG = grid1.Text
grid1.col = 3
Champs_sourceG = grid1.Text
grid1.col = 4
Champs_destinationG = grid1.Text

Frm_transfert_document_maj1.Show 1
Call Afficher_grid
Call grid.SetSelection(1, Row_Actif, 1, Row_Actif)
Call Afficher_grid1
Call grid1.SetSelection(1, Row_Actif1, 1, Row_Actif1)
Call Afficher_grid2
Call grid2.SetSelection(1, Row_Actif2, 1, Row_Actif2)

End Sub

Private Sub cmdEdit2_Click()

grid.Row = grid.ActiveRow
grid.col = 3
Type_documentG1 = grid.Text
grid.col = 4
Type_documentG2 = grid.Text
'Grid.col = 4
'NumeroG = Grid.Text
grid.col = 6
NumeroG = grid.Text

Typ_tableG = "L"

grid2.Row = grid2.ActiveRow
grid2.col = 1
ID_transfert_ChampsG = grid2.Text
grid2.col = 3
Champs_sourceG = grid2.Text
grid2.col = 4
Champs_destinationG = grid2.Text

Frm_transfert_document_maj1.Show 1
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

Private Sub Grid2_dblClick(ByVal col As Long, ByVal Row As Long)
If Row > 0 Then
    Row_Actif = grid.ActiveRow
    Row_Actif1 = grid1.ActiveRow
    Row_Actif2 = grid2.ActiveRow
    cmdEdit2_Click
End If
End Sub

