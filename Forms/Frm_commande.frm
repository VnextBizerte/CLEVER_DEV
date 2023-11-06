VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form Frm_commande 
   Caption         =   "Bon de commande"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   Icon            =   "Frm_commande.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1215
      Left            =   11400
      TabIndex        =   1
      Top             =   360
      Width           =   2055
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   15495
      _Version        =   393216
      _ExtentX        =   27331
      _ExtentY        =   10610
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
      SpreadDesigner  =   "Frm_commande.frx":1272
   End
End
Attribute VB_Name = "Frm_commande"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Importer_Excel()

Dim rs2 As New Recordset
Dim sql As String
Dim i As Long
Dim j As Long
Dim List() As String
Dim ListCount As Integer
Dim fromRight As Long
Dim myPath As String
Dim xlfile As String
Dim handle As Integer
Dim f

ReDim List(1)
    

With CommonDialog1
    .FileName = "*.xls"
    .DialogTitle = "Select Excel file"
    .Filter = "Excel files|(*.xls)"
    .FilterIndex = 0
    .InitDir = App.Path
    .Flags = cdlOFNHideReadOnly
    .ShowOpen
    fromRight = InStrRev(.FileName, "\", , vbTextCompare)
    If fromRight > 1 Then
        myPath = Left(.FileName, fromRight)
    End If
    grid(0).MaxRows = 0
    f = grid.GetExcelSheetList(.FileName, List, ListCount, (myPath & "log.txt"), handle, True)
    f = grid.ImportExcelSheet(handle, 0)
    
    xlfile = .FileName
    'grid.MaxCols = 27
    'grid.RowsFrozen = 1
    'For i = 1 To 27
    '    grid(0).Row = 0
    '    grid(0).col = i
    '    grid(0).Text = " "
    'Next i
End With

End Sub

Private Sub Command1_Click()

Call Importer_Excel

End Sub
