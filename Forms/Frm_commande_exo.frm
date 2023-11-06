VERSION 5.00
Begin VB.Form Frm_commande_exo 
   Caption         =   "Commande"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1215
      Left            =   9000
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.PictureBox CommonDialog1 
      Height          =   480
      Left            =   1320
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   2
      Top             =   1080
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1335
      Left            =   11880
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "Frm_commande_exo"
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
    grid.MaxRows = 0
    f = grid.GetExcelSheetList(.FileName, List, ListCount, (myPath & "log.txt"), handle, True)
    f = grid.ImportExcelSheet(handle, 0)
    
    xlfile = .FileName
    'grid(0).MaxCols = 27
    'grid(0).RowsFrozen = 1
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

Private Sub Exporter(Exercice As String, trimestre As String)

Dim sql As String
Dim rs As New Recordset
Dim rs_société As New Recordset
Dim rs_commande As New Recordset
Dim rs_Trimestre As New Recordset
Dim compteur As Integer
Dim Nombre_total_BDC As Integer
Dim FileName As String
Dim rep
Dim max
Dim PrixHT
Dim PrixTVA

On Error GoTo Erreur


'Entête
'-----------------------------------------------------
Exp.Text = "EF"
Exp.Text = Exp.Text & Format(Trim$(rs_société("Matricule Fiscal")), "0000000")
Exp.Text = Exp.Text & Trim$(rs_société("Clé du matricule fiscal"))
Exp.Text = Exp.Text & Trim$(rs_société("Catégorie contribuable"))
Exp.Text = Exp.Text & Format(rs_société("Numéro de établissement"), "000")
Exp.Text = Exp.Text & Trim$(CboExercice.Text)
Exp.Text = Exp.Text & "T" & Trim$(CboTrimestre.Text)
Exp.Text = Exp.Text & Trim$(rs_société("Nom")) & String(40 - Len(Trim$(rs_société("nom"))), " ")
Exp.Text = Exp.Text & Trim$(rs_société("Activité contribuable déclarant")) & String(40 - Len(Trim$(rs_société("Activité contribuable déclarant"))), " ")
Exp.Text = Exp.Text & Trim$(rs_société("Ville")) & String(40 - Len(Trim$(rs_société("Ville"))), " ")
Exp.Text = Exp.Text & Trim$(rs_société("Rue")) & String(72 - Len(Trim$(rs_société("Rue"))), " ")
Exp.Text = Exp.Text & Format(rs_société("Numéro"), "0000")
Exp.Text = Exp.Text & Format(rs_société("Code postal"), "0000")


'-----------------------------------------------------
'LIGNE
'-----------------------------------------------------
compteur = 0
PrixHT = 0
PrixTVA = 0

While Not rs_commande.EOF
    compteur = compteur + 1
    PrixHT = PrixHT + rs_commande("Prix d’achat (HT)")
    PrixTVA = PrixTVA + rs_commande("Montant TVA suspendue")
    Exp.Text = Exp.Text & Chr(13) & Chr(10)
    Exp.Text = Exp.Text & "DF"
    Exp.Text = Exp.Text & Format(Trim$(rs_société("Matricule Fiscal")), "0000000")
    Exp.Text = Exp.Text & Trim$(rs_société("Clé du matricule fiscal"))
    Exp.Text = Exp.Text & Trim$(rs_société("Catégorie contribuable"))
    Exp.Text = Exp.Text & Format(rs_société("Numéro de établissement"), "000")
    Exp.Text = Exp.Text & Trim$(CboExercice.Text)
    Exp.Text = Exp.Text & "T" & Trim$(CboTrimestre.Text)
    Exp.Text = Exp.Text & Format(compteur, "000000")
    Exp.Text = Exp.Text & Trim$(rs_Trimestre("Numéro autorisation d’achat en suspension de TVA")) & String(30 - Len(Trim$(rs_Trimestre("Numéro autorisation d’achat en suspension de TVA"))), " ")
    Exp.Text = Exp.Text & Format(Trim$(rs_commande("Numéro bon de commandes")), "0000000000000")
    Exp.Text = Exp.Text & Mid(rs_commande("Date bon de commandes"), 1, 2) & Mid(rs_commande("Date bon de commandes"), 4, 2) & Mid(rs_commande("Date bon de commandes"), 7, 4)
    Exp.Text = Exp.Text & Trim$(rs_commande("Matricule Fiscal Fournisseur")) & String(13 - Len(Trim$(rs_commande("Matricule Fiscal Fournisseur"))), " ")
    Exp.Text = Exp.Text & Trim$(rs_commande("Nom_fournisseur")) & String(40 - Len(Trim$(rs_commande("Nom_fournisseur"))), " ")
    Exp.Text = Exp.Text & Trim$(rs_commande("Numéro facture")) & String(30 - Len(Trim$(rs_commande("Numéro facture"))), " ")
    Exp.Text = Exp.Text & Mid(rs_commande("Date facture"), 1, 2) & Mid(rs_commande("Date facture"), 4, 2) & Mid(rs_commande("Date facture"), 7, 4)
    Exp.Text = Exp.Text & Format(1000 * rs_commande("Prix d’achat (HT)"), "000000000000000")
    Exp.Text = Exp.Text & Format(1000 * rs_commande("Montant TVA suspendue"), "000000000000000")
    Exp.Text = Exp.Text & "<"
    Exp.Text = Exp.Text & Trim$(Replace(rs_commande("Objet facture"), Chr(13), " ")) & String(320 - Len(Trim$(rs_commande("Objet facture"))), " ")
    Exp.Text = Exp.Text & "/>"
    rs_commande.MoveNext
Wend

'-----------------------------------------------------
'PIED
    Exp.Text = Exp.Text & Chr(13) & Chr(10)
    Exp.Text = Exp.Text & "TF"
    Exp.Text = Exp.Text & Format(Trim$(rs_société("Matricule Fiscal")), "0000000")
    Exp.Text = Exp.Text & Trim$(rs_société("Clé du matricule fiscal"))
    Exp.Text = Exp.Text & Trim$(rs_société("Catégorie contribuable"))
    Exp.Text = Exp.Text & Format(rs_société("Numéro de établissement"), "000")
    Exp.Text = Exp.Text & Trim$(CboExercice.Text)
    Exp.Text = Exp.Text & "T" & Trim$(CboTrimestre.Text)
    Exp.Text = Exp.Text & Format(Nombre_total_BDC, "000000")
    Exp.Text = Exp.Text & String(142, " ")
    Exp.Text = Exp.Text & Format(1000 * PrixHT, "000000000000000")
    Exp.Text = Exp.Text & Format(1000 * PrixTVA, "000000000000000")

comm:
With CommonDialog1
    .CancelError = True
    .FileName = "BCD_" & "T" & Trim$(CboTrimestre.Text) & "_" & Mid(CboExercice.Text, 3, 2)
    .DialogTitle = "Exportation de fichiers"
    .InitDir = App.Path & "\fichiers"
    .Flags = cdlOFNHideReadOnly
    .ShowOpen
    FileName = .FileName
End With

If Dir(FileName) <> "" Then
    rep = MsgBox("Ce fichier existe déjà, voulez-vous l'ecraser ?", vbYesNo)
    If rep = 6 Then
        Open FileName For Output As #1
        Print #1, Exp.Text
        Close #1
    Else
        GoTo comm
    End If
Else
    Open FileName For Output As #1
    Print #1, Exp.Text
    Close #1
End If


sql = "insert into Journal  (Exercice,Trimestre,[date],Nombre_ligne,[Prix d’achat (HT)],[Montant TVA suspendue]) values (" _
& CboExercice.Text & "," & CboTrimestre.Text & ",'" & Date & " " & Time & "'," & compteur & "," & replPrixHT & "," & PrixTVA & ") "
db.Execute sql

sql = "select max(code) from journal "
rs.Open sql, db, adOpenStatic, adLockOptimistic
If IsNull(rs(0)) Then
    max = 1
Else
    max = rs(0)
End If

sql = "insert into [Bon de commande journal] " _
& " (Code_journal,Code,Code_trimestre,Code_fournisseur,[Numéro bon de commandes]," _
& " [Date bon de commandes],[Numéro facture],[Date facture],[Prix d’achat (HT)],[Montant TVA suspendue]," _
& " [Objet facture],[Nom_Fournisseur],[Matricule Fiscal Fournisseur],[Numéro autorisation d’achat en suspension de TVA])" _
& " select " & max & ",C.Code,C.Code_trimestre,C.Code_fournisseur,C.[Numéro bon de commandes]," _
& " C.[Date bon de commandes],C.[Numéro facture],C.[Date facture],C.[Prix d’achat (HT)],C.[Montant TVA suspendue]," _
& " C.[Objet facture],F.nom,F.[Matricule Fiscal],'" _
& rs_Trimestre("Numéro autorisation d’achat en suspension de TVA") & "' from [Bon de commande] C " _
& " inner Join Fournisseur F on F.code=C.code_fournisseur " _
& " where code_trimestre=" & rs_Trimestre("code")

db.Execute sql

rs.Close

rs_société.Close
rs_commande.Close
rs_Trimestre.Close

MsgBox "Le fichier des bons de commandes est exporté avec succés. " & Chr(13) & FileName


Exit Sub

Erreur:
If Err = 32755 Then Exit Sub
MsgBox Err.Description

End Sub


Private Sub Command2_Click()

Call Exporter("2021", "1")

End Sub
