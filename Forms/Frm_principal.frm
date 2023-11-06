VERSION 5.00
Begin VB.MDIForm Frm_principal 
   BackColor       =   &H8000000C&
   Caption         =   "CLEVER_DEV"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   Icon            =   "Frm_principal.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mn_document 
      Caption         =   "&Document"
      Begin VB.Menu mn_page 
         Caption         =   "&Page"
      End
      Begin VB.Menu mn_grille 
         Caption         =   "&Grille"
      End
      Begin VB.Menu mn_transfert 
         Caption         =   "&Transfert"
      End
   End
   Begin VB.Menu mn_BCC 
      Caption         =   "Bon de commande"
   End
   Begin VB.Menu mn_quitter 
      Caption         =   "Quitter"
   End
End
Attribute VB_Name = "Frm_principal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()

With db
    .CursorLocation = adUseClient
    .ConnectionTimeout = 20
    .CommandTimeout = 200
    '.Open "Provider=sqloledb;User ID=sa ;password=NeverForget ;Data Source=VW002\DEV ;initial catalog =CLEVER "
    .Open "Provider=sqloledb;User ID=sa ;password=$$vnext2023 ;Data Source=VS026 ;initial catalog =CLEVER_EVAL "
End With

End Sub

Private Sub mn_BCC_Click()
Frm_commande_exo.Show
End Sub

Private Sub mn_grille_Click()
Frm_grille.Show
End Sub

Private Sub mn_page_Click()
Frm_page.Show
End Sub

Private Sub mn_transfert_Click()
Frm_transfert_document.Show
End Sub

