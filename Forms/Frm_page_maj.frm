VERSION 5.00
Begin VB.Form Frm_page_maj 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAGE"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12300
   Icon            =   "Frm_page_maj.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   12300
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   6855
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   12135
      Begin VB.ComboBox CboFields 
         Height          =   315
         Index           =   2
         Left            =   7920
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   960
         Width           =   3255
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H8000000F&
         DataField       =   "TD_ID"
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
         Index           =   16
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   79
         Top             =   1320
         Width           =   1695
      End
      Begin VB.ComboBox CboListe3 
         Height          =   315
         Left            =   2520
         TabIndex        =   78
         Text            =   "Combo1"
         Top             =   1320
         Width           =   4815
      End
      Begin VB.ListBox List3 
         Height          =   255
         Left            =   1680
         TabIndex        =   77
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check2"
         DataField       =   "EXIST_STATUT"
         Height          =   195
         Index           =   29
         Left            =   11280
         TabIndex        =   75
         Top             =   3960
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "EXIST_CLIENT"
         Height          =   195
         Index           =   28
         Left            =   11280
         TabIndex        =   73
         Top             =   3600
         Width           =   255
      End
      Begin VB.ComboBox CboFields 
         DataField       =   "TABLE_ID"
         Height          =   315
         Index           =   1
         Left            =   2520
         TabIndex        =   3
         Top             =   1680
         Width           =   4815
      End
      Begin VB.CheckBox ChkFields 
         DataField       =   "B_INFORMATION"
         Height          =   255
         Index           =   27
         Left            =   8280
         TabIndex        =   70
         Top             =   5085
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         DataField       =   "B_CLIENT"
         Height          =   195
         Index           =   26
         Left            =   8280
         TabIndex        =   69
         Top             =   4725
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         DataField       =   "B_LISTE_REGLEMENT"
         Height          =   195
         Index           =   25
         Left            =   8280
         TabIndex        =   68
         Top             =   4365
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         DataField       =   "B_REGLEMENT"
         Height          =   195
         Index           =   24
         Left            =   8280
         TabIndex        =   67
         Top             =   3960
         Width           =   375
      End
      Begin VB.CheckBox ChkFields 
         DataField       =   "B_BASE_TVA"
         Height          =   255
         Index           =   23
         Left            =   8280
         TabIndex        =   66
         Top             =   3600
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check9"
         DataField       =   "B_TRANSFERT"
         Height          =   195
         Index           =   22
         Left            =   5400
         TabIndex        =   60
         Top             =   6525
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check8"
         DataField       =   "B_TRANSFERER_BON_LIVRAISON"
         Height          =   195
         Index           =   21
         Left            =   5400
         TabIndex        =   59
         Top             =   6165
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         DataField       =   "B_TRANSFERER_FACTURE"
         Height          =   195
         Index           =   20
         Left            =   5400
         TabIndex        =   58
         Top             =   5805
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         DataField       =   "B_TRANSFERER_COMMANDE"
         Height          =   255
         Index           =   19
         Left            =   5400
         TabIndex        =   57
         Top             =   5445
         Width           =   375
      End
      Begin VB.CheckBox ChkFields 
         DataField       =   "B_DUPLIQUER"
         Height          =   255
         Index           =   18
         Left            =   5400
         TabIndex        =   56
         Top             =   5085
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         DataField       =   "B_ANNULER"
         Height          =   255
         Index           =   17
         Left            =   5400
         TabIndex        =   55
         Top             =   4725
         Width           =   495
      End
      Begin VB.CheckBox ChkFields 
         DataField       =   "B_VALIDER"
         Height          =   195
         Index           =   16
         Left            =   5400
         TabIndex        =   54
         Top             =   4365
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "ChkFields"
         DataField       =   "B_IMPRIMER"
         Height          =   255
         Index           =   15
         Left            =   5400
         TabIndex        =   53
         Top             =   3960
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         DataField       =   "B_LISTER"
         Height          =   255
         Index           =   14
         Left            =   5400
         TabIndex        =   44
         Top             =   3600
         Width           =   375
      End
      Begin VB.ComboBox CboListe2 
         Height          =   315
         Left            =   2520
         TabIndex        =   6
         Text            =   "CboListe2"
         Top             =   2760
         Width           =   4815
      End
      Begin VB.ListBox List2 
         Height          =   255
         Left            =   1680
         TabIndex        =   41
         Top             =   2760
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H8000000F&
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
         Index           =   15
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H8000000F&
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
         Index           =   14
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   38
         Top             =   2400
         Width           =   1695
      End
      Begin VB.ListBox List1 
         Height          =   255
         Left            =   1680
         TabIndex        =   37
         Top             =   2400
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox CboListe1 
         Height          =   315
         Left            =   2520
         TabIndex        =   5
         Text            =   "CboListe1"
         Top             =   2400
         Width           =   4815
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "B_DERNIER"
         Height          =   195
         Index           =   13
         Left            =   2520
         TabIndex        =   34
         Top             =   6525
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "B_PRECEDENT"
         Height          =   195
         Index           =   12
         Left            =   2520
         TabIndex        =   32
         Top             =   6165
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "B_SUIVANT"
         Height          =   195
         Index           =   11
         Left            =   2520
         TabIndex        =   30
         Top             =   5805
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "B_PREMIER"
         Height          =   195
         Index           =   10
         Left            =   2520
         TabIndex        =   28
         Top             =   5445
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "B_ACTUALISER"
         Height          =   195
         Index           =   9
         Left            =   2520
         TabIndex        =   26
         Top             =   5085
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "B_RECHERCHER"
         Height          =   195
         Index           =   8
         Left            =   2520
         TabIndex        =   24
         Top             =   4725
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "B_SUPPRIMER"
         Height          =   195
         Index           =   7
         Left            =   2520
         TabIndex        =   22
         Top             =   4365
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "B_MODIFIER"
         Height          =   195
         Index           =   6
         Left            =   2520
         TabIndex        =   20
         Top             =   4005
         Width           =   255
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "B_AJOUTER"
         Height          =   195
         Index           =   5
         Left            =   2520
         TabIndex        =   18
         Top             =   3645
         Width           =   255
      End
      Begin VB.ComboBox CboFields 
         Height          =   315
         Index           =   0
         Left            =   2520
         TabIndex        =   1
         Text            =   "CboFields"
         Top             =   960
         Width           =   3135
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
         Left            =   11040
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   6120
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
         Left            =   9960
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   6120
         Width           =   855
      End
      Begin VB.CheckBox ChkFields 
         Caption         =   "Check1"
         DataField       =   "ACTIF_PAGE"
         Height          =   195
         Index           =   4
         Left            =   2520
         TabIndex        =   10
         Top             =   3120
         Width           =   255
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H00FFFFFF&
         DataField       =   "NOM_PAGE"
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
      Begin VB.ComboBox CboListe 
         Height          =   315
         Left            =   2520
         TabIndex        =   4
         Text            =   "CboListe"
         Top             =   2040
         Width           =   4815
      End
      Begin VB.ListBox List 
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   2040
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox txtFields 
         BackColor       =   &H8000000F&
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
         Index           =   3
         Left            =   7440
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2040
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         Caption         =   "Table Update:"
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
         Index           =   34
         Left            =   5880
         TabIndex        =   80
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Document:"
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
         Index           =   33
         Left            =   240
         TabIndex        =   76
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Exist Statut:"
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
         Index           =   32
         Left            =   9120
         TabIndex        =   74
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Exist client:"
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
         Index           =   31
         Left            =   9120
         TabIndex        =   72
         Top             =   3600
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
         Index           =   30
         Left            =   240
         TabIndex        =   71
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Information:"
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
         Index           =   29
         Left            =   6240
         TabIndex        =   65
         Top             =   5040
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Client:"
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
         Index           =   28
         Left            =   6240
         TabIndex        =   64
         Top             =   4680
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Liste réglement:"
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
         Index           =   27
         Left            =   6240
         TabIndex        =   63
         Top             =   4320
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Réglement:"
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
         Index           =   26
         Left            =   6240
         TabIndex        =   62
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Base TVA:"
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
         Index           =   25
         Left            =   6240
         TabIndex        =   61
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Transfert:"
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
         Index           =   24
         Left            =   3240
         TabIndex        =   52
         Top             =   6480
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Transfert BL:"
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
         Index           =   23
         Left            =   3240
         TabIndex        =   51
         Top             =   6120
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Transfert facture:"
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
         Index           =   22
         Left            =   3240
         TabIndex        =   50
         Top             =   5760
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Transfert commande:"
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
         Index           =   21
         Left            =   3240
         TabIndex        =   49
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Dupliquer:"
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
         Left            =   3240
         TabIndex        =   48
         Top             =   5040
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Annuler:"
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
         Left            =   3240
         TabIndex        =   47
         Top             =   4680
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Valider:"
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
         Left            =   3240
         TabIndex        =   46
         Top             =   4320
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Imprimer:"
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
         Left            =   3240
         TabIndex        =   45
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Lister:"
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
         Left            =   3240
         TabIndex        =   43
         Top             =   3600
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Grid lister:"
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
         TabIndex        =   42
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Grid rechercher:"
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
         TabIndex        =   39
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label LblFields 
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Index           =   0
         Left            =   2520
         TabIndex        =   36
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label lblLabels 
         Caption         =   "Dernier:"
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
         TabIndex        =   35
         Top             =   6480
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Précédent:"
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
         TabIndex        =   33
         Top             =   6120
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Suivant:"
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
         TabIndex        =   31
         Top             =   5760
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Premier:"
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
         TabIndex        =   29
         Top             =   5400
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Actualiser:"
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
         TabIndex        =   27
         Top             =   5040
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Rechercher:"
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
         TabIndex        =   25
         Top             =   4680
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Supprimer:"
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
         TabIndex        =   23
         Top             =   4320
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Modifier:"
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
         TabIndex        =   21
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Ajouter:"
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
         TabIndex        =   19
         Top             =   3600
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
         Index           =   5
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   1935
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
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   1815
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
         Left            =   240
         TabIndex        =   14
         Top             =   3120
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Grid:"
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
         TabIndex        =   13
         Top             =   2040
         Width           =   1935
      End
   End
End
Attribute VB_Name = "Frm_page_maj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Table = "PAGE"

Private Sub Cbofields_Click(Index As Integer)
If Index = 0 Then
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
    If Insert_TB(Me, Table, 1, 1, 0, 0, 0, 0, "", 1) Then
        Unload Me
    End If
Else
    If Update_TB(Me, Table, 1, 1, 0, 0, 0, " ID_PAGE='" & baseG & "' ", baseG, 1) Then
        Unload Me
    End If
End If

Exit Sub

erreur:
MsgBox Err.Description

End Sub


Private Sub Combo1_Change()

End Sub

Private Sub Form_Load()
Dim Sql As String
Dim rs As New Recordset

Call remplir_combo

If baseG <> "" Then
    Sql = "select * from PAGE P " _
    & " inner join TYPE_DOCUMENT T on P.TD_ID=T.ID_TD where ID_PAGE='" & baseG & "' "
    rs.Open Sql, db, adOpenStatic, adLockOptimistic
    txtFields(1).Text = rs("NOM_PAGE")
    CboFields(0).Text = rs("TABLE_TD")
    CboFields(1).Text = IfNull(rs("TABLE_TD_ID"), "")
    CboFields(2).Text = IfNull(rs("TABLE_UPDATE_TD"), "")
    txtFields(3).Text = IfNull(rs("GRID_PAGE_ID"), "")
    txtFields(14).Text = IfNull(rs("GRID_PAGE_RECHERCHER_ID"), "")
    txtFields(15).Text = IfNull(rs("GRID_PAGE_LISTER_ID"), "")
    txtFields(16).Text = IfNull(rs("TD_ID"), "")
    If rs("EXIST_CLIENT") Then
        ChkFields(28).Value = 1
    Else
        ChkFields(28).Value = 0
    End If
    If rs("EXIST_STATUT") Then
        ChkFields(29).Value = 1
    Else
        ChkFields(29).Value = 0
    End If
    
    If rs("ACTIF_PAGE") Then
        ChkFields(4).Value = 1
    Else
        ChkFields(4).Value = 0
    End If
    If rs("B_AJOUTER") Then
        ChkFields(5).Value = 1
    Else
        ChkFields(5).Value = 0
    End If
    If rs("B_MODIFIER") Then
        ChkFields(6).Value = 1
    Else
        ChkFields(6).Value = 0
    End If
    If rs("B_SUPPRIMER") Then
        ChkFields(7).Value = 1
    Else
        ChkFields(7).Value = 0
    End If
    If rs("B_RECHERCHER") Then
        ChkFields(8).Value = 1
    Else
        ChkFields(8).Value = 0
    End If
    If rs("B_ACTUALISER") Then
        ChkFields(9).Value = 1
    Else
        ChkFields(9).Value = 0
    End If
    If rs("B_PREMIER") Then
        ChkFields(10).Value = 1
    Else
        ChkFields(10).Value = 0
    End If
    If rs("B_SUIVANT") Then
        ChkFields(11).Value = 1
    Else
        ChkFields(11).Value = 0
    End If
    If rs("B_PRECEDENT") Then
        ChkFields(12).Value = 1
    Else
        ChkFields(12).Value = 0
    End If
    If rs("B_DERNIER") Then
        ChkFields(13).Value = 1
    Else
        ChkFields(13).Value = 0
    End If
    If rs("B_LISTER") Then
        ChkFields(14).Value = 1
    Else
        ChkFields(14).Value = 0
    End If
    If rs("B_IMPRIMER") Then
        ChkFields(15).Value = 1
    Else
        ChkFields(15).Value = 0
    End If
    If rs("B_VALIDER") Then
        ChkFields(16).Value = 1
    Else
        ChkFields(16).Value = 0
    End If
    If rs("B_ANNULER") Then
        ChkFields(17).Value = 1
    Else
        ChkFields(17).Value = 0
    End If
    If rs("B_DUPLIQUER") Then
        ChkFields(18).Value = 1
    Else
        ChkFields(18).Value = 0
    End If
    If rs("B_TRANSFERER_COMMANDE") Then
        ChkFields(19).Value = 1
    Else
        ChkFields(19).Value = 0
    End If
    If rs("B_TRANSFERER_FACTURE") Then
        ChkFields(20).Value = 1
    Else
        ChkFields(20).Value = 0
    End If
    If rs("B_TRANSFERER_BON_LIVRAISON") Then
        ChkFields(21).Value = 1
    Else
        ChkFields(21).Value = 0
    End If
    If rs("B_TRANSFERT") Then
        ChkFields(22).Value = 1
    Else
        ChkFields(22).Value = 0
    End If
    If rs("B_BASE_TVA") Then
        ChkFields(23).Value = 1
    Else
        ChkFields(23).Value = 0
    End If
    If rs("B_REGLEMENT") Then
        ChkFields(24).Value = 1
    Else
        ChkFields(24).Value = 0
    End If
    If rs("B_CLIENT") Then
        ChkFields(26).Value = 1
    Else
        ChkFields(26).Value = 0
    End If
    If rs("B_INFORMATION") Then
        ChkFields(27).Value = 1
    Else
        ChkFields(27).Value = 0
    End If
        
    LblFields(0).Caption = baseG
Else
    LblFields(0).Caption = baseG
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

Private Sub txtFields_Change(Index As Integer)

Dim Sql As String
Dim rs As New Recordset
If Index = 3 Or Index = 14 Or Index = 15 Then
    If Val(txtFields(Index).Text) <> 0 Then
        Sql = "select NOM_GRID from GRID where ID_GRID=" & txtFields(Index).Text
        rs.Open Sql, db, adOpenStatic, adLockOptimistic
        If Index = 3 Then CboListe.Text = rs(0)
        If Index = 14 Then CboListe1.Text = rs(0)
        If Index = 15 Then CboListe2.Text = rs(0)
        rs.Close
    End If
End If

If Index = 16 Then
    Sql = "select NOM_TD from TYPE_DOCUMENT where ID_TD=" & txtFields(Index).Text
    rs.Open Sql, db, adOpenStatic, adLockOptimistic
    CboListe3.Text = rs(0)
End If

End Sub

Private Sub remplir_combo()
Dim Sql As String
Dim rs As New Recordset
Dim i As Integer

CboListe.Clear
List.Clear
List.AddItem 0
CboListe.AddItem ""
Sql = "select ID_GRID,NOM_GRID from GRID order by NOM_GRID "
rs.Open Sql, db, adOpenStatic, adLockOptimistic
While Not rs.EOF
  List.AddItem rs(0)
  CboListe.AddItem rs(1)
  rs.MoveNext
Wend
rs.Close

CboListe1.Clear
List1.Clear
List1.AddItem 0
CboListe1.AddItem ""
Sql = "select ID_GRID,NOM_GRID from GRID order by NOM_GRID "
rs.Open Sql, db, adOpenStatic, adLockOptimistic
While Not rs.EOF
  List1.AddItem rs(0)
  CboListe1.AddItem rs(1)
  rs.MoveNext
Wend
rs.Close

CboListe2.Clear
List2.Clear
List2.AddItem 0
CboListe2.AddItem ""
Sql = "select ID_GRID,NOM_GRID from GRID order by NOM_GRID "
rs.Open Sql, db, adOpenStatic, adLockOptimistic
While Not rs.EOF
  List2.AddItem rs(0)
  CboListe2.AddItem rs(1)
  rs.MoveNext
Wend
rs.Close


'CboFields(0).Clear
'
'Sql = "select name from sys.sysobjects where xtype='U' or xtype='V' order by name"
'rs.Open Sql, db, adOpenStatic, adLockOptimistic
'While Not rs.EOF
'  CboFields(0).AddItem rs(0)
'  rs.MoveNext
'Wend
'rs.Close


CboListe3.Clear
List3.Clear
List3.AddItem 0
CboListe3.AddItem ""
Sql = "select ID_TD,NOM_TD from TYPE_DOCUMENT order by NOM_TD "
rs.Open Sql, db, adOpenStatic, adLockOptimistic
While Not rs.EOF
  List3.AddItem rs(0)
  CboListe3.AddItem rs(1)
  rs.MoveNext
Wend
rs.Close


'CboFields(2).Clear
'
'Sql = "select name from sys.sysobjects where xtype='U' order by name"
'rs.Open Sql, db, adOpenStatic, adLockOptimistic
'While Not rs.EOF
'  CboFields(2).AddItem rs(0)
'  rs.MoveNext
'Wend
'rs.Close

End Sub

Private Sub remplir_combo1()
Dim Sql As String
Dim rs As New Recordset

'CboFields(1).Clear
'If CboFields(0).Text <> "" Then
'    Sql = "select CHAMPS from VUE_TABLES where [TABLE]='" & CboFields(0).Text & "' "
'    rs.Open Sql, db, adOpenStatic, adLockOptimistic
'    While Not rs.EOF
'        CboFields(1).AddItem rs(0)
'        rs.MoveNext
'    Wend
'    rs.Close
'End If

End Sub


Private Sub CboListe_Click()
List.ListIndex = CboListe.ListIndex

If CboListe.ListIndex <> 0 Then
    txtFields(3).Text = List.Text
Else
    txtFields(3).Text = ""
End If
End Sub

Private Sub CboListe1_Click()
List1.ListIndex = CboListe1.ListIndex

If CboListe1.ListIndex <> 0 Then
    txtFields(14).Text = List1.Text
Else
    txtFields(14).Text = ""
End If
End Sub


Private Sub CboListe2_Click()
List2.ListIndex = CboListe2.ListIndex

If CboListe2.ListIndex <> 0 Then
    txtFields(15).Text = List2.Text
Else
    txtFields(15).Text = ""
End If
End Sub

Private Sub CboListe3_Click()
List3.ListIndex = CboListe3.ListIndex

If CboListe3.ListIndex <> 0 Then
    txtFields(16).Text = List3.Text
Else
    txtFields(16).Text = ""
End If
End Sub

