VERSION 5.00
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "todg6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FormPainel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Painel de Criptomoedas"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormPainel.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   550
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.StatusBar StatusBarPainel 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   7965
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   8996
            Text            =   "CarteiraCrypto Ver. 1.0.0 - Desenvolvido por Arthur Santos Maciel"
            TextSave        =   "CarteiraCrypto Ver. 1.0.0 - Desenvolvido por Arthur Santos Maciel"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Visible         =   0   'False
            Object.Width           =   6720
            Text            =   "Erro de conexão com a API!!!"
            TextSave        =   "Erro de conexão com a API!!!"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   8996
            TextSave        =   "12:21"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   8996
            TextSave        =   "30/05/2024"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TrueOleDBGrid60.TDBGrid GridPainel 
      Height          =   6255
      Left            =   240
      OleObjectBlob   =   "FormPainel.frx":06DA
      TabIndex        =   4
      Top             =   1290
      Width           =   14865
   End
   Begin VB.Frame fraPainel 
      Appearance      =   0  'Flat
      Caption         =   "Controles do Painel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1065
      Left            =   240
      TabIndex        =   0
      Top             =   210
      Width           =   14865
      Begin VB.Frame fraMoeda 
         Appearance      =   0  'Flat
         Caption         =   "Moeda Comprada"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Index           =   1
         Left            =   3810
         TabIndex        =   3
         Top             =   240
         Width           =   3675
         Begin VB.CheckBox chkMoeda 
            Caption         =   "PAXG"
            Height          =   195
            Index           =   2
            Left            =   1050
            TabIndex        =   10
            Top             =   240
            Width           =   705
         End
         Begin VB.CheckBox chkMoeda 
            Caption         =   "ETH"
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   9
            Top             =   480
            Width           =   585
         End
         Begin VB.CheckBox chkMoeda 
            Caption         =   "BTC"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   8
            Top             =   240
            Width           =   585
         End
      End
      Begin VB.Frame fraMoeda 
         Appearance      =   0  'Flat
         Caption         =   "Moeda Utilizada"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Top             =   240
         Width           =   3675
         Begin VB.OptionButton optMoeda 
            Appearance      =   0  'Flat
            Caption         =   "USDT"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   30
            TabIndex        =   7
            Top             =   480
            Width           =   705
         End
         Begin VB.OptionButton optMoeda 
            Appearance      =   0  'Flat
            Caption         =   "BRL"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   30
            TabIndex        =   6
            Top             =   240
            Width           =   585
         End
      End
      Begin VB.CommandButton cmdPesquisar 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Pesquisar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   13710
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   660
         Width           =   1110
      End
   End
   Begin VB.Label lblLegenda 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"FormPainel.frx":5919
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   240
      TabIndex        =   11
      Top             =   7590
      Width           =   11895
   End
   Begin VB.Image ImagePlanoFundo 
      Height          =   8250
      Left            =   0
      Picture         =   "FormPainel.frx":59AC
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "FormPainel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkMoeda_Click(index As Integer)
100     Call ControlaMarcacaoFiltros(CBool(chkMoeda(index).Value), False, index)
End Sub

Private Sub cmdPesquisar_Click()
100     If VerificaOptionMarcado = False Then
101         MsgBox "É necessário realizar pelo menos um filtro no frame " & Chr(34) & "Moeda Utilizada" & Chr(34), vbInformation
102         Exit Sub
103     End If
104     If VerificaCheckboxMarcado = False Then
105         MsgBox "É necessário realizar pelo menos um filtro no frame " & Chr(34) & "Moeda Comprada" & Chr(34), vbInformation
106         Exit Sub
107     End If
108     AguardeProcessamento True
109     Call RealizaRequisicao(MontaQueryFiltrosPainel)
110     Call mdlPainel.PreencheGridPainel
111     AguardeProcessamento False
End Sub

Private Sub optMoeda_Click(index As Integer)
100     Call ControlaMarcacaoFiltros(CBool(optMoeda(index).Value), True, index)
End Sub
