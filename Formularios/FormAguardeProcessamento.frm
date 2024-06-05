VERSION 5.00
Begin VB.Form FormAguardeProcessamento 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3450
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormAguardeProcessamento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   50
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblMensagem 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Aguarde o Processamento..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   3225
   End
End
Attribute VB_Name = "FormAguardeProcessamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

