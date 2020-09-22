VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Connect As?"
   ClientHeight    =   750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   2760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClient 
      Caption         =   "Client"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdServer 
      Caption         =   "&Server"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***remember only one server***

Private Sub cmdClient_Click()
    frmClient.Show 'shows approiate form
    Unload Me 'unloads unneeded form
    
End Sub

Private Sub cmdServer_Click()
    frmServer.Show 'shows approiate form
    Unload Me 'unloads unneeded form
    
End Sub

Private Sub Form_Load()
    'centers form
    Left = (Screen.Width - Me.Width) \ 2
    Top = (Screen.Height - Me.Height) \ 2
End Sub
