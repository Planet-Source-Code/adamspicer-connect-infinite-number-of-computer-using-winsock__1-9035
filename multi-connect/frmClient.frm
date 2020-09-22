VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmClient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client"
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4155
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   4155
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   240
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lblConnect 
      Caption         =   "Label1"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
End
Attribute VB_Name = "frmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'centers form
    Left = (Screen.Width - Me.Width) \ 2
    Top = (Screen.Height - Me.Height) \ 2
    
    'displays status of winsock
    lblConnect.Caption = "Not Connected"
    
    'connects to IP address 127.0.0.1
    'through port 55
    Winsock.Connect "127.0.0.1", 55
    
End Sub

Private Sub Winsock_Close()
    'if the connection is closed... displays...
    lblConnect.Caption = "Not Connected"
End Sub

Private Sub Winsock_Connect()
    'tells u that you were connected efficiently
    lblConnect.Caption = "You Are Now Connected to Server!"
    
End Sub

