VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Number of Computers Connected:"
      Height          =   855
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   3015
      Begin VB.Label lblConnected 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2655
      End
   End
   Begin MSWinsockLib.Winsock Winsock 
      Index           =   0
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*|*|*|*|*|*|*|*|*|*| MultiConnect |*|*|*|*|*|*|*|*|*|*|*
'*|*|*|*|*|*|*|*|*|**|*|  Notes  |*|*|*|*|*|*|*|*|*|*|*|*
'The Winsock control on the Server form is in an array
'The reason for the array is because Winsock does not
'support multi-connections alone (simply at least)
'So what we have to do instead is load a new winsock control
'Winsock(0) will always be listening for someone wanting to connect
'while Winsock(x) will establish the actual connection
'You can have an infinite number of connections

Dim intConnection As Integer

Private Sub Form_Load()
    'centers form
    Left = (Screen.Width - Me.Width) \ 2
    Top = (Screen.Height - Me.Height) \ 2
    
    'displays the connection status
    lblConnected.Caption = "0"

    'winsock will listen on port 55
    Winsock(0).LocalPort = 55
    Winsock(0).Listen
End Sub

Private Sub Winsock_Close(Index As Integer)
    lblConnected.Caption = lblConnected.Caption - 1 'takes one off the connection status
End Sub

Private Sub Winsock_ConnectionRequest(Index As Integer, ByVal requestID As Long)
    'adds 1 to the intConnection
    'to keep winsock(0) listening on port 55
    intConnection = intConnection + 1
    
    'loads a NEW winsock with an array > 0
    Load Winsock(intConnection)
    'lets winsock connect to client
    Winsock(intConnection).Accept requestID
    'shows how many computers are connected
    lblConnected.Caption = lblConnected.Caption + 1
End Sub

