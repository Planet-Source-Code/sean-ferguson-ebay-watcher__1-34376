VERSION 5.00
Begin VB.Form frmNotify 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Bid Notification"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   Icon            =   "frmNotify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   6015
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   2400
      Width           =   3615
   End
   Begin VB.Label currentPrice 
      Height          =   255
      Left            =   1200
      TabIndex        =   8
      Top             =   1680
      Width           =   4695
   End
   Begin VB.Label numberOfBids 
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label itemName 
      Height          =   255
      Left            =   1200
      TabIndex        =   6
      Top             =   960
      Width           =   4695
   End
   Begin VB.Label itemNumber 
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   600
      Width           =   4695
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Current Price:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "An item you are watching has received a new bid!"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "# of Bids:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Item Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Item #:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   975
   End
End
Attribute VB_Name = "frmNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then Cancel = 1
End Sub
