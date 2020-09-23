VERSION 5.00
Begin VB.Form frmAddItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Item"
   ClientHeight    =   735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6120
   Icon            =   "frmAddItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   735
   ScaleWidth      =   6120
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   163
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   163
      Width           =   1215
   End
   Begin VB.TextBox txtItemID 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   193
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "Item ID:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmAddItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error Resume Next
    Dim lngItem As Integer
    If frmWatchConfig.lstWatch.ListCount > 0 Then
        For lngItem = 0 To (frmWatchConfig.lstWatch.ListCount - 1)
            If frmWatchConfig.lstWatch.List(lngItem) = txtItemID.Text Then
                MsgBox "This item is already on your watch list.", vbOKOnly + vbInformation, "Add Item"
                Exit Sub
            End If
        Next lngItem
    End If
    frmWatchConfig.lstWatch.AddItem txtItemID.Text
    Unload Me
End Sub
