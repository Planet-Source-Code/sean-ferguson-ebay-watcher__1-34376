VERSION 5.00
Begin VB.Form frmWatchConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Items Being Watched"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   Icon            =   "frmWatchConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstWatch 
      Height          =   3765
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   4080
      Width           =   7695
   End
   Begin VB.CommandButton cmdRemove 
      Cancel          =   -1  'True
      Caption         =   "&Remove Item"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   3720
      Width           =   3855
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add Item"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3720
      Width           =   3855
   End
End
Attribute VB_Name = "frmWatchConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    On Error Resume Next
    frmAddItem.Show vbModal, Me
End Sub

Private Sub cmdClose_Click()
    On Error Resume Next
    Dim watchItem, itemWatches
    If MsgBox("Do you want to save your watch list?", vbYesNo + vbQuestion, "Save Watch List") = vbYes Then
        On Error GoTo deleteWatchesKey
        itemWatches = GetAllKeys(HKEY_CURRENT_USER, "Software\PCSCT Software\eBay Watcher\Watches")
        If getUBound(itemWatches) > -1 Then
            For Each watchItem In itemWatches
                DeleteKey HKEY_CURRENT_USER, "Software\PCSCT Software\eBay Watcher\Watches\" & watchItem
            Next
        End If
deleteWatchesKey:
        On Error Resume Next
        DeleteKey HKEY_CURRENT_USER, "Software\PCSCT Software\eBay Watcher\Watches"
        If lstWatch.ListCount > 0 Then
            Dim lngItem As Integer
            For lngItem = 0 To (lstWatch.ListCount - 1)
                SaveSettingString HKEY_CURRENT_USER, "Software\PCSCT Software\eBay Watcher\Watches\" & lstWatch.List(lngItem), "Item", lstWatch.List(lngItem)
            Next lngItem
        End If
    End If
    Unload Me
End Sub

Private Sub cmdRemove_Click()
    On Error Resume Next
    If lstWatch.ListIndex > -1 Then
        If MsgBox("Are you sure you want to remove this item from your watch list?", vbYesNo + vbQuestion, "Remove Item") = vbYes Then
            lstWatch.RemoveItem lstWatch.ListIndex
        End If
    Else
        MsgBox "You must select an item to remove.", vbOKOnly + vbInformation, "Remove Item"
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo endLoad
    Dim itemWatches, watchItem
    Dim lI As ListItem
    itemWatches = GetAllKeys(HKEY_CURRENT_USER, "Software\PCSCT Software\eBay Watcher\Watches")
    If getUBound(itemWatches) > -1 Then
        For Each watchItem In itemWatches
            lstWatch.AddItem watchItem
        Next
    End If
    
endLoad:
    ' Exit Sub
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then Cancel = 1
End Sub
