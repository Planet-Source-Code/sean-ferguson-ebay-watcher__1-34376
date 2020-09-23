VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmWatch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "eBay Watcher"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "frmWatch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   7695
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   3720
      Width           =   3855
   End
   Begin VB.CommandButton cmdGoTo 
      Caption         =   "&Go to Item Page"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3720
      Width           =   3855
   End
   Begin MSComctlLib.ListView lstWatching 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   6588
      View            =   3
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item #"
         Object.Width           =   2998
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Item Name"
         Object.Width           =   6703
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Bids"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Price"
         Object.Width           =   2204
      EndProperty
   End
End
Attribute VB_Name = "frmWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub cmdGoTo_Click()
    On Error Resume Next
    If isSelectedItem Then
        openURL "http://cgi.ebay.com/aw-cgi/eBayISAPI.dll?ViewItem&item=" & lstWatching.SelectedItem.Text
    End If
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then Unload Me: End
    Load frmResource
    If LCase(Command()) = "/show" Then Me.Visible = True
    updateItems
    frmResource.tmrCheckBids.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If UnloadMode <> 1 Then
        If Me.Visible = False Then
            Unload frmResource
            End
        Else
            Cancel = 1
        End If
    Else
        Unload frmResource
        End
    End If
End Sub

Private Sub lstWatching_BeforeLabelEdit(Cancel As Integer)
    Cancel = 1
End Sub

Private Sub lstWatching_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    On Error Resume Next
    lstWatching.SortKey = (ColumnHeader.Index - 1)
    lstWatching.SortOrder = IIf(lstWatching.SortOrder = lvwAscending, lvwDescending, lvwAscending)
End Sub

Private Sub lstWatching_DblClick()
    cmdGoTo_Click
End Sub

Private Sub lstWatching_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdGoTo_Click
    End If
End Sub

Private Function isSelectedItem() As Boolean
    On Error GoTo handleError
    isSelectedItem = lstWatching.SelectedItem.Selected
    Exit Function
    
handleError:
    isSelectedItem = False
    Exit Function
End Function

Public Function updateItems()
    On Error Resume Next
    Dim itemWatches, watchItem
    showAlerts = False
    lstWatching.ListItems.Clear
    Dim lI As ListItem
    itemWatches = GetAllKeys(HKEY_CURRENT_USER, "Software\PCSCT Software\eBay Watcher\Watches")
    If getUBound(itemWatches) > -1 Then
        For Each watchItem In itemWatches
            Set lI = lstWatching.ListItems.Add(, , watchItem)
            lI.ListSubItems.Add , , ""
            lI.ListSubItems.Add , , ""
            lI.ListSubItems.Add , , ""
        Next
    End If
    frmResource.updateBids
End Function
