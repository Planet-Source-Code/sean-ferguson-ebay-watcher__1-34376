VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmResource 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   615
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   1590
   Icon            =   "frmResource.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   615
   ScaleWidth      =   1590
   ShowInTaskbar   =   0   'False
   Begin InetCtlsObjects.Inet iNet 
      Left            =   480
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer tmrCheckBids 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   0
      Top             =   0
   End
   Begin VB.Menu mnuOptions 
      Caption         =   ""
      Begin VB.Menu mnuShow 
         Caption         =   "Show Watch Window"
      End
      Begin VB.Menu mnuGetItemInfo 
         Caption         =   "Update Item Info"
      End
      Begin VB.Menu mnuSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetUpWatch 
         Caption         =   "Set-Up Watch Items"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lngMinutes As Integer

Private Sub Form_Load()
    On Error Resume Next
    AddToTray frmWatch.Icon, frmWatch.Caption, Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Select Case RespondToTray(X)
        Case 1
            frmWatch.Show
        Case 2
            PopupMenu mnuOptions
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    RemoveFromTray
    Unload Me
End Sub

Private Sub mnuExit_Click()
    Unload frmWatch
End Sub

Private Sub mnuGetItemInfo_Click()
    updateBids
End Sub

Private Sub mnuSetUpWatch_Click()
    On Error Resume Next
    tmrCheckBids.Enabled = False
    frmWatchConfig.Show vbModal, frmWatch
    tmrCheckBids.Enabled = True
    frmWatch.updateItems
End Sub

Private Sub mnuShow_Click()
    On Error Resume Next
    frmWatch.Show
End Sub

Private Sub tmrCheckBids_Timer()
    lngMinutes = lngMinutes + 1
    If lngMinutes >= 5 Then
        lngMinutes = 0
        updateBids
    End If
End Sub

Public Function updateBids()
    On Error GoTo handleError
    Dim strPageURL As String
    Dim iInfo() As String
    Dim itemID As Long
    Dim strData As String
    If frmWatch.lstWatching.ListItems.Count > 0 Then
        mnuSetUpWatch.Enabled = False
        tmrCheckBids.Enabled = False
        mnuGetItemInfo.Enabled = False
        If frmWatch.lstWatching.ListItems.Count > 0 Then
            For itemID = 1 To frmWatch.lstWatching.ListItems.Count
                If frmWatch.lstWatching.ListItems(itemID).ListSubItems(2) = "" Then frmWatch.lstWatching.ListItems(itemID).Tag = "-1" Else frmWatch.lstWatching.ListItems(itemID).Tag = frmWatch.lstWatching.ListItems(itemID).ListSubItems(2)
                frmWatch.lstWatching.ListItems(itemID).ListSubItems(1).Text = "Updating..."
                frmWatch.lstWatching.ListItems(itemID).ListSubItems(2).Text = ""
                frmWatch.lstWatching.ListItems(itemID).ListSubItems(3).Text = ""
            Next itemID
            For itemID = 1 To frmWatch.lstWatching.ListItems.Count
                strPageURL = "http://cgi.ebay.com/aw-cgi/eBayISAPI.dll?ViewItem&item=" & frmWatch.lstWatching.ListItems(itemID)
                strData = iNet.openURL(strPageURL)
                iInfo() = returnInfoArray(strData)
                frmWatch.lstWatching.ListItems(itemID).ListSubItems(1).Text = iInfo(1)
                frmWatch.lstWatching.ListItems(itemID).ListSubItems(2).Text = iInfo(2)
                frmWatch.lstWatching.ListItems(itemID).ListSubItems(3).Text = iInfo(3)
                If CLng(Val(frmWatch.lstWatching.ListItems(itemID).Tag)) > -1 Then
                    If CLng(Val(iInfo(2))) > CLng(Val(frmWatch.lstWatching.ListItems(itemID).Tag)) Then
                        alertChange CInt(itemID)
                    End If
                End If
                frmWatch.lstWatching.ListItems(itemID).Tag = ""
            Next itemID
        End If
        mnuSetUpWatch.Enabled = True
        tmrCheckBids.Enabled = True
        mnuGetItemInfo.Enabled = True
        showAlerts = True
    End If
    Exit Function
    
handleError:
    On Error GoTo handleError2
    For itemID = 1 To frmWatch.lstWatching.ListItems.Count
        frmWatch.lstWatching.ListItems(itemID).ListSubItems(1).Text = ""
        frmWatch.lstWatching.ListItems(itemID).ListSubItems(2).Text = ""
        frmWatch.lstWatching.ListItems(itemID).ListSubItems(3).Text = ""
    Next itemID
    mnuSetUpWatch.Enabled = True
    tmrCheckBids.Enabled = True
    mnuGetItemInfo.Enabled = True
    showAlerts = True
    Exit Function
    
handleError2:
    mnuSetUpWatch.Enabled = True
    tmrCheckBids.Enabled = True
    mnuGetItemInfo.Enabled = True
    showAlerts = True
    Exit Function
End Function

Private Function alertChange(lngItemID As Integer)
    On Error Resume Next
    If showAlerts = False Then Exit Function
    Load frmNotify
    frmNotify.itemNumber = frmWatch.lstWatching.ListItems(lngItemID).Text
    frmNotify.itemName = frmWatch.lstWatching.ListItems(lngItemID).ListSubItems(1).Text
    frmNotify.numberOfBids = frmWatch.lstWatching.ListItems(lngItemID).ListSubItems(2).Text
    frmNotify.currentPrice = frmWatch.lstWatching.ListItems(lngItemID).ListSubItems(3).Text
    frmNotify.Show vbModal
End Function

