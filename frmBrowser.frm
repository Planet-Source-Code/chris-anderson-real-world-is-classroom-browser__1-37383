VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmBrowser 
   Caption         =   "Classroom Browser"
   ClientHeight    =   7605
   ClientLeft      =   3060
   ClientTop       =   3345
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   6480
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cboAddress 
      Height          =   315
      Left            =   600
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   3795
   End
   Begin VB.CommandButton cmdCollapseExpand 
      Caption         =   "<>"
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   720
      Width           =   255
   End
   Begin ComctlLib.ListView lstSites 
      Height          =   3735
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   6588
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin SHDocVwCtl.WebBrowser brwWebBrowser 
      Height          =   3735
      Left            =   2565
      TabIndex        =   0
      Top             =   720
      Width           =   2880
      ExtentX         =   5080
      ExtentY         =   6588
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer timTimer 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   5880
      Top             =   1500
   End
   Begin VB.Label lblAddress 
      Caption         =   "&Address:"
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Tag             =   "&Address:"
      Top             =   5040
      Visible         =   0   'False
      Width           =   3075
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'accessible sites controller
Private mSiteController As cAccessibleSites

'module-level variable for the current address
Private mAddress As String

Dim mbDontNavigateNow As Boolean

'-------------------------------------------------------------
' Properties
'-------------------------------------------------------------

'-----------------------------
'Property:  Address
'Purpose:
'   Manage the address that will be used when this form is displayed
'Public - read-write
'-----------------------------
Public Property Get Address() As String
    
    'return the module-level variable
    Address = mAddress

End Property
Public Property Let Address(ByVal NewValue As String)

    'set the module-level variable
    mAddress = NewValue
    
End Property

'-------------------------------------------------------------
' Event Handlers
'-------------------------------------------------------------

Private Sub brwWebBrowser_BeforeNavigate2(ByVal pDisp As Object, _
    URL As Variant, Flags As Variant, TargetFrameName As Variant, _
    PostData As Variant, Headers As Variant, Cancel As Boolean)
    
    'validate that the page that will be navigated to should be accessible
    If Not mSiteController.isAccessible(cboAddress.Text) Then
        'cancel navigation
        Cancel = True
        'reset the address bar
        If cboAddress.ListCount > 0 Then
            cboAddress.ListIndex = 0
        End If
        'show an error message
        MsgBox "This site is not accessible."
    End If

End Sub

Private Sub brwWebBrowser_DownloadComplete()
On Error Resume Next
    
    'set the form caption to the name of the page downloaded
    Me.Caption = brwWebBrowser.LocationName
    
End Sub

Private Sub brwWebBrowser_NavigateComplete2( _
    ByVal pDisp As Object, URL As Variant)

    'set the form caption to the page now loaded
    Me.Caption = brwWebBrowser.LocationName
    
    'whether this location was found in the combo box
    Dim bFound As Boolean
    
    'loop through the combo box
    Dim i As Integer
    For i = 0 To cboAddress.ListCount - 1
        'if the item matches the page now loaded
        If cboAddress.List(i) = brwWebBrowser.LocationURL Then
            'mark it as found
            bFound = True
            'exit the loop
            Exit For
        End If
    Next i
    
    'set the flag indicating not to navigate
    mbDontNavigateNow = True
    
    'if the page already existed in the combo box
    If bFound Then
        'remove it
        cboAddress.RemoveItem i
    End If
    
    'add the page now loaded to the combo box
    cboAddress.AddItem brwWebBrowser.LocationURL, 0
    cboAddress.ListIndex = 0
    
    'reset the flag indicating the combo box can be used
    'for navigation now
    mbDontNavigateNow = False
    
End Sub

Private Sub cmdCollapseExpand_Click()

    'reverse the current visibility of the sites list
    lstSites.Visible = Not lstSites.Visible
    
    'explicitly call a form resize to resize other controsl
    Form_Resize

End Sub

Private Sub Form_Load()
    
    'line up the address box
    cboAddress.Move 50, lblAddress.Top + lblAddress.Height + 15
    
    'show this form
    Me.Show
    
    'load the accessible site list
    Set mSiteController = New cAccessibleSites
    
    LoadAccessibleSites
    
    'get the home page for this app
    Address = mSiteController.getHomePage
    
    'get the list of accessible sites
    Dim sSites() As String
    sSites = mSiteController.getAccessibleList
    
    'if the address is filled in
    If Len(Address) > 0 Then
        'set the combo box
        With cboAddress
            .Text = Address
            .AddItem .Text
        End With
        'try to navigate to the starting address
        NavigateToPage Address
    Else
        NavigateToPage "about:blank"
    End If

End Sub

Private Sub Form_Resize()
On Error GoTo ErrHdl

    'calculate the space needed for the top of the browser
    Dim iFormOverhead As Integer
    iFormOverhead = 150
    
    'enforce a minimum height
    Dim iMinHeight As Integer
    iMinHeight = iFormOverhead + 500
    If Me.Height < iMinHeight Then
        Me.Height = iMinHeight
    End If
    
    'center the collapse/expand button verticall
    cmdCollapseExpand.Top = (Me.ScaleHeight / 2) - (cmdCollapseExpand.Height / 2)
    
    'if the sites list is visible
    If lstSites.Visible = True Then
        'resize the sites list
        With lstSites
            '.Left =
            '.Width =
            .Top = 0
            .Height = Me.ScaleHeight
        End With
        
        'move the collapse/expand button
        cmdCollapseExpand.Left = lstSites.Left + lstSites.Width + 100
    
        'resize the browser window
        With brwWebBrowser
            .Left = cmdCollapseExpand.Left + cmdCollapseExpand.Width + _
                100
            .Width = Me.ScaleWidth - cmdCollapseExpand.Left - _
                cmdCollapseExpand.Width - 100
            .Top = 0
            .Height = Me.ScaleHeight
        End With
        
    Else
        'move the collapse/expand button
        cmdCollapseExpand.Left = 100
        
        'place and resize the browser window
        With brwWebBrowser
            .Left = cmdCollapseExpand.Left + cmdCollapseExpand.Width + _
                100
            .Width = Me.ScaleWidth - cmdCollapseExpand.Left - _
                cmdCollapseExpand.Width - 100
            .Top = 0
            .Height = Me.ScaleHeight
        End With
        
    End If
    
Exit Sub
ErrHdl:

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'clean up
    Set mSiteController = Nothing
    
End Sub

Private Sub lstSites_Click()
    
    'browse to the url corresponding to the selected item
    NavigateToPage (lstSites.SelectedItem.Tag)
    
End Sub

Private Sub timTimer_Timer()

    'if the browser is busy
    If brwWebBrowser.Busy = True Then
        'show the 'working' caption
        Me.Caption = "Working..."
    Else
        'keep the timer enabled
        timTimer.Enabled = False
        'set the form caption to the page that is loaded
        Me.Caption = brwWebBrowser.LocationName
    End If
    
End Sub

'-------------------------------------------------------------
' Core Functions
'-------------------------------------------------------------

Private Sub LoadAccessibleSites()
On Error GoTo ErrHdl

    'get an array of sites
    Dim sSites() As String
    sSites = mSiteController.getAccessibleList
    
    'setup the columns in the list
    lstSites.ColumnHeaders.Add , , "Web Site"
    
    Dim li As ListItem
    
    'if the array was not empty
    If Not IsEmpty(sSites) Then
        
        'loop through the list of sites
        Dim i As Integer
        For i = 0 To UBound(sSites, 2)
            'add the site to the list box
            'display name
            Set li = lstSites.ListItems.Add(, , sSites(0, i))
        
            li.Tag = sSites(1, i)
        
        Next i
        
    End If
    
    'clean up
    Erase sSites
    Set li = Nothing
    
Exit Sub
ErrHdl:
    'clean up
    Erase sSites
    Set li = Nothing
    
End Sub

Private Sub NavigateToPage(ByVal pURL As String)

    'if the form is not in a navigation state already
    If mbDontNavigateNow = False Then
        'enable the timer
        timTimer.Enabled = True
        'start navigating the internal browser
        brwWebBrowser.Navigate pURL
    End If
    
End Sub

'-------------------------------------------------------------
' Utility Functions
'-------------------------------------------------------------
