Attribute VB_Name = "gMain"
'-------------------------------------------------------------
' Module:    gMain
'
' Purpose:  Contains core functionality of the application, and
'           application-specific utility functions.
'
' Dependencies:
'
' For issues and suggestions mail to:
'   Chris Anderson
'   cander@realworldis.com
'-------------------------------------------------------------
Option Explicit

Private Const MODULE_NAME As String = "gMain"

'global instance of the main form
Public fBrowser As frmBrowser

'-------------------------------------------------------------
' Sub Main() Function
'-------------------------------------------------------------

Sub Main()
    
    'instantiate the browser window form
    Set fBrowser = New frmBrowser
    
    'show it
    fBrowser.Show
    
End Sub



