VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAccessibleSites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'-------------------------------------------------------------
' Module:    cAccessibleSites
'
' Purpose:  Provides the interface to browse the list
' of application-specific sites
'
' Dependencies: Microsoft XML parser v3.0
'
' For issues and suggestions mail to:
'   Chris Anderson
'   cander@realworldis.com
'-------------------------------------------------------------

Option Explicit

'xml set of accessible sites
Private xmlSites As MSXML2.DOMDocument30

'-------------------------------------------------------------
' Public Functions
'-------------------------------------------------------------

'browse the list of sites

Public Function getAccessibleList() As String()
On Error GoTo ErrHdl

    'return array
    Dim sSites() As String

    'select the root node
    Dim rootNode As MSXML2.IXMLDOMNode
    Set rootNode = xmlSites.selectSingleNode("sites")
    
    'if the root node was selected
    If Not (rootNode Is Nothing) Then
    
        'get the collection of site nodes
        Dim siteNodes As MSXML2.IXMLDOMNodeList
        Set siteNodes = rootNode.selectNodes("site")
    
        'determine how many sites are accessible
        Dim iSiteCount As Integer
        iSiteCount = siteNodes.length
    
        'redimension the array for return
        ReDim sSites(1, iSiteCount - 1)
    
        'loop through the list of sites
        Dim i As Integer
        For i = 0 To siteNodes.length - 1
            'pick out the name
            sSites(0, i) = _
                siteNodes.Item(i).selectSingleNode("name").Text
            'pick out the URL
            sSites(1, i) = _
                stripURL(siteNodes.Item(i).selectSingleNode("url").Text)
        Next i
    
        'return the array
        getAccessibleList = sSites
        
    End If
    
Exit Function
ErrHdl:

End Function

Public Function getHomePage() As String
On Error GoTo ErrHdl

    'get the home page value
    getHomePage = stripURL(xmlSites.selectSingleNode("sites/homepage").Text)
    
Exit Function
ErrHdl:
    getHomePage = ""

End Function

Public Function isAccessible(ByVal pURL As String) As Boolean
On Error GoTo ErrHdl

    'select the root node
    Dim rootNode As MSXML2.IXMLDOMNode
    Set rootNode = xmlSites.selectSingleNode("sites")
    
    'construct the XPath query string
    Dim sQuery As String
    sQuery = "site/url[contains(., '" & pURL & "')]"
    
    'determine if this is in the list of site
    Dim matchingNodes As MSXML2.IXMLDOMNodeList
    Set matchingNodes = rootNode.selectNodes(sQuery)
    
    Dim iCount As Integer
    iCount = matchingNodes.length
    
    'if there was at least on match
    If iCount > 0 Then
        'return true
        isAccessible = True
    Else
        'return false
        isAccessible = False
    End If
    
Exit Function
ErrHdl:
    'return false
    isAccessible = False

End Function

'-------------------------------------------------------------
' Class Functions
'-------------------------------------------------------------

Private Sub Class_Initialize()
    
    'instantiate the module XML doc
    Set xmlSites = New MSXML2.DOMDocument30
    xmlSites.setProperty "SelectionLanguage", "XPath"
    
    'read in a list of accessible sites
    xmlSites.Load App.Path & "\SiteList.xml"
    
    'if the site list was not found
    If xmlSites.xml = "" Then
        'raise an error
        MsgBox "There are no web sites allowed at this time."
    End If
    
End Sub

Private Sub Class_Terminate()

    'close the list of accessible sites
    Set xmlSites = Nothing
    
End Sub

'-------------------------------------------------------------
' Utility Functions
'-------------------------------------------------------------
Private Function stripURL(ByVal pURL As String) As String
On Error GoTo ErrHdl

    Dim sURL As String
    sURL = pURL
    
    'strip off the '[CDATA[' and the ']]'
    sURL = Mid$(sURL, 8)
    sURL = Mid$(sURL, 1, Len(sURL) - 2)
    
    'return the cleaned url
    stripURL = sURL
    
Exit Function
ErrHdl:
    stripURL = ""
    
End Function

Private Function wrapURL(ByVal pURL As String) As String

    'add the '[CDATA[' and the ']]'
    wrapURL = "[CDATA[" & pURL & "]]"
    
End Function
