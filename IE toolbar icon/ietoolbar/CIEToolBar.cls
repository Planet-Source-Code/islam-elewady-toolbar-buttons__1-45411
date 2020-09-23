VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIEToolBar"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"CIEToolBarButtons"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
' This is the top of the Hierarchy
' the object model like this
'Toolbar ' a Global Multiuse
'      ----Toolbarbuttons
'      -------ToolBarButton

'author: islam elewady
'date: mar 8 2003
'email: islam@mshawki.com

'Please have a look at the description of each
'method ,property in the Object Browser
Private mvarCIEToolBarButtons As CIEToolBarButtons
Private Sub Class_Initialize()
    Set mvarCIEToolBarButtons = New CIEToolBarButtons
End Sub
Public Property Get IEToolBarButtons() As CIEToolBarButtons
Attribute IEToolBarButtons.VB_Description = "The access point for the collection of extension buttons in IE."
    Set IEToolBarButtons = mvarCIEToolBarButtons
End Property
Private Sub Class_Terminate()
    Set mvarCIEToolBarButtons = Nothing
End Sub



