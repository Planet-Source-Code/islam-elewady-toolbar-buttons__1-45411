VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIEToolBarButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Attribute VB_Ext_KEY = "Member0" ,"CIEToolbarButton"
Option Explicit
'this is the collection of all the extension
'toolbar buttons
'found in the standard toolbar
'you access a specific toolbar button by its
'GUID which is unique and is returned
'when you enumerate through the collection

'author: islam elewady
'date: mar 8 2003
'email: islam@mshawki.com

'Please have a look at the description of each
'method property in the Object Browser
Const mcsKeyName = "Software\Microsoft\Internet Explorer\Extensions\"
Private mCol As Collection
Public Property Get ToolbarButton(ByVal GUID As String) As CIEToolbarButton
Attribute ToolbarButton.VB_Description = "Access a specific toolbar button with the GUID ."
Attribute ToolbarButton.VB_UserMemId = 0
    Set ToolbarButton = mCol(GUID)
End Property
Public Property Get Count() As Long
Attribute Count.VB_Description = "The count of the toolbar buttons found."
    Count = mCol.Count
End Property
Private Sub Class_Initialize()
   Set mCol = New Collection
   'LoadCollection
End Sub
Private Sub Class_Terminate()
   Set mCol = Nothing
End Sub
Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "440"
    Set NewEnum = mCol.[_NewEnum]
End Property
Public Sub UninstallButton(Button As CIEToolbarButton)
Attribute UninstallButton.VB_Description = "Uninstalls an existing button.The Value contains the guid to the button you want to install."
Button.Uninstall
End Sub
Public Sub InstallButton(NewButton As CIEToolbarButton)
Attribute InstallButton.VB_Description = "Installs a new created button."
NewButton.Install
End Sub
Public Sub Refresh()
Attribute Refresh.VB_Description = "Refreshes the collection of buttons.Should be called before accessing a specific button to regenerate the buttons."
ClearCol
LoadCollection
End Sub
Private Sub ClearCol()
Dim i As Long

For i = 0 To mCol.Count - 1
   mCol.Remove 1
Next

End Sub
Private Sub LoadCollection()
Dim sKeys() As String
Dim nKeysCount As Long

Dim sKeyGuid As String
Dim oNewMember As CIEToolbarButton
Dim tenkey As EnumRegistryRootKeys
Dim bcurUser As Boolean
Dim i As Integer, j As Integer
'first enumerate local machine keys
'then current user
For j = 0 To 1
If j = 0 Then
   tenkey = rrkHKeyLocalMachine
   bcurUser = False
Else
   bcurUser = True
   tenkey = rrkHKeyCurrentUser
End If
modRegistry.RegistryEnumerateSubKeys tenkey, mcsKeyName, sKeys, nKeysCount
   For i = 0 To nKeysCount - 1
      sKeyGuid = sKeys(i)
      If IsValidGuid(sKeyGuid) Then
      Set oNewMember = New CIEToolbarButton
      oNewMember.CurrentUserButton = bcurUser
      oNewMember.LoadButton sKeyGuid
      
      mCol.Add oNewMember, sKeyGuid
      End If
   Next
Next
End Sub
Private Function IsValidGuid(ByVal TestStr As String) As Boolean

'this is done as i was almost asleep

Dim pos As Long
Dim isValid As Boolean
Dim sTest As String
sTest = TestStr
pos = InStr(sTest, "-")

If pos = 10 Then
   pos = InStr(pos + 1, sTest, "-", vbTextCompare)
   
   If pos = 15 Then
      pos = InStr(pos + 1, sTest, "-", vbTextCompare)
      If pos = 20 Then
         
         pos = InStr(pos + 1, sTest, "-", vbTextCompare)
         If pos = 25 Then
         
            pos = Len(Mid$(sTest, 26))
            If pos = 13 Then
               isValid = True
            End If
         End If
     End If
   End If

Else
   isValid = False
End If
IsValidGuid = isValid
End Function
