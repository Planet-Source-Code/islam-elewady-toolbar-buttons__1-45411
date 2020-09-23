VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CIEToolbarButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit
'this is the encapsulation class
'of the Internet Explorer Standard toolbar
'button.
'author: islam elewady
'date: mar 8 2003
'email: islam@mshawki.com

'Please have a look at the description of each
'method property in the Object Browser

Private Type PicIcon
   Size As Long
   Type As PictureTypeConstants
   hIcon As Long
End Type
Private Type DispGUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Enum enumIEButtonFunctions
    ComObject = 1
    ExplorerBar = 2
    ScriptFile = 3
    ExecutableFile = 4
End Enum
Public Enum enumAddToMenu
    HelpMenu = 1
    ToolsMenu = 2
End Enum

Const mcsKeyName = "Software\Microsoft\Internet Explorer\Extensions\"
Private Declare Function CoCreateGuid Lib "ole32" (id As Any) As Long
Private Declare Function ExtractIconEx Lib "shell32.dll" Alias "ExtractIconExA" (ByVal lpszFile As String, ByVal nIconIndex As Long, phiconLarge As Long, phiconSmall As Long, ByVal nIcons As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicIcon, RefIID As DispGUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Private mvarDefaultVisible As Boolean
Private mvarButtonText As String
Private mvarHotIcon As String
Private mvarIcon As String
Private mvarButtonFunction As enumIEButtonFunctions
Private mvarButtonValue As String

Private mvarAddToMenu As enumAddToMenu
Private mvarMenuText As String
Private mvarMenuStatusBarText As String
Private mvarButtonToolTip As String

Private mcRegistryKey As EnumRegistryRootKeys
Private msCurKey As String
Private msGuid As String
Public Property Get ToolTip() As String
Attribute ToolTip.VB_Description = "The tool tip will show up when you hover your mouse over the button."
   ToolTip = mvarButtonToolTip
End Property
Public Property Let ToolTip(ByVal vData As String)
   mvarButtonToolTip = vData
End Property

Public Property Get AddToMenu() As enumAddToMenu
Attribute AddToMenu.VB_Description = "Defaults to the Tools menu.It can be Help menu to add the menu item to the Help menu"
   AddToMenu = mvarAddToMenu
End Property
Public Property Let AddToMenu(ByVal vData As enumAddToMenu)
   mvarAddToMenu = vData
End Property
Public Function GetIconPic() As StdPicture
Dim spath As String
Dim lIcon As Long, R As Long
Dim pos As Long, resNo As Long
Dim IID_IDispatch As DispGUID
Dim pic As PicIcon
Dim IPic As IPicture

pos = InStr(mvarIcon, ",")
If pos Then
      spath = Left$(mvarIcon, pos - 1)
      'this is the supposed Resource ID
      resNo = Right(mvarIcon, Len(mvarIcon) - pos)
      'as the resource id returns nothing ...see beginning
      'for now i will return the first found
      'icon resource in the module
      ExtractIconEx spath, 0&, lIcon, ByVal 0&, 1
      
   With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    'Fill picture info
    With pic
        .Size = Len(pic) ' Length of structure
        .Type = vbPicTypeIcon ' Type of Picture (icon)
        .hIcon = lIcon     'handle to loaded icon
    End With
    'create Ole compatible picture
   R = OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)
   Set GetIconPic = IPic
      'free the resource
   DestroyIcon lIcon
Else
   'no resource id is needed, just use
   'the vba function to load the icon
      Set GetIconPic = LoadPicture(mvarIcon)
End If
End Function
Public Function GetHotIconPic() As StdPicture
'i have a problem here
'if the icon is in a resource the
' format of the hoticon string should be
' filename,resourceid
'the resource id after the filename is not working
'to retrieve the icon with ExtractIconEx or
'LoadImage or FindResource or anything
'if you got an idea drop me an email please.
'for now the function will return the first
'icon resource found in the module.
Dim spath As String
Dim lIcon As Long, R As Long
Dim pos As Long, resNo As Long
Dim IID_IDispatch As DispGUID
Dim pic As PicIcon
Dim IPic As IPicture


pos = InStr(mvarHotIcon, ",")
If pos Then
      spath = Left$(mvarHotIcon, pos - 1)
      resNo = Right(mvarHotIcon, Len(mvarIcon) - pos)

      ExtractIconEx spath, 0&, lIcon, ByVal 0&, 1
   With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    'Fill picture info
    With pic
        .Size = Len(pic) ' Length of structure
        .Type = vbPicTypeIcon ' Type of Picture (icon)
        .hIcon = lIcon 'handle to icon
    End With
    'create Ole compatible picture
   R = OleCreatePictureIndirect(pic, IID_IDispatch, 1, IPic)
   Set GetHotIconPic = IPic
'free the resource
   DestroyIcon lIcon
Else
   'no resource id is needed, just use
   'the vba function to load the icon
      Set GetHotIconPic = LoadPicture(mvarHotIcon)
End If
End Function


Public Property Let MenuStatusBarText(ByVal vData As String)
Attribute MenuStatusBarText.VB_Description = "The text would appear in the status bar of IE when you hover over the menu item."
    mvarMenuStatusBarText = vData
End Property
Public Property Get MenuStatusBarText() As String
    MenuStatusBarText = mvarMenuStatusBarText
End Property
Public Property Get GUID() As String
   GUID = msGuid
End Property
Public Property Let MenuText(ByVal vData As String)
Attribute MenuText.VB_Description = "required if you want a menu item to appear for the button appended to the Tools menu or as specified by AddToMenu"
    mvarMenuText = vData
End Property
Public Property Get MenuText() As String
    MenuText = mvarMenuText
End Property
Public Property Let Value(ByVal vData As String)
Attribute Value.VB_Description = "It Depends on the value of functionlity .Set to the GUID of the COM object, CLSID of the Explorer Bar you want to open,the full path of the script that will be run,the full path of the .exe file that will be run, or the GUID of a button you want to unins"
    mvarButtonValue = vData
End Property
Public Property Get Value() As String
    Value = mvarButtonValue
End Property
Friend Sub Uninstall()
Attribute Uninstall.VB_Description = "set buttonvalue to the GUID of a button you want to uninstall."
If Len(mvarButtonValue) <> 0 Then
   If ValidGUIDFormat Then
      UnRegisterButton
   End If
End If
End Sub
Friend Sub LoadButton(ByVal sKeyGuid As String)
Dim sValuesNames() As String
Dim sValue As String
Dim nValuesCount As Long
Dim sValueName As String
Dim i As Long

 modRegistry.RegistryEnumerateValues mcRegistryKey, _
            mcsKeyName & sKeyGuid, sValuesNames, nValuesCount
            For i = 0 To nValuesCount - 1
               sValueName = sValuesNames(i)
               sValue = modRegistry.RegistryGetKeyValue(mcRegistryKey, mcsKeyName & sKeyGuid, sValueName)
               SetProperty sValueName, sValue
            Next
msGuid = sKeyGuid
End Sub
Private Sub SetProperty(ByVal PropertyName As String, ByVal PropertyValue As Variant)
Dim sPropName As String
sPropName = LCase$(PropertyName)
Select Case sPropName
   Case "buttontext"
      mvarButtonText = PropertyValue
   Case "default visible"
      If UCase$(PropertyValue) = "YES" Then
         mvarDefaultVisible = True
      Else
         mvarDefaultVisible = False
      End If
   Case "tooltip"
      mvarButtonToolTip = PropertyValue
   Case "menutext"
      mvarMenuText = PropertyValue
   Case "menustatusbar"
      mvarMenuStatusBarText = PropertyValue
   Case "menucustomize"
      If PropertyValue = "help" Then
         mvarAddToMenu = HelpMenu
      Else
         mvarAddToMenu = ToolsMenu
      End If
   Case "exec"
      mvarButtonFunction = ExecutableFile
      mvarButtonValue = PropertyValue
   Case "script"
      mvarButtonFunction = ScriptFile
      mvarButtonValue = PropertyValue
   Case "clsidextension"
      mvarButtonFunction = ComObject
      mvarButtonValue = PropertyValue
   Case "bandclsid"
      mvarButtonFunction = ExplorerBar
      mvarButtonValue = PropertyValue
   Case "hoticon"
      mvarHotIcon = PropertyValue
   Case "icon"
      mvarIcon = PropertyValue
End Select

   

End Sub
Private Function ValidGUIDFormat() As Boolean
Dim pos As Long
Dim isValid As Boolean
Dim sTest As String
sTest = mvarButtonValue
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
ValidGUIDFormat = isValid
End Function
Private Function CreateGUID() As String
    Dim id(0 To 15) As Byte
    Dim Cnt As Long, GUID As String
    If CoCreateGuid(id(0)) = 0 Then
        For Cnt = 0 To 15
            CreateGUID = CreateGUID + IIf(id(Cnt) < 16, "0", "") + Hex$(id(Cnt))
        Next Cnt
        CreateGUID = Left$(CreateGUID, 8) + "-" + Mid$(CreateGUID, 9, 4) + "-" + Mid$(CreateGUID, 13, 4) + "-" + Mid$(CreateGUID, 17, 4) + "-" + Right$(CreateGUID, 12)
   End If
End Function
Friend Sub Install()
Attribute Install.VB_Description = "installs the button with specified properties."
If AllValuesSet Then
   RegisterButton
End If
End Sub
Private Sub RegisterButton()

'first step
CreateCurRegistryKey
'second step
AddValueStringToRegistry "Default Visible", IIf(mvarDefaultVisible = True, "Yes", "No")
'third step
AddValueStringToRegistry "ButtonText", mvarButtonText
'fourth step
AddValueStringToRegistry "HotIcon", mvarHotIcon
'fifth step
AddValueStringToRegistry "Icon", mvarIcon

'menu addin
If Len(Trim$(mvarMenuText)) > 0 Then
   AddValueStringToRegistry "MenuText", mvarMenuText
   If mvarAddToMenu = HelpMenu Then
AddValueStringToRegistry "MenuCustomize", "help"
   End If
   If Len(Trim$(mvarMenuStatusBarText)) > 0 Then
AddValueStringToRegistry "MenuStatusBar", mvarMenuStatusBarText
   End If
End If
If Len(Trim$(mvarButtonToolTip)) > 0 Then
   AddValueStringToRegistry "ToolTip", mvarButtonToolTip
End If
'The details
Select Case mvarButtonFunction
   Case ComObject
AddValueStringToRegistry "ClsID", "{1FBA04EE-3024-11d2-8F1F-0000F87ABD16}"
AddValueStringToRegistry "ClsidExtension", mvarButtonValue
   Case ExplorerBar
AddValueStringToRegistry "ClsID", "{E0DD6CAB-2D10-11D2-8F1A-0000F87ABD16}"
AddValueStringToRegistry "BandCLSID", mvarButtonValue
   Case ScriptFile
AddValueStringToRegistry "ClsID", "{1FBA04EE-3024-11D2-8F1F-0000F87ABD16}"
AddValueStringToRegistry "Script", mvarButtonValue
   Case ExecutableFile
AddValueStringToRegistry "ClsID", "{1FBA04EE-3024-11D2-8F1F-0000F87ABD16}"
AddValueStringToRegistry "Exec", mvarButtonValue
End Select


End Sub
Private Sub AddValueStringToRegistry(ByVal ValueName As String, ByVal ValueData As String)
modRegistry.RegistrySetKeyValue mcRegistryKey, msCurKey, ValueName, ValueData, rrkRegSZ
End Sub
Private Sub CreateCurRegistryKey()

msGuid = "{" & CreateGUID & "}"
msCurKey = mcsKeyName & msGuid
modRegistry.RegistryCreateNewKey mcRegistryKey, msCurKey
End Sub
Private Sub UnRegisterButton()
Dim sGuid As String
Dim KeyName As String
sGuid = mvarButtonValue
KeyName = mcsKeyName & sGuid

modRegistry.RegistryDeleteKey rrkHKeyLocalMachine, KeyName
modRegistry.RegistryDeleteKey rrkHKeyCurrentUser, KeyName
End Sub
Private Function AllValuesSet() As Boolean
   AllValuesSet = True
End Function
Public Property Let Functionlity(ByVal vData As enumIEButtonFunctions)
Attribute Functionlity.VB_Description = "the functionality can be one of the enum."
    mvarButtonFunction = vData
End Property
Public Property Get Functionlity() As enumIEButtonFunctions
     Functionlity = mvarButtonFunction
End Property
Public Property Let Icon(ByVal vData As String)
Attribute Icon.VB_Description = "The full path of the .ico file that contains the three grayscale icons.\r\n16x16 16-shade (Required) \r\n20x20 16-shade (Optional) \r\n20x20 256-shade (Required) \r\nIt can be in the format filename,resourceid to load the icon from resource at run time."
    mvarIcon = vData
End Property
Public Property Get Icon() As String
   Icon = mvarIcon
End Property
Public Property Let HotIcon(ByVal vData As String)
Attribute HotIcon.VB_Description = "the full path of the .ico file that contains the three color icons.   16x16 16-color icon (Required) \r\n20x20 16-color icon (Optional) \r\n20x20 256-color icon (Required) \r\nIt can be in the format filename,resourceid to load the icon from resource at run t"
   mvarHotIcon = vData
End Property
Public Property Get HotIcon() As String
   HotIcon = mvarHotIcon
End Property
Public Property Let Text(ByVal vData As String)
Attribute Text.VB_Description = "the label you want for the toolbar button."
    mvarButtonText = vData
End Property
Public Property Get Text() As String
    Text = mvarButtonText
End Property
Public Property Let DefaultVisible(ByVal vData As Boolean)
Attribute DefaultVisible.VB_Description = "To make the toolbar button to appear on the Internet Explorer toolbar by default, set Default Visible to ""Yes"", otherwise set Default Visible to ""No""."
    mvarDefaultVisible = vData
End Property
Public Property Get DefaultVisible() As Boolean
    DefaultVisible = mvarDefaultVisible
End Property
Public Property Get CurrentUserButton() As Boolean
Attribute CurrentUserButton.VB_Description = "if True only the current logged user would see the button.otherwise, it will appear for all users."
If mcRegistryKey = rrkHKeyCurrentUser Then
   CurrentUserButton = True
Else
   CurrentUserButton = False
End If
End Property
Public Property Let CurrentUserButton(ByVal vData As Boolean)
If vData = True Then
   mcRegistryKey = rrkHKeyCurrentUser
Else
   mcRegistryKey = rrkHKeyLocalMachine
End If
End Property
Private Sub Class_Initialize()
mvarAddToMenu = ToolsMenu
mcRegistryKey = rrkHKeyLocalMachine
End Sub
