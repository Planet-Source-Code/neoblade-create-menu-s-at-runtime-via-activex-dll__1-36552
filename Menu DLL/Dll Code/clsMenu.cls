VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Caption As String
Public Id As Long
Public Pos As Long
Public Hwnd As Long

Private SubMenuInfo() As clsMenu
Private SubMenuPropId As Long


Public Sub AddMenu(newCaption As String)
Dim MenuHwnd As Long, SubMenuHwnd As Long, NewMenuId As Long, ret As Long

'''Set New Menu ID
NewMenuId = UBound(MenuId) + 1
ReDim Preserve MenuId(NewMenuId)
'''sets the menuid to the current caption for event purposes
MenuId(NewMenuId) = newCaption

'if menucount is <=0 then make top menu a popup
'else just add menu to popup
If Count <= 0 Then
    ''Get Menu from stored hwnd
    MenuHwnd = GetMenu(FormHwnd)

    'Create new blank popup menu to fill
    SubMenuHwnd = CreatePopupMenu

    'add new menu to created popup
    Call AppendMenu(SubMenuHwnd, MF_STRING, NewMenuId, newCaption)

    'modify origin menu so it is now a popup link
    Call ModifyMenu(MenuHwnd, Id, MF_STRING Or MF_POPUP, SubMenuHwnd, Caption)

Else
    'get handle for the menu via handle of popup and position
    SubMenuHwnd = GetSubMenu(Hwnd, Pos)
    'add menu to popup
    Call AppendMenu(SubMenuHwnd, MF_STRING, NewMenuId, newCaption)
End If

''refresh menubar to show changes
Call DrawMenuBar(FormHwnd)

''set all stored data
ReDim Preserve SubMenuInfo(SubMenuPropId)

Set SubMenuInfo(SubMenuPropId) = New clsMenu

SubMenuInfo(SubMenuPropId).Caption = newCaption
SubMenuInfo(SubMenuPropId).Id = NewMenuId
SubMenuInfo(SubMenuPropId).Pos = SubMenuPropId
SubMenuInfo(SubMenuPropId).Hwnd = SubMenuHwnd

'set next propertyid
SubMenuPropId = UBound(SubMenuInfo) + 1


End Sub


Public Function Count()
Count = GetMenuItemCount(GetSubMenu(Hwnd, Pos))
End Function

Public Property Get SubMenu(Index As Long) As clsMenu
'sets the current submenu trying to be recieved, to the submenuinfo associated with it
Set SubMenu = SubMenuInfo(Index)
End Property

