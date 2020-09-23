VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenuBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'for every submenu created a new class is initialized to store
'all the data
'ie submenu(0) = clsmenu(0)
'   submenu(0).submenu(0) = clsmenu(0).clsmenu(0)
'
'
'
'
'
'
Option Explicit

Private MenuInfo() As clsMenu
Private MenuPropID As Long

Public Event MenuClicked(MenuId As Long, Caption As String)

Private Sub Hook(FormHwnd As Long)
   gOldProc& = GetWindowLong(FormHwnd, GWL_WNDPROC)
   
   ''Now replace the old window procedure
   Call SetWindowLong(FormHwnd, GWL_WNDPROC, AddressOf MenuProc)

    
End Sub


Public Sub Create(Hwnd As Long)
Dim MenuHwnd As Long, SubMenuHwnd As Long

FormHwnd = Hwnd

''destroys bar created
Call DestroyMenu(GetMenu(FormHwnd))

''creates a new menubar
MenuHwnd = CreateMenu

'sets the bar to form hwnd provided
Call SetMenu(FormHwnd, MenuHwnd)

'refreshs menubar
Call DrawMenuBar(FormHwnd)


'resets the menuinfo to 0 to start
ReDim Preserve MenuInfo(0)
'resets the initial menu id to 69, so no conflicts are caused
ReDim Preserve MenuId(69)


'''hooks the formhwnd provided, so menu clicks can be recorded
Hook FormHwnd
End Sub

Public Sub AddMenu(Caption As String)
Dim MenuHwnd As Long, NewMenuId As Long

''Get Menu from formhwnd
MenuHwnd = GetMenu(FormHwnd)

'''Set New Menu ID
NewMenuId = UBound(MenuId) + 1
ReDim Preserve MenuId(NewMenuId)
'''sets the menuid to the current caption for event purposes
MenuId(NewMenuId) = Caption

'''add item to menu
Call AppendMenu(MenuHwnd, MF_POPUP Or MF_STRING Or MF_OWNERDRAW, NewMenuId, Caption)

'''set the menu against the form hwnd
Call SetMenu(FormHwnd, MenuHwnd)

''refresh the menubar
Call DrawMenuBar(FormHwnd)

''set all stored data
ReDim Preserve MenuInfo(MenuPropID)

Set MenuInfo(MenuPropID) = New clsMenu

MenuInfo(MenuPropID).Caption = Caption
MenuInfo(MenuPropID).Id = NewMenuId
MenuInfo(MenuPropID).Pos = MenuPropID
MenuInfo(MenuPropID).Hwnd = MenuHwnd
'MenuInfo(MenuPropID).Count = GetMenuItemCount(Hwnd)

MenuPropID = UBound(MenuInfo) + 1

End Sub

Public Property Get SubMenu(Index As Long) As clsMenu
'sets the current submenu trying to be recieved, to the menuinfo associated with it
Set SubMenu = MenuInfo(Index)
End Property


Public Sub ClickMenu(ClickedMenuId As Long)
'raises the mouse click event, so it can be linked from the module
RaiseEvent MenuClicked(ClickedMenuId, MenuId(ClickedMenuId))
End Sub


Private Sub Class_Initialize()
'set the menuevent object to be equal to the current
'object created so it can handle the click events
Set MenuEvent = Me
End Sub
