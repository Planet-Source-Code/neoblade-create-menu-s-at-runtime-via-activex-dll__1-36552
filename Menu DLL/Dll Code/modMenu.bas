Attribute VB_Name = "modMenu"
Option Explicit

'Menu Functions

Public Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function DrawMenuBar Lib "user32" (ByVal Hwnd As Long) As Long
Public Declare Function SetMenu Lib "user32" (ByVal Hwnd As Long, ByVal hMenu As Long) As Long
Public Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Public Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Public Declare Function CreateMenu Lib "user32" () As Long
Public Declare Function LoadMenu Lib "user32" Alias "LoadMenuA" (ByVal hInstance As Long, ByVal lpString As String) As Long
Public Declare Function GetMenu Lib "user32" (ByVal Hwnd As Long) As Long
Public Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal B As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bypos As Long, lpcMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal Hwnd As Long, ByVal lptpm As Any) As Long

'''Window Functions
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'Error Functions
Public Declare Function GetLastError Lib "kernel32" () As Long



Public Const MF_CHECKED = &H8&
Public Const MF_APPEND = &H100&
Public Const MF_OWNERDRAW = &H100&
Public Const MF_MENUBREAK = &H40&
Public Const MF_MENUBARBREAK = &H20&
Public Const MF_BYPOSITION = &H400&
Public Const MF_POPUP = &H10&
Public Const MF_STRING = &H0&
Public Const MF_DISABLED = &H2&
Public Const MF_GRAYED = &H1&
Public Const MF_SEPARATOR = &H800&


Public Const MIIM_STATE = &H1
Public Const MIIM_ID = &H2
Public Const MIIM_TYPE = &H10

Public Const TPM_RETURNCMD = &H100&
Public Const TPM_RIGHTBUTTON = &H2&
Public Const TPM_LEFTALIGN = &H0&

Public Const GWL_WNDPROC = (-4)

Public Const WM_COMMAND = &H111
Public Const WM_CLOSE = &H10

Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type


Private Type POINTAPI
    X As Long
    Y As Long
End Type

Public MousePoint As POINTAPI

''''''''''''''''''''

''Variable to hold the address of the old window procedure
Public gOldProc As Long

'''''''''''''''''''''

''Holds all the menu ids staring at 70
Public MenuId() As String

'''Holds all the info for the different menus
Public MenuEvent As clsMenuBar

'Holds the handle for the main form
Public FormHwnd As Long





''''''''''''''''''''''''''''''''''''''''''''''''''

''''Where the windows messages are processed
Public Function MenuProc(ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

       Select Case wMsg&
           Case WM_CLOSE:
               ''When the user closes the window, this frees up the resources and gives the handling back to the origibal window
               Call SetWindowLong(Hwnd&, GWL_WNDPROC, gOldProc&)
          
           Case WM_COMMAND:
           '''the wm_command event is when a user click on a menu
           '''''the wparam is the menuid
                    
                  MenuEvent.ClickMenu wParam&
                      
                  
       
       End Select

    ''Call original window procedure for default processing.
    MenuProc = CallWindowProc(gOldProc&, Hwnd&, wMsg&, wParam&, lParam&)

End Function



