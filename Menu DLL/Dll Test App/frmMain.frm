VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Dynamic Menus via Dll"
   ClientHeight    =   1950
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   4980
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   45
   End
   Begin VB.Menu mnuTest 
      Caption         =   "Test"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'NOTE: A menu must be added to the form using traditional methods
'any entries in this menu will be deleted, it just needs the object initially

'Make sure withevents is declared to get the mouse events
Private WithEvents MenuBar As clsMenuBar
Attribute MenuBar.VB_VarHelpID = -1


Private Sub Form_Load()

Set MenuBar = New clsMenuBar


With MenuBar
.Create Me.hWnd
.AddMenu "MENU 0"
.AddMenu "MENU 1"
.AddMenu "MENU 2"

.SubMenu(0).AddMenu "MENU 0 - SUBMENU 0"

.SubMenu(1).AddMenu "MENU 1 - SUBMENU 0"
.SubMenu(1).AddMenu "MENU 1 - SUBMENU 1"
.SubMenu(1).SubMenu(1).AddMenu "MENU 1 - SUB-SUBMENU 0"
.SubMenu(1).AddMenu "MENU 1 - SUBMENU 2"

.SubMenu(2).AddMenu "MENU 2 - SUBMENU 0"
.SubMenu(2).SubMenu(0).AddMenu "MENU 2 - SUB-SUBMENU 0"
.SubMenu(2).SubMenu(0).SubMenu(0).AddMenu "MENU 2 - SUB-SUB-SUBMENU 0"

lblCount.Caption = .SubMenu(1).Caption & " contains " & .SubMenu(1).Count & " menu's"

End With
End Sub


Private Sub MenuBar_MenuClicked(MenuId As Long, Caption As String)
MsgBox "Caption: " & Caption & vbCrLf & "MenuId: " & MenuId
End Sub
