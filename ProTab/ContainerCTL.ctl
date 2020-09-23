VERSION 5.00
Begin VB.UserControl ContainerCTL 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipControls    =   0   'False
   ControlContainer=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ContainerCTL.ctx":0000
End
Attribute VB_Name = "ContainerCTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

Private enbl As Boolean ', allTabs As Boolean
Private autoColor As Boolean
Private BC As OLE_COLOR

Private Function getContainerBack() As OLE_COLOR
    getContainerBack = GetBkColor(GetDC(UserControl.ContainerHwnd))
End Function
    
'End Function
'----------------------------------------------------------------
'----------------------------------------------------------------
'----------------------------------------------------------------
Public Property Get Enabled() As Boolean
    Enabled = enbl
End Property
Public Property Let Enabled(ByVal nV As Boolean)
    enbl = nV
    reDraw
    PropertyChanged "Enabled"
End Property

Public Property Get AutoBackColor() As Boolean
    AutoBackColor = autoColor
End Property
Public Property Let AutoBackColor(ByVal nV As Boolean)
    autoColor = nV
    reDraw
    PropertyChanged "AutoBackColor"
End Property

'Public Property Get ForAllTabs() As Boolean
'    ForAllTabs = allTabs
'End Property
'Public Property Let ForAllTabs(ByVal nV As Boolean)
'    allTabs = nV
'    reDraw
'    PropertyChanged "ForAllTabs"
'End Property

Public Property Get BackColor() As OLE_COLOR
    BackColor = BC
End Property
Public Property Let BackColor(ByVal nV As OLE_COLOR)
    BC = nV
    UserControl.BackColor = BC
    reDraw
    PropertyChanged "BackColor"
End Property





'----------------------------------------------------------------
'----------------------------------------------------------------
'----------------------------------------------------------------
'
Private Sub UserControl_InitProperties()
    BC = vbButtonFace
    enbl = True
    autoColor = True
    'allTabs = False
End Sub

Private Sub UserControl_Paint()
    reDraw
End Sub

'
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    BC = PropBag.ReadProperty("BackColor", vbButtonFace)
    enbl = PropBag.ReadProperty("Enabled", True)
    autoColor = PropBag.ReadProperty("AutoBackColor", True)
    'allTabs = PropBag.ReadProperty("ForAllTabs", False)
    reDraw
End Sub
'
Private Sub UserControl_Resize()
    reDraw
End Sub
'
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "BackColor", BC, vbButtonFace
    PropBag.WriteProperty "Enabled", enbl, True
    PropBag.WriteProperty "AutoBackColor", autoColor, True
    'PropBag.WriteProperty "ForAllTabs", allTabs, False
End Sub
'
'----------------------------------------------------------------
'----------------------------------------------------------------
'----------------------------------------------------------------
Private Sub reDraw()
    If Ambient.UserMode = True Then
        UserControl.BorderStyle = 0
    Else
        UserControl.BorderStyle = 1
    End If
    UserControl.Enabled = enbl

    PropertyChanged "BackColor"
    
    If autoColor = True Then
        UserControl.BackColor = getContainerBack
        BC = getContainerBack
    Else
        UserControl.BackColor = BC
    End If
    'UserControl.Refresh
End Sub


