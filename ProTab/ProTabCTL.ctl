VERSION 5.00
Begin VB.UserControl ProTabCTL 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   238
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForwardFocus    =   -1  'True
   PropertyPages   =   "ProTabCTL.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "ProTabCTL.ctx":0058
   Begin VB.Timer tmrSlide 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   960
      Top             =   1020
   End
End
Attribute VB_Name = "ProTabCTL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long


'for color schemes
Private mConColor() As Long
'
Public Enum ePTAppearance
    ptAppearanceFlat
    ptAppearance3D
End Enum
Private mePTAppearance As ePTAppearance


Public Enum ePTStyles
    ptRectangle
    ptRoundedRectangle
    ptRoundedRectangle2
    ptCornerCutLeft
    ptCornerCutRight
    ptCoolLeft
    ptCoolRight
    ptDistorted
    ptVerticalLine
    ptRoundMenu
    ptDistortedMenu
    ptXP
    ptSSTab
    ptFlatButton
    ptProTab
End Enum
Private mePTStylesNormal As ePTStyles, mePTStylesActive As ePTStyles
'
'
Public Enum ePTOrientation
    ptOrientationTop
    ptOrientationBottom
End Enum
Private mePTOrientation As ePTOrientation
'
Public Enum ePTTabStyle
    ptStandard
    ptGraphical
End Enum
Private mePTTabStyle As ePTTabStyle
'
Public Enum ePTHorTextAlg
    ptLeft
    ptCenter
    ptRight
End Enum
Private mePTHorTextAlg As ePTHorTextAlg
'
Public Enum ePTVertTextAlg
    ptTop
    ptMiddle
    ptBottom
End Enum
Private mePTVertTextAlg As ePTVertTextAlg
'
Public Enum ePTPicAlig
    ptPicLeftEdge
    ptPicRightEdge
    ptPicCenter
    ptPicLeftOfCapton
    ptPicRightOfCaption
End Enum
Private mePTPicAlig As ePTPicAlig
'
Public Enum ePTSlideButtStyle
    ptArrow
    ptTrinangle
    ptFilledArrow
End Enum
Private mePTSlideButtStyle As ePTSlideButtStyle
'
Public Enum ePTSlideAlign
    ptScrollLeft
    ptScrollRight
End Enum
Private mePTSlideAlign As ePTSlideAlign
'
'color schemes
Public Enum ePTColorScheme
    ptColorUser
    ptColorNoteOne
End Enum
Private mePTColorScheme As ePTColorScheme


Private meTTStyle As eTTStyle
'
Private mTT As New clsToolTip
'properties:
Private BC  As OLE_COLOR
Private BCActive As OLE_COLOR
Private BCHover As OLE_COLOR
Private BCDisabled As OLE_COLOR
''
Private shadC As OLE_COLOR
'
Private FC As OLE_COLOR
Private FCActive As OLE_COLOR
Private FCHover As OLE_COLOR
Private FCDisabled As OLE_COLOR
'
Private controlBorderC As OLE_COLOR '
Private controlBC As OLE_COLOR '
Private tabAreaBC As OLE_COLOR '
Private controlShadC As OLE_COLOR
Private controlHLightC As OLE_COLOR
'
Private slideBDRC As OLE_COLOR
Private slideBDRCDisabled As OLE_COLOR
Private slideHLC As OLE_COLOR
Private slideShadC As OLE_COLOR
'
Private slideBC As OLE_COLOR
Private slideBCHover As OLE_COLOR
Private slideBCDown As OLE_COLOR
Private slideBCDisabled As OLE_COLOR
'
Private slideFillC As OLE_COLOR
Private slideFillCHover As OLE_COLOR
Private slideFillCDown As OLE_COLOR
Private slideFillCDisabled As OLE_COLOR
'
Private lstCaptions As New Collection
Private lstDisabled As New Collection
Private lstPositions As New Collection
Private lstSizes As New Collection
Private lstLeftSpc As New Collection
Private lstRightSpc As New Collection
Private strToolTips() As String
'
Private mFont As New StdFont '
Private mFontHover As New StdFont '
Private mFontActive As New StdFont '
Private mFontDisabled As New StdFont '
'
Private icWid As Integer, icHeig As Integer '
Private tabHeig As Integer '
Private tabHeigActive As Integer
'
Private enbl As Boolean '
Private shSlideButtons As Boolean '
Private hSlideBack As Boolean
Private efectSel As Boolean '
Private drawClArea As Boolean
Private aSize As Boolean
Private scrHover As Boolean, scrFlat As Boolean
'
Private mSpacing As Integer '
Private buttSpacing As Integer '
Private startX As Integer '
Private selIndex As Integer, hoverIndex As Integer
'
Private tabIcons() As StdPicture
Private tabEnabl() As Boolean
Private tabTags() As Variant
Private tabVis() As Boolean
Private tabCnt As Integer
'
'
'
Private mX As Long
Private lstRghtSpc As Integer
Private hoverSlide As Integer, downSlide As Integer
Private scrlSide As Integer, firstItemL As Integer, firstItemR As Integer
Private scrLEnb As Boolean, scrREnb As Boolean
Private mStartX As Long
Private currColScheme As Byte
'redraw active tab
Private needRedraw As Boolean, isHover As Boolean
Private RDX As Integer

'for contained controls
'   store control name, index and tabIndex: ctl1(0)1 or clt1()1 if there is no index
Private ctlLst() As New Collection ', visibleCTLs As New Collection
'
'private varijables for drawing
Private ucSW As Integer, ucSH As Integer, bWid As Integer
Private mBDRC As OLE_COLOR, mBC As OLE_COLOR, mShadC As OLE_COLOR
'chache

'constants
Private Const cSize1 As Byte = 2
Private Const cSize2 As Byte = 4
Private Const cSizeXP As Byte = 3

'objects

'---------------------------------------------
'   e   v   e   n   t   s

Public Event Click()
Public Event DblClick()
Public Event TabChange(ByRef LastTab As Integer, ByRef NewTab As Integer)
Public Event BeforeTabChange(ByRef LastTab As Integer, ByRef NewTab As Integer)
Public Event TabClick(ByRef TabIndex As Integer)


Public Event AfterScroll()
'Public Event AfterScrollLeft()
'Public Event AfterScrollRight()
Public Event Scroll()

Public Event KeyDown(ByRef KeyCode As Integer, ByRef Shift As Integer)
Public Event KeyPress(ByRef KeyAscii As Integer)
Public Event KeyUp(ByRef KeyCode As Integer, ByRef Shift As Integer)

Public Event MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
Public Event MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
Public Event MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)

Public Event Resize()




'
'*************************************************************************************
'*************************************************************************************
'                           properties
'*************************************************************************************
'*************************************************************************************
'
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property
Public Property Get ToolTipStyle() As eTTStyle
    ToolTipStyle = meTTStyle
End Property
Public Property Let ToolTipStyle(ByVal nV As eTTStyle)
    meTTStyle = nV
    PropertyChanged "ToolTipStyle"
End Property

'Appearance
Public Property Get Appearance() As ePTAppearance
    Appearance = mePTAppearance
End Property
Public Property Let Appearance(ByVal nV As ePTAppearance)
    mePTAppearance = nV
    reDraw
    PropertyChanged "Appearance"
End Property
'
'
'colors
Public Property Get ColorScheme() As ePTColorScheme
    ColorScheme = mePTColorScheme
End Property
Public Property Let ColorScheme(ByVal nV As ePTColorScheme)
    mePTColorScheme = nV
    reDraw
    PropertyChanged "ColorScheme"
End Property
'   t   a   b   s
Public Property Get ActiveTab() As Integer
    ActiveTab = selIndex
End Property
Public Property Let ActiveTab(ByVal mTabIndex As Integer)
    If mTabIndex < 0 Or mTabIndex >= tabCnt Then
        Err.Raise 380 ' invalid property value
        Exit Property
    End If
    handleControls selIndex, mTabIndex
    selIndex = mTabIndex
    If selIndex >= 0 Then
        Dim mTabHeig1 As Integer
        If tabHeig > tabHeigActive Then mTabHeig1 = tabHeig Else mTabHeig1 = tabHeigActive
        'set tab position to see it whole
        If firstItemL > mTabIndex Then
            firstItemL = mTabIndex + 1
            If firstItemL < 1 Then firstItemL = 1
        ElseIf mTabIndex >= firstItemL + lstSizes.Count - 2 Or (getStyleByIndex(mTabIndex) = ptCoolRight And mTabIndex >= firstItemL + lstSizes.Count - 3) Then  'Or mX > UserControl.ScaleWidth Then
            Dim isMov As Boolean
            isMov = False
            If mTabIndex <> tabCnt - 1 Then
                Do While mTabIndex >= firstItemL + lstSizes.Count - 2
                    firstItemL = firstItemL + 1
                    isMov = True
                Loop
            End If
            If mX > UserControl.ScaleWidth And isMov = False Then
                firstItemL = firstItemL + 1
            End If
        End If
        If mStartX > 0 Then mStartX = 0
    End If
    'reDraw
    reDraw
    PropertyChanged "ActiveTab"
End Property

'caption
Public Property Get TabCaption(ByVal mTabIndex As Integer) As String
    If mTabIndex > -1 And mTabIndex < lstCaptions.Count Then
        TabCaption = lstCaptions.Item(mTabIndex + 1)
    Else
        Err.Raise 9 ' Subscript Out of Range
        Exit Property
    End If
End Property
Public Property Let TabCaption(ByVal mTabIndex As Integer, ByVal mTabCaption As String)
    If mTabIndex > -1 And mTabIndex < lstCaptions.Count Then
        replaceDataInCollection lstCaptions, mTabIndex + 1, mTabCaption
        reDraw
        PropertyChanged "TabCaption"
    Else
        Err.Raise 9 ' Subscript Out of Range
        Exit Property
    End If
End Property
'tag
Public Property Get TabTag(ByVal mTabIndex As Integer) As Variant
    If mTabIndex > -1 And mTabIndex < lstCaptions.Count Then
        TabTag = tabTags(mTabIndex)
    Else
        Err.Raise 9 ' Subscript Out of Range
        Exit Property
    End If
End Property
Public Property Let TabTag(ByVal mTabIndex As Integer, ByVal mTabTag As Variant)
    If mTabIndex > -1 And mTabIndex < lstCaptions.Count Then
        tabTags(mTabIndex) = mTabTag
        PropertyChanged "TabTag"
    Else
        Err.Raise 9 ' Subscript Out of Range
        Exit Property
    End If
End Property
'tooltip
Public Property Get TabToolTip(ByVal mTabIndex As Integer) As String
    If mTabIndex > -1 And mTabIndex < lstCaptions.Count Then
        TabToolTip = strToolTips(mTabIndex)
    Else
        Err.Raise 9 ' Subscript Out of Range
        Exit Property
    End If
End Property
Public Property Let TabToolTip(ByVal mTabIndex As Integer, ByVal mTabToolTip As String)
    If mTabIndex > -1 And mTabIndex < lstCaptions.Count Then
        strToolTips(mTabIndex) = mTabToolTip
        'reDraw
        PropertyChanged "TabToolTip"
    Else
        Err.Raise 9 ' Subscript Out of Range
        Exit Property
    End If
End Property
'icon
Public Property Get TabIcon(ByVal mTabIndex As Integer) As StdPicture
    If mTabIndex > -1 And mTabIndex < lstCaptions.Count Then
        Set TabIcon = tabIcons(mTabIndex)
    Else
        Err.Raise 9 ' Subscript Out of Range
        Exit Property
    End If
End Property
Public Property Set TabIcon(ByVal mTabIndex As Integer, ByVal mTabIcon As StdPicture)
    If mTabIndex > -1 And mTabIndex < lstCaptions.Count Then
        Set tabIcons(mTabIndex) = mTabIcon
        reDraw
        PropertyChanged "TabIcon"
    Else
        Err.Raise 9 ' Subscript Out of Range
        Exit Property
    End If
End Property

'tab enabled
Public Property Get TabEnabled(ByVal mTabIndex As Integer) As Boolean
    If mTabIndex > -1 And mTabIndex < lstCaptions.Count Then
        TabEnabled = tabEnabl(mTabIndex)
        'reDraw
    Else
        Err.Raise 9 ' Subscript Out of Range
        Exit Property
    End If
End Property
Public Property Let TabEnabled(ByVal mTabIndex As Integer, ByVal mNewVal As Boolean)
  If mTabIndex > -1 And mTabIndex < tabCnt Then
    tabEnabl(mTabIndex) = mNewVal
    reDraw
    PropertyChanged "TabEnabled"
  Else
    Err.Raise 9 ' Subscript Out of Range
    Exit Property    'if out of range
  End If
End Property
'tab visible
Public Property Get TabVisible(ByVal mTabIndex As Integer) As Boolean
    If mTabIndex > -1 And mTabIndex < lstCaptions.Count Then
        TabVisible = tabVis(mTabIndex)
       'reDraw
    Else
        Err.Raise 9 ' Subscript Out of Range
        Exit Property
    End If
End Property
Public Property Let TabVisible(ByVal mTabIndex As Integer, ByVal mNewVal As Boolean)
  If mTabIndex > -1 And mTabIndex < tabCnt Then
    tabVis(mTabIndex) = mNewVal
    reDraw
    PropertyChanged "TabVisible"
  Else
    Err.Raise 9 ' Subscript Out of Range
    Exit Property    'if out of range
  End If
End Property
'-----------------------------------------------------------
' --------------------- s c r o l l s ----------------------
'back
Public Property Get ScrollAreaTransparent() As Boolean
    ScrollAreaTransparent = hSlideBack
End Property
Public Property Let ScrollAreaTransparent(ByVal nV As Boolean)
    hSlideBack = nV
    reDraw
    PropertyChanged "ScrollAreaTransparent"
End Property
'ScrollBorderColor
Public Property Get ScrollBorderColor() As OLE_COLOR
    ScrollBorderColor = slideBDRC
End Property
Public Property Let ScrollBorderColor(ByVal nV As OLE_COLOR)
    slideBDRC = nV
    reDraw
    PropertyChanged "ScrollBorderColor"
End Property
'ScrollBorderColor
Public Property Get ScrollBorderColorDisabled() As OLE_COLOR
    ScrollBorderColorDisabled = slideBDRCDisabled
End Property
Public Property Let ScrollBorderColorDisabled(ByVal nV As OLE_COLOR)
    slideBDRCDisabled = nV
    reDraw
    PropertyChanged "ScrollBorderColorDisabled"
End Property
'ScrollHighLightColor
Public Property Get ScrollHighlightColor() As OLE_COLOR
    ScrollHighlightColor = slideHLC
End Property
Public Property Let ScrollHighlightColor(ByVal nV As OLE_COLOR)
    slideHLC = nV
    reDraw
    PropertyChanged "ScrollHighlightColor"
End Property
'ScrollShadowColor
Public Property Get ScrollShadowColor() As OLE_COLOR
    ScrollShadowColor = slideShadC
End Property
Public Property Let ScrollShadowColor(ByVal nV As OLE_COLOR)
    slideShadC = nV
    reDraw
    PropertyChanged "ScrollShadowColor"
End Property
'---'
'ScrollBackColor
Public Property Get ScrollBackColor() As OLE_COLOR
    ScrollBackColor = slideBC
End Property
Public Property Let ScrollBackColor(ByVal nV As OLE_COLOR)
    slideBC = nV
    reDraw
    PropertyChanged "ScrollBackColor"
End Property
'ScrollBackColorHover
Public Property Get ScrollBackColorHover() As OLE_COLOR
    ScrollBackColorHover = slideBCHover
End Property
Public Property Let ScrollBackColorHover(ByVal nV As OLE_COLOR)
    slideBCHover = nV
    'reDraw
    PropertyChanged "ScrollBackColorHover"
End Property
'ScrollBackColorDown
Public Property Get ScrollBackColorDown() As OLE_COLOR
    ScrollBackColorDown = slideBCDown
End Property
Public Property Let ScrollBackColorDown(ByVal nV As OLE_COLOR)
    slideBCDown = nV
    'reDraw
    PropertyChanged "ScrollBackColorDown"
End Property
'ScrollBackColorDisabled
Public Property Get ScrollBackColorDisabled() As OLE_COLOR
    ScrollBackColorDisabled = slideBCDisabled
End Property
Public Property Let ScrollBackColorDisabled(ByVal nV As OLE_COLOR)
    slideBCDisabled = nV
    reDraw
    PropertyChanged "ScrollBackColorDisabled"
End Property
'-- fill color --
'ScrollFillColor
Public Property Get ScrollFillColor() As OLE_COLOR
    ScrollFillColor = slideFillC
End Property
Public Property Let ScrollFillColor(ByVal nV As OLE_COLOR)
    slideFillC = nV
    reDraw
    PropertyChanged "ScrollFillColor"
End Property
'ScrollFillColorHover
Public Property Get ScrollFillColorHover() As OLE_COLOR
    ScrollFillColorHover = slideFillCHover
End Property
Public Property Let ScrollFillColorHover(ByVal nV As OLE_COLOR)
    slideFillCHover = nV
    'reDraw
    PropertyChanged "ScrollFillColorHover"
End Property
'ScrollFillColorDown
Public Property Get ScrollFillColorDown() As OLE_COLOR
    ScrollFillColorDown = slideFillCDown
End Property
Public Property Let ScrollFillColorDown(ByVal nV As OLE_COLOR)
    slideFillCDown = nV
    'reDraw
    PropertyChanged "ScrollFillColorDown"
End Property
'ScrollFillColorDisabled
Public Property Get ScrollFillColorDisabled() As OLE_COLOR
    ScrollFillColorDisabled = slideFillCDisabled
End Property
Public Property Let ScrollFillColorDisabled(ByVal nV As OLE_COLOR)
    slideFillCDisabled = nV
    reDraw
    PropertyChanged "ScrollFillColorDisabled"
End Property
'---------------- s t y l e s ------------------------------
'STYLE > normal
Public Property Get IconAlign() As ePTPicAlig
    IconAlign = mePTPicAlig
End Property
Public Property Let IconAlign(ByVal nV As ePTPicAlig)
    mePTPicAlig = nV
    reDraw
    PropertyChanged "IconAlign"
End Property
'STYLE > normal
Public Property Get StyleNormal() As ePTStyles
    StyleNormal = mePTStylesNormal
End Property
Public Property Let StyleNormal(ByVal nV As ePTStyles)
    mePTStylesNormal = nV
    If (mePTStylesActive = ptCoolRight Or mePTStylesActive = ptCoolLeft Or mePTStylesActive = ptDistorted Or mePTStylesActive = ptRoundMenu Or mePTStylesActive = ptDistortedMenu) And nV <> ptVerticalLine Then
        mePTStylesActive = nV
    End If
    reDraw
    PropertyChanged "StyleNormal"
End Property
'STYLE > Active
Public Property Get StyleActive() As ePTStyles
    StyleActive = mePTStylesActive
End Property
Public Property Let StyleActive(ByVal nV As ePTStyles)
    mePTStylesActive = nV
    If (mePTStylesNormal = ptCoolRight Or mePTStylesNormal = ptCoolLeft Or mePTStylesNormal = ptDistorted Or mePTStylesNormal = ptRoundMenu Or mePTStylesNormal = ptDistortedMenu) And (nV <> mePTStylesNormal) Then
        mePTStylesNormal = nV
    End If
    
    reDraw
    PropertyChanged "StyleActive"
End Property
'ORIENTATION
Public Property Get Orientation() As ePTOrientation
    Orientation = mePTOrientation
End Property
Public Property Let Orientation(ByVal nV As ePTOrientation)
    mePTOrientation = nV
    reDraw
    PropertyChanged "Orientation"
End Property
'TEXTALIGN>Horizontal
Public Property Get TextAlignHorizontal() As ePTHorTextAlg
    TextAlignHorizontal = mePTHorTextAlg
End Property
Public Property Let TextAlignHorizontal(ByVal nV As ePTHorTextAlg)
    mePTHorTextAlg = nV
    reDraw
    PropertyChanged "TextAlignHorizontal"
End Property
'TEXTALIGN>Vertical
Public Property Get TextAlignVertical() As ePTVertTextAlg
    TextAlignVertical = mePTVertTextAlg
End Property
Public Property Let TextAlignVertical(ByVal nV As ePTVertTextAlg)
    mePTVertTextAlg = nV
    reDraw
    PropertyChanged "TextAlignVertical"
End Property
'ScrollStyle
Public Property Get ScrollStyle() As ePTSlideButtStyle
    ScrollStyle = mePTSlideButtStyle
End Property
Public Property Let ScrollStyle(ByVal nV As ePTSlideButtStyle)
    mePTSlideButtStyle = nV
    reDraw
    PropertyChanged "ScrollStyle"
End Property
'scroll align
Public Property Get ScrollAlign() As ePTSlideAlign
    ScrollAlign = mePTSlideAlign
End Property
Public Property Let ScrollAlign(ByVal nV As ePTSlideAlign)
    mePTSlideAlign = nV
    reDraw
    PropertyChanged "ScrollAlign"
End Property
'+++++++++++++++++++++++++++++++++++++++++++
'   c  o  l  o  r  s
'+++++++++++++++++++++++++++++++++++++++++++
' control
'controlBorderColor
Public Property Get BorderColor() As OLE_COLOR
    BorderColor = controlBorderC
End Property
Public Property Let BorderColor(ByVal nV As OLE_COLOR)
    controlBorderC = nV
    reDraw
    PropertyChanged "BorderColor"
End Property
'controlBackColor
Public Property Get BackColor() As OLE_COLOR
    If mePTColorScheme = ptColorUser Then
        BackColor = controlBC
    ElseIf mePTColorScheme = ptColorNoteOne Then
        If selIndex <= 7 Then
            BackColor = mConColor(selIndex)
        Else
            Dim tmpInd As Integer
            'tmpInd = Int(mInd / 7)
            tmpInd = selIndex
            Do While tmpInd > 7
                tmpInd = tmpInd - 7 - 1
            Loop
            BackColor = mConColor(tmpInd)
        End If
    End If
End Property
Public Property Let BackColor(ByVal nV As OLE_COLOR)
    controlBC = nV
    reDraw
    PropertyChanged "BackColor"
End Property
'ShadowColor
Public Property Get ShadowColor() As OLE_COLOR
    ShadowColor = controlShadC
End Property
Public Property Let ShadowColor(ByVal nV As OLE_COLOR)
    controlShadC = nV
    reDraw
    PropertyChanged "ShadowColor"
End Property
'HighlightColor
Public Property Get HighlightColor() As OLE_COLOR
    HighlightColor = controlHLightC
End Property
Public Property Let HighlightColor(ByVal nV As OLE_COLOR)
    controlHLightC = nV
    reDraw
    PropertyChanged "HighlightColor"
End Property
'controlBackColor
Public Property Get TabAreaBackColor() As OLE_COLOR
    TabAreaBackColor = tabAreaBC
End Property
Public Property Let TabAreaBackColor(ByVal nV As OLE_COLOR)
    tabAreaBC = nV
    reDraw
    PropertyChanged "TabAreaBackColor"
End Property
'TABCOLOR>normal
Public Property Get TabColor() As OLE_COLOR
    TabColor = BC
End Property
Public Property Let TabColor(ByVal nV As OLE_COLOR)
    BC = nV
    reDraw
    PropertyChanged "TabColor"
End Property
'TABCOLOR>Active
Public Property Get TabColorActive() As OLE_COLOR
    TabColorActive = BCActive
End Property
Public Property Let TabColorActive(ByVal nV As OLE_COLOR)
    BCActive = nV
    reDraw
    PropertyChanged "TabColorActive"
End Property
'TABCOLOR>hover
Public Property Get TabColorHover() As OLE_COLOR
    TabColorHover = BCHover
End Property
Public Property Let TabColorHover(ByVal nV As OLE_COLOR)
    BCHover = nV
    'reDraw
    PropertyChanged "TabColorHover"
End Property
'TABCOLOR>disabled
Public Property Get TabColorDisabled() As OLE_COLOR
    TabColorDisabled = BCDisabled
End Property
Public Property Let TabColorDisabled(ByVal nV As OLE_COLOR)
    BCDisabled = nV
    reDraw
    PropertyChanged "TabColorDisabled"
End Property
''+++++++++++++++++++++++
'ShadowColor>normal
Public Property Get TabShadowColor() As OLE_COLOR
    TabShadowColor = shadC
End Property
Public Property Let TabShadowColor(ByVal nV As OLE_COLOR)
    shadC = nV
    reDraw
    PropertyChanged "TabShadowColor"
End Property

'FORECOLOR>normal
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = FC
End Property
Public Property Let ForeColor(ByVal nV As OLE_COLOR)
    FC = nV
    reDraw
    PropertyChanged "ForeColor"
End Property
'FORECOLOR>Active
Public Property Get ForeColorActive() As OLE_COLOR
    ForeColorActive = FCActive
End Property
Public Property Let ForeColorActive(ByVal nV As OLE_COLOR)
    FCActive = nV
    reDraw
    PropertyChanged "ForeColorActive"
End Property
'FORECOLOR>hover
Public Property Get ForeColorHover() As OLE_COLOR
    ForeColorHover = FCHover
End Property
Public Property Let ForeColorHover(ByVal nV As OLE_COLOR)
    FCHover = nV
    'reDraw
    PropertyChanged "ForeColorHover"
End Property
'FORECOLOR>disabled
Public Property Get ForeColorDisabled() As OLE_COLOR
    ForeColorDisabled = FCDisabled
End Property
Public Property Let ForeColorDisabled(ByVal nV As OLE_COLOR)
    FCDisabled = nV
    reDraw
    PropertyChanged "ForeColorDisabled"
End Property
'+++++++++++++++++++++++++++++++++++++++++++
'   f  o  n  t
'+++++++++++++++++++++++++++++++++++++++++++
'font>normal
Public Property Get Font() As StdFont
    Set Font = mFont
End Property
Public Property Set Font(ByVal nV As StdFont)
    Set mFont = nV
    Set UserControl.Font = nV
    reDraw
    PropertyChanged "Font"
End Property
'font>Active
Public Property Get FontActive() As StdFont
    Set FontActive = mFontActive
End Property
Public Property Set FontActive(ByVal nV As StdFont)
    Set mFontActive = nV
    reDraw
    PropertyChanged "FontActive"
End Property
'font>hover
Public Property Get FontHover() As StdFont
    Set FontHover = mFontHover
End Property
Public Property Set FontHover(ByVal nV As StdFont)
    Set mFontHover = nV
    'reDraw
    PropertyChanged "FontHover"
End Property
'font>disabled
Public Property Get FontDisabled() As StdFont
    Set FontDisabled = mFontDisabled
End Property
Public Property Set FontDisabled(ByVal nV As StdFont)
    Set mFontDisabled = nV
    reDraw
    PropertyChanged "FontDisabled"
End Property
'+++++++++++++++++++++++++++++++++++++++++++
'   i  n  t  e  g  e  r  s
'+++++++++++++++++++++++++++++++++++++++++++
'TabCount
Public Property Get TabCount() As Integer
    TabCount = tabCnt
End Property
Public Property Let TabCount(ByVal nV As Integer)
    If nV < 1 Or nV > 500 Then
        Err.Raise 380 ' invalid property value
        Exit Property
    End If
    tabCnt = nV
    setTabCount
    If selIndex > tabCnt - 1 Then selIndex = tabCnt - 1
    mStartX = 0
    reDraw
    ActiveTab = selIndex
    PropertyChanged "TabCount"
    'UserControl_ReadProperties
End Property
'iconWidth
Public Property Get IconWidth() As Integer
    IconWidth = icWid
End Property
Public Property Let IconWidth(ByVal nV As Integer)
    icWid = nV
    reDraw
    PropertyChanged "IconWidth"
End Property
'iconHeight
Public Property Get IconHeight() As Integer
    IconHeight = icHeig
End Property
Public Property Let IconHeight(ByVal nV As Integer)
    icHeig = nV
    reDraw
    PropertyChanged "IconHeight"
End Property
'TabHeight
Public Property Get TabHeight() As Integer
    TabHeight = tabHeig
End Property
Public Property Let TabHeight(ByVal nV As Integer)
    tabHeig = nV
    If mePTStylesNormal = ptDistorted Or mePTStylesActive = ptDistorted Then
        tabHeigActive = nV
    End If
    If tabHeig > tabHeigActive Then tabHeigActive = tabHeig
    reDraw
    PropertyChanged "TabHeight"
End Property
'TabHeight
Public Property Get TabHeightActive() As Integer
    TabHeightActive = tabHeigActive
End Property
Public Property Let TabHeightActive(ByVal nV As Integer)
    tabHeigActive = nV
    If mePTStylesNormal = ptDistorted Or mePTStylesActive = ptDistorted Then
        tabHeigActive = nV
    End If
    If tabHeigActive < tabHeig Then tabHeig = tabHeigActive
    reDraw
    PropertyChanged "TabHeightActive"
End Property
'LeftSpacing
Public Property Get LeftSpacing() As Integer
    LeftSpacing = startX
End Property
Public Property Let LeftSpacing(ByVal nV As Integer)
    startX = nV
    reDraw
    PropertyChanged "LeftSpacing"
End Property
'TabSpacing
Public Property Get TabSpacing() As Integer
    TabSpacing = buttSpacing
End Property
Public Property Let TabSpacing(ByVal nV As Integer)
    If nV < 0 Then
        Err.Raise 380 ' invalid property value
        Exit Property
    End If
    buttSpacing = nV
    reDraw
    PropertyChanged "TabSpacing"
End Property
'FontSpacing
Public Property Get FontSpacing() As Integer
    FontSpacing = mSpacing
End Property
Public Property Let FontSpacing(ByVal nV As Integer)
    mSpacing = nV
    reDraw
    PropertyChanged "FontSpacing"
End Property
'+++++++++++++++++++++++++++++++++++++++++++
'   b  o  o  l  s
'+++++++++++++++++++++++++++++++++++++++++++
'DrawClientArea
Public Property Get DrawClientArea() As Boolean
    DrawClientArea = drawClArea
End Property
Public Property Let DrawClientArea(ByVal nV As Boolean)
    drawClArea = nV
    reDraw
    PropertyChanged "DrawClientArea"
End Property
'Enabled
Public Property Get Enabled() As Boolean
    Enabled = enbl
End Property
Public Property Let Enabled(ByVal nV As Boolean)
    enbl = nV
    reDraw
    PropertyChanged "Enabled"
End Property
'ScrollHoverButton
Public Property Get ScrollHoverButton() As Boolean
    ScrollHoverButton = scrHover
End Property
Public Property Let ScrollHoverButton(ByVal nV As Boolean)
    scrHover = nV
    reDraw
    PropertyChanged "ScrollHoverButton"
End Property

'AutoSize
Public Property Get AutoSize() As Boolean
    AutoSize = aSize
End Property
Public Property Let AutoSize(ByVal nV As Boolean)
    aSize = nV
    reDraw
    PropertyChanged "AutoSize"
End Property
'EnableFontMoving
Public Property Get EnableFontMoving() As Boolean
    EnableFontMoving = efectSel
End Property
Public Property Let EnableFontMoving(ByVal nV As Boolean)
    efectSel = nV
    reDraw
    PropertyChanged "EnableFontMoving"
End Property
'EnableFontMoving
Public Property Get ShowScroll() As Boolean
    ShowScroll = shSlideButtons
End Property
Public Property Let ShowScroll(ByVal nV As Boolean)
    shSlideButtons = nV
    reDraw
    PropertyChanged "ShowScroll"
End Property
'*************************************************************************************
'*************************************************************************************
'*************************************************************************************
'*************************************************************************************
'*************************************************************************************

Private Sub tmrSlide_Timer()
    If scrlSide = -1 Then
        firstItemL = firstItemL + 1
    Else
        If firstItemL >= 1 Then firstItemL = firstItemL - 1
        If firstItemL < 1 Then firstItemL = 1
    End If
    RaiseEvent Scroll
    reDraw
    If tmrSlide.Interval > 10 Then tmrSlide.Interval = tmrSlide.Interval - 2
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_Initialize()

    currColScheme = 0
    firstItemL = 1
    UserControl_InitProperties
    'MsgBox "INTALIZE"
    'label1.Tag
End Sub

Private Sub UserControl_InitProperties()
    meTTStyle = TTStyle_Balloon
    
    mePTStylesNormal = ptVerticalLine
    mePTStylesActive = ptCoolLeft
    '
    mePTOrientation = ptOrientationTop
    
    mePTTabStyle = ptStandard
    mePTHorTextAlg = ptCenter
    mePTVertTextAlg = ptMiddle
    mePTPicAlig = ptPicLeftEdge
    '
    mePTColorScheme = ptColorNoteOne
    '
    'mePTButtAlign = 1
    scrHover = True
    mePTAppearance = ptAppearance3D
    '
    BC = vbButtonFace
    BCActive = vbButtonFace
    BCHover = vbButtonFace
    BCDisabled = vbButtonFace
    controlHLightC = vbWhite
    controlShadC = &H80808
    '
'    borderC = &H808080
'    borderCActive = &H808080
'    borderCHover = &H808080
'    borderCDisabled = &H808080
    '
    shadC = vbWhite
'    shadCActive = vbWhite
'    shadCHover = vbWhite
'    shadCDisabled = vbWhite
    '
    FC = vbBlack
    FCActive = &H808080
    FCHover = &H404040
    FCDisabled = vbBlack
    '
    hSlideBack = True
    '
    slideBDRC = &H404040
    slideBDRCDisabled = &H404040
    slideHLC = vbWhite
    slideShadC = &H404040
    '
    slideBC = vbHighlight
    slideBCHover = vbHighlight
    slideBCDown = vbHighlight
    slideBCDisabled = vbHighlight
    '
    slideFillC = vbBlack
    slideFillCHover = vbBlack
    slideFillCDisabled = &H808080
    slideFillCDown = vbBlack
    '
    mePTSlideAlign = ptScrollRight
    mePTSlideButtStyle = ptArrow

'
    controlBorderC = &H808080
    controlBC = vbButtonFace
    tabAreaBC = vbHighlight
    '
    icWid = 16
    icHeig = 16
    tabHeig = 18
    tabHeigActive = 20
    '
    enbl = True

    shSlideButtons = False
    efectSel = True
    aSize = False
    '
    drawClArea = True
    '
    mSpacing = 0
    selIndex = 0
    startX = 10
    tabCnt = 4
    hoverIndex = -1
    '
    setTabCount
    '
    'hoverSlide = -1
    'downSlide = -1
    reDraw
End Sub








Private Sub reDraw() '(Optional onlyTabs As Boolean = False)
    If aSize = True Then mStartX = 0

    Dim i As Integer, mTabHeig As Integer, mTmpInd As Integer
    Dim mWINC As Integer, tmpMX As Integer
    
    checkProperties
    mX = startX
    
    'mObjInfo.hWnd = UserControl.hWnd
    
    'UserControl.AutoRedraw = True
    'If onlyTabs <> True Then
        UserControl.Cls
        If UserControl.BackColor <> controlBC And mePTColorScheme <> ptColorNoteOne Then
            UserControl.BackColor = controlBC
            setContainerBackColor controlBC
        End If
    'End If
    Set lstSizes = Nothing
    Set lstPositions = Nothing
    'color scheme
    If mePTColorScheme = ptColorNoteOne Then
        'load color scheme
        If currColScheme <> 1 Then loadColors 1
        Dim mColorInd As Byte
        If selIndex <= 7 Then
            mColorInd = selIndex
        Else
            Dim tmpInd As Integer
            'tmpInd = Int(mInd / 7)
            tmpInd = selIndex
            Do While tmpInd > 7
                tmpInd = tmpInd - 7 - 1
            Loop
            mColorInd = tmpInd
        End If
        'fill control with selected tab color
        If mePTStylesActive <> ptFlatButton And mePTStylesNormal <> ptFlatButton And mePTStylesActive <> ptRoundMenu And mePTStylesNormal <> ptRoundMenu Then
            If UserControl.BackColor <> mConColor(mColorInd) Then
                UserControl.BackColor = mConColor(mColorInd)
                setContainerBackColor mConColor(mColorInd)
            End If
        End If
    End If
    'fix tab heig for some styles
    If mePTStylesNormal = ptDistortedMenu Or mePTStylesActive = ptDistortedMenu Then tabHeig = tabHeigActive
    'find max tab heig
    If tabHeigActive > tabHeig Then
        mTabHeig = tabHeigActive
    Else
        mTabHeig = tabHeig
    End If
    '
    'draw tab area
    UserControl.FillStyle = 1
    If mePTOrientation = ptOrientationTop Then
        UserControl.Line (0, 0)-(UserControl.ScaleWidth, mTabHeig), tabAreaBC, BF
        'draw tab area border
        If mePTStylesActive <> ptFlatButton And mePTStylesNormal <> ptFlatButton _
                        And mePTStylesActive <> ptRoundMenu And mePTStylesNormal <> ptRoundMenu _
                        And mePTStylesActive <> ptDistortedMenu And mePTStylesNormal <> ptDistortedMenu Then UserControl.Line (0, mTabHeig)-(UserControl.ScaleWidth, mTabHeig), controlBorderC
        '
        
        If drawClArea = True And mePTStylesActive <> ptFlatButton And mePTStylesNormal <> ptFlatButton Then
            UserControl.Line (0, mTabHeig)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), controlBorderC, B
            'draw shadow
            If mePTAppearance <> ptAppearanceFlat Then
                If mePTStylesNormal <> ptXP And mePTStylesActive <> ptXP Then
                    UserControl.Line (1, mTabHeig + 1)-(UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2), controlHLightC, B
                    If mePTStylesNormal = ptSSTab And mePTStylesActive = ptSSTab Then UserControl.Line (2, mTabHeig + 2)-(UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3), controlHLightC, B

                    UserControl.Line (UserControl.ScaleWidth - 2, mTabHeig + 1)-(UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2), controlShadC, B
                    If mePTStylesNormal = ptSSTab And mePTStylesActive = ptSSTab Then UserControl.Line (UserControl.ScaleWidth - 3, mTabHeig + 2)-(UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3), controlShadC, B

                    UserControl.Line (1, UserControl.ScaleHeight - 2)-(UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 2), controlShadC, B
                    If mePTStylesNormal = ptSSTab And mePTStylesActive = ptSSTab Then UserControl.Line (2, UserControl.ScaleHeight - 3)-(UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 3), controlShadC, B
                End If
            End If
        Else
            UserControl.Height = (Abs(tabHeigActive - tabHeig) + tabHeig + 1) * Screen.TwipsPerPixelY
        End If
    Else
        UserControl.Line (0, UserControl.ScaleHeight)-(UserControl.ScaleWidth, UserControl.ScaleHeight - mTabHeig - 1), tabAreaBC, BF
        'draw tab area border
        If mePTStylesActive <> ptFlatButton And mePTStylesNormal <> ptFlatButton _
                        And mePTStylesActive <> ptRoundMenu And mePTStylesNormal <> ptRoundMenu _
                        And mePTStylesActive <> ptDistortedMenu And mePTStylesNormal <> ptDistortedMenu Then UserControl.Line (0, UserControl.ScaleHeight - mTabHeig - 1)-(UserControl.ScaleWidth, UserControl.ScaleHeight - mTabHeig - 1), controlBorderC
        '
        If drawClArea = True And mePTStylesActive <> ptFlatButton And mePTStylesNormal <> ptFlatButton Then
            UserControl.Line (0, UserControl.ScaleHeight - mTabHeig - 1)-(UserControl.ScaleWidth - 1, 0), controlBorderC, B
            'draw shadow
            If mePTAppearance <> ptAppearanceFlat Then
                If mePTStylesNormal <> ptXP And mePTStylesActive <> ptXP Then
                     UserControl.Line (1, UserControl.ScaleHeight - 1 - mTabHeig - 1)-(UserControl.ScaleWidth - 2, 1), controlHLightC, B
                     If mePTStylesNormal = ptSSTab And mePTStylesActive = ptSSTab Then UserControl.Line (2, UserControl.ScaleHeight - 1 - mTabHeig - 2)-(UserControl.ScaleWidth - 3, 2), controlHLightC, B
                     
                     UserControl.Line (UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 1 - mTabHeig - 1)-(UserControl.ScaleWidth - 2, 1), controlShadC, B
                     If mePTStylesNormal = ptSSTab And mePTStylesActive = ptSSTab Then UserControl.Line (UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 1 - mTabHeig - 2)-(UserControl.ScaleWidth - 3, 2), controlShadC, B
                    
                     UserControl.Line (1, UserControl.ScaleHeight - 1 - mTabHeig - 1)-(UserControl.ScaleWidth - 2, UserControl.ScaleHeight - 1 - mTabHeig - 1), controlShadC
                     If mePTStylesNormal = ptSSTab And mePTStylesActive = ptSSTab Then UserControl.Line (2, UserControl.ScaleHeight - 1 - mTabHeig - 2)-(UserControl.ScaleWidth - 3, UserControl.ScaleHeight - 1 - mTabHeig - 2), controlShadC, B
                End If
            End If
        Else
            UserControl.Height = (Abs(tabHeigActive - tabHeig) + tabHeig + 1) * Screen.TwipsPerPixelY
        End If
        
    End If
    '
    needRedraw = False
    UserControl.FillStyle = 0
    'for some styles make some changes
    If mePTStylesActive = ptCoolLeft Or mePTStylesNormal = ptCoolLeft Then
        buttSpacing = 0
    ElseIf mePTStylesActive = ptXP Or mePTStylesNormal = ptXP And shSlideButtons = True And mePTSlideAlign = ptScrollLeft Then
        mX = mX + 3 + buttSpacing
    End If
    'if scroll is on left side then increase start x
    If mePTSlideAlign = ptScrollLeft And shSlideButtons = True Then
        mX = mX + Int(mTabHeig * 2 / 1.5) + 1
    End If
    'increase startx position for some styles
    If mePTStylesNormal = ptRoundedRectangle2 Or mePTStylesActive = ptRoundedRectangle2 Then
        mX = mX + 3
    ElseIf mePTStylesNormal = ptCoolLeft Or mePTStylesActive = ptCoolLeft Then
        mX = mX + (tabHeigActive) * Cos(60 * 3.14 / 180) * 2
    End If
    'if firtsItemL is hidden then
    If tabVis(firstItemL - 1) = False Then
        For i = firstItemL To lstCaptions.Count
            If tabVis(i - 1) = True Then Exit For
        Next i
        
        If i <= lstCaptions.Count Then
            firstItemL = i '- 1
            'MsgBox i - 1
        Else
            For i = firstItemL To 1 Step -1
                If tabVis(i - 1) = True Then Exit For
            Next i
            firstItemL = i '- 1
        End If
        'MsgBox firstItemL
    End If
    '
    If firstItemL = 0 Then Exit Sub
    For i = firstItemL To lstCaptions.Count
        If tabVis(i - 1) <> False Then
            'set color values for current tab
            'If mePTColorScheme = ptColorNoteOne Or i - 2 = selIndex Or i - 2 = hoverIndex Or i - 1 = hoverIndex Or i - 1 = selIndex Then
                setValues i - 1, mBDRC, mBC, mShadC
           ' End If
            'make some position fix if all tab isn't at same style
            mWINC = 0
            If i - 1 = selIndex And i - 1 <> 0 Then
                mWINC = Abs(getWidInc(0) - getWidInc(i - 1))
            ElseIf i - 1 = 0 And selIndex = 0 Then
                If tabCnt > 1 Then mWINC = Abs(getWidInc(0) - getWidInc(1))
            End If
            If (i - 1 = selIndex And mePTStylesNormal <> ptCoolLeft And mePTStylesActive = ptCoolLeft) Or (i - 1 <> selIndex And mePTStylesActive <> ptCoolLeft And mePTStylesNormal = ptCoolLeft) Then
                If mePTStylesActive <> mePTStylesNormal Then mX = mX - getWidInc(i - 1)  ' + mWINC - 2 '+ 10
            End If
            '
            If i > 1 And (mePTStylesNormal = ptVerticalLine Or mePTStylesActive = ptVerticalLine) And getStyleByIndex(i - 1) = ptDistorted Then
                mX = mX - getWidInc(i - 1) / 2
            ElseIf i > 1 And (mePTStylesNormal = ptDistorted Or mePTStylesActive = ptDistorted) And getStyleByIndex(i - 1) = ptVerticalLine Then
                mX = mX - getWidInc(i - 1) / 2
            End If
            lstPositions.Add mX
            If i - 1 = selIndex Then
                If mePTStylesActive = ptRectangle Then
                    drawRectangleTab i - 1
                ElseIf mePTStylesActive = ptRoundedRectangle Then
                    drawRoundedTab i - 1
                ElseIf mePTStylesActive = ptCornerCutLeft Then
                    drawCornerCutLeft i - 1
                ElseIf mePTStylesActive = ptCornerCutRight Then
                    drawCornerCutRight i - 1
                ElseIf mePTStylesActive = ptCoolLeft Then
                    drawCoolLeftTab i - 1
                ElseIf mePTStylesActive = ptCoolRight Then
                    drawCoolRightTab i - 1
                    needRedraw = True
                    mTmpInd = i
                ElseIf mePTStylesActive = ptDistorted Then
                    drawDistortedTab i - 1
                ElseIf mePTStylesActive = ptVerticalLine Then
                    drawVerLineTab i - 1
                ElseIf mePTStylesActive = ptRoundMenu Then
                    drawRoundMenuTab i - 1
                ElseIf mePTStylesActive = ptXP Then
                    drawRoundedXPTab i - 1
                ElseIf mePTStylesActive = ptSSTab Then
                    drawSSTab i - 1
                ElseIf mePTStylesActive = ptFlatButton Then
                    drawFlatButtonTab i - 1
                ElseIf mePTStylesActive = ptProTab Then
                    drawProTabTab i - 1
                ElseIf mePTStylesActive = ptRoundedRectangle2 Then
                    drawRoundedRectangle2Tab i - 1
                    needRedraw = True
                    mTmpInd = i
                ElseIf mePTStylesActive = ptDistortedMenu Then
                    drawDistortedMenuTab i - 1
                End If
            Else
                If mePTStylesNormal = ptRectangle Then
                    drawRectangleTab i - 1
                ElseIf mePTStylesNormal = ptRoundedRectangle Then
                    drawRoundedTab i - 1
                ElseIf mePTStylesNormal = ptCornerCutLeft Then
                    drawCornerCutLeft i - 1
                ElseIf mePTStylesNormal = ptCornerCutRight Then
                    drawCornerCutRight i - 1
                ElseIf mePTStylesNormal = ptCoolLeft Then
                    drawCoolLeftTab i - 1
                ElseIf mePTStylesNormal = ptCoolRight Then
                    drawCoolRightTab i - 1
                    needRedraw = True
                    'tmpInd = i
                ElseIf mePTStylesNormal = ptDistorted Then
                    drawDistortedTab i - 1
                ElseIf mePTStylesNormal = ptVerticalLine Then
                    drawVerLineTab i - 1
                ElseIf mePTStylesNormal = ptRoundMenu Then
                    drawRoundMenuTab i - 1
                ElseIf mePTStylesNormal = ptXP Then
                    drawRoundedXPTab i - 1
                ElseIf mePTStylesNormal = ptSSTab Then
                    drawSSTab i - 1
                ElseIf mePTStylesNormal = ptFlatButton Then
                    drawFlatButtonTab i - 1
                ElseIf mePTStylesNormal = ptProTab Then
                    drawProTabTab i - 1
    
                ElseIf mePTStylesNormal = ptRoundedRectangle2 Then
                    drawRoundedRectangle2Tab i - 1
                ElseIf mePTStylesNormal = ptDistortedMenu Then
                    drawDistortedMenuTab i - 1
                End If
            End If
            
            '
            'If mX < 0 Then lstPositions.Add "-100" Else lstPositions.Add mX
            mX = mX + buttSpacing
            If mX > UserControl.ScaleWidth Then Exit For
        Else
            lstPositions.Add -50
            lstSizes.Add -50
'            lstPositions.Add -50
            'MsgBox i
        End If
        'DoEvents
    Next i
    'auto size
    If aSize = True Then UserControl.Width = (mX + 1) * Screen.TwipsPerPixelX
    'redraw if need
    If needRedraw = True Then
        tmpMX = mX
        mX = RDX
        setValues mTmpInd - 1, mBDRC, mBC, mShadC
        If mePTStylesActive = ptCoolRight Then
            drawCoolRightTab selIndex
        ElseIf mePTStylesActive = ptRoundedRectangle2 Then
            drawRoundedRectangle2Tab selIndex
        End If
        mX = tmpMX
    End If
    'draw scroll buttons
    If shSlideButtons = True Then drawSlideButtons
End Sub
'/////////////////////////////////////////////////////////////////////
'////////////////s l i d e    b u t t o n s///////////////////////////
'/////////////////////////////////////////////////////////////////////
Private Sub drawSlideButtons()
    Dim mBDRCL1 As OLE_COLOR, mBDRCR1 As OLE_COLOR, mBC1 As OLE_COLOR, mFillC1 As OLE_COLOR, mBDRCL11 As OLE_COLOR, mBDRCR11 As OLE_COLOR
    Dim mBDRCL2 As OLE_COLOR, mBDRCR2 As OLE_COLOR, mBC2 As OLE_COLOR, mFillC2 As OLE_COLOR, mBDRCL21 As OLE_COLOR, mBDRCR21 As OLE_COLOR
    Dim slideHeig As Integer, mTabHeig As Integer
    Dim mStartX1 As Integer, mY1 As Integer, i As Integer, ukWid As Long
    '
    Const fillHeig As Byte = 8
    Const fillWid As Byte = 4
    
    'ucSW = UserControl.ScaleWidth
    'ucSH = UserControl.ScaleHeight
    ukWid = 0
    For i = 1 To lstSizes.Count
        ukWid = ukWid + 1
    Next i
    '
    If downSlide = 1 Then
        mBDRCL1 = slideShadC
        mBDRCL11 = slideBDRC
        mBDRCR1 = slideHLC
        mBDRCR11 = slideBC
        mBC1 = slideBCDown
        mFillC1 = slideFillCDown
    ElseIf hoverSlide = 1 Then
        mBDRCL1 = slideBC
        mBDRCL11 = slideHLC
        mBDRCR1 = slideBDRC
        mBDRCR11 = slideShadC
        mBC1 = slideBCHover
        mFillC1 = slideFillCHover
    Else
        mBDRCL1 = slideBC
        mBDRCL11 = slideHLC
        mBDRCR1 = slideBDRC
        mBDRCR11 = slideShadC
        mBC1 = slideBC
        mFillC1 = slideFillC
    End If
    If downSlide = 2 Then
        mBDRCL2 = slideShadC
        mBDRCL21 = slideBDRC
        mBDRCR2 = slideHLC
        mBDRCR21 = slideBC
        mBC2 = slideBCDown
        mFillC2 = slideFillCDown
    ElseIf hoverSlide = 2 Then
        mBDRCL2 = slideBC
        mBDRCL21 = slideHLC
        mBDRCR2 = slideBDRC
        mBDRCR21 = slideShadC
        mBC2 = slideBCHover
        mFillC2 = slideFillCHover
    Else
        mBDRCL2 = slideBC
        mBDRCL21 = slideHLC
        mBDRCR2 = slideBDRC
        mBDRCR21 = slideShadC
        mBC2 = slideBC
        mFillC2 = slideFillC
    End If
    '
    scrLEnb = True
    scrREnb = True
    '
    If tabHeig > tabHeigActive Then mTabHeig = tabHeig Else mTabHeig = tabHeigActive
    If mTabHeig Mod 2 <> 0 Then mTabHeig = mTabHeig + 1
    slideHeig = Int(mTabHeig / 1.5)
    If slideHeig Mod 2 <> 0 Then slideHeig = slideHeig + 1
    '
    If mePTSlideAlign = ptScrollLeft Then
        mStartX1 = 0
        
        If firstItemL = 1 Then
            mBDRCL1 = slideBDRCDisabled
            mBDRCR1 = slideBDRCDisabled
            mBDRCL11 = slideBCDisabled
            mBDRCR11 = slideBCDisabled
            mBC1 = slideBCDisabled
            mFillC1 = slideFillCDisabled
            scrLEnb = False
            If tmrSlide.Enabled = True Then tmrSlide.Enabled = False
        End If
        
        '
        If mX < ucSW - 5 Then
            mBDRCL2 = slideBDRCDisabled
            mBDRCR2 = slideBDRCDisabled
            mBDRCL21 = slideBCDisabled
            mBDRCR21 = slideBCDisabled
            mBC2 = slideBCDisabled
            mFillC2 = slideFillCDisabled
            scrREnb = False
            If tmrSlide.Enabled = True Then tmrSlide.Enabled = False
        End If
    Else
        mStartX1 = ucSW - 5 - slideHeig * 2
        Dim mXI As Integer
        If mePTStylesNormal = ptCoolLeft Or mePTStylesActive = ptCoolLeft Then
            mXI = 12
        ElseIf mePTStylesNormal = ptXP Or mePTStylesActive = ptXP Then
            mXI = -3
        End If
        'MsgBox lstCaptions.Count & vbCrLf & firstItemL & "+" & lstSizes.Count
        If mX < UserControl.ScaleWidth - mTabHeig * 2 / 1.5 - 10 Then
            mBDRCL2 = slideBDRCDisabled
            mBDRCR2 = slideBDRCDisabled
            mBC2 = slideBCDisabled
            mFillC2 = slideFillCDisabled
            scrREnb = False
            If tmrSlide.Enabled = True Then tmrSlide.Enabled = False
        End If
        '
        If firstItemL = 1 Then
            mBDRCL1 = slideBDRCDisabled
            mBDRCR1 = slideBDRCDisabled
            mBC1 = slideBCDisabled
            mFillC1 = slideFillCDisabled
            scrLEnb = False
            If tmrSlide.Enabled = True Then tmrSlide.Enabled = False
        End If
    End If
    '
    If mePTOrientation = ptOrientationTop Then mY1 = 0 Else mY1 = ucSH - mTabHeig
    If mePTSlideAlign = ptScrollLeft Then
        If hSlideBack = False Then UserControl.Line (mStartX1, mY1)-(slideHeig + 1 + mStartX1 + 2 + slideHeig, mY1 + mTabHeig - 2), tabAreaBC, BF
    Else
        If hSlideBack = False Then UserControl.Line (mStartX1, mY1)-(slideHeig + 1 + mStartX1 + 5 + slideHeig, mY1 + mTabHeig - 2), tabAreaBC, BF
    End If
    'draw box
    '   1
    
    UserControl.Line (mStartX1 + 2, mY1 + mTabHeig / 2 - slideHeig / 2)-(mStartX1 + 2 + slideHeig, mY1 + mTabHeig / 2 + slideHeig / 2), mBC1, BF
    If scrHover <> True Or downSlide = 1 Or hoverSlide = 1 Then
        '
        If scrHover <> True Or scrLEnb = True Then
            If mePTAppearance <> ptAppearanceFlat Then
                UserControl.Line (mStartX1 + 2, mY1 + mTabHeig / 2 - slideHeig / 2)-(mStartX1 + 2, mY1 + mTabHeig / 2 + slideHeig / 2), mBDRCL1
                UserControl.Line (mStartX1 + 2, mY1 + mTabHeig / 2 - slideHeig / 2)-(mStartX1 + 2 + slideHeig, mY1 + mTabHeig / 2 - slideHeig / 2), mBDRCL1
                UserControl.Line (mStartX1 + 3, mY1 + mTabHeig / 2 - slideHeig / 2 + 1)-(mStartX1 + 3, mY1 + mTabHeig / 2 + slideHeig / 2 - 1), mBDRCL11
                UserControl.Line (mStartX1 + 3, mY1 + mTabHeig / 2 - slideHeig / 2 + 1)-(mStartX1 + 2 + slideHeig - 1, mY1 + mTabHeig / 2 - slideHeig / 2 + 1), mBDRCL11
                '
                UserControl.Line (mStartX1 + 2, mY1 + mTabHeig / 2 + slideHeig / 2)-(mStartX1 + 2 + slideHeig, mY1 + mTabHeig / 2 + slideHeig / 2), mBDRCR1
                UserControl.Line (mStartX1 + 2 + slideHeig, mY1 + mTabHeig / 2 + slideHeig / 2)-(mStartX1 + 2 + slideHeig, mY1 + mTabHeig / 2 - slideHeig / 2 - 1), mBDRCR1
                UserControl.Line (mStartX1 + 3, mY1 + mTabHeig / 2 + slideHeig / 2 - 1)-(mStartX1 + 2 + slideHeig, mY1 + mTabHeig / 2 + slideHeig / 2 - 1), mBDRCR11
                UserControl.Line (mStartX1 + 1 + slideHeig, mY1 + mTabHeig / 2 + slideHeig / 2 - 1)-(mStartX1 + 1 + slideHeig, mY1 + mTabHeig / 2 - slideHeig / 2), mBDRCR11
            Else
                UserControl.FillStyle = 1
                If scrLEnb = True Then
                    UserControl.Line (mStartX1 + 2, mY1 + mTabHeig / 2 - slideHeig / 2)-(mStartX1 + 2 + slideHeig, mY1 + mTabHeig / 2 + slideHeig / 2), slideBDRC, B
                Else
                    UserControl.Line (mStartX1 + 2, mY1 + mTabHeig / 2 - slideHeig / 2)-(mStartX1 + 2 + slideHeig, mY1 + mTabHeig / 2 + slideHeig / 2), slideBDRCDisabled, B
                End If
                UserControl.FillStyle = 0
            End If
        End If
    End If
        '   2
    
    UserControl.Line (slideHeig + 1 + mStartX1 + 2, mY1 + mTabHeig / 2 - slideHeig / 2)-(slideHeig + 1 + mStartX1 + 2 + slideHeig, mY1 + mTabHeig / 2 + slideHeig / 2), mBC2, BF
    If scrHover <> True Or downSlide = 2 Or hoverSlide = 2 Then
        If scrHover <> True Or scrREnb = True Then
            If mePTAppearance <> ptAppearanceFlat Then
                UserControl.Line (slideHeig + 1 + mStartX1 + 2, mY1 + mTabHeig / 2 - slideHeig / 2)-(slideHeig + 1 + mStartX1 + 2, mY1 + mTabHeig / 2 + slideHeig / 2), mBDRCL2
                UserControl.Line (slideHeig + 1 + mStartX1 + 2, mY1 + mTabHeig / 2 - slideHeig / 2)-(slideHeig + 1 + mStartX1 + 2 + slideHeig, mY1 + mTabHeig / 2 - slideHeig / 2), mBDRCL2
                UserControl.Line (slideHeig + 2 + mStartX1 + 2, mY1 + mTabHeig / 2 - slideHeig / 2 + 1)-(slideHeig + 2 + mStartX1 + 2, mY1 + mTabHeig / 2 + slideHeig / 2 - 1), mBDRCL21
                UserControl.Line (slideHeig + 2 + mStartX1 + 2, mY1 + mTabHeig / 2 - slideHeig / 2 + 1)-(slideHeig + 1 + mStartX1 + 2 + slideHeig - 1, mY1 + mTabHeig / 2 - slideHeig / 2 + 1), mBDRCL21
                '
                UserControl.Line (slideHeig + 1 + mStartX1 + 2, mY1 + mTabHeig / 2 + slideHeig / 2)-(slideHeig + 1 + mStartX1 + 2 + slideHeig, mY1 + mTabHeig / 2 + slideHeig / 2), mBDRCR2
                UserControl.Line (slideHeig + 1 + mStartX1 + 2 + slideHeig, mY1 + mTabHeig / 2 + slideHeig / 2)-(slideHeig + 1 + mStartX1 + 2 + slideHeig, mY1 + mTabHeig / 2 - slideHeig / 2 - 1), mBDRCR2
            
                UserControl.Line (slideHeig + 1 + mStartX1 + 2 + 1, mY1 + mTabHeig / 2 + slideHeig / 2 - 1)-(slideHeig + 1 + mStartX1 + 2 + slideHeig - 1, mY1 + mTabHeig / 2 + slideHeig / 2 - 1), mBDRCR21
                UserControl.Line (slideHeig + 1 + mStartX1 + 1 + slideHeig, mY1 + mTabHeig / 2 + slideHeig / 2 - 1)-(slideHeig + 1 + mStartX1 + 1 + slideHeig, mY1 + mTabHeig / 2 - slideHeig / 2), mBDRCR21
                
            Else
                UserControl.FillStyle = 1
                If scrREnb = True Then
                    UserControl.Line (slideHeig + 1 + mStartX1 + 2, mY1 + mTabHeig / 2 - slideHeig / 2)-(slideHeig + 1 + mStartX1 + 2 + slideHeig, mY1 + mTabHeig / 2 + slideHeig / 2), slideBDRC, B
                Else
                    UserControl.Line (slideHeig + 1 + mStartX1 + 2, mY1 + mTabHeig / 2 - slideHeig / 2)-(slideHeig + 1 + mStartX1 + 2 + slideHeig, mY1 + mTabHeig / 2 + slideHeig / 2), slideBDRCDisabled, B
                End If
                UserControl.FillStyle = 0
            End If
        End If
    End If
    '
    If mePTSlideButtStyle = ptArrow Then
        '1
        UserControl.Line (mStartX1 + 2 + slideHeig / 2 + fillWid / 2, mY1 + mTabHeig / 2 - fillHeig / 2)-(mStartX1 + 2 + slideHeig / 2 - fillWid / 2 - 1, mY1 + mTabHeig / 2 + 1), mFillC1
        UserControl.Line (mStartX1 + 2 + slideHeig / 2 + fillWid / 2, mY1 + mTabHeig / 2 - fillHeig / 2 + 1)-(mStartX1 + 2 + slideHeig / 2 - fillWid / 2, mY1 + mTabHeig / 2 + 1), mFillC1
        '
        UserControl.Line (mStartX1 + 2 + slideHeig / 2 + fillWid / 2, mY1 + mTabHeig / 2 + fillHeig / 2)-(mStartX1 + 2 + slideHeig / 2 - fillWid / 2, mY1 + mTabHeig / 2), mFillC1
        UserControl.Line (mStartX1 + 2 + slideHeig / 2 + fillWid / 2, mY1 + mTabHeig / 2 + fillHeig / 2 - 1)-(mStartX1 + 2 + slideHeig / 2 - fillWid / 2, mY1 + mTabHeig / 2), mFillC1
        '
        '2
        UserControl.Line (slideHeig + 1 + mStartX1 + 2 + slideHeig / 2 - fillWid / 2, mY1 + mTabHeig / 2 - fillHeig / 2)-(slideHeig + 1 + mStartX1 + 2 + slideHeig / 2 + fillWid / 2 + 1, mY1 + mTabHeig / 2 + 1), mFillC2
        UserControl.Line (slideHeig + 1 + mStartX1 + 2 + slideHeig / 2 - fillWid / 2, mY1 + mTabHeig / 2 - fillHeig / 2 + 1)-(slideHeig + 1 + mStartX1 + 2 + slideHeig / 2 + fillWid / 2, mY1 + mTabHeig / 2 + 1), mFillC2
        '
        UserControl.Line (slideHeig + 1 + mStartX1 + 2 + slideHeig / 2 - fillWid / 2, mY1 + mTabHeig / 2 + fillHeig / 2)-(slideHeig + 1 + mStartX1 + 2 + slideHeig / 2 + fillWid / 2, mY1 + mTabHeig / 2), mFillC2
        UserControl.Line (slideHeig + 1 + mStartX1 + 2 + slideHeig / 2 - fillWid / 2, mY1 + mTabHeig / 2 + fillHeig / 2 - 1)-(slideHeig + 1 + mStartX1 + 2 + slideHeig / 2 + fillWid / 2, mY1 + mTabHeig / 2), mFillC2
    ElseIf mePTSlideButtStyle = ptTrinangle Then
        '1
        UserControl.Line (mStartX1 + 2 + slideHeig / 2 + fillWid / 2, mY1 + mTabHeig / 2 - fillHeig / 2)-(mStartX1 + 2 + slideHeig / 2 - fillWid / 2 - 1, mY1 + mTabHeig / 2 + 1), mFillC1
        UserControl.Line (mStartX1 + 2 + slideHeig / 2 + fillWid / 2, mY1 + mTabHeig / 2 - fillHeig / 2)-(mStartX1 + 2 + slideHeig / 2 + fillWid / 2, mY1 + mTabHeig / 2 + fillHeig / 2), mFillC1
        UserControl.Line (mStartX1 + 2 + slideHeig / 2 + fillWid / 2, mY1 + mTabHeig / 2 + fillHeig / 2)-(mStartX1 + 2 + slideHeig / 2 - fillWid / 2, mY1 + mTabHeig / 2), mFillC1
        '
        '2
        UserControl.Line (slideHeig + 1 + mStartX1 + 2 + slideHeig / 2 - fillWid / 2, mY1 + mTabHeig / 2 - fillHeig / 2)-(slideHeig + 1 + mStartX1 + 2 + slideHeig / 2 + fillWid / 2 + 1, mY1 + mTabHeig / 2 + 1), mFillC2
        UserControl.Line (slideHeig + 1 + mStartX1 + 2 + slideHeig / 2 - fillWid / 2, mY1 + mTabHeig / 2 - fillHeig / 2)-(slideHeig + 1 + mStartX1 + 2 + slideHeig / 2 - fillWid / 2, mY1 + mTabHeig / 2 + fillHeig / 2), mFillC2
        UserControl.Line (slideHeig + 1 + mStartX1 + 2 + slideHeig / 2 - fillWid / 2, mY1 + mTabHeig / 2 + fillHeig / 2)-(slideHeig + 1 + mStartX1 + 2 + slideHeig / 2 + fillWid / 2, mY1 + mTabHeig / 2), mFillC2
        'UserControl.Line (slideHeig + 1 + mStartX1 + 2 + slideHeig / 2 - fillWid / 2, mY1 + mTabHeig / 2 + fillHeig / 2 - 1)-(slideHeig + 1 + mStartX1 + 2 + slideHeig / 2 + fillWid / 2, mY1 + mTabHeig / 2), mFillC2
    Else
        '1
        UserControl.Line (mStartX1 + 2 + slideHeig / 2 + fillWid / 2, mY1 + mTabHeig / 2 - fillHeig / 2)-(mStartX1 + 2 + slideHeig / 2 - fillWid / 2 - 1, mY1 + mTabHeig / 2 + 1), mFillC1
        UserControl.Line (mStartX1 + 2 + slideHeig / 2 + fillWid / 2, mY1 + mTabHeig / 2 - fillHeig / 2)-(mStartX1 + 2 + slideHeig / 2 + fillWid / 2, mY1 + mTabHeig / 2 + fillHeig / 2), mFillC1
        UserControl.Line (mStartX1 + 2 + slideHeig / 2 + fillWid / 2, mY1 + mTabHeig / 2 + fillHeig / 2)-(mStartX1 + 2 + slideHeig / 2 - fillWid / 2, mY1 + mTabHeig / 2), mFillC1
        UserControl.FillColor = mFillC1
        ExtFloodFill UserControl.hdc, mStartX1 + 2 + slideHeig / 2 + fillWid / 2 - 1, mY1 + mTabHeig / 2 + 1, UserControl.Point(mStartX1 + 2 + slideHeig / 2 + fillWid / 2 - 1, mY1 + mTabHeig / 2 + 1), 1
        '2
        UserControl.Line (slideHeig + 1 + mStartX1 + 2 + slideHeig / 2 - fillWid / 2, mY1 + mTabHeig / 2 - fillHeig / 2)-(slideHeig + 1 + mStartX1 + 2 + slideHeig / 2 + fillWid / 2 + 1, mY1 + mTabHeig / 2 + 1), mFillC2
        UserControl.Line (slideHeig + 1 + mStartX1 + 2 + slideHeig / 2 - fillWid / 2, mY1 + mTabHeig / 2 - fillHeig / 2)-(slideHeig + 1 + mStartX1 + 2 + slideHeig / 2 - fillWid / 2, mY1 + mTabHeig / 2 + fillHeig / 2), mFillC2
        UserControl.Line (slideHeig + 1 + mStartX1 + 2 + slideHeig / 2 - fillWid / 2, mY1 + mTabHeig / 2 + fillHeig / 2)-(slideHeig + 1 + mStartX1 + 2 + slideHeig / 2 + fillWid / 2, mY1 + mTabHeig / 2), mFillC2
        UserControl.FillColor = mFillC2
        ExtFloodFill UserControl.hdc, slideHeig + 1 + mStartX1 + 2 + slideHeig / 2 - fillWid / 2 + 1, mY1 + mTabHeig / 2 + 1, UserControl.Point(slideHeig + 1 + mStartX1 + 2 + slideHeig / 2 - fillWid / 2 + 1, mY1 + mTabHeig / 2 + 1), 1
    End If
End Sub


'/////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////
Private Sub drawRectangleTab(ByVal mInd As Integer)
    Dim mTabHeig As Integer, mYInc As Integer
    Dim fntMovPC As Byte, pcMov As Byte
    '
    Dim havePIC As Boolean
    'setValues mInd, mBDRC, mBC, mShadC
    'ucSW = UserControl.ScaleWidth
    'ucSH = UserControl.ScaleHeight
    'find tab width
    bWid = UserControl.TextWidth(lstCaptions.Item(mInd + 1)) + mSpacing * 2 + 10
    If mePTStylesActive <> ptCoolLeft And mePTStylesNormal <> ptCoolLeft Then mX = mX + getWidInc(mInd)
    'if there is picture then increase tab width
    havePIC = False
    If Not tabIcons(mInd) Is Nothing Then
        If mePTPicAlig <> ptPicCenter Then bWid = bWid + icWid + mSpacing + 5 '* 2
        havePIC = True
    End If
    'add width to collection
    lstSizes.Add bWid
    '
    If bWid + mX < 0 Or mX > ucSW Then
        mX = mX + bWid
        Exit Sub
    End If
    '
    'If UserControl.ForeColor <> mFC Then UserControl.ForeColor = mFC
    'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
    'find tab height and y position
    mTabHeig = tabHeig
    mYInc = Abs(tabHeigActive - tabHeig)
    If mInd = selIndex Then
        mTabHeig = tabHeigActive
        If tabHeigActive > tabHeig Then mYInc = 0
    Else
        If tabHeig > tabHeigActive Then mYInc = 0
    End If
    'draw tab
    If mePTOrientation = ptOrientationTop Then
        UserControl.Line (mX, mYInc)-(mX + bWid, mTabHeig + mYInc), mBC, BF
        'draw shadow
        If mePTAppearance = ptAppearance3D Then
            'center
            UserControl.Line (mX + 1, 1 + mYInc)-(mX + bWid - 1, 1 + mYInc), controlHLightC
            'right
            UserControl.Line (mX + bWid - 1, 1 + mYInc)-(mX + bWid - 1, mTabHeig + mYInc + 1), controlShadC
            'left
            UserControl.Line (mX + 1, 1 + mYInc)-(mX + 1, mTabHeig + mYInc + 1), controlHLightC
        End If
        'border
        UserControl.FillStyle = 1
        UserControl.Line (mX, mYInc)-(mX + bWid, mTabHeig + mYInc), mBDRC, B
        UserControl.FillStyle = 0
        'find font y-position
        If mePTVertTextAlg = ptTop Then
            UserControl.CurrentY = 2 + mYInc
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 + mYInc
        Else
            UserControl.CurrentY = mTabHeig - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 + mYInc
        End If
    Else
        UserControl.Line (mX, ucSH - 1 - mYInc)-(mX + bWid, ucSH - mTabHeig - 1 - mYInc), mBC, BF
        If mePTAppearance = ptAppearance3D Then
            'center
            UserControl.Line (mX + 1, ucSH - 1 - 1 - mYInc)-(mX + bWid - 1, ucSH - 1 - 1 - mYInc), controlShadC
            'right
            UserControl.Line (mX + bWid - 1, ucSH - 1 - 1 - mYInc)-(mX + bWid - 1, ucSH - 1 - mTabHeig - mYInc - 1), controlShadC
            'left
            UserControl.Line (mX + 1, ucSH - 1 - 1 - mYInc)-(mX + 1, ucSH - 1 - mTabHeig - mYInc - 1), controlHLightC
        End If
        'border
        UserControl.FillStyle = 1
        UserControl.Line (mX, ucSH - 1 - mYInc)-(mX + bWid, ucSH - mTabHeig - 1 - mYInc), mBDRC, B
        UserControl.FillStyle = 0
        'find font y-position
        If mePTVertTextAlg = ptBottom Then
            UserControl.CurrentY = ucSH - mTabHeig + 2 - mYInc '+ UserControl.TextHeight(lstCaptions.Item(mInd + 1)) + 5
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = ucSH - mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 - mYInc
        Else
            UserControl.CurrentY = ucSH - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 - mYInc
        End If
    End If
    'draw picture - if exist
    putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig
    'find font x-position
    UserControl.CurrentX = getFontX(bWid, mInd, havePIC, mePTPicAlig)
    If mInd = selIndex And efectSel = True Then
        If mePTOrientation = ptOrientationTop Then
            UserControl.CurrentY = UserControl.CurrentY - 1
        Else
            UserControl.CurrentY = UserControl.CurrentY + 1
        End If
        UserControl.CurrentX = UserControl.CurrentX - 1
    End If
    'print caption
    UserControl.Print lstCaptions.Item(mInd + 1)
    'if this is active tab then hide line at bottom of tab
    hideBottom mInd, mX, mYInc, bWid, mTabHeig, mBC, controlHLightC, controlShadC

    'find position of new (next) tab - if exist
    mX = mX + bWid
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub drawFlatButtonTab(ByVal mInd As Integer)
    Dim mTabHeig As Integer, mYInc As Integer
    Dim fntMovPC As Byte, pcMov As Byte
    '
    Dim havePIC As Boolean
    'find tab width
    bWid = UserControl.TextWidth(lstCaptions.Item(mInd + 1)) + mSpacing * 2 + 10
    If mePTStylesActive <> ptCoolLeft And mePTStylesNormal <> ptCoolLeft Then mX = mX + getWidInc(mInd)
    'if there is picture then increase tab width
    havePIC = False
    If Not tabIcons(mInd) Is Nothing Then
        If mePTPicAlig <> ptPicCenter Then bWid = bWid + icWid + mSpacing + 5 '* 2
        havePIC = True
    End If
    'add width to collection
    lstSizes.Add bWid
    '
    If bWid + mX < 0 Or mX > ucSW Then
        mX = mX + bWid
        Exit Sub
    End If
    'find tab height and y position
    mTabHeig = tabHeig
    mYInc = Abs(tabHeigActive - tabHeig)
    If mInd = selIndex Then
        mTabHeig = tabHeigActive
        If tabHeigActive > tabHeig Then mYInc = 0
    Else
        If tabHeig > tabHeigActive Then mYInc = 0
    End If
    'draw tab
    If mePTOrientation = ptOrientationTop Then
        'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
        UserControl.Line (mX, mYInc)-(mX + bWid, mTabHeig + mYInc), mBC, BF
        UserControl.Line (mX, mYInc)-(mX + bWid, mTabHeig + mYInc), mBDRC, B
        'draw shadow
        'find font y-position
        If mePTVertTextAlg = ptTop Then
            UserControl.CurrentY = 2 + mYInc
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 + mYInc
        Else
            UserControl.CurrentY = mTabHeig - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 + mYInc
        End If
    Else
        'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
        UserControl.Line (mX, ucSH - 1 - mYInc)-(mX + bWid, ucSH - mTabHeig - 1 - mYInc), mBC, BF
        UserControl.Line (mX, ucSH - 1 - mYInc)-(mX + bWid, ucSH - mTabHeig - 1 - mYInc), mBDRC, B
        'find font y-position
        If mePTVertTextAlg = ptBottom Then
            UserControl.CurrentY = ucSH - mTabHeig + 2 - mYInc '+ UserControl.TextHeight(lstCaptions.Item(mInd + 1)) + 5
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = ucSH - mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 - mYInc
        Else
            UserControl.CurrentY = ucSH - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 - mYInc
        End If
    End If
    'draw picture - if exist
    putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig
    'find font x-position
    UserControl.CurrentX = getFontX(bWid, mInd, havePIC, mePTPicAlig)
    If mInd = selIndex And efectSel = True Then
        If mePTOrientation = ptOrientationTop Then
            UserControl.CurrentY = UserControl.CurrentY - 1
        Else
            UserControl.CurrentY = UserControl.CurrentY + 1
        End If
        UserControl.CurrentX = UserControl.CurrentX - 1
    End If
    'print caption
    UserControl.Print lstCaptions.Item(mInd + 1)
    'find position of new (next) tab - if exist
    mX = mX + bWid
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub drawProTabTab(ByVal mInd As Integer)
    Dim mTabHeig As Integer, mYInc As Integer
    Dim fntMovPC As Byte, pcMov As Byte, mWidInc1 As Integer
    '
    Dim havePIC As Boolean

    'find tab width
    bWid = UserControl.TextWidth(lstCaptions.Item(mInd + 1)) + mSpacing * 2 + 10
    If mInd = selIndex Then
        bWid = bWid + tabHeigActive * Cos(60 * 3.14 / 180) * 2
        mWidInc1 = tabHeigActive * Cos(60 * 3.14 / 180)
    Else
        bWid = bWid + tabHeig * Cos(60 * 3.14 / 180) * 2
        mWidInc1 = tabHeig * Cos(60 * 3.14 / 180)
    End If
    'if there is picture then increase tab width
    havePIC = False
    If Not tabIcons(mInd) Is Nothing Then
        If mePTPicAlig <> ptPicCenter Then bWid = bWid + icWid + mSpacing + 5 '* 2
        havePIC = True
    End If
    'add width to collection
    lstSizes.Add bWid
    '
    If bWid + mX < 0 Or mX > ucSW Then
        mX = mX + bWid
        Exit Sub
    End If
    'find tab height and y position
    mTabHeig = tabHeig
    mYInc = Abs(tabHeigActive - tabHeig)
    If mInd = selIndex Then
        mTabHeig = tabHeigActive
        If tabHeigActive > tabHeig Then mYInc = 0
    Else
        If tabHeig > tabHeigActive Then mYInc = 0
    End If
    '
    If mInd <> selIndex Then mYInc = 0 'mYInc - 2
    'draw tab
    If mePTOrientation = ptOrientationTop Then
        'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
        'draw selected tab
        If mInd = selIndex Then
            'draw left line
            UserControl.Line (mX, mYInc + mTabHeig)-(mX + mWidInc1, mYInc), mBDRC
            'center
            UserControl.Line (mX + mWidInc1, mYInc)-(mX + bWid - mWidInc1, mYInc), mBDRC
            'right line
            UserControl.Line (mX + bWid - mWidInc1, mYInc)-(mX + bWid, mYInc + mTabHeig), mBDRC
            'fill
            ExtFloodFill UserControl.hdc, mX + 2, mYInc + mTabHeig - 2, UserControl.Point(mX + 2, mYInc + mTabHeig - 2), 1
            '---fill at right side---
            'ExtFloodFill usercontrol.hdc, mX + bWid - 2, mYInc + mTabHeig - 1, UserControl.Point(mX + bWid - 2, mYInc + mTabHeig - 1), 1
            'draw shadow
            If mePTAppearance = ptAppearance3D Then
                UserControl.Line (mX + 1, mYInc + mTabHeig)-(mX + mWidInc1 + 1, mYInc), controlHLightC
                UserControl.Line (mX + mWidInc1, mYInc + 1)-(mX + bWid - mWidInc1, mYInc + 1), controlHLightC
                UserControl.Line (mX + bWid - mWidInc1 - 1, mYInc)-(mX + bWid - 1, mYInc + mTabHeig), controlShadC
                'redraw center because shadow
                UserControl.Line (mX + mWidInc1, mYInc)-(mX + bWid - mWidInc1, mYInc), mBDRC
            End If
        ElseIf mInd = selIndex + 1 Then
            'left
            UserControl.Line (mX, mYInc)-(mX + mWidInc1, mYInc + mTabHeig), mBDRC
            'center
            UserControl.Line (mX, mYInc)-(mX + bWid, mYInc), mBDRC
            'right
            UserControl.Line (mX + bWid - mWidInc1, mYInc + mTabHeig)-(mX + bWid, mYInc), mBDRC
            'bottom
            If tabHeig = tabHeigActive Then
                UserControl.Line (mX + mWidInc1, mYInc + mTabHeig)-(mX + bWid - mWidInc1 + 1, mYInc + mTabHeig), controlBorderC
            Else
                UserControl.Line (mX + mWidInc1, mYInc + mTabHeig)-(mX + bWid - mWidInc1 + 1, mYInc + mTabHeig), mBDRC
            End If
            'fill
            ExtFloodFill UserControl.hdc, mX + 2, mYInc + 1, UserControl.Point(mX + 2, mYInc + 1), 1
            '---fill at right side---
            'ExtFloodFill usercontrol.hdc, mX + bWid - 2, mYInc + 1, UserControl.Point(mX + bWid - 2, mYInc + 1), 1
            'draw shadow
            If mePTAppearance = ptAppearance3D Then
                UserControl.Line (mX + 1, mYInc)-(mX + mWidInc1 + 1, mYInc + mTabHeig), controlHLightC
                UserControl.Line (mX + 1, mYInc + 1)-(mX + bWid, mYInc + 1), controlHLightC
                UserControl.Line (mX + bWid - mWidInc1 - 1, mYInc + mTabHeig - 1)-(mX + bWid - 1, mYInc), controlShadC
                'redraw center because shadow
                UserControl.Line (mX, mYInc)-(mX + bWid, mYInc), mBDRC
            End If
        Else
            UserControl.Line (mX, mYInc + mTabHeig)-(mX + mWidInc1, mYInc), mBDRC
            UserControl.Line (mX + mWidInc1, mYInc)-(mX + bWid, mYInc), mBDRC
            UserControl.Line (mX + bWid - mWidInc1, mYInc + mTabHeig)-(mX + bWid, mYInc), mBDRC
            'bottom
            If tabHeig = tabHeigActive Then
                UserControl.Line (mX, mYInc + mTabHeig)-(mX + bWid - mWidInc1 + 1, mYInc + mTabHeig), controlBorderC
            Else
                UserControl.Line (mX, mYInc + mTabHeig)-(mX + bWid - mWidInc1 + 1, mYInc + mTabHeig), mBDRC
            End If
            ExtFloodFill UserControl.hdc, mX + 2, mYInc + mTabHeig - 2, UserControl.Point(mX + 2, mYInc + mTabHeig - 2), 1
            '---fill at right side---
            'ExtFloodFill usercontrol.hdc, mX + bWid - 2, mYInc + 1, UserControl.Point(mX + bWid - 2, mYInc + 1), 1
            'draw shadow
            If mePTAppearance = ptAppearance3D Then
                UserControl.Line (mX + 1, mYInc + mTabHeig - 1)-(mX + mWidInc1 + 1, mYInc), controlHLightC
                UserControl.Line (mX + mWidInc1, mYInc + 1)-(mX + bWid, mYInc + 1), controlHLightC
                UserControl.Line (mX + bWid - mWidInc1 - 1, mYInc + mTabHeig - 1)-(mX + bWid - 1, mYInc), controlShadC
            End If
        End If
        'find font y-position
        If mePTVertTextAlg = ptTop Then
            UserControl.CurrentY = 2 + mYInc
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 + mYInc
        Else
            UserControl.CurrentY = mTabHeig - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 + mYInc
        End If
    Else
        'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
        If mInd = selIndex Then
            UserControl.Line (mX, ucSH - 1 - mYInc - mTabHeig)-(mX + mWidInc1, ucSH - 1 - mYInc), mBDRC
            UserControl.Line (mX + mWidInc1, ucSH - 1 - mYInc)-(mX + bWid - mWidInc1, ucSH - 1 - mYInc), mBDRC
            UserControl.Line (mX + bWid - mWidInc1, ucSH - 1 - mYInc)-(mX + bWid, ucSH - 1 - mYInc - mTabHeig), mBDRC
            ExtFloodFill UserControl.hdc, mX + 2, ucSH - 1 - mYInc - mTabHeig + 2, UserControl.Point(mX + 2, ucSH - 1 - mYInc - mTabHeig + 2), 1
            '---fill at right side---
            'ExtFloodFill usercontrol.hdc, mX + bWid - 2, ucSH - 1 - mYInc - mTabHeig + 1, UserControl.Point(mX + bWid - 2, ucSH - 1 - mYInc - mTabHeig + 1), 1
            'draw shadow
            If mePTAppearance = ptAppearance3D Then
                UserControl.Line (mX + 1, ucSH - 1 - mYInc - mTabHeig)-(mX + mWidInc1 + 1, ucSH - 1 - mYInc), controlHLightC
                UserControl.Line (mX + mWidInc1, ucSH - 1 - mYInc - 1)-(mX + bWid - mWidInc1, ucSH - 1 - mYInc - 1), controlShadC
                UserControl.Line (mX + bWid - mWidInc1 - 1, ucSH - 1 - mYInc)-(mX + bWid - 1, ucSH - 1 - mYInc - mTabHeig), controlShadC
                'redraw center because shadow
                UserControl.Line (mX + mWidInc1, ucSH - 1 - mYInc)-(mX + bWid - mWidInc1, ucSH - 1 - mYInc), mBDRC
            End If
        ElseIf mInd = selIndex + 1 Then
            UserControl.Line (mX, ucSH - 1 - mYInc)-(mX + mWidInc1, ucSH - 1 - mYInc - mTabHeig), mBDRC
            UserControl.Line (mX, ucSH - 1 - mYInc)-(mX + bWid, ucSH - 1 - mYInc), mBDRC
            UserControl.Line (mX + mWidInc1, ucSH - 1 - mYInc - mTabHeig)-(mX + bWid - mWidInc1, ucSH - 1 - mYInc - mTabHeig), mBDRC
            UserControl.Line (mX + bWid - mWidInc1, ucSH - 1 - mYInc - mTabHeig)-(mX + bWid, ucSH - 1 - mYInc), mBDRC
            ExtFloodFill UserControl.hdc, mX + 2, ucSH - 1 - mYInc - 1, UserControl.Point(mX + 2, ucSH - 1 - mYInc - 1), 1
            '---fill at right side---
            'ExtFloodFill usercontrol.hdc, mX + bWid - 2, ucSH - 1 - mYInc - 1, UserControl.Point(mX + bWid - 2, ucSH - 1 - mYInc - 1), 1
            'draw shadow
            If mePTAppearance = ptAppearance3D Then
                UserControl.Line (mX + 1, ucSH - 1 - mYInc)-(mX + mWidInc1 + 1, ucSH - 1 - mYInc - mTabHeig), controlHLightC
                UserControl.Line (mX + 1, ucSH - 1 - mYInc - 1)-(mX + bWid, ucSH - 1 - mYInc - 1), controlShadC
                UserControl.Line (mX + bWid - mWidInc1 - 1, ucSH - 1 - mYInc - mTabHeig)-(mX + bWid - 1, ucSH - 1 - mYInc), controlShadC
                'redraw center
                UserControl.Line (mX, ucSH - 1 - mYInc)-(mX + bWid, ucSH - 1 - mYInc), mBDRC
                UserControl.Line (mX + mWidInc1, ucSH - 1 - mYInc - mTabHeig)-(mX + bWid - mWidInc1, ucSH - 1 - mYInc - mTabHeig), mBDRC
            End If
        Else
            UserControl.Line (mX, ucSH - 1 - mYInc - mTabHeig)-(mX + mWidInc1, ucSH - 1 - mYInc), mBDRC
            UserControl.Line (mX + mWidInc1, ucSH - 1 - mYInc)-(mX + bWid, ucSH - 1 - mYInc), mBDRC
            UserControl.Line (mX + bWid - mWidInc1, ucSH - 1 - mYInc - mTabHeig)-(mX + bWid, ucSH - 1 - mYInc), mBDRC
            UserControl.Line (mX, ucSH - 1 - mYInc - mTabHeig)-(mX + bWid - mWidInc1, ucSH - 1 - mYInc - mTabHeig), mBDRC
            ExtFloodFill UserControl.hdc, mX + 2, ucSH - 1 - mYInc - mTabHeig + 2, UserControl.Point(mX + 2, ucSH - 1 - mYInc - mTabHeig + 2), 1
            '---fill at right side---
            'ExtFloodFill usercontrol.hdc, mX + bWid - 2, ucSH - 1 - mYInc - 1, UserControl.Point(mX + bWid - 2, ucSH - 1 - mYInc - 1), 1
            'draw shadow
            If mePTAppearance = ptAppearance3D Then
                UserControl.Line (mX + 1, ucSH - 1 - mYInc - mTabHeig)-(mX + mWidInc1 + 1, ucSH - 1 - mYInc), controlHLightC
                UserControl.Line (mX + mWidInc1, ucSH - 1 - mYInc - 1)-(mX + bWid, ucSH - 1 - mYInc - 1), controlShadC
                UserControl.Line (mX + bWid - mWidInc1 - 1, ucSH - 1 - mYInc - mTabHeig)-(mX + bWid - 1, ucSH - 1 - mYInc), controlShadC
                'redraw center
                UserControl.Line (mX, ucSH - 1 - mYInc - mTabHeig)-(mX + bWid - mWidInc1, ucSH - 1 - mYInc - mTabHeig), mBDRC
            End If
        End If
        'find font y-position
        If mePTVertTextAlg = ptBottom Then
            UserControl.CurrentY = ucSH - mTabHeig + 2 - mYInc '+ UserControl.TextHeight(lstCaptions.Item(mInd + 1)) + 5
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = ucSH - mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 - mYInc
        Else
            UserControl.CurrentY = ucSH - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 - mYInc
        End If
    End If
    'draw picture - if exist
    putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig
    'find font x-position
    UserControl.CurrentX = getFontX(bWid, mInd, havePIC, mePTPicAlig)
    If mInd = selIndex And efectSel = True Then
        If mePTOrientation = ptOrientationTop Then
            UserControl.CurrentY = UserControl.CurrentY - 1
        Else
            UserControl.CurrentY = UserControl.CurrentY + 1
        End If
        UserControl.CurrentX = UserControl.CurrentX - 1
    End If
    'print caption
    UserControl.Print lstCaptions.Item(mInd + 1)
    'if this is active tab then hide line at bottom of tab
    hideBottom mInd, mX, mYInc, bWid, mTabHeig, mBC, controlHLightC, controlShadC
    'find position of new (next) tab - if exist
    mX = mX + bWid
    If mInd <> tabCnt - 1 And mePTStylesNormal = mePTStylesActive Then
        mX = mX - mWidInc1
    'ElseIf mInd = tabCnt - 1 And mInd = selIndex Then
        'mX = mX - 1
    End If
    
    If mInd = selIndex - 1 Then mX = mX - Abs(tabHeig - tabHeigActive)
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub drawRoundMenuTab(ByVal mInd As Integer)
    Dim mTabHeig As Integer, mYInc As Integer
    Dim fntMovPC As Byte, pcMov As Byte
    '
    Dim havePIC As Boolean
    'setValues mInd, mBDRC, mBC, mShadC
    'ucSW = UserControl.ScaleWidth
    'ucSH = UserControl.ScaleHeight
    'find tab width
    bWid = UserControl.TextWidth(lstCaptions.Item(mInd + 1)) + mSpacing * 2 + 10
    If mePTStylesActive <> ptCoolLeft And mePTStylesNormal <> ptCoolLeft Then mX = mX + getWidInc(mInd)
    'if there is picture then increase tab width
    havePIC = False
    If Not tabIcons(mInd) Is Nothing Then
        If mePTPicAlig <> ptPicCenter Then bWid = bWid + icWid + mSpacing + 5 '* 2
        havePIC = True
    End If
    'add width to collection
    lstSizes.Add bWid
    '
    If bWid + mX < 0 Or mX > ucSW Then
        mX = mX + bWid
        Exit Sub
    End If
    'find tab height and y position
    mTabHeig = tabHeig
    mYInc = Abs(tabHeigActive - tabHeig)
    If mInd = selIndex Then
        mTabHeig = tabHeigActive
        If tabHeigActive > tabHeig Then mYInc = 0
    Else
        If tabHeig > tabHeigActive Then mYInc = 0
    End If
    If mInd = 0 Or mInd = tabCnt - 1 Then bWid = bWid + mTabHeig / 2
    '--draw tab--
    If mePTOrientation = ptOrientationTop Then
        'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
        If mInd = 0 Or mInd = tabCnt - 1 Then
            If mInd = 0 Then
                UserControl.Circle (mX + mTabHeig / 2, mYInc + mTabHeig / 2), mTabHeig / 2, mBC
                UserControl.FillStyle = 1
                UserControl.Circle (mX + mTabHeig / 2, mYInc + mTabHeig / 2), mTabHeig / 2, mBDRC
                UserControl.FillStyle = 0
                UserControl.Line (mX + mTabHeig / 2, mYInc)-(mX + bWid, mTabHeig + mYInc), mBC, BF
                UserControl.Line (mX + mTabHeig / 2, mYInc)-(mX + bWid, mTabHeig + mYInc), mBDRC, B
                '
                UserControl.Line (mX + mTabHeig / 2, mYInc + 1)-(mX + mTabHeig / 2, mYInc + mTabHeig), mBC
                'shadow
                If mePTAppearance = ptAppearance3D Then
                    'center
                    UserControl.Line (mX + mTabHeig / 2, 1 + mYInc)-(mX + bWid, 1 + mYInc), controlHLightC
                    UserControl.Line (mX + mTabHeig / 2 - 3, mYInc + mTabHeig - 1)-(mX + bWid, mYInc + mTabHeig - 1), controlShadC
                    'left
                    UserControl.Circle (mX + mTabHeig / 2, mYInc + mTabHeig / 2), mTabHeig / 2 - 1, controlHLightC, 3.14 / 2, (200 / 180) * 3.14
                    UserControl.Circle (mX + mTabHeig / 2, mYInc + mTabHeig / 2), mTabHeig / 2 - 1, controlHLightC, (200 / 180) * 3.14, 1.5 * 3.14
                    'right
                    UserControl.Line (mX + bWid - 1, mYInc + 1)-(mX + bWid - 1, mYInc + mTabHeig - 1), controlShadC
                End If
                UserControl.Circle (mX + mTabHeig / 2, mYInc + mTabHeig / 2), mTabHeig / 2, mBDRC, (90 / 180) * 3.14, 1.5 * 3.14
            Else
                UserControl.Circle (mX + bWid - mTabHeig / 2, mYInc + mTabHeig / 2), mTabHeig / 2, mBC
                UserControl.FillStyle = 1
                UserControl.Circle (mX + bWid - mTabHeig / 2, mYInc + mTabHeig / 2), mTabHeig / 2, mBDRC
                UserControl.FillStyle = 0
                UserControl.Line (mX, mYInc)-(mX + bWid - mTabHeig / 2, mTabHeig + mYInc), mBC, BF
                UserControl.Line (mX, mYInc)-(mX + bWid - mTabHeig / 2, mTabHeig + mYInc), mBDRC, B
                '
                UserControl.Line (mX + bWid - mTabHeig / 2, mYInc + 1)-(mX + bWid - mTabHeig / 2, mYInc + mTabHeig), mBC
                'shadow
                If mePTAppearance = ptAppearance3D Then
                    'right
                    UserControl.Circle (mX + bWid - mTabHeig / 2, mYInc + mTabHeig / 2), mTabHeig / 2 - 1, controlShadC, (340 / 180) * 3.14, 3.14 / 2
                    UserControl.Circle (mX + bWid - mTabHeig / 2, mYInc + mTabHeig / 2), mTabHeig / 2 - 1, controlShadC, (270 / 180) * 3.14, (340 / 180) * 3.14
                    'center
                    UserControl.Line (mX + 1, 1 + mYInc)-(mX + bWid - mTabHeig / 2 + 1, 1 + mYInc), controlHLightC
                    UserControl.Line (mX + 1, mYInc + mTabHeig - 1)-(mX + bWid - mTabHeig / 2 + 3, mYInc + mTabHeig - 1), controlShadC
                    'left
                    UserControl.Line (mX + 1, 1 + mYInc)-(mX + 1, mYInc + mTabHeig - 1), controlHLightC

                End If
                
                UserControl.Circle (mX + bWid - mTabHeig / 2, mYInc + mTabHeig / 2), mTabHeig / 2, mBDRC, (270 / 180) * 3.14, 3.14 / 2
            End If
        Else
            UserControl.Line (mX, mYInc)-(mX + bWid, mTabHeig + mYInc), mBC, BF
            UserControl.Line (mX, mYInc)-(mX + bWid, mTabHeig + mYInc), mBDRC, B
            'draw shadow
            If mePTAppearance = ptAppearance3D Then
                'center
                UserControl.Line (mX + 1, 1 + mYInc)-(mX + bWid, 1 + mYInc), controlHLightC
                UserControl.Line (mX + 1, mYInc + mTabHeig - 1)-(mX + bWid, mYInc + mTabHeig - 1), controlShadC
                'left
                UserControl.Line (mX + 1, 1 + mYInc)-(mX + 1, mYInc + mTabHeig - 1), controlHLightC
                'right
                UserControl.Line (mX + bWid - 1, mYInc + 2)-(mX + bWid - 1, mYInc + mTabHeig - 1), controlShadC
            End If
        End If
        'find font y-position
        If mePTVertTextAlg = ptTop Then
            UserControl.CurrentY = 2 + mYInc
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 + mYInc
        Else
            UserControl.CurrentY = mTabHeig - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 + mYInc
        End If
    Else
        'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
        If mInd = 0 Or mInd = tabCnt - 1 Then
            If mInd = 0 Then
                UserControl.Circle (mX + mTabHeig / 2, ucSH - 1 - mYInc - mTabHeig / 2), mTabHeig / 2, mBC
                UserControl.FillStyle = 1
                UserControl.Circle (mX + mTabHeig / 2, ucSH - 1 - mYInc - mTabHeig / 2), mTabHeig / 2, mBDRC
                UserControl.FillStyle = 0
                UserControl.Line (mX + mTabHeig / 2, ucSH - 1 - mYInc)-(mX + bWid, ucSH - 1 - mTabHeig - mYInc), mBC, BF
                UserControl.Line (mX + mTabHeig / 2, ucSH - 1 - mYInc)-(mX + bWid, ucSH - 1 - mTabHeig - mYInc), mBDRC, B
                '
                UserControl.Line (mX + mTabHeig / 2, ucSH - 1 - mYInc - 1)-(mX + mTabHeig / 2, ucSH - 1 - mYInc - mTabHeig), mBC
                'shadow
                If mePTAppearance = ptAppearance3D Then
                    'center
                    UserControl.Line (mX + mTabHeig / 2 - 2, ucSH - 1 - 1 - mYInc)-(mX + bWid, ucSH - 1 - 1 - mYInc), controlShadC
                    UserControl.Line (mX + mTabHeig / 2, ucSH - 1 - mYInc - mTabHeig + 1)-(mX + bWid, ucSH - 1 - mYInc - mTabHeig + 1), controlHLightC
                    'left
                    UserControl.Circle (mX + mTabHeig / 2, ucSH - 1 - mYInc - mTabHeig / 2), mTabHeig / 2 - 1, controlHLightC, 3.14 / 2, (200 / 180) * 3.14
                    UserControl.Circle (mX + mTabHeig / 2, ucSH - 1 - mYInc - mTabHeig / 2), mTabHeig / 2 - 1, controlHLightC, (200 / 180) * 3.14, 1.5 * 3.14
                    'right
                    UserControl.Line (mX + bWid - 1, ucSH - 1 - mYInc - 1)-(mX + bWid - 1, ucSH - 1 - mYInc - mTabHeig + 1), controlShadC
                End If
                '
                UserControl.Circle (mX + mTabHeig / 2, ucSH - 1 - mYInc - mTabHeig / 2), mTabHeig / 2, mBDRC, (90 / 180) * 3.14, 1.5 * 3.14
            Else
                UserControl.Circle (mX + bWid - mTabHeig / 2, ucSH - 1 - mYInc - mTabHeig / 2), mTabHeig / 2, mBC
                UserControl.FillStyle = 1
                UserControl.Circle (mX + bWid - mTabHeig / 2, ucSH - 1 - mYInc - mTabHeig / 2), mTabHeig / 2, mBDRC
                UserControl.FillStyle = 0
                UserControl.Line (mX, ucSH - 1 - mYInc)-(mX + bWid - mTabHeig / 2, ucSH - 1 - mTabHeig - mYInc), mBC, BF
                UserControl.Line (mX, ucSH - 1 - mYInc)-(mX + bWid - mTabHeig / 2, ucSH - 1 - mTabHeig - mYInc), mBDRC, B
                '
                UserControl.Line (mX + bWid - mTabHeig / 2, ucSH - 1 - mYInc - 1)-(mX + bWid - mTabHeig / 2, ucSH - 1 - mYInc - mTabHeig), mBC
                'shadow
                If mePTAppearance = ptAppearance3D Then
                    'center
                    UserControl.Line (mX + 1, ucSH - 1 - 1 - mYInc)-(mX + bWid - mTabHeig / 2 + 3, ucSH - 1 - 1 - mYInc), controlShadC
                    UserControl.Line (mX + 1, ucSH - 1 - mYInc - mTabHeig + 1)-(mX + bWid - mTabHeig / 2, ucSH - 1 - mYInc - mTabHeig + 1), controlHLightC
                    'right
                    UserControl.Circle (mX + bWid - mTabHeig / 2, ucSH - 1 - mYInc - mTabHeig / 2), mTabHeig / 2 - 1, controlShadC, (340 / 180) * 3.14, 3.14 / 2
                    UserControl.Circle (mX + bWid - mTabHeig / 2, ucSH - 1 - mYInc - mTabHeig / 2), mTabHeig / 2 - 1, controlShadC, (270 / 180) * 3.14, (340 / 180) * 3.14
                    'left
                    UserControl.Line (mX + 1, ucSH - 1 - 1 - mYInc)-(mX + 1, ucSH - 1 - mYInc - mTabHeig + 1), controlHLightC
                End If
                
                UserControl.Circle (mX + bWid - mTabHeig / 2, ucSH - 1 - mYInc - mTabHeig / 2), mTabHeig / 2, mBDRC, (270 / 180) * 3.14, 3.14 / 2
            End If
        Else
            UserControl.Line (mX, ucSH - 1 - mYInc)-(mX + bWid, ucSH - 1 - mTabHeig - mYInc), mBC, BF
            UserControl.Line (mX, ucSH - 1 - mYInc)-(mX + bWid, ucSH - 1 - mTabHeig - mYInc), mBDRC, B
            'draw shadow
            If mePTAppearance = ptAppearance3D Then
                'center
                UserControl.Line (mX + 1, ucSH - 1 - 1 - mYInc)-(mX + bWid, ucSH - 1 - 1 - mYInc), controlShadC
                UserControl.Line (mX + 1, ucSH - 1 - mYInc - mTabHeig + 1)-(mX + bWid, ucSH - 1 - mYInc - mTabHeig + 1), controlHLightC
                'left
                UserControl.Line (mX + 1, ucSH - 1 - 1 - mYInc)-(mX + 1, ucSH - 1 - mYInc - mTabHeig + 1), controlHLightC
                'right
                UserControl.Line (mX + bWid - 1, ucSH - 1 - mYInc - 2)-(mX + bWid - 1, ucSH - 1 - mYInc - mTabHeig + 1), controlShadC
            End If
        End If
        'find font y-position
        If mePTVertTextAlg = ptBottom Then
            UserControl.CurrentY = ucSH - mTabHeig + 2 - mYInc '+ UserControl.TextHeight(lstCaptions.Item(mInd + 1)) + 5
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = ucSH - mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 - mYInc
        Else
            UserControl.CurrentY = ucSH - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 - mYInc
        End If
    End If
    'draw picture - if exist
    putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig
    'find font x-position
    UserControl.CurrentX = getFontX(bWid, mInd, havePIC, mePTPicAlig)
    If mInd = selIndex And efectSel = True Then
        If mePTOrientation = ptOrientationTop Then
            UserControl.CurrentY = UserControl.CurrentY - 1
        Else
            UserControl.CurrentY = UserControl.CurrentY + 1
        End If
        UserControl.CurrentX = UserControl.CurrentX - 1
    End If
    'print caption
    UserControl.Print lstCaptions.Item(mInd + 1)

    'find position of new (next) tab - if exist
    mX = mX + bWid
End Sub
'/////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////////////
Private Sub drawDistortedMenuTab(ByVal mInd As Integer)
    Dim mTabHeig As Integer, mYInc As Integer
    Dim fntMovPC As Byte, pcMov As Byte
    Const dSize As Byte = 5
    '
    Dim havePIC As Boolean
    'setValues mInd, mBDRC, mBC, mShadC
    '
    'ucSW = UserControl.ScaleWidth
    'ucSH = UserControl.ScaleHeight
    'find tab width
    bWid = UserControl.TextWidth(lstCaptions.Item(mInd + 1)) + mSpacing * 2 + 10
    If mePTStylesActive <> ptCoolLeft And mePTStylesNormal <> ptCoolLeft Then mX = mX + getWidInc(mInd)
    'if there is picture then increase tab width
    havePIC = False
    If Not tabIcons(mInd) Is Nothing Then
        If mePTPicAlig <> ptPicCenter Then bWid = bWid + icWid + mSpacing + 5 '* 2
        havePIC = True
    End If
    'add width to collection
    lstSizes.Add bWid
    '
    If bWid + mX < 0 Or mX > ucSW Then
        mX = mX + bWid
        Exit Sub
    End If
    'find tab height and y position
    mTabHeig = tabHeig
    mYInc = Abs(tabHeigActive - tabHeig)
    If mInd = selIndex Then
        mTabHeig = tabHeigActive
        If tabHeigActive > tabHeig Then mYInc = 0
    Else
        If tabHeig > tabHeigActive Then mYInc = 0
    End If
    '
    'draw tab
    If mePTOrientation = ptOrientationTop Then
        'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
        If mInd = firstItemL - 1 Then
            UserControl.Line (mX, mYInc)-(mX, mTabHeig + mYInc), mBDRC
            UserControl.Line (mX, mYInc)-(mX + bWid - dSize, mYInc), mBDRC
            UserControl.Line (mX + bWid - dSize, mYInc)-(mX + bWid, mYInc + dSize), mBDRC
            UserControl.Line (mX + bWid, mYInc + dSize)-(mX + bWid, mYInc + mTabHeig - dSize), mBDRC
            UserControl.Line (mX + bWid, mYInc + mTabHeig - dSize)-(mX + bWid + dSize, mYInc + mTabHeig), mBDRC
            UserControl.Line (mX + bWid + dSize, mYInc + mTabHeig)-(mX, mYInc + mTabHeig), mBDRC
            'draw shadow
            If mePTAppearance = ptAppearance3D Then
                'center
                UserControl.Line (mX + 1, mYInc + 1)-(mX + bWid - dSize, mYInc + 1), controlHLightC
                UserControl.Line (mX + bWid + dSize, mYInc + mTabHeig - 1)-(mX + 1, mYInc + mTabHeig - 1), controlShadC
                'right
                UserControl.Line (mX + bWid - dSize - 1, mYInc)-(mX + bWid - 1, mYInc + dSize), controlShadC
                UserControl.Line (mX + bWid - 1, mYInc + dSize)-(mX + bWid - 1, mYInc + mTabHeig - dSize), controlShadC
                UserControl.Line (mX + bWid - 1, mYInc + mTabHeig - dSize)-(mX + bWid + dSize - 1, mYInc + mTabHeig), controlShadC
                    
                UserControl.Line (mX + bWid - dSize + 1, mYInc)-(mX + bWid + 1, mYInc + dSize), controlHLightC
                UserControl.Line (mX + bWid + 1, mYInc + dSize)-(mX + bWid + 1, mYInc + mTabHeig - dSize), controlHLightC
                UserControl.Line (mX + bWid + 1, mYInc + mTabHeig - dSize)-(mX + bWid + dSize + 1, mYInc + mTabHeig), controlHLightC
                'left
                UserControl.Line (mX + 1, mYInc + 1)-(mX + 1, mTabHeig + mYInc - 1), controlHLightC
            End If
        ElseIf mInd = tabCnt - 1 Then
            UserControl.Line (mX + bWid, mYInc)-(mX + bWid, mTabHeig + mYInc), mBDRC
            UserControl.Line (mX - dSize, mYInc)-(mX + bWid, mYInc), mBDRC

            UserControl.Line (mX - dSize, mYInc)-(mX, mYInc + dSize), mBDRC
            UserControl.Line (mX, mYInc + dSize)-(mX, mYInc + mTabHeig - dSize), mBDRC
            UserControl.Line (mX, mYInc + mTabHeig - dSize)-(mX + dSize, mYInc + mTabHeig), mBDRC
           ' UserControl.Line (mX + dSize, mYInc + mTabHeig)-(mX, mYInc + mTabHeig), mBDRC
            'draw shadow
            If mePTAppearance = ptAppearance3D Then
                'right
                UserControl.Line (mX + bWid - 1, mYInc)-(mX + bWid - 1, mTabHeig + mYInc), controlShadC
                'Left
                    UserControl.Line (mX - dSize - 1, mYInc)-(mX - 1, mYInc + dSize), controlShadC
                    UserControl.Line (mX - 1, mYInc + dSize)-(mX - 1, mYInc + mTabHeig - dSize), controlShadC
                    UserControl.Line (mX - 1, mYInc + mTabHeig - dSize)-(mX + dSize - 1, mYInc + mTabHeig), controlShadC

                    UserControl.Line (mX - dSize + 1, mYInc)-(mX + 1, mYInc + dSize), controlHLightC
                    UserControl.Line (mX + 1, mYInc + dSize)-(mX + 1, mYInc + mTabHeig - dSize), controlHLightC
                    UserControl.Line (mX + 1, mYInc + mTabHeig - dSize)-(mX + dSize + 1, mYInc + mTabHeig), controlHLightC
                'center
                    UserControl.Line (mX - dSize + 2, mYInc + 1)-(mX + bWid - 1, mYInc + 1), controlHLightC
                    UserControl.Line (mX + dSize, mYInc + mTabHeig - 1)-(mX + bWid, mYInc + mTabHeig - 1), controlShadC
            End If
           ' reDraw
            UserControl.Line (mX - dSize, mYInc)-(mX + bWid, mYInc), mBDRC
            UserControl.Line (mX + dSize, mTabHeig + mYInc)-(mX + bWid, mTabHeig + mYInc), mBDRC
        Else
            UserControl.Line (mX - dSize, mYInc)-(mX + bWid - dSize, mYInc), mBDRC
            'left
            UserControl.Line (mX - dSize, mYInc)-(mX, mYInc + dSize), mBDRC
            UserControl.Line (mX, mYInc + dSize)-(mX, mYInc + mTabHeig - dSize), mBDRC
            UserControl.Line (mX, mYInc + mTabHeig - dSize)-(mX + dSize, mYInc + mTabHeig), mBDRC
            'UserControl.Line (mX + dSize, mYInc + mTabHeig)-(mX, mYInc + mTabHeig), mBDRC
            'right
            UserControl.Line (mX + bWid - dSize, mYInc)-(mX + bWid, mYInc + dSize), mBDRC
            UserControl.Line (mX + bWid, mYInc + dSize)-(mX + bWid, mYInc + mTabHeig - dSize), mBDRC
            UserControl.Line (mX + bWid, mYInc + mTabHeig - dSize)-(mX + bWid + dSize, mYInc + mTabHeig), mBDRC
            UserControl.Line (mX + bWid + dSize, mYInc + mTabHeig)-(mX + dSize - 1, mYInc + mTabHeig), mBDRC
            'draw shadow
            If mePTAppearance = ptAppearance3D Then
                'left
                UserControl.Line (mX - dSize - 1, mYInc)-(mX - 1, mYInc + dSize), controlShadC
                UserControl.Line (mX - 1, mYInc + dSize)-(mX - 1, mYInc + mTabHeig - dSize), controlShadC
                UserControl.Line (mX - 1, mYInc + mTabHeig - dSize)-(mX + dSize - 1, mYInc + mTabHeig), controlShadC
                'UserControl.Line (mX + dSize - 1, mYInc + mTabHeig)-(mX - 1, mYInc + mTabHeig), controlShadC
                '
                UserControl.Line (mX - dSize + 1, mYInc)-(mX + 1, mYInc + dSize), controlHLightC
                UserControl.Line (mX + 1, mYInc + dSize)-(mX + 1, mYInc + mTabHeig - dSize), controlHLightC
                UserControl.Line (mX + 1, mYInc + mTabHeig - dSize)-(mX + dSize + 1, mYInc + mTabHeig), controlHLightC
                'UserControl.Line (mX + dSize + 1, mYInc + mTabHeig)-(mX + 1, mYInc + mTabHeig), controlHLightC
                'right
                UserControl.Line (mX + bWid - dSize + 1, mYInc)-(mX + bWid + 1, mYInc + dSize), controlHLightC
                UserControl.Line (mX + bWid + 1, mYInc + dSize)-(mX + bWid + 1, mYInc + mTabHeig - dSize), controlHLightC
                UserControl.Line (mX + bWid + 1, mYInc + mTabHeig - dSize)-(mX + bWid + dSize + 1, mYInc + mTabHeig), controlHLightC
                'UserControl.Line (mX + bWid + dSize + 1, mYInc + mTabHeig)-(mX + 1, mYInc + mTabHeig), controlHLightC
                '
                UserControl.Line (mX + bWid - dSize - 1, mYInc)-(mX + bWid - 1, mYInc + dSize), controlShadC
                UserControl.Line (mX + bWid - 1, mYInc + dSize)-(mX + bWid - 1, mYInc + mTabHeig - dSize), controlShadC
                UserControl.Line (mX + bWid - 1, mYInc + mTabHeig - dSize)-(mX + bWid + dSize - 1, mYInc + mTabHeig), controlShadC
                'UserControl.Line (mX + bWid + dSize - 1, mYInc + mTabHeig)-(mX - 1 + dSize, mYInc + mTabHeig), controlShadC
                'center
                UserControl.Line (mX - dSize + 2, mYInc + 1)-(mX + bWid - dSize + 1, mYInc + 1), controlHLightC
                UserControl.Line (mX + dSize, mTabHeig + mYInc - 1)-(mX + bWid + dSize, mTabHeig + mYInc - 1), controlShadC
            End If
            UserControl.Line (mX - dSize - 1, mYInc)-(mX + bWid - dSize, mYInc), mBDRC
            'UserControl.Line (mX + dSize, mTabHeig + mYInc)-(mX + bWid + dSize, mTabHeig + mYInc), mBDRC
        End If
        ExtFloodFill UserControl.hdc, mX + 2, mYInc + mTabHeig / 2, UserControl.Point(mX + 2, mYInc + mTabHeig / 2), 1
        'find font y-position
        If mePTVertTextAlg = ptTop Then
            UserControl.CurrentY = 2 + mYInc
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 + mYInc
        Else
            UserControl.CurrentY = mTabHeig - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 + mYInc
        End If
    Else
        'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
        If mInd = firstItemL - 1 Then
            UserControl.Line (mX, ucSH - 1 - mYInc)-(mX, ucSH - 1 - mTabHeig - mYInc), mBDRC
            UserControl.Line (mX, ucSH - 1 - mYInc)-(mX + bWid - dSize, ucSH - 1 - mYInc), mBDRC
            UserControl.Line (mX + bWid - dSize, ucSH - 1 - mYInc)-(mX + bWid, ucSH - 1 - mYInc - dSize), mBDRC
            UserControl.Line (mX + bWid, ucSH - 1 - mYInc - dSize)-(mX + bWid, ucSH - 1 - mYInc - mTabHeig + dSize), mBDRC
            UserControl.Line (mX + bWid, ucSH - 1 - mYInc - mTabHeig + dSize)-(mX + bWid + dSize, ucSH - 1 - mYInc - mTabHeig), mBDRC
            UserControl.Line (mX + bWid + dSize, ucSH - 1 - mYInc - mTabHeig)-(mX, ucSH - 1 - mYInc - mTabHeig), mBDRC
            'draw shadow
            If mePTAppearance = ptAppearance3D Then
                'center
                UserControl.Line (mX + 1, ucSH - 1 - mYInc - 1)-(mX + bWid - dSize, ucSH - 1 - mYInc - 1), controlShadC
                UserControl.Line (mX + bWid + dSize, ucSH - 1 - mYInc - mTabHeig + 1)-(mX + 1, ucSH - 1 - mYInc - mTabHeig + 1), controlHLightC
                'right
                UserControl.Line (mX + bWid - dSize - 1, ucSH - 1 - mYInc)-(mX + bWid - 1, ucSH - 1 - mYInc - dSize), controlShadC
                UserControl.Line (mX + bWid - 1, ucSH - 1 - mYInc - dSize)-(mX + bWid - 1, ucSH - 1 - mYInc - mTabHeig + dSize), controlShadC
                UserControl.Line (mX + bWid - 1, ucSH - 1 - mYInc - mTabHeig + dSize)-(mX + bWid + dSize - 1, ucSH - 1 - mYInc - mTabHeig), controlShadC
                
                UserControl.Line (mX + bWid - dSize + 1, ucSH - 1 - mYInc)-(mX + bWid + 1, ucSH - 1 - mYInc - dSize), controlHLightC
                UserControl.Line (mX + bWid + 1, ucSH - 1 - mYInc - dSize)-(mX + bWid + 1, ucSH - 1 - mYInc - mTabHeig + dSize), controlHLightC
                UserControl.Line (mX + bWid + 1, ucSH - 1 - mYInc - mTabHeig + dSize)-(mX + bWid + dSize + 1, ucSH - 1 - mYInc - mTabHeig), controlHLightC
                'left
                UserControl.Line (mX + 1, ucSH - 1 - mYInc - 1)-(mX + 1, ucSH - 1 - mTabHeig - mYInc + 1), controlHLightC
            End If
        ElseIf mInd = tabCnt - 1 Then
            UserControl.Line (mX + bWid, ucSH - 1 - mYInc)-(mX + bWid, ucSH - 1 - mTabHeig - mYInc), mBDRC
            UserControl.Line (mX - dSize, ucSH - 1 - mYInc)-(mX + bWid, ucSH - 1 - mYInc), mBDRC

            UserControl.Line (mX - dSize, ucSH - 1 - mYInc)-(mX, ucSH - 1 - mYInc - dSize), mBDRC
            UserControl.Line (mX, ucSH - 1 - mYInc - dSize)-(mX, ucSH - 1 - mYInc - mTabHeig + dSize), mBDRC
            UserControl.Line (mX, ucSH - 1 - mYInc - mTabHeig + dSize)-(mX + dSize, ucSH - 1 - mYInc - mTabHeig), mBDRC
            'UserControl.Line (mX + dSize, ucSH - 1 - mYInc - mTabHeig)-(mX, ucSH - 1 - mYInc - mTabHeig), mBDRC
            'draw shadow
            If mePTAppearance = ptAppearance3D Then
                'right
                UserControl.Line (mX + bWid - 1, ucSH - 1 - mYInc)-(mX + bWid - 1, ucSH - 1 - mTabHeig - mYInc), controlShadC
                'left
                UserControl.Line (mX - dSize - 1, ucSH - 1 - mYInc)-(mX - 1, ucSH - 1 - mYInc - dSize), controlShadC
                UserControl.Line (mX - 1, ucSH - 1 - mYInc - dSize)-(mX - 1, ucSH - 1 - mYInc - mTabHeig + dSize), controlShadC
                UserControl.Line (mX - 1, ucSH - 1 - mYInc - mTabHeig + dSize)-(mX + dSize - 1, ucSH - 1 - mYInc - mTabHeig), controlShadC

                UserControl.Line (mX - dSize + 1, ucSH - 1 - mYInc)-(mX + 1, ucSH - 1 - mYInc - dSize), controlHLightC
                UserControl.Line (mX + 1, ucSH - 1 - mYInc - dSize)-(mX + 1, ucSH - 1 - mYInc - mTabHeig + dSize), controlHLightC
                UserControl.Line (mX + 1, ucSH - 1 - mYInc - mTabHeig + dSize)-(mX + dSize + 1, ucSH - 1 - mYInc - mTabHeig), controlHLightC
                'center
                UserControl.Line (mX - dSize + 2, ucSH - 1 - mYInc - 1)-(mX + bWid - 1, ucSH - 1 - mYInc - 1), controlShadC
                UserControl.Line (mX + dSize, ucSH - 1 - mYInc - mTabHeig + 1)-(mX + bWid, ucSH - 1 - mYInc - mTabHeig + 1), controlHLightC
            End If
           ' reDraw
            UserControl.Line (mX - dSize, ucSH - 1 - mYInc)-(mX + bWid, ucSH - 1 - mYInc), mBDRC
            UserControl.Line (mX + dSize, ucSH - 1 - mTabHeig - mYInc)-(mX + bWid, ucSH - 1 - mTabHeig - mYInc), mBDRC
        Else
            UserControl.Line (mX - dSize, ucSH - 1 - mYInc)-(mX + bWid - dSize, ucSH - 1 - mYInc), mBDRC
            'left
            UserControl.Line (mX - dSize, ucSH - 1 - mYInc)-(mX, ucSH - 1 - mYInc - dSize), mBDRC
            UserControl.Line (mX, ucSH - 1 - mYInc - dSize)-(mX, ucSH - 1 - mYInc - mTabHeig + dSize), mBDRC
            UserControl.Line (mX, ucSH - 1 - mYInc - mTabHeig + dSize)-(mX + dSize, ucSH - 1 - mYInc - mTabHeig), mBDRC
            'UserControl.Line (mX + dSize, ucSH - 1 - mYInc - mTabHeig)-(mX, ucSH - 1 - mYInc - mTabHeig), mBDRC
            'right
            UserControl.Line (mX + bWid - dSize, ucSH - 1 - mYInc)-(mX + bWid, ucSH - 1 - mYInc - dSize), mBDRC
            UserControl.Line (mX + bWid, ucSH - 1 - mYInc - dSize)-(mX + bWid, ucSH - 1 - mYInc - mTabHeig + dSize), mBDRC
            UserControl.Line (mX + bWid, ucSH - 1 - mYInc - mTabHeig + dSize)-(mX + bWid + dSize, ucSH - 1 - mYInc - mTabHeig), mBDRC
            UserControl.Line (mX + bWid + dSize, ucSH - 1 - mYInc - mTabHeig)-(mX + dSize - 1, ucSH - 1 - mYInc - mTabHeig), mBDRC
            'draw shadow
            If mePTAppearance = ptAppearance3D Then
                'left
                UserControl.Line (mX - dSize - 1, ucSH - 1 - mYInc)-(mX - 1, ucSH - 1 - mYInc - dSize), controlShadC
                UserControl.Line (mX - 1, ucSH - 1 - mYInc - dSize)-(mX - 1, ucSH - 1 - mYInc - mTabHeig + dSize), controlShadC
                UserControl.Line (mX - 1, ucSH - 1 - mYInc - mTabHeig + dSize)-(mX + dSize - 1, ucSH - 1 - mYInc - mTabHeig), controlShadC
                'UserControl.Line (mX + dSize - 1, ucSH - 1 - mYInc - mTabHeig)-(mX - 1, ucSH - 1 - mYInc - mTabHeig), controlShadC
                '
                UserControl.Line (mX - dSize + 1, ucSH - 1 - mYInc)-(mX + 1, ucSH - 1 - mYInc - dSize), controlHLightC
                UserControl.Line (mX + 1, ucSH - 1 - mYInc - dSize)-(mX + 1, ucSH - 1 - mYInc - mTabHeig + dSize), controlHLightC
                UserControl.Line (mX + 1, ucSH - 1 - mYInc - mTabHeig + dSize)-(mX + dSize + 1, ucSH - 1 - mYInc - mTabHeig), controlHLightC
                    'UserControl.Line (mX + dSize + 1, ucSH - 1 - mYInc - mTabHeig)-(mX + 1, ucSH - 1 - mYInc - mTabHeig), controlHLightC
                'right
                UserControl.Line (mX + bWid - dSize + 1, ucSH - 1 - mYInc)-(mX + bWid + 1, ucSH - 1 - mYInc - dSize), controlHLightC
                UserControl.Line (mX + bWid + 1, ucSH - 1 - mYInc - dSize)-(mX + bWid + 1, ucSH - 1 - mYInc - mTabHeig + dSize), controlHLightC
                UserControl.Line (mX + bWid + 1, ucSH - 1 - mYInc - mTabHeig + dSize)-(mX + bWid + dSize + 1, ucSH - 1 - mYInc - mTabHeig), controlHLightC
                'UserControl.Line (mX + bWid + dSize + 1, ucSH - 1 - mYInc * mTabHeig)-(mX + 1, ucSH - 1 - mYInc - mTabHeig), controlHLightC
                '
                UserControl.Line (mX + bWid - dSize - 1, ucSH - 1 - mYInc)-(mX + bWid - 1, ucSH - 1 - mYInc - dSize), controlShadC
                UserControl.Line (mX + bWid - 1, ucSH - 1 - mYInc - dSize)-(mX + bWid - 1, ucSH - 1 - mYInc - mTabHeig + dSize), controlShadC
                UserControl.Line (mX + bWid - 1, ucSH - 1 - mYInc - mTabHeig + dSize)-(mX + bWid + dSize - 1, ucSH - 1 - mYInc - mTabHeig), controlShadC
                    'UserControl.Line (mX + bWid + dSize - 1, ucSH - 1 - mYInc - mTabHeig)-(mX - 1, ucSH - 1 - mYInc - mTabHeig), controlShadC
                'center
                UserControl.Line (mX - dSize + 2, ucSH - 1 - mYInc - 1)-(mX + bWid - dSize, ucSH - 1 - mYInc - 1), controlShadC
                UserControl.Line (mX + dSize, ucSH - 1 - mTabHeig - mYInc + 1)-(mX + bWid + dSize - 2, ucSH - 1 - mTabHeig - mYInc + 1), controlHLightC

            End If
            UserControl.Line (mX - dSize - 1, ucSH - 1 - mYInc)-(mX + bWid - dSize + 2, ucSH - 1 - mYInc), mBDRC
            'UserControl.Line (mX + dSize, ucSH - 1 - mYInc)-(mX + bWid + dSize + 2, ucSH - 1 - mYInc), mBDRC
        End If
        ExtFloodFill UserControl.hdc, mX + 2, ucSH - 1 - mYInc - mTabHeig / 2, UserControl.Point(mX + 2, ucSH - 1 - mYInc - mTabHeig / 2), 1
        'find font y-position
        If mePTVertTextAlg = ptBottom Then
            UserControl.CurrentY = ucSH - mTabHeig + 2 - mYInc '+ UserControl.TextHeight(lstCaptions.Item(mInd + 1)) + 5
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = ucSH - mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 - mYInc
        Else
            UserControl.CurrentY = ucSH - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 - mYInc
        End If
    End If
    'draw picture - if exist
    putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig
    'find font x-position
    UserControl.CurrentX = getFontX(bWid, mInd, havePIC, mePTPicAlig)
    If mInd = selIndex And efectSel = True Then
        If mePTOrientation = ptOrientationTop Then
            UserControl.CurrentY = UserControl.CurrentY - 1
        Else
            UserControl.CurrentY = UserControl.CurrentY + 1
        End If
        UserControl.CurrentX = UserControl.CurrentX - 1
    End If
    'print caption
    UserControl.Print lstCaptions.Item(mInd + 1)
    'if this is active tab then hide line at bottom of tab
   ' hideBottom mInd, mX, mYInc, bWid, mTabHeig, mBC, controlHLightC, controlShadC
'    If mInd = selIndex Then
'        If mePTOrientation = ptOrientationTop Then
'            UserControl.Line (mX + 1, mTabHeig + mYInc)-(mX + bWid - 1, mTabHeig + mYInc + 1), mBC, BF
'
'        Else
'            UserControl.Line (mX + 1, ucSH - mTabHeig - 1 - mYInc)-(mX + bWid - 1, ucSH - mTabHeig - 2 - mYInc), mBC, BF
'        End If
'    End If
    'find position of new (next) tab - if exist
    mX = mX + bWid
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub drawRoundedTab(ByVal mInd As Integer)
    Dim mTabHeig As Integer, mYInc As Integer
    Dim fntMovPC As Byte, pcMov As Byte
    Dim cSize As Byte
    '
    Dim havePIC As Boolean
    'setValues mInd, mBDRC, mBC, mShadC
    'ucSW = UserControl.ScaleWidth
    'ucSH = UserControl.ScaleHeight
    'find tab width
    bWid = UserControl.TextWidth(lstCaptions.Item(mInd + 1)) + mSpacing * 2 + 10
    If mePTStylesActive <> ptCoolLeft And mePTStylesNormal <> ptCoolLeft Then mX = mX + getWidInc(mInd)
    'if there is picture then increase tab width
    havePIC = False
    If Not tabIcons(mInd) Is Nothing Then
        If mePTPicAlig <> ptPicCenter Then bWid = bWid + icWid + mSpacing + 5 '* 2
        havePIC = True
    End If
    'add width to collection
    lstSizes.Add bWid
    '
    If bWid + mX < 0 Or mX > ucSW Then
        mX = mX + bWid
        Exit Sub
    End If
    'find tab height and y position
    mTabHeig = tabHeig
    mYInc = Abs(tabHeigActive - tabHeig)
    If mInd = selIndex Then
        mTabHeig = tabHeigActive
        If tabHeigActive > tabHeig Then mYInc = 0
    Else
        If tabHeig > tabHeigActive Then mYInc = 0
    End If
    'find corner size
    cSize = (mTabHeig / 10) + 2
    'draw tab
    'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
    If mePTOrientation = ptOrientationTop Then
        UserControl.Line (mX, mYInc + mTabHeig - 1)-(mX, mYInc + cSize - 1), mBDRC
        UserControl.Circle (mX + cSize, mYInc + cSize), cSize, mBDRC, 3.14 / 2, 3.14
        UserControl.Line (mX + cSize, mYInc)-(mX + bWid - cSize + 1, mYInc), mBDRC
        UserControl.Circle (mX - cSize + bWid, mYInc + cSize), cSize, mBDRC, 0, 3.14 / 2
        UserControl.Line (mX + bWid, mYInc + cSize)-(mX + bWid, mYInc + mTabHeig), mBDRC
        'fill tab
        If mX < 0 And mX + bWid > 0 Then
            ExtFloodFill UserControl.hdc, mX + bWid - 2, mTabHeig - 2 + mYInc, UserControl.Point(mX + bWid - 2, mTabHeig - 2 + mYInc), 1
        ElseIf mX < ucSW Then
            ExtFloodFill UserControl.hdc, mX + 2, mTabHeig - 2 + mYInc, UserControl.Point(mX + 2, mTabHeig - 2 + mYInc), 1
        End If
        'draw shadow
        If mePTAppearance = ptAppearance3D Then
            'left
            UserControl.Line (mX + 1, mYInc + mTabHeig - 1)-(mX + 1, mYInc + cSize - 1), controlHLightC
            UserControl.Circle (mX + cSize, mYInc + cSize), cSize - 1, controlHLightC, 3.14 / 2, 3.14
            'center
            UserControl.Line (mX + cSize, mYInc + 1)-(mX + bWid - cSize + 1, mYInc + 1), controlHLightC
            'right
            UserControl.Circle (mX - cSize + bWid, mYInc + cSize), cSize - 1, controlShadC, 0, 3.14 / 2
            UserControl.Line (mX + bWid - 1, mYInc + cSize)-(mX + bWid - 1, mYInc + mTabHeig), controlShadC
        End If
        'font y
        If mePTVertTextAlg = ptTop Then
            UserControl.CurrentY = 2 + mYInc
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 + mYInc
        Else
            UserControl.CurrentY = mTabHeig - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 + mYInc
        End If
    Else
        UserControl.Line (mX, ucSH - 1 - mYInc - mTabHeig)-(mX, ucSH - 1 - mYInc - cSize + 1), mBDRC
        UserControl.Circle (mX + cSize, ucSH - 1 - mYInc - cSize), cSize, mBDRC, 3.14, 3.14 * 1.5
        UserControl.Line (mX + cSize, ucSH - 1 - mYInc)-(mX + bWid - cSize + 1, ucSH - 1 - mYInc), mBDRC
        UserControl.Circle (mX - cSize + bWid, ucSH - 1 - mYInc - cSize), cSize, mBDRC, 3.14 * 1.5, 0
        UserControl.Line (mX + bWid, ucSH - 1 - mYInc - cSize)-(mX + bWid, ucSH - 1 - mYInc - mTabHeig), mBDRC
        'fill tab
        If mX < 0 And mX + bWid > 0 Then
            ExtFloodFill UserControl.hdc, mX + bWid - 2, ucSH - 1 - mTabHeig + 2 - mYInc, UserControl.Point(mX + bWid - 2, ucSH - 1 - mTabHeig + 2 - mYInc), 1
        ElseIf mX < ucSW Then
            ExtFloodFill UserControl.hdc, mX + 2, ucSH - 1 - mTabHeig + 2 - mYInc, UserControl.Point(mX + 2, ucSH - 1 - mTabHeig + 2 - mYInc), 1
        End If
        'draw shadow
        If mePTAppearance = ptAppearance3D Then
            'left
            UserControl.Line (mX + 1, ucSH - 1 - mYInc - mTabHeig)-(mX + 1, ucSH - 1 - mYInc - cSize + 1), controlHLightC
            UserControl.Circle (mX + cSize, ucSH - 1 - mYInc - cSize), cSize - 1, controlHLightC, 3.14, 3.14 * 1.5
            'center
            UserControl.Line (mX + cSize, ucSH - 1 - mYInc - 1)-(mX + bWid - cSize + 1, ucSH - 1 - mYInc - 1), controlShadC
            'right
            UserControl.Circle (mX - cSize + bWid, ucSH - 1 - mYInc - cSize), cSize - 1, controlShadC, 3.14 * 1.5, 0
            UserControl.Line (mX + bWid - 1, ucSH - 1 - mYInc - cSize)-(mX + bWid - 1, ucSH - 1 - mYInc - mTabHeig), controlShadC
        End If
        'find font y-position
        If mePTVertTextAlg = ptBottom Then
            UserControl.CurrentY = ucSH - mTabHeig + 2 - mYInc '+ UserControl.TextHeight(lstCaptions.Item(mInd + 1)) + 5
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = ucSH - mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 - mYInc
        Else
            UserControl.CurrentY = ucSH - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 - mYInc
        End If
    End If
    'draw picture - if exist
    putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig
    'find font x-position
    UserControl.CurrentX = getFontX(bWid, mInd, havePIC, mePTPicAlig)
    If mInd = selIndex And efectSel = True Then
        If mePTOrientation = ptOrientationTop Then
            UserControl.CurrentY = UserControl.CurrentY - 1
        Else
            UserControl.CurrentY = UserControl.CurrentY + 1
        End If
        UserControl.CurrentX = UserControl.CurrentX - 1
    End If
    'print caption
    UserControl.Print lstCaptions.Item(mInd + 1)
    'if this is active tab then hide line at bottom of tab
    hideBottom mInd, mX, mYInc, bWid, mTabHeig, mBC, controlHLightC, controlShadC
    'find position of new (next) tab - if exist
    mX = mX + bWid
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub drawRoundedXPTab(ByVal mInd As Integer)
    Dim mTabHeig As Integer, mYInc As Integer, i As Integer
    Dim fntMovPC As Byte, pcMov As Byte
    Dim drawAll As Boolean, havePIC As Boolean
    If mInd = firstItemL - 1 Then mX = mX - 3
    
    'setValues mInd, mBDRC, mBC, mShadC
    mShadC = shadC
    'ucSW = UserControl.ScaleWidth
    'ucSH = UserControl.ScaleHeight
    'find tab width
    bWid = UserControl.TextWidth(lstCaptions.Item(mInd + 1)) + mSpacing * 2 + 10 + getWidInc(mInd)
    'if there is picture then increase tab width
    havePIC = False
    If Not tabIcons(mInd) Is Nothing Then
        If mePTPicAlig <> ptPicCenter Then bWid = bWid + icWid + mSpacing + 5 '* 2
        havePIC = True
    End If
    If mePTStylesActive <> mePTStylesNormal Then drawAll = True
    
    If mInd = selIndex Or mInd = selIndex + 1 Then
        If mInd = selIndex Then bWid = bWid + buttSpacing * 2
        mX = mX - buttSpacing '- 3 '- mSpacing - 6
    End If
    'add width to collection
    lstSizes.Add bWid
    '
    If bWid + mX - 1 < 0 Or mX > ucSW Then
        mX = mX + bWid
        Exit Sub
    End If
    'find tab height and y position
    mTabHeig = tabHeig
    mYInc = Abs(tabHeigActive - tabHeig)
    If mInd = selIndex Then
        mTabHeig = tabHeigActive
        If tabHeigActive > tabHeig Then mYInc = 0
    Else
        If tabHeig > tabHeigActive Then mYInc = 0
    End If
    'draw tab
    'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
    If mePTOrientation = ptOrientationTop Then
        If mInd <> selIndex + 1 Or drawAll = True Or mInd = firstItemL - 1 Then
            UserControl.Line (mX, mYInc + mTabHeig)-(mX, mYInc + cSizeXP - 1), mBDRC
            UserControl.Circle (mX + cSizeXP, mYInc + cSizeXP), cSizeXP, mBDRC, 3.14 / 2, 3.14
            UserControl.Line (mX + cSizeXP, mYInc)-(mX + bWid - cSizeXP + 1, mYInc), mBDRC
        Else
            UserControl.Line (mX + 1, mYInc)-(mX + bWid - cSizeXP + 1, mYInc), mBDRC
        End If
        UserControl.Circle (mX - cSizeXP + bWid, mYInc + cSizeXP), cSizeXP, mBDRC, 0, 3.14 / 2
        UserControl.Line (mX + bWid, mYInc + cSizeXP)-(mX + bWid, mYInc + mTabHeig), mBDRC
        'fill tab
        If mX < 0 And mX + bWid > 0 Then
            ExtFloodFill UserControl.hdc, mX + bWid - 2, mTabHeig - 2 + mYInc, UserControl.Point(mX + bWid - 2, mTabHeig - 2 + mYInc), 1
        ElseIf mX < ucSW Then
            If mInd <> selIndex + 1 Or drawAll = True Then
                ExtFloodFill UserControl.hdc, mX + 2, mTabHeig - 2 + mYInc, UserControl.Point(mX + 2, mTabHeig - 2 + mYInc), 1
            Else
                ExtFloodFill UserControl.hdc, mX + 4, mTabHeig - 2 + mYInc, UserControl.Point(mX + 4, mTabHeig - 2 + mYInc), 1
            End If
        End If
        'draw shadow
        If mInd = selIndex Or mInd = hoverIndex Then
            For i = 0 To cSizeXP - 1
                If mInd <> selIndex + 1 Or drawAll = True Or mInd = firstItemL - 1 Then
                    If i = 0 Then
                        UserControl.Line (mX + cSizeXP - i - 1, mYInc + i)-(mX + bWid - cSizeXP + i + 2, mYInc + i), mShadC
                    Else
                        UserControl.Line (mX + cSizeXP - i - 1, mYInc + i)-(mX + bWid - cSizeXP + i + 2, mYInc + i), getColorMix(mShadC, RGB(255, 255, 0), 2)
                    End If
                Else
                    If i = 0 Then
                        UserControl.Line (mX + 1, mYInc + i)-(mX + bWid - cSizeXP + i + 2, mYInc + i), mShadC
                    Else
                        UserControl.Line (mX + 1, mYInc + i)-(mX + bWid - cSizeXP + i + 2, mYInc + i), getColorMix(mShadC, RGB(255, 255, 0), 2)
                    End If
                End If
            Next i
            If mShadC <> mBC Then
                If mInd <> selIndex + 1 Or drawAll = True Or mInd = firstItemL - 1 Then UserControl.Circle (mX + cSizeXP, mYInc + cSizeXP), cSizeXP, mShadC, 3.14 / 2, 3.14
                UserControl.Circle (mX - cSizeXP + bWid, mYInc + cSizeXP), cSizeXP, mShadC, 0, 3.14 / 2
            End If
        End If
        'font y
        If mePTVertTextAlg = ptTop Then
            UserControl.CurrentY = 2 + mYInc
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 + mYInc
        Else
            UserControl.CurrentY = mTabHeig - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 + mYInc
        End If
    Else
        If mInd <> selIndex + 1 Or drawAll = True Or mInd = firstItemL - 1 Then
            UserControl.Line (mX, ucSH - 1 - mYInc - mTabHeig)-(mX, ucSH - 1 - mYInc - cSizeXP + 1), mBDRC
            UserControl.Circle (mX + cSizeXP, ucSH - 1 - mYInc - cSizeXP), cSizeXP, mBDRC, 3.14, 1.5 * 3.14
            UserControl.Line (mX + cSizeXP, ucSH - 1 - mYInc)-(mX + bWid - cSizeXP + 1, ucSH - 1 - mYInc), mBDRC
        Else
            UserControl.Line (mX + 1, ucSH - 1 - mYInc)-(mX + bWid - cSizeXP + 1, ucSH - 1 - mYInc), mBDRC
        End If
        UserControl.Circle (mX - cSizeXP + bWid, ucSH - 1 - mYInc - cSizeXP), cSizeXP, mBDRC, 1.5 * 3.14, 0
        UserControl.Line (mX + bWid, ucSH - 1 - mYInc - cSizeXP)-(mX + bWid, ucSH - 1 - mYInc - mTabHeig), mBDRC
        'fill tab
        If mX < 0 And mX + bWid > 0 Then
            ExtFloodFill UserControl.hdc, mX + bWid - 2, ucSH - 1 - mTabHeig + 2 - mYInc, UserControl.Point(mX + bWid - 2, ucSH - 1 - mTabHeig + 2 - mYInc), 1
        ElseIf mX < ucSW Then
            If mInd <> selIndex + 1 Or drawAll = True Or mInd = firstItemL - 1 Then
                ExtFloodFill UserControl.hdc, mX + 2, ucSH - 1 - mTabHeig + 2 - mYInc, UserControl.Point(mX + 2, ucSH - 1 - mTabHeig + 2 - mYInc), 1
            Else
                ExtFloodFill UserControl.hdc, mX + 4, ucSH - 1 - mTabHeig + 2 - mYInc, UserControl.Point(mX + 4, ucSH - 1 - mTabHeig + 2 - mYInc), 1
            End If
        End If
        'draw shadow
        If mInd = selIndex Or mInd = hoverIndex Then
            For i = 0 To cSizeXP - 1
                If mInd <> selIndex + 1 Or drawAll = True Or mInd = firstItemL - 1 Then
                    If i = 0 Then
                        UserControl.Line (mX + cSizeXP - i - 1, ucSH - 1 - mYInc - i)-(mX + bWid - cSizeXP + i + 2, ucSH - 1 - mYInc - i), mShadC
                    Else
                        'MsgBox mBC
                        UserControl.Line (mX + cSizeXP - i - 1, ucSH - 1 - mYInc - i)-(mX + bWid - cSizeXP + i + 2, ucSH - 1 - mYInc - i), getColorMix(mShadC, RGB(255, 255, 0), 2)
                    End If
                Else
                    If i = 0 Then
                        UserControl.Line (mX + 1, ucSH - 1 - mYInc - i)-(mX + bWid - cSizeXP + i + 2, ucSH - 1 - mYInc - i), mShadC
                    Else
                        'MsgBox mBC
                        UserControl.Line (mX + 1, ucSH - 1 - mYInc - i)-(mX + bWid - cSizeXP + i + 2, ucSH - 1 - mYInc - i), getColorMix(mShadC, RGB(255, 255, 0), 2)
                    End If
                End If
            Next i
        
            If mShadC <> mBC Then
                If mInd <> selIndex + 1 Or drawAll = True Then UserControl.Circle (mX + cSizeXP, ucSH - 1 - mYInc - cSizeXP), cSizeXP, mShadC, 3.14, 1.5 * 3.14
                UserControl.Circle (mX - cSizeXP + bWid, ucSH - 1 - mYInc - cSizeXP), cSizeXP, mShadC, 1.5 * 3.14, 0
            End If
        End If
        'find font y-position
        If mePTVertTextAlg = ptBottom Then
            UserControl.CurrentY = ucSH - mTabHeig + 2 - mYInc '+ UserControl.TextHeight(lstCaptions.Item(mInd + 1)) + 5
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = ucSH - mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 - mYInc
        Else
            UserControl.CurrentY = ucSH - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 - mYInc
        End If
    End If
    'draw picture - if exist
    putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig
    'find font x-position
    UserControl.CurrentX = getFontX(bWid, mInd, havePIC, mePTPicAlig)
    If mInd = selIndex And efectSel = True Then
        If mePTOrientation = ptOrientationTop Then
            UserControl.CurrentY = UserControl.CurrentY - 1
        Else
            UserControl.CurrentY = UserControl.CurrentY + 1
        End If
        UserControl.CurrentX = UserControl.CurrentX - 1
    End If
    'print caption
    UserControl.Print lstCaptions.Item(mInd + 1)
    'if this is active tab then hide line at bottom of tab
    If mInd = selIndex Then
        If mePTOrientation = ptOrientationTop Then
            UserControl.Line (mX + 1, mTabHeig + mYInc)-(mX + bWid, mTabHeig + mYInc + 1), mBC, BF
        Else
            UserControl.Line (mX + 1, ucSH - mTabHeig - 1 - mYInc)-(mX + bWid, ucSH - mTabHeig - 2 - mYInc), mBC, BF
        End If
    End If
    'find position of new (next) tab - if exist
    mX = mX + bWid '+ 1
    'If drawAll = False Then mX = mX - 3 - buttSpacing ' * 2
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub drawCornerCutLeft(ByVal mInd As Integer)
    Dim cutWid As Integer, cutHeig As Integer, mTabHeig As Integer, mYInc As Integer
    Dim fntMovPC As Byte, pcMov As Byte
    Dim havePIC As Boolean
    'setValues mInd, mBDRC, mBC, mShadC
    'set control properties
    'ucSW = UserControl.ScaleWidth
    'ucSH = UserControl.ScaleHeight
    'find tab height
    mTabHeig = tabHeig
    mYInc = Abs(tabHeigActive - tabHeig)
    If mInd = selIndex Then
        mTabHeig = tabHeigActive
        If tabHeigActive > tabHeig Then mYInc = 0
    Else
        If tabHeig > tabHeigActive Then mYInc = 0
    End If
    '
    cutWid = mTabHeig / 3 * Cos(60 * 3.14 / 180)
    cutHeig = mTabHeig / 3
    'find tab width
    bWid = UserControl.TextWidth(lstCaptions.Item(mInd + 1)) + mSpacing * 2 + 10
    If mePTStylesActive <> ptCoolLeft And mePTStylesNormal <> ptCoolLeft Then mX = mX + getWidInc(mInd)
    'if there is picture then increase tab wid
    havePIC = False
    If Not tabIcons(mInd) Is Nothing Then
        If mePTPicAlig <> ptPicCenter Then bWid = bWid + icWid + mSpacing '* 2
        havePIC = True
    End If
    lstSizes.Add bWid
    '
    If bWid + mX < 0 Or mX > ucSW Then
        mX = mX + bWid
        Exit Sub
    End If
    'draw tab
    If mePTOrientation = ptOrientationTop Then
        UserControl.Line (mX, cutHeig + mYInc)-(mX, mTabHeig + mYInc), mBDRC
        UserControl.Line (mX, cutHeig + mYInc)-(mX + cutWid, mYInc), mBDRC
        UserControl.Line (mX + cutWid, mYInc)-(mX + bWid, mYInc), mBDRC
        UserControl.Line (mX + bWid, mYInc)-(mX + bWid, mTabHeig + mYInc), mBDRC
        UserControl.Line (mX, mTabHeig + mYInc)-(mX + bWid, mTabHeig + mYInc), mBDRC
        'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
        'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
        If mX < 0 And mX + bWid > 0 Then
            ExtFloodFill UserControl.hdc, mX + bWid - 2, mTabHeig - 2 + mYInc, UserControl.Point(mX + bWid - 2, mTabHeig - 2 + mYInc), 1
        ElseIf mX < ucSW Then
            ExtFloodFill UserControl.hdc, mX + 2, mTabHeig - 2 + mYInc, UserControl.Point(mX + 2, mTabHeig - 2 + mYInc), 1
        End If
        'draw shadow
        If mePTAppearance = ptAppearance3D Then
            'center
            UserControl.Line (mX + cutWid + 1, 1 + mYInc)-(mX + bWid, 1 + mYInc), controlHLightC
            'right
            UserControl.Line (mX + bWid - 1, 1 + mYInc)-(mX + bWid - 1, mTabHeig + mYInc), controlShadC
            'left
            UserControl.Line (mX + 1, cutHeig + mYInc)-(mX + 1, mTabHeig + mYInc), controlHLightC
            UserControl.Line (mX + 1, cutHeig + mYInc)-(mX + cutWid + 1, mYInc), controlHLightC
        End If
        '
        If mePTVertTextAlg = ptTop Then
            UserControl.CurrentY = 2 + mYInc 'UserControl.TextHeight(lstCaptions.Item(mInd + 1)) + 5
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 + mYInc
        Else
            UserControl.CurrentY = mTabHeig - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 + mYInc
        End If
    Else
        UserControl.Line (mX, ucSH - 1 - cutHeig - mYInc)-(mX, ucSH - 1 - mTabHeig - mYInc), mBDRC
        UserControl.Line (mX, ucSH - 1 - cutHeig - mYInc)-(mX + cutWid, ucSH - 1 - mYInc), mBDRC
        UserControl.Line (mX + cutWid, ucSH - 1 - mYInc)-(mX + bWid, ucSH - 1 - mYInc), mBDRC
        UserControl.Line (mX + bWid, ucSH - 1 - mYInc)-(mX + bWid, ucSH - 1 - mTabHeig - mYInc), mBDRC
        UserControl.Line (mX, ucSH - 1 - mTabHeig - mYInc)-(mX + bWid, ucSH - 1 - mTabHeig - mYInc), mBDRC
        'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
        'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
        If mX < 0 And mX + bWid > 0 Then
            ExtFloodFill UserControl.hdc, mX + bWid - 2, ucSH - 1 - mTabHeig + 2 - mYInc, UserControl.Point(mX + bWid - 2, ucSH - 1 - mTabHeig + 2 - mYInc), 1
        ElseIf mX < ucSW Then
            ExtFloodFill UserControl.hdc, mX + 2, ucSH - 1 - mTabHeig + 2 - mYInc, UserControl.Point(mX + 2, ucSH - 1 - mTabHeig + 2 - mYInc), 1
        End If
        'draw shadow
        If mePTAppearance = ptAppearance3D Then
            'center
            UserControl.Line (mX + cutWid + 1, ucSH - 1 - 1 - mYInc)-(mX + bWid, ucSH - 1 - 1 - mYInc), controlShadC
            'right
            UserControl.Line (mX + bWid - 1, ucSH - 1 - 1 - mYInc)-(mX + bWid - 1, ucSH - 1 - mTabHeig - mYInc), controlShadC
            'left
            UserControl.Line (mX + 1, ucSH - 1 - cutHeig - mYInc)-(mX + 1, ucSH - 1 - mTabHeig - mYInc), controlHLightC
            UserControl.Line (mX + 1, ucSH - 1 - cutHeig - mYInc)-(mX + cutWid + 1, ucSH - 1 - mYInc), controlHLightC
        End If
        '
        If mePTVertTextAlg = ptBottom Then
            UserControl.CurrentY = ucSH - mTabHeig + 2 - mYInc '+ UserControl.TextHeight(lstCaptions.Item(mInd + 1)) + 5
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = ucSH - mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 - mYInc
        Else
            UserControl.CurrentY = ucSH - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 - mYInc
        End If
    End If
    'draw picture if exist
    putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig
    'find font x-pos
    UserControl.CurrentX = getFontX(bWid, mInd, havePIC, mePTPicAlig) 'getFontX(bWid, mInd)
    'if this is active tab then make some effect
    If mInd = selIndex And efectSel = True Then
        If mePTOrientation = ptOrientationTop Then
            UserControl.CurrentY = UserControl.CurrentY - 1
        Else
            UserControl.CurrentY = UserControl.CurrentY + 1
        End If
        UserControl.CurrentX = UserControl.CurrentX - 1
    End If
    'print caption
    UserControl.Print lstCaptions.Item(mInd + 1)
    'hide line at bottom of active tab
    hideBottom mInd, mX, mYInc, bWid, mTabHeig, mBC, controlHLightC, controlShadC
    'find position ef next tab
    mX = mX + bWid
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub drawCornerCutRight(ByVal mInd As Integer)
    Dim cutWid As Integer, cutHeig As Integer, mTabHeig As Integer, mYInc As Integer
    Dim fntMovPC As Byte, pcMov As Byte
    '
    Dim havePIC As Boolean
    'setValues mInd, mBDRC, mBC, mShadC
    '
    'ucSW = UserControl.ScaleWidth
    'ucSH = UserControl.ScaleHeight
    'find tab height
    mTabHeig = tabHeig
    mYInc = Abs(tabHeigActive - tabHeig)
    If mInd = selIndex Then
        mTabHeig = tabHeigActive
        If tabHeigActive > tabHeig Then mYInc = 0
    Else
        If tabHeig > tabHeigActive Then mYInc = 0
    End If
    '
    cutWid = mTabHeig / 3 * Cos(60 * 3.14 / 180)
    cutHeig = mTabHeig / 3
    'find tab width
    bWid = UserControl.TextWidth(lstCaptions.Item(mInd + 1)) + mSpacing * 2 + 10
    If mePTStylesActive <> ptCoolLeft And mePTStylesNormal <> ptCoolLeft Then mX = mX + getWidInc(mInd)
    'if there is picture then increase tab wid
    havePIC = False
    If Not tabIcons(mInd) Is Nothing Then
        If mePTPicAlig <> ptPicCenter Then bWid = bWid + icWid + mSpacing '* 2
        havePIC = True
    End If
    lstSizes.Add bWid
    '
    If bWid + mX < 0 Or mX > ucSW Then
        mX = mX + bWid
        Exit Sub
    End If
    'draw tab
    If mePTOrientation = ptOrientationTop Then
        '
        UserControl.Line (mX + bWid, cutHeig + mYInc)-(mX + bWid, mTabHeig + mYInc), mBDRC
        UserControl.Line (mX + bWid, cutHeig + mYInc)-(mX - cutWid + bWid, mYInc), mBDRC
        UserControl.Line (mX - cutWid + bWid, mYInc)-(mX, mYInc), mBDRC
        UserControl.Line (mX, mYInc)-(mX, mTabHeig + mYInc), mBDRC
        UserControl.Line (mX, mTabHeig + mYInc)-(mX + bWid, mTabHeig + mYInc), mBDRC
        'fill tab
        'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
        If mX < 0 And mX + bWid > 0 Then
            ExtFloodFill UserControl.hdc, mX + bWid - 2, mTabHeig - 2 + mYInc, UserControl.Point(mX + bWid - 2, mTabHeig - 2 + mYInc), 1
        ElseIf mX < ucSW Then
            ExtFloodFill UserControl.hdc, mX + 2, mTabHeig - 2 + mYInc, UserControl.Point(mX + 2, mTabHeig - 2 + mYInc), 1
        End If
        'draw shadow
        If mePTAppearance = ptAppearance3D Then
            'center
            UserControl.Line (mX - cutWid + bWid - 1, 1 + mYInc)-(mX + 1, 1 + mYInc), controlHLightC
            'left
            UserControl.Line (mX + 1, 1 + mYInc)-(mX + 1, mTabHeig + mYInc), controlHLightC
            'right
            UserControl.Line (mX + bWid - 1, cutHeig + mYInc)-(mX + bWid - 1, mTabHeig + mYInc), controlShadC
            UserControl.Line (mX + bWid - 1, cutHeig + mYInc)-(mX - cutWid + bWid - 1, 0 + mYInc), controlShadC

        End If
        '
        If mePTVertTextAlg = ptTop Then
            UserControl.CurrentY = 2 + mYInc 'UserControl.TextHeight(lstCaptions.Item(mInd + 1)) + 5
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 + mYInc
        Else
            UserControl.CurrentY = mTabHeig - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 + mYInc
        End If
    Else
        UserControl.Line (mX + bWid, ucSH - 1 - cutHeig - mYInc)-(mX + bWid, ucSH - 1 - mTabHeig - mYInc), mBDRC
        UserControl.Line (mX + bWid, ucSH - 1 - cutHeig - mYInc)-(mX - cutWid + bWid, ucSH - 1 - mYInc), mBDRC
        UserControl.Line (mX - cutWid + bWid, ucSH - 1 - mYInc)-(mX, ucSH - 1 - mYInc), mBDRC
        UserControl.Line (mX, ucSH - 1 - mYInc)-(mX, ucSH - 1 - mTabHeig - mYInc), mBDRC
        UserControl.Line (mX, ucSH - 1 - mTabHeig - mYInc)-(mX + bWid, ucSH - 1 - mTabHeig - mYInc), mBDRC
        'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
        'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
        If mX < 0 And mX + bWid > 0 Then
            ExtFloodFill UserControl.hdc, mX + bWid - 2, ucSH - 1 - mTabHeig + 2 - mYInc, UserControl.Point(mX + bWid - 2, ucSH - 1 - mTabHeig + 2 - mYInc), 1
        ElseIf mX < ucSW Then
            ExtFloodFill UserControl.hdc, mX + 2, ucSH - 1 - mTabHeig + 2 - mYInc, UserControl.Point(mX + 2, ucSH - 1 - mTabHeig + 2 - mYInc), 1
        End If
        'draw shadow
        If mePTAppearance = ptAppearance3D Then
            'center
            UserControl.Line (mX - cutWid + bWid - 1, ucSH - 1 - 1 - mYInc)-(mX + 1, ucSH - 1 - 1 - mYInc), controlShadC
            'left
            UserControl.Line (mX + 1, ucSH - 1 - 1 - mYInc)-(mX + 1, ucSH - 1 - mTabHeig - mYInc), controlHLightC
            'right
            UserControl.Line (mX + bWid - 1, ucSH - 1 - cutHeig - mYInc)-(mX + bWid - 1, ucSH - 1 - mTabHeig - mYInc), controlShadC
            UserControl.Line (mX + bWid - 1, ucSH - 1 - cutHeig - mYInc)-(mX - cutWid + bWid - 1, ucSH - 1 - mYInc), controlShadC
        End If
        '
        If mePTVertTextAlg = ptBottom Then
            UserControl.CurrentY = ucSH - mTabHeig + 2 - mYInc
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = ucSH - mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 - mYInc
        Else
            UserControl.CurrentY = ucSH - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 - mYInc
        End If
    End If
    'draw picture if exist
    putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig
    'find font x
    UserControl.CurrentX = getFontX(bWid, mInd, havePIC, mePTPicAlig) 'getFontX(bWid, mInd)
    'if this is active tab then make som efect
    If mInd = selIndex And efectSel = True Then
        If mePTOrientation = ptOrientationTop Then
            UserControl.CurrentY = UserControl.CurrentY - 1
        Else
            UserControl.CurrentY = UserControl.CurrentY + 1
        End If
        UserControl.CurrentX = UserControl.CurrentX - 1
    End If
    'print caption
    UserControl.Print lstCaptions.Item(mInd + 1)
    'hide line at bottom of active tab
    hideBottom mInd, mX, mYInc, bWid, mTabHeig, mBC, controlHLightC, controlShadC
    'fond position of next tab
    mX = mX + bWid
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub drawVerLineTab(ByVal mInd As Integer)
'''''    Dim ucSW As Integer, ucSH As Integer, bWid As Integer
    
'''''    Dim mBDRC As OLE_COLOR, mBC As OLE_COLOR, mShadC As OLE_COLOR, mFC As OLE_COLOR
    Dim mTabHeig As Integer, mYInc As Integer
    ''Dim fntMovPC As byte, pcMov As byte
    'Dim havePIC As Boolean
    '
    'setValues mInd, mBDRC, mBC, mShadC
    'set control font
    'Set UserControl.Font = mFNT
    'If UserControl.ForeColor <> mFC Then UserControl.ForeColor = mFC
    'ucSW = UserControl.ScaleWidth
    'ucSH = UserControl.ScaleHeight
    'find tab width
    bWid = UserControl.TextWidth(lstCaptions.Item(mInd + 1)) + mSpacing * 2 + 10  'getwidinc
    '
    lstSizes.Add bWid
    '
    If bWid + mX < 0 Or mX > ucSW Then
        mX = mX + bWid
        Exit Sub
    End If
    'find tab height
    mTabHeig = tabHeig
    mYInc = Abs(tabHeigActive - tabHeig)
    If mInd = selIndex Then
        mTabHeig = tabHeigActive
        If tabHeigActive > tabHeig Then mYInc = 0
    Else
        If tabHeig > tabHeigActive Then mYInc = 0
    End If
    'draw tab
    If mePTOrientation = ptOrientationTop Then
        If mInd <> lstCaptions.Count - 1 And mInd <> selIndex - 1 Then
            UserControl.Line (mX + bWid, 2 + mYInc)-(mX + bWid, mTabHeig - 2 + mYInc), mBDRC
            'center
            If mePTAppearance = ptAppearance3D Then UserControl.Line (mX + bWid + 1, 2 + mYInc)-(mX + bWid + 1, mTabHeig - 2 + mYInc), controlHLightC
        End If
        'find font y
        If mePTVertTextAlg = ptTop Then
            UserControl.CurrentY = 2 + mYInc 'UserControl.TextHeight(lstCaptions.Item(mInd + 1)) + 5
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 + mYInc
        Else
            UserControl.CurrentY = mTabHeig - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 + mYInc
        End If
    Else
        If mInd <> lstCaptions.Count - 1 And mInd <> selIndex - 1 Then
            UserControl.Line (mX + bWid, ucSH - 2 - mYInc)-(mX + bWid, ucSH - mTabHeig + 2 - mYInc), mBDRC
            'center
            If mePTAppearance = ptAppearance3D Then UserControl.Line (mX + bWid + 1, ucSH - 2 - mYInc)-(mX + bWid + 1, ucSH - mTabHeig + 2 - mYInc), controlHLightC
        End If
        'find font y
        If mePTVertTextAlg = ptBottom Then
            UserControl.CurrentY = ucSH - mTabHeig + 2 - mYInc '+ UserControl.TextHeight(lstCaptions.Item(mInd + 1)) + 5
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = ucSH - mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 - mYInc
        Else
            UserControl.CurrentY = ucSH - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 - mYInc
        End If
    End If
    '
    'putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig
    '
    'UserControl.CurrentX = getFontX(bWid, mInd, havePIC, mePTPicAlig)
    'find font x
    
    UserControl.CurrentX = getFontX(bWid, mInd, False) '+ 5
    'print caption
    UserControl.Print lstCaptions.Item(mInd + 1)
    'find x-position of next tab
    mX = mX + bWid
End Sub
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub drawRoundedRectangle2Tab(ByVal mInd As Integer)
    Dim mTabHeig As Integer, mYInc As Integer
    Dim havePIC As Boolean, fntMovPC As Integer, pcMov As Byte
    'setValues mInd, mBDRC, mBC, mShadC
    If mInd = selIndex Then RDX = mX
    'set control properties
   ' If UserControl.ForeColor <> mFC Then UserControl.ForeColor = mFC
    'ucSW = UserControl.ScaleWidth
    'ucSH = UserControl.ScaleHeight
    'find tab width
    bWid = UserControl.TextWidth(lstCaptions.Item(mInd + 1)) + mSpacing * 2 + 10  'getwidinc
    'if there is picture then increase tab width
    havePIC = False
    If Not tabIcons(mInd) Is Nothing Then
        If mePTPicAlig <> ptPicCenter Then bWid = bWid + icWid + mSpacing + 5 '* 2
        havePIC = True
    End If
    lstSizes.Add bWid
    '
    If bWid + mX < 0 Or mX > ucSW Then
        mX = mX + bWid
        Exit Sub
    End If
    'find tab height
    mTabHeig = tabHeig
    mYInc = Abs(tabHeigActive - tabHeig)
    If mInd = selIndex Then
        mTabHeig = tabHeigActive
        If tabHeigActive > tabHeig Then mYInc = 0
    Else
        If tabHeig > tabHeigActive Then mYInc = 0
    End If
    'draw tab
    If mePTOrientation = ptOrientationTop Then
        UserControl.Line (mX + bWid, mYInc + 5)-(mX + bWid, mYInc + mTabHeig - 1), mBDRC
        UserControl.Circle (mX + bWid - 5, mYInc + 5), 5, mBDRC, 0, (90 / 180 * 3.14)
        UserControl.PSet (mX + bWid + 1, mYInc + mTabHeig - 1), mBDRC
        'UserControl.Circle (mX + bWid + 5, mYInc + mTabHeig - 3), 5, mBDRC, 3.14, (220 / 180) * 3.14
            
        UserControl.Line (mX, mYInc + 5)-(mX, mYInc + mTabHeig - 1), mBDRC
        UserControl.Circle (mX + 5, mYInc + 5), 5, mBDRC, 3.14 / 2, 3.14
        UserControl.PSet (mX - 1, mYInc + mTabHeig - 1), mBDRC
        'UserControl.Circle (mX - 5, mYInc + mTabHeig - 3), 5, mBDRC, 320 / 180 * 3.14, 0
        
        UserControl.Line (mX + 5, mYInc)-(mX - 3 + bWid, mYInc), mBDRC
        
        'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
        ExtFloodFill UserControl.hdc, mX + 1, mYInc + mTabHeig / 2, UserControl.Point(mX + 1, mYInc + mTabHeig / 2), 1
        '
        If mePTAppearance = ptAppearance3D Then
            'right
            UserControl.Line (mX + bWid - 1, mYInc + 5)-(mX + bWid - 1, mYInc + mTabHeig - 1), controlShadC
            UserControl.Circle (mX + bWid - 5, mYInc + 5), 4, controlShadC, 0, (90 / 180 * 3.14)
            UserControl.PSet (mX + bWid, mYInc + mTabHeig - 1), controlShadC
            UserControl.PSet (mX + bWid - 1, mYInc + mTabHeig - 2), controlShadC
            UserControl.PSet (mX + bWid + 1, mYInc + mTabHeig), controlShadC
            'UserControl.Circle (mX + bWid + 5, mYInc + mTabHeig - 3), 6, controlShadC, 3.14, (220 / 180) * 3.14
            'left
            UserControl.Line (mX + 1, mYInc + 5)-(mX + 1, mYInc + mTabHeig - 1), controlHLightC
            UserControl.Circle (mX + 5, mYInc + 5), 4, controlHLightC, 3.14 / 2, 3.14
            UserControl.PSet (mX, mYInc + mTabHeig - 1), controlHLightC
            UserControl.PSet (mX - 1, mYInc + mTabHeig), controlHLightC
            UserControl.PSet (mX + 1, mYInc + mTabHeig + 1), controlHLightC
            'UserControl.Circle (mX - 5, mYInc + mTabHeig - 3), 6, controlHLightC, 320 / 180 * 3.14, 0
            'center
            UserControl.Line (mX + 5, mYInc + 1)-(mX - 3 + bWid, mYInc + 1), controlHLightC

        End If
        'hide bottom
        If mInd = selIndex Then
            UserControl.Line (mX, mYInc + mTabHeig)-(mX + bWid + 1, mYInc + mTabHeig), mBC
            If mePTColorScheme = ptColorUser Then
                UserControl.Line (mX - 1, mYInc + mTabHeig + 1)-(mX + bWid + 2, mYInc + mTabHeig + 1), controlBC
            Else
                UserControl.Line (mX - 1, mYInc + mTabHeig + 1)-(mX + bWid + 2, mYInc + mTabHeig + 1), mBC
            End If
        End If
        
        'find font y
        If mePTVertTextAlg = ptTop Then
            UserControl.CurrentY = 2 + mYInc 'UserControl.TextHeight(lstCaptions.Item(mInd + 1)) + 5
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 + mYInc
        Else
            UserControl.CurrentY = mTabHeig - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 + mYInc
        End If
    Else
        UserControl.Line (mX + bWid, ucSH - 1 - mYInc - 5)-(mX + bWid, ucSH - 1 - mYInc - mTabHeig + 1), mBDRC
        UserControl.Circle (mX + bWid - 5, ucSH - 1 - mYInc - 5), 5, mBDRC, 1.5 * 3.14, 0
        UserControl.PSet (mX + bWid + 1, ucSH - 1 - mYInc - mTabHeig + 1), mBDRC
            
        UserControl.Line (mX, ucSH - 1 - mYInc - 5)-(mX, ucSH - 1 - mYInc - mTabHeig + 1), mBDRC
        UserControl.Circle (mX + 5, ucSH - 1 - mYInc - 5), 5, mBDRC, 3.14, 3.14 * 1.5
        UserControl.PSet (mX - 1, ucSH - 1 - mYInc - mTabHeig + 1), mBDRC
       
        
        UserControl.Line (mX + 5, ucSH - 1 - mYInc)-(mX - 3 + bWid, ucSH - 1 - mYInc), mBDRC
        
        'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
        ExtFloodFill UserControl.hdc, mX + 1, ucSH - 1 - mYInc - mTabHeig / 2, UserControl.Point(mX + 1, ucSH - 1 - mYInc - mTabHeig / 2), 1
        '
        If mePTAppearance = ptAppearance3D Then
            'right
            UserControl.Line (mX + bWid - 1, ucSH - 1 - mYInc - 5)-(mX + bWid - 1, ucSH - 1 - mYInc - mTabHeig + 1), controlShadC
            UserControl.Circle (mX + bWid - 5, ucSH - 1 - mYInc - 5), 4, controlShadC, 1.5 * 3.14, 0
            UserControl.PSet (mX + bWid, ucSH - 1 - mYInc - mTabHeig + 1), controlShadC
            UserControl.PSet (mX + bWid + 1, ucSH - 1 - mYInc - mTabHeig), controlShadC
            'left
            UserControl.Line (mX + 1, ucSH - 1 - mYInc - 5)-(mX + 1, ucSH - 1 - mYInc - mTabHeig + 1), controlHLightC
            UserControl.Circle (mX + 5, ucSH - 1 - mYInc - 5), 4, controlHLightC, 3.14, 3.14 * 1.5
            UserControl.PSet (mX, ucSH - 1 - mYInc - mTabHeig + 1), controlHLightC
            UserControl.PSet (mX - 1, ucSH - 1 - mYInc - mTabHeig), controlHLightC
            'UserControl.PSet (mX - 2, ucSH - 1 - mYInc - mTabHeig - 1), controlHLightC
            'center
            UserControl.Line (mX + 5, ucSH - 1 - mYInc - 1)-(mX - 3 + bWid, ucSH - 1 - mYInc - 1), controlShadC
        End If
        'hide bottom
        If mInd = selIndex Then
            UserControl.Line (mX, ucSH - 1 - mYInc - mTabHeig)-(mX + bWid + 1, ucSH - 1 - mYInc - mTabHeig), mBC
            If mePTColorScheme = ptColorUser Then
                UserControl.Line (mX - 1, ucSH - 1 - mYInc - mTabHeig - 1)-(mX + bWid + 2, ucSH - 1 - mYInc - mTabHeig - 1), controlBC
            Else
                UserControl.Line (mX - 1, ucSH - 1 - mYInc - mTabHeig - 1)-(mX + bWid + 2, ucSH - 1 - mYInc - mTabHeig - 1), mBC
            End If
        End If
        'find font y
        If mePTVertTextAlg = ptBottom Then
            UserControl.CurrentY = ucSH - mTabHeig + 2 - mYInc '+ UserControl.TextHeight(lstCaptions.Item(mInd + 1)) + 5
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = ucSH - mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 - mYInc
        Else
            UserControl.CurrentY = ucSH - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 - mYInc
        End If
    End If
    '
''''''    'putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig
''''''    '
''''''    'UserControl.CurrentX = getFontX(bWid, mInd, havePIC, mePTPicAlig)
''''''    'find font x
''''''
''''''    UserControl.CurrentX = getFontX(bWid, mInd, False) '+ 5
''''''
''''''    'print caption
''''''    UserControl.Print lstCaptions.Item(mInd + 1)
    'draw picture - if exist
    putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig
    'find font x-position
    UserControl.CurrentX = getFontX(bWid, mInd, havePIC, mePTPicAlig)
    If mInd = selIndex And efectSel = True Then
        If mePTOrientation = ptOrientationTop Then
            UserControl.CurrentY = UserControl.CurrentY - 1
        Else
            UserControl.CurrentY = UserControl.CurrentY + 1
        End If
        UserControl.CurrentX = UserControl.CurrentX - 1
    End If
    'print caption
    UserControl.Print lstCaptions.Item(mInd + 1)

    'find x-position of next tab
    mX = mX + bWid
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub drawDistortedTab(ByVal mInd As Integer)
    Dim cutWidLFT As Integer, cutWidRGHT As Integer, cutHeigLFT As Integer, cutHeigRGHT As Integer, mTabHeig As Integer, mYInc As Integer
    Dim mWidInc As Integer
    Dim fntMovPC As Byte, pcMov As Byte, i As Integer, z As Integer
    '
    Dim havePIC As Boolean
    'setValues mInd, mBDRC, mBC, mShadC
    'ucSW = UserControl.ScaleWidth
    'ucSH = UserControl.ScaleHeight
    'find tab height
    mTabHeig = tabHeig
    
    'find visible tab area
    mWidInc = Int(getWidInc(mInd) / 2)
    '   if tab is active then whole tab area is visible
    If mInd = selIndex Then
        cutWidLFT = mWidInc
        cutHeigLFT = mTabHeig
        cutWidRGHT = cutWidLFT
        cutHeigRGHT = mTabHeig
    '   else use default values
    Else
        cutWidLFT = buttSpacing
        cutHeigLFT = buttSpacing / Cos(60 * 3.14 / 180)
        cutWidRGHT = mWidInc
        cutHeigRGHT = mTabHeig
    End If
    '   if this is first tab then whole tab is visible
    If mInd = firstItemL - 1 Then
        cutHeigLFT = mTabHeig
        cutWidLFT = mWidInc
    End If
    '   if this is tab at left of active tab then cut right side
    If mInd = selIndex - 1 Then
        cutWidRGHT = buttSpacing + 1
        cutHeigRGHT = buttSpacing / Cos(60 * 3.14 / 180) + 1
    End If
    '   if this is tab at right of active tab then cut left side
    If mInd = selIndex + 1 And mInd <> firstItemL - 1 Then
        cutWidLFT = buttSpacing
        cutHeigLFT = buttSpacing / Cos(60 * 3.14 / 180)
    End If
    '   if there is other styles then show whole tab
    If mePTStylesActive <> ptDistorted Or mePTStylesNormal <> ptDistorted Then
        cutWidLFT = mWidInc
        cutHeigLFT = mTabHeig
        cutWidRGHT = cutWidLFT
        cutHeigRGHT = mTabHeig
    End If
    'make som check
    If cutWidLFT > mWidInc Then cutWidLFT = mWidInc
    If cutWidRGHT > mWidInc Then cutWidRGHT = mWidInc
    If cutHeigLFT > mTabHeig Then cutHeigLFT = mTabHeig
    If cutHeigRGHT > mTabHeig Then cutHeigRGHT = mTabHeig
    'find tab width
    bWid = UserControl.TextWidth(lstCaptions.Item(mInd + 1)) + mSpacing * 2 + 10 + getWidInc(mInd)
    'if there is picture then increase tab width
    havePIC = False
    If Not tabIcons(mInd) Is Nothing Then
        If mePTPicAlig <> ptPicCenter Then bWid = bWid + icWid + mSpacing '* 2
        havePIC = True
    End If
    lstSizes.Add bWid
    '
    If bWid + mX < 0 Or mX > ucSW Then
        mX = mX + bWid
        Exit Sub
    End If
    If buttSpacing < 4 Then buttSpacing = 4
    
    'draw tab
    If mePTOrientation = ptOrientationTop Then
        '
        UserControl.Line (mX + mWidInc - cutWidLFT, mYInc + cutHeigLFT)-(mX + mWidInc, mYInc), mBDRC
        UserControl.Line (mX + mWidInc, mYInc)-(mX + bWid - mWidInc, mYInc), mBDRC
        UserControl.Line (mX + bWid - mWidInc, mYInc)-(mX + bWid - mWidInc + cutWidRGHT, mYInc + cutHeigRGHT), mBDRC
        'close tab area to fill
        UserControl.Line (mX + bWid - mWidInc + cutWidRGHT, mYInc + cutHeigRGHT)-(mX + bWid - mWidInc + cutWidRGHT, mYInc + mTabHeig), mBDRC
        'hide what need to be hidden
        If mInd = selIndex Then
            For i = 0 To mTabHeig
                If i Mod 2 = 1 Then z = z + 1
                UserControl.Line (mX + mWidInc - z + 1, mYInc + i + 1)-(mX + bWid - mWidInc + z - 1, mYInc + i + 1), tabAreaBC
            Next i
        End If
        '
        'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
        'find fill point
        '   for active tab
        If mInd = selIndex Then
            'if position is at left and not whole tab is visible
            If mX < mWidInc - cutWidLFT And mX + bWid > 0 Then
                ExtFloodFill UserControl.hdc, mX + bWid - 2, mYInc + tabHeig - 1, UserControl.Point(mX + bWid - 2, mYInc + tabHeig - 1), 1
                ExtFloodFill UserControl.hdc, 1, mYInc + tabHeig - 1, UserControl.Point(1, mYInc + tabHeig - 1), 1
            'else
            ElseIf mX < ucSW - mWidInc + cutWidLFT Then
                ExtFloodFill UserControl.hdc, mX + 2 - cutWidLFT + mWidInc, mYInc + cutHeigLFT - 2, UserControl.Point(mX + 2 - cutWidLFT + mWidInc, mYInc + cutHeigLFT - 2), 1
                ExtFloodFill UserControl.hdc, mX + 2 + mWidInc + lstRghtSpc, mYInc + cutHeigLFT - 1, UserControl.Point(mX + 2 + mWidInc + lstRghtSpc, mYInc + cutHeigLFT - 1), 1
            End If
        '   for tab at left of active tab
        ElseIf mInd = selIndex - 1 Then
            'if position is at left and not whole tab is visible
            If mX < mWidInc - cutWidLFT And mX + bWid > 0 Then
                ExtFloodFill UserControl.hdc, mX + bWid + cutWidRGHT - mWidInc - 3, mYInc + cutHeigRGHT - 1, UserControl.Point(mX + bWid + cutWidRGHT - mWidInc - 3, mYInc + cutHeigRGHT - 1), 1
            ElseIf mX < ucSW - mWidInc + cutWidLFT Then
                ExtFloodFill UserControl.hdc, mX + 2 - cutWidLFT + mWidInc, mYInc + cutHeigLFT - 3, UserControl.Point(mX + 2 - cutWidLFT + mWidInc, mYInc + cutHeigLFT - 1), 1
            End If
        '   for all other tabs
        Else
            'if position is at left and not whole tab is visible
            If mX < mWidInc - cutWidLFT And mX + bWid > 0 And mInd <> selIndex - 1 Then
                ExtFloodFill UserControl.hdc, mX + bWid - 2, mYInc + tabHeig - 1, UserControl.Point(mX + bWid - 2, mYInc + tabHeig - 1), 1
            ElseIf mX < ucSW - mWidInc + cutWidLFT Then
                ExtFloodFill UserControl.hdc, mX + 2 - cutWidLFT + mWidInc, mYInc + cutHeigLFT - 1, UserControl.Point(mX + 2 - cutWidLFT + mWidInc, mYInc + cutHeigLFT - 1), 1
            End If
        End If
        'hide liine for close fill area
        UserControl.Line (mX + bWid - mWidInc + cutWidRGHT, mYInc + cutHeigRGHT)-(mX + bWid - mWidInc + cutWidRGHT, mYInc + mTabHeig), tabAreaBC
        lstRghtSpc = cutWidRGHT
        'draw shadow
        If mePTAppearance = ptAppearance3D Then
            'left
            UserControl.Line (mX + mWidInc - cutWidLFT + 1, mYInc + cutHeigLFT)-(mX + mWidInc + 1, mYInc), controlHLightC
            'center
            UserControl.Line (mX + mWidInc + 1, mYInc + 1)-(mX + bWid - mWidInc - 1, mYInc + 1), controlHLightC
            'right
            UserControl.Line (mX + bWid - mWidInc - 1, mYInc)-(mX + bWid - mWidInc + cutWidRGHT - 1, mYInc + cutHeigRGHT), controlShadC
        End If
        'redraw bottom
        If mInd <> selIndex And mInd <> selIndex + 1 Then UserControl.Line (mX, mYInc + mTabHeig)-(mX + 2, mYInc + mTabHeig), mBDRC
        '
        UserControl.Line (mX + mWidInc, mYInc)-(mX + bWid - mWidInc, mYInc), mBDRC
        If mePTVertTextAlg = ptTop Then
            UserControl.CurrentY = 2 + mYInc 'UserControl.TextHeight(lstCaptions.Item(mInd + 1)) + 5
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 + mYInc
        Else
            UserControl.CurrentY = mTabHeig - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 + mYInc
        End If
    Else
        UserControl.Line (mX + mWidInc - cutWidLFT, ucSH - 1 - mYInc - cutHeigLFT)-(mX + mWidInc, ucSH - 1 - mYInc), mBDRC
        UserControl.Line (mX + mWidInc, ucSH - 1 - mYInc)-(mX + bWid - mWidInc, ucSH - 1 - mYInc), mBDRC
        UserControl.Line (mX + bWid - mWidInc, ucSH - 1 - mYInc)-(mX + bWid - mWidInc + cutWidRGHT, ucSH - 1 - mYInc - cutHeigRGHT), mBDRC
        'close tab area to fill
        UserControl.Line (mX + bWid - mWidInc + cutWidRGHT, ucSH - 1 - mYInc - cutHeigRGHT)-(mX + bWid - mWidInc + cutWidRGHT, ucSH - 1 - mYInc - mTabHeig), mBDRC
        'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
        'hide what need to be hidden
        If mInd = selIndex Then
            For i = 0 To mTabHeig
                If i Mod 2 = 1 Then z = z + 1
                UserControl.Line (mX + mWidInc - z + 1, ucSH - mYInc - i - 2)-(mX + bWid - mWidInc + z - 1, ucSH - mYInc - i - 2), tabAreaBC
            Next i
        End If
        'find fill point
        '   for active tab
        If mInd = selIndex Then
            If mX < mWidInc - cutWidLFT And mX + bWid > 0 Then
                ExtFloodFill UserControl.hdc, mX + bWid - 2, ucSH - 1 - mYInc - tabHeig + 1, UserControl.Point(mX + bWid - 2, ucSH - 1 - mYInc - tabHeig + 1), 1
                ExtFloodFill UserControl.hdc, 1, ucSH - 1 - mYInc - tabHeig + 1, UserControl.Point(1, ucSH - 1 - mYInc - tabHeig + 1), 1
            ElseIf mX < ucSW - mWidInc + cutWidLFT Then
                ExtFloodFill UserControl.hdc, mX + 2 - cutWidLFT + mWidInc, ucSH - 1 - mYInc - cutHeigLFT + 2, UserControl.Point(mX + 2 - cutWidLFT + mWidInc, ucSH - 1 - mYInc - cutHeigLFT + 2), 1
                ExtFloodFill UserControl.hdc, mX + 2 + mWidInc + lstRghtSpc, ucSH - 1 - mYInc - cutHeigLFT + 1, UserControl.Point(mX + 2 + mWidInc + lstRghtSpc, ucSH - 1 - mYInc - cutHeigLFT + 1), 1
            End If
        ElseIf mInd = selIndex - 1 Then
            If mX < mWidInc - cutWidLFT And mX + bWid > 0 Then
                ExtFloodFill UserControl.hdc, mX + bWid + cutWidRGHT - mWidInc - 3, ucSH - 1 - mYInc - cutHeigRGHT + 1, UserControl.Point(mX + bWid + cutWidRGHT - mWidInc - 3, ucSH - 1 - mYInc - cutHeigRGHT + 1), 1
            ElseIf mX < ucSW - mWidInc + cutWidLFT Then
                ExtFloodFill UserControl.hdc, mX + 2 - cutWidLFT + mWidInc, ucSH - 1 - mYInc - cutHeigLFT + 3, UserControl.Point(mX + 2 - cutWidLFT + mWidInc, ucSH - 1 - mYInc - cutHeigLFT + 1), 1
            End If
        Else
            If mX < mWidInc - cutWidLFT And mX + bWid > 0 And mInd <> selIndex - 1 Then
                ExtFloodFill UserControl.hdc, mX + bWid - 2, ucSH - 1 - mYInc - tabHeig + 1, UserControl.Point(mX + bWid - 2, ucSH - 1 - mYInc - tabHeig + 1), 1
            ElseIf mX < ucSW - mWidInc + cutWidLFT Then
                ExtFloodFill UserControl.hdc, mX + 2 - cutWidLFT + mWidInc, ucSH - 1 - mYInc - cutHeigLFT + 1, UserControl.Point(mX + 2 - cutWidLFT + mWidInc, ucSH - 1 - mYInc - cutHeigLFT + 1), 1
            End If
        End If
        UserControl.Line (mX + bWid - mWidInc + cutWidRGHT, ucSH - 1 - mYInc - cutHeigRGHT)-(mX + bWid - mWidInc + cutWidRGHT, ucSH - 1 - mYInc - mTabHeig), tabAreaBC
        lstRghtSpc = cutWidRGHT
        'draw shadow
        If mePTAppearance = ptAppearance3D Then
            'left
            UserControl.Line (mX + mWidInc - cutWidLFT + 1, ucSH - 1 - mYInc - cutHeigLFT)-(mX + mWidInc + 1, ucSH - 1 - mYInc), controlHLightC
            'center
            UserControl.Line (mX + mWidInc + 1, ucSH - 1 - mYInc - 1)-(mX + bWid - mWidInc - 1, ucSH - 1 - mYInc - 1), controlShadC
            'right
            UserControl.Line (mX + bWid - mWidInc - 1, ucSH - 1 - mYInc)-(mX + bWid - mWidInc + cutWidRGHT - 1, ucSH - 1 - mYInc - cutHeigRGHT), controlShadC
        End If
        'redraw bottom
        If mInd <> selIndex And mInd <> selIndex + 1 Then UserControl.Line (mX, ucSH - 1 - mYInc - mTabHeig)-(mX + 2, ucSH - 1 - mYInc - mTabHeig), mBDRC
        '
        UserControl.Line (mX + mWidInc, ucSH - 1 - mYInc)-(mX + bWid - mWidInc, ucSH - 1 - mYInc), mBDRC
        '
        If mePTVertTextAlg = ptBottom Then
            UserControl.CurrentY = ucSH - mTabHeig + 2 - mYInc '+ UserControl.TextHeight(lstCaptions.Item(mInd + 1)) + 5
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = ucSH - mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 - mYInc
        Else
            UserControl.CurrentY = ucSH - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 - mYInc
        End If
    End If
    'draw picture if exist
    putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig
    'find font x
    UserControl.CurrentX = getFontX(bWid, mInd, havePIC, mePTPicAlig) 'getFontX(bWid, mInd)
    'if this is active tab then make som efect
    If mInd = selIndex And efectSel = True Then
        If mePTOrientation = ptOrientationTop Then
            UserControl.CurrentY = UserControl.CurrentY - 1
        Else
            UserControl.CurrentY = UserControl.CurrentY + 1
        End If
        UserControl.CurrentX = UserControl.CurrentX - 1
    End If
    'print caption
    UserControl.Print lstCaptions.Item(mInd + 1)
    'hide line at bottom of active tab
    hideBottom mInd, mX, mYInc, bWid, mTabHeig, mBC, controlHLightC, controlShadC
    'if there is other styles then increase position of next tab because
    '   next code decrease is for same value (decrease only when all styles
    '   are distorted
    If mePTStylesActive <> ptDistorted Or mePTStylesNormal <> ptDistorted Then
        mX = mX + Abs(buttSpacing - mWidInc) + mWidInc
    End If
    'decrease position of next tab to create effect like
    '   tab on tab
    mX = mX + bWid - mWidInc - Abs(buttSpacing - mWidInc)
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub drawCoolLeftTab(ByVal mInd As Integer)
    Dim mTabHeig As Integer, mYInc As Integer
    Dim fntMovPC As Byte, pcMov As Byte
    Dim cSize3 As Byte
    Dim havePIC As Boolean, drawAll As Boolean
    Dim mWINC As Integer, i As Integer
    
    '
    mWINC = 0
    If mInd = selIndex And mInd <> 0 Then
        mWINC = Abs(getWidInc(0) - getWidInc(mInd))
    ElseIf mInd = 0 And selIndex = 0 Then
        If tabCnt > 1 Then mWINC = Abs(getWidInc(0) - getWidInc(1))
    End If
    '
    '

    'find tab height
    mTabHeig = tabHeig
    mYInc = Abs(tabHeigActive - tabHeig)
    If mInd = selIndex Then
        mTabHeig = tabHeigActive
        If tabHeigActive > tabHeig Then mYInc = 0
    Else
        If tabHeig > tabHeigActive Then mYInc = 0
    End If
    cSize3 = Int(mTabHeig / 10) + 1
    '
    'find tab width
    bWid = UserControl.TextWidth(lstCaptions.Item(mInd + 1)) + mSpacing * 2 + 10 + getWidInc(mInd) '+ cSize1 + cSize2  'getwidinc
    'if there is picture then increase tab width
    havePIC = False
    If Not tabIcons(mInd) Is Nothing Then
        If mePTPicAlig <> ptPicCenter Then bWid = bWid + icWid + mSpacing * 2
        havePIC = True
    End If
    'add tab width to collection
    lstSizes.Add bWid
    If bWid + mX < 0 Or mX > ucSW Then
        mX = mX + bWid
        mX = mX - getWidInc(mInd) + cSize2 - cSize1 - 1 + mWINC
        Exit Sub
    End If
    '
    If mePTStylesNormal <> ptCoolLeft Or mePTStylesActive <> ptCoolLeft Then drawAll = True Else drawAll = False
    'draw tab
    If mePTOrientation = ptOrientationTop Then
        'draw left side
        If mInd = firstItemL - 1 Or mInd = selIndex Or drawAll = True Then
            UserControl.Line (mX, mYInc + mTabHeig - 1)-(mX + cSize1, mYInc + mTabHeig - 1), mBDRC
            UserControl.Line (mX + cSize1 - 2, mYInc + mTabHeig)-(mX + cSize1 + getWidInc(mInd, cSize2 / 3) - 2, mYInc + cSize2 / 3), mBDRC
        End If
        UserControl.Circle (mX + cSize1 + getWidInc(mInd, cSize2) + cSize2 - 1, mYInc + cSize2 + 1), cSize2, mBDRC, 3.14 / 2, (150 / 180) * 3.14
        UserControl.Line (mX + cSize1 + getWidInc(mInd, cSize2) + cSize2, mYInc)-(mX + bWid - 1, mYInc), mBDRC
        UserControl.Circle (mX + bWid - cSize3, mYInc + cSize3), cSize3, mBDRC, 0, 3.14 / 2
        '
        If mInd <> selIndex - 1 Or drawAll = True Then
            UserControl.Line (mX + bWid, mYInc + cSize3)-(mX + bWid, mYInc + mTabHeig), mBDRC
        End If
        'close tab area to fill
        If mInd <> selIndex And drawAll <> True Then
            If mInd = selIndex - 1 Then UserControl.Line (mX + bWid, mYInc + cSize3)-(mX + bWid, mYInc + mTabHeig), mBDRC
        Else
            'before fill hide all what dont need be visible
            For i = 0 To mTabHeig - cSize2
                UserControl.Line (mX + cSize1 + 1 + i, mYInc + mTabHeig - 1 - i)-(mX + getWidInc(mInd, cSize2 / 3) + 5, mYInc + mTabHeig - 1 - i), tabAreaBC
            Next i
        End If

        'find fill point
        '   for active tab
        If mInd = selIndex Or drawAll = True Then
            'if position is at left and not whole tab is visible
            'MsgBox usercontrol.hdc & vbCrLf & UserControl.hdc
            If mX < 0 And mX + bWid > 0 Then
                ExtFloodFill UserControl.hdc, mX + bWid - 2, mYInc + tabHeig - 1, UserControl.Point(mX + bWid - 2, mYInc + tabHeig - 1), 1
                ExtFloodFill UserControl.hdc, 1, mYInc + tabHeig - 1, UserControl.Point(1, mYInc + tabHeig - 1), 1
            'else
            ElseIf mX < ucSW - getWidInc(mInd) Then
                ExtFloodFill UserControl.hdc, mX + 2, mYInc + mTabHeig - 1, UserControl.Point(mX + 2, mYInc + mTabHeig - 1), 1
                ExtFloodFill UserControl.hdc, mX + 2 + getWidInc(mInd) + lstRghtSpc, mYInc + mTabHeig - 1, UserControl.Point(mX + 2 + getWidInc(mInd) + lstRghtSpc, mYInc + mTabHeig - 1), 1
            End If
        '   for tab at left of active tab
        ElseIf mInd = selIndex - 1 Then
            'if position is at left and not whole tab is visible
            If mX < 0 And mX + bWid > 0 Then
                ExtFloodFill UserControl.hdc, mX + bWid - 2, mYInc + 2, UserControl.Point(mX + bWid - 2, mYInc + 2), 1
            ElseIf mX < ucSW - getWidInc(mInd) Then
                ExtFloodFill UserControl.hdc, mX + 2 + getWidInc(mInd), mYInc + mTabHeig - 2, UserControl.Point(mX + 2 + getWidInc(mInd), mYInc + mTabHeig - 2), 1
            End If
        '   for all other tabs
        Else
            'if position is at left and not whole tab is visible
            If mX < getWidInc(mInd) And mX + bWid > 0 And mInd <> selIndex - 1 Then
                ExtFloodFill UserControl.hdc, mX + bWid - 2, mYInc + 2, UserControl.Point(mX + bWid - 2, mYInc + 2), 1
            ElseIf mX < ucSW - getWidInc(mInd) Then
                ExtFloodFill UserControl.hdc, mX + 2 + getWidInc(mInd), mYInc + mTabHeig - 2, UserControl.Point(mX + 2 + getWidInc(mInd), mYInc + mTabHeig - 2), 1
            End If
        End If

        If mInd <> selIndex Then
            If mInd = selIndex - 1 Then UserControl.Line (mX + bWid, mYInc + cSize3)-(mX + bWid, mYInc + mTabHeig), mBC
            'If mInd <> 0 Then UserControl.Line (mX + getWidInc(mInd, cSize2) + cSize1 + 1, mYInc + cSize2)-(mX + getWidInc(mInd, cSize2) + cSize1 + 1, mYInc + mTabHeig), mBC
        End If
'        'draw shadow
        If mePTAppearance = ptAppearance3D Then
            'left shadow
            If mInd = firstItemL - 1 Or mInd = selIndex Or drawAll = True Then
                UserControl.Line (mX + 1, mYInc + mTabHeig - 1)-(mX + cSize1 + 1, mYInc + mTabHeig - 1), controlHLightC
                UserControl.Line (mX + cSize1 - 2 + 1, mYInc + mTabHeig)-(mX + cSize1 + getWidInc(mInd, cSize2 / 3) - 2, mYInc + cSize2 / 3 + 1), controlHLightC
                'If mInd <> selIndex Then UserControl.Line (mX, mTabHeig + mYInc)-(mX + bWid, mTabHeig + mYInc), controlHLightC
            End If
            'UserControl.Circle (mX + cSize1 + getWidInc(mInd, cSize2) + cSize2, mYInc + cSize2), cSize2 - 1, mShadC, 3.14 / 2, (140 / 180) * 3.14
            UserControl.Circle (mX + cSize1 + getWidInc(mInd, cSize2) + cSize2 - 1, mYInc + cSize2 + 1), cSize2 - 1, controlHLightC, 3.14 / 2, (150 / 180) * 3.14
            'center
            UserControl.Line (mX + cSize1 + getWidInc(mInd, cSize2) + cSize2, mYInc + 1)-(mX + bWid - 1, mYInc + 1), controlHLightC
            'right
            UserControl.Circle (mX + bWid - cSize3 + 1, mYInc + cSize3), cSize3 - 1, controlShadC, 0, 3.14 / 2
            If mInd <> selIndex - 1 Or drawAll = True Then
                UserControl.Line (mX + bWid - 1, mYInc + cSize3)-(mX + bWid - 1, mYInc + mTabHeig), controlShadC
            End If
        End If

        'redraw
        If mInd = firstItemL - 1 Or mInd = selIndex Or drawAll = True Then UserControl.Line (mX, mYInc + mTabHeig - 1)-(mX + cSize1, mYInc + mTabHeig - 1), mBDRC
        'find font y-pos
        If mePTVertTextAlg = ptTop Then
            UserControl.CurrentY = 2 + mYInc 'UserControl.TextHeight(lstCaptions.Item(mInd + 1)) + 5
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 + mYInc
        Else
            UserControl.CurrentY = mTabHeig - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 + mYInc
        End If
    Else
        'cSize2 = cSize2 + 1
        'draw left side
        If mInd = firstItemL - 1 Or mInd = selIndex Or drawAll = True Then
            UserControl.Line (mX, ucSH - 1 - mYInc - mTabHeig + 1)-(mX + cSize1, ucSH - 1 - mYInc - mTabHeig + 1), mBDRC
            UserControl.Line (mX + cSize1 - 2, ucSH - 1 - mYInc - mTabHeig)-(mX + cSize1 + getWidInc(mInd, cSize2 / 3) - 2, ucSH - 1 - mYInc - cSize2 / 3), mBDRC
        End If

        UserControl.Circle (mX + cSize1 + getWidInc(mInd, cSize2) + cSize2 + 1, ucSH - 1 - mYInc - cSize2), cSize2, mBDRC, (210 / 180) * 3.14, 1.5 * 3.14
        UserControl.PSet (mX + cSize1 + getWidInc(mInd, cSize2 / 3) - 2, ucSH - 1 - mYInc - cSize2 / 3 - 1), mBDRC
        UserControl.PSet (mX + cSize1 + getWidInc(mInd, cSize2 / 3) - 2 + 1, ucSH - 1 - mYInc - cSize2 / 3 - 1), mBC
        '
        UserControl.Line (mX + cSize1 + getWidInc(mInd, cSize2) + cSize2, ucSH - 1 - mYInc)-(mX + bWid - 1, ucSH - 1 - mYInc), mBDRC
        UserControl.Circle (mX + bWid - cSize3, ucSH - 1 - mYInc - cSize3), cSize3, mBDRC, 1.5 * 3.14, 0
        If mInd <> selIndex - 1 Or drawAll = True Then
            UserControl.Line (mX + bWid, ucSH - 1 - mYInc - cSize3)-(mX + bWid, ucSH - 1 - mYInc - mTabHeig), mBDRC
        End If
        'close tab area to fill
        If mInd <> selIndex Then
            If mInd = selIndex - 1 Then UserControl.Line (mX + bWid, ucSH - 1 - mYInc - cSize3)-(mX + bWid, ucSH - 1 - mYInc - mTabHeig), mBDRC
        Else
            'before fill hide all what dont need be visible
            For i = 0 To mTabHeig - cSize2
                UserControl.Line (mX + cSize1 + 1 + i, ucSH - 1 - mYInc - mTabHeig + 1 + i)-(mX + getWidInc(mInd, cSize2 / 3), ucSH - 1 - mYInc - mTabHeig + 1 + i), tabAreaBC
            Next i
        End If
        'find fill point
        '   for active tab
        If mInd = selIndex Or drawAll = True Then
            'if position is at left and not whole tab is visible
            If mX < 0 And mX + bWid > 0 Then
                ExtFloodFill UserControl.hdc, mX + bWid - 2, ucSH - 1 - mYInc - tabHeig + 1, UserControl.Point(mX + bWid - 2, ucSH - 1 - mYInc - tabHeig + 1), 1
                ExtFloodFill UserControl.hdc, 1, ucSH - 1 - mYInc - tabHeig + 1, UserControl.Point(1, ucSH - 1 - mYInc - tabHeig + 1), 1
            'else
            ElseIf mX < ucSW - getWidInc(mInd) Then
                ExtFloodFill UserControl.hdc, mX + 2, ucSH - 1 - mYInc - mTabHeig + 1, UserControl.Point(mX + 2, ucSH - 1 - mYInc - mTabHeig + 1), 1
                ExtFloodFill UserControl.hdc, mX + 2 + getWidInc(mInd) + lstRghtSpc, ucSH - 1 - mYInc - mTabHeig + 1, UserControl.Point(mX + 2 + getWidInc(mInd) + lstRghtSpc, ucSH - 1 - mYInc - mTabHeig + 1), 1
            End If
        '   for tab at left of active tab
        ElseIf mInd = selIndex - 1 Then
            'if position is at left and not whole tab is visible
            If mX < 0 And mX + bWid > 0 Then
                ExtFloodFill UserControl.hdc, mX + bWid - 2, ucSH - 1 - mYInc - 2, UserControl.Point(mX + bWid - 2, ucSH - 1 - mYInc - 2), 1
            ElseIf mX < ucSW - getWidInc(mInd) Then
                ExtFloodFill UserControl.hdc, mX + 2 + getWidInc(mInd), ucSH - 1 - mYInc - mTabHeig + 2, UserControl.Point(mX + 2 + getWidInc(mInd), ucSH - 1 - mYInc - mTabHeig + 2), 1
            End If
        '   for all other tabs
        Else
            'if position is at left and not whole tab is visible
            If mX < getWidInc(mInd) And mX + bWid > 0 And mInd <> selIndex - 1 Then
                ExtFloodFill UserControl.hdc, mX + bWid - 2, ucSH - 1 - mYInc - 2, UserControl.Point(mX + bWid - 2, ucSH - 1 - mYInc - 2), 1
            ElseIf mX < ucSW - getWidInc(mInd) Then
                ExtFloodFill UserControl.hdc, mX + 2 + getWidInc(mInd), ucSH - 1 - mYInc - mTabHeig + 2, UserControl.Point(mX + 2 + getWidInc(mInd), ucSH - 1 - mYInc - mTabHeig + 2), 1
            End If
        End If

        If mInd <> selIndex Then
            If mInd = selIndex - 1 Then UserControl.Line (mX + bWid, ucSH - 1 - mYInc - cSize3)-(mX + bWid, ucSH - 1 - mYInc - mTabHeig), mBC
            'If mInd <> 0 Then UserControl.Line (mX + getWidInc(mInd, cSize2) + cSize1 + 1, mYInc + cSize2)-(mX + getWidInc(mInd, cSize2) + cSize1 + 1, mYInc + mTabHeig), mBC
        End If
        'draw shadow
        If mePTAppearance = ptAppearance3D Then
            'left
            If mInd = firstItemL - 1 Or mInd = selIndex Or drawAll = True Then
                'UserControl.Line (mX+1, mYInc + mTabHeig - 1)-(mX + cSize1+1, mYInc + mTabHeig - 1), mShadC
                UserControl.Line (mX + cSize1 - 2 + 1, ucSH - 1 - mYInc - mTabHeig)-(mX + cSize1 + getWidInc(mInd, cSize2 / 3) - 2 + 1, ucSH - 1 - mYInc - cSize2 / 3), controlHLightC
                'If mInd <> selIndex Then UserControl.Line (mX, ucSH - 1 - mTabHeig - mYInc)-(mX + bWid, ucSH - 1 - mTabHeig - mYInc), controlHLightC
            End If
            UserControl.PSet (mX + cSize1 + getWidInc(mInd, cSize2 / 3) - 2, ucSH - 1 - mYInc - cSize2 / 3 - 2), controlHLightC
            UserControl.PSet (mX + cSize1 + getWidInc(mInd, cSize2 / 3), ucSH - 1 - mYInc - cSize2 / 3 - 1), controlHLightC
            UserControl.PSet (mX + cSize1 + getWidInc(mInd, cSize2 / 3) - 1, ucSH - 1 - mYInc - cSize2 / 3 - 1), controlHLightC
            UserControl.PSet (mX + cSize1 + getWidInc(mInd, cSize2 / 3) + 1, ucSH - 1 - mYInc - cSize2 / 3), controlHLightC
           'center
            UserControl.Line (mX + cSize1 + getWidInc(mInd, cSize2) + cSize2 + 1, ucSH - 1 - mYInc - 1)-(mX + bWid - 1, ucSH - 1 - mYInc - 1), controlShadC
            'right
            UserControl.Circle (mX + bWid - cSize3, ucSH - 1 - mYInc - cSize3), cSize3 - 1, controlShadC, 1.5 * 3.14, 0
            If mInd <> selIndex - 1 Or drawAll = True Then
                UserControl.Line (mX + bWid - 1, ucSH - 1 - mYInc - cSize3)-(mX + bWid - 1, ucSH - 1 - mYInc - mTabHeig), controlShadC
            End If
        End If
        UserControl.PSet (mX + cSize1 + getWidInc(mInd, cSize2 / 3) - 2, ucSH - 1 - mYInc - cSize2 / 3 - 1), mBDRC
        'find font y-pos
        If mePTVertTextAlg = ptBottom Then
            UserControl.CurrentY = ucSH - mTabHeig + 2 - mYInc '+ UserControl.TextHeight(lstCaptions.Item(mInd + 1)) + 5
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = ucSH - mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 - mYInc
        Else
            UserControl.CurrentY = ucSH - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 - mYInc
        End If
    End If
    'draw picture - if exist
    If mePTPicAlig = ptPicCenter Then
        putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig, mX + getWidInc(mInd) + (bWid - getWidInc(mInd)) / 2 - icWid / 2
        fntMovPC = 0
    ElseIf mePTPicAlig = ptPicLeftEdge Then
        putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig, mX + getWidInc(mInd)
        fntMovPC = 0
    ElseIf mePTPicAlig = ptPicLeftOfCapton Then
        If mePTHorTextAlg = ptLeft Then
            putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig, mX + getWidInc(mInd)
            fntMovPC = icWid + mSpacing
        ElseIf mePTHorTextAlg = ptCenter Then
            putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig, mX + getWidInc(mInd) + (bWid - getWidInc(mInd)) / 2 - UserControl.TextWidth(lstCaptions.Item(mInd + 1)) / 2 - icWid - mSpacing
            fntMovPC = 0 '(icWid + mSpacing) / 2
        ElseIf mePTHorTextAlg = ptRight Then
            putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig, mX + bWid - icWid - mSpacing * 2 - UserControl.TextWidth(lstCaptions.Item(mInd + 1))
            fntMovPC = 0 'icWid + mSpacing
        End If
    ElseIf mePTPicAlig = ptPicRightEdge Then
        putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig, mX + bWid - icWid - 5
        fntMovPC = 0
    ElseIf mePTPicAlig = ptPicRightOfCaption Then
        If mePTHorTextAlg = ptLeft Then
            putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig, mX + getWidInc(mInd) + UserControl.TextWidth(lstCaptions.Item(mInd + 1)) + mSpacing
            fntMovPC = 0 'icWid + mSpacing
        ElseIf mePTHorTextAlg = ptCenter Then
            putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig, mX + getWidInc(mInd) / 2 + (bWid - getWidInc(mInd)) / 2 + UserControl.TextWidth(lstCaptions.Item(mInd + 1)) / 2 + mSpacing '+ icWid / 2
            fntMovPC = -(icWid + mSpacing) / 2
        ElseIf mePTHorTextAlg = ptRight Then
            putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig, mX + bWid - icWid - mSpacing
            fntMovPC = icWid + mSpacing
        End If
    End If
    If tabIcons(mInd) Is Nothing Then fntMovPC = 0
    'find font x-pos
    If mePTHorTextAlg = ptCenter Then
        UserControl.CurrentX = mX + getWidInc(mInd) + (bWid - getWidInc(mInd)) / 2 - UserControl.TextWidth(lstCaptions.Item(mInd + 1)) / 2 + fntMovPC
    ElseIf mePTHorTextAlg = ptLeft Then
        UserControl.CurrentX = mX + getWidInc(mInd) + fntMovPC
    Else
        UserControl.CurrentX = mX + bWid - UserControl.TextWidth(lstCaptions.Item(mInd + 1)) - fntMovPC - mSpacing
    End If
    'make som efect if this is active tab
    If mInd = selIndex And efectSel = True Then
        If mePTOrientation = ptOrientationTop Then
            UserControl.CurrentY = UserControl.CurrentY - 1
        Else
            UserControl.CurrentY = UserControl.CurrentY + 1
        End If
        UserControl.CurrentX = UserControl.CurrentX - 1
    End If
    'print caption
    UserControl.Print lstCaptions.Item(mInd + 1)
    'hide line at bottom of active tab
    hideBottom mInd, mX, mYInc, bWid, mTabHeig, mBC, controlHLightC, controlShadC
    'find position of next tab
    
    mX = mX + bWid
    If drawAll <> True And mInd <> tabCnt - 1 Then mX = mX - getWidInc(mInd) + cSize2 - cSize1 + mWINC
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub drawCoolRightTab(ByVal mInd As Integer)
    Dim mTabHeig As Integer, mYInc As Integer
    Dim fntMovPC As Byte, pcMov As Byte
    Dim cSize3 As Byte
    '
    Dim havePIC As Boolean, drawAll As Boolean
    Dim mWINC As Integer, i As Integer
    mWINC = 0
    '
    If mInd = selIndex Then RDX = mX
    '
    If mInd = selIndex And mInd <> 0 Or mInd = tabCnt - 1 Then
        mWINC = Abs(getWidInc(0) - getWidInc(mInd))
    ElseIf mInd = 0 And selIndex = 0 Then
        If tabCnt > 1 Then mWINC = Abs(getWidInc(0) - getWidInc(1))
    End If
    '
    'setValues mInd, mBDRC, mBC, mShadC
    'set control propertues
    'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
    '
    'ucSW = UserControl.ScaleWidth
    'ucSH = UserControl.ScaleHeight
    '
    'find tab height
    mTabHeig = tabHeig
    mYInc = Abs(tabHeigActive - tabHeig)
    If mInd = selIndex Then
        mTabHeig = tabHeigActive
        If tabHeigActive > tabHeig Then mYInc = 0
    Else
        If tabHeig > tabHeigActive Then mYInc = 0
    End If
    '
    cSize3 = Int(mTabHeig / 10) + 1
    '
    'find tab width
    bWid = UserControl.TextWidth(lstCaptions.Item(mInd + 1)) + mSpacing * 2 + 10 + getWidInc(mInd) '+ cSize1 + cSize2  'getwidinc
    'if there is picture then increase tab width
    havePIC = False
    If Not tabIcons(mInd) Is Nothing Then
        If mePTPicAlig <> ptPicCenter Then bWid = bWid + icWid + mSpacing * 2
        havePIC = True
    End If
    'add tab width to collection
    lstSizes.Add bWid
    If bWid + mX < 0 Or mX > ucSW Then
        mX = mX + bWid
        mX = mX - getWidInc(mInd) + cSize2 - cSize1 - 1 + mWINC
        Exit Sub
    End If
    '
    If mePTStylesNormal <> ptCoolRight Or mePTStylesActive <> ptCoolRight Then drawAll = True Else drawAll = False
    'draw tab
    If mePTOrientation = ptOrientationTop Then
        'draw right side
        If mInd = selIndex Or drawAll = True Or mInd = tabCnt - 1 Then
            UserControl.Line (mX + bWid, mYInc + mTabHeig - 1)-(mX + bWid - cSize1, mYInc + mTabHeig - 1), mBDRC
            UserControl.Line (mX + bWid - cSize1 + 2, mYInc + mTabHeig)-(mX + bWid - cSize1 - getWidInc(mInd, cSize2 / 3) + 2, mYInc + cSize2 / 3), mBDRC
        End If
        UserControl.PSet (mX + bWid - cSize1 - cSize2 - getWidInc(mInd, cSize2) + 1, mYInc + 1), mBDRC
        UserControl.PSet (mX + bWid - cSize1 - cSize2 - getWidInc(mInd, cSize2) + 2, mYInc + 1), mBDRC
        UserControl.PSet (mX + bWid - cSize1 - cSize2 - getWidInc(mInd, cSize2) + 3, mYInc + 2), mBDRC
        UserControl.PSet (mX + bWid - cSize1 - cSize2 - getWidInc(mInd, cSize2) + 4, mYInc + 2), mBDRC
        '
        UserControl.Line (mX + bWid - cSize1 - getWidInc(mInd, cSize2) - cSize2, mYInc)-(mX + 1, mYInc), mBDRC
        UserControl.Circle (mX + cSize3, mYInc + cSize3), cSize3, mBDRC, 3.14 / 2, 3.14
        '
        If mInd <> selIndex + 1 Or drawAll = True Then
            UserControl.Line (mX, mYInc + cSize3)-(mX, mYInc + mTabHeig), mBDRC
        End If
        'close tab area to fill
        If mInd <> selIndex And drawAll <> True And mInd <> tabCnt - 1 Then
            UserControl.Line (mX + bWid - getWidInc(mInd, cSize2) - cSize1, mYInc + cSize2 - 1)-(mX + bWid - getWidInc(mInd, cSize2) - cSize1, mYInc + mTabHeig), mBDRC
        Else
            'before fill hide all what don' need be visible
            For i = 0 To mTabHeig - cSize2
                UserControl.Line (mX + bWid - cSize1 - 1 - i, mYInc + mTabHeig - 1 - i)-(mX + 10, mYInc + mTabHeig - 1 - i), mBC
            Next i
        End If
        
        'fill
        ExtFloodFill UserControl.hdc, mX + 2, mYInc + 2, UserControl.Point(mX + 2, mYInc + 2), 1
        'draw shadow
        If mePTAppearance = ptAppearance3D Then
            'right
            If mInd = selIndex Or drawAll = True Or mInd = tabCnt - 1 Then
                UserControl.Line (mX + bWid - 1, mYInc + mTabHeig - 1)-(mX + bWid - cSize1 - 1, mYInc + mTabHeig - 1), controlShadC
                UserControl.Line (mX + bWid - cSize1 + 2 - 1, mYInc + mTabHeig)-(mX + bWid - cSize1 - getWidInc(mInd, cSize2 / 3) + 2, mYInc + cSize2 / 3 + 1), controlShadC
                If mInd <> selIndex Then UserControl.Line (mX + bWid, mTabHeig + mYInc)-(mX, mTabHeig + mYInc), controlBorderC
            End If
            'UserControl.PSet (mX + bWid - cSize1 - cSize2 - getWidInc(mInd, cSize2) + 1, mYInc + 2), controlShadC
            'UserControl.PSet (mX + bWid - cSize1 - cSize2 - getWidInc(mInd, cSize2) + 2, mYInc + 2), controlShadC
            'UserControl.PSet (mX + bWid - cSize1 - cSize2 - getWidInc(mInd, cSize2) + 3, mYInc + 3), controlShadC
            'UserControl.PSet (mX + bWid - cSize1 - cSize2 - getWidInc(mInd, cSize2) + 4, mYInc + 3), controlShadC
            'center
            UserControl.Line (mX + bWid - cSize1 - getWidInc(mInd, cSize2) - cSize2, mYInc + 1)-(mX + 1, mYInc + 1), controlHLightC
            'left
            UserControl.Circle (mX, mYInc + cSize3), cSize3 - 1, controlHLightC, 0, 3.14 / 2
            If mInd <> selIndex - 1 Or drawAll = True Then
                UserControl.Line (mX + 1, mYInc + cSize3)-(mX + 1, mYInc + mTabHeig), controlHLightC
            End If
        End If
        'redraw
        If mInd = selIndex Or drawAll = True Then UserControl.Line (mX + bWid, mYInc + mTabHeig - 1)-(mX + bWid - cSize1, mYInc + mTabHeig - 1), mBDRC
        'find font y-pos
        If mePTVertTextAlg = ptTop Then
            UserControl.CurrentY = 2 + mYInc
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 + mYInc
        Else
            UserControl.CurrentY = mTabHeig - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 + mYInc
        End If
    Else
        'draw right side
        If mInd = selIndex Or drawAll = True Then
            UserControl.Line (mX + bWid, ucSH - 1 - mYInc - mTabHeig + 1)-(mX + bWid - cSize1, ucSH - 1 - mYInc - mTabHeig + 1), mBDRC
            UserControl.Line (mX + bWid - cSize1 + 2, ucSH - 1 - mYInc - mTabHeig)-(mX + bWid - cSize1 - getWidInc(mInd, cSize2 / 3) + 2, ucSH - 1 - mYInc - cSize2 / 3), mBDRC
        End If
        'UserControl.Circle (mX + bWid - cSize1 - getWidInc(mInd, cSize2) - cSize2 + 1, mYInc + cSize2 + 1), cSize2, vbWhite, (60 / 180) * 3.14, 3.14 / 2
        UserControl.PSet (mX + bWid - cSize1 - cSize2 - getWidInc(mInd, cSize2) + 1, ucSH - 1 - mYInc - 1), mBDRC
        UserControl.PSet (mX + bWid - cSize1 - cSize2 - getWidInc(mInd, cSize2) + 2, ucSH - 1 - mYInc - 1), mBDRC
        UserControl.PSet (mX + bWid - cSize1 - cSize2 - getWidInc(mInd, cSize2) + 3, ucSH - 1 - mYInc - 2), mBDRC
        UserControl.PSet (mX + bWid - cSize1 - cSize2 - getWidInc(mInd, cSize2) + 4, ucSH - 1 - mYInc - 2), mBDRC
        '
        UserControl.Line (mX + bWid - cSize1 - getWidInc(mInd, cSize2) - cSize2, ucSH - 1 - mYInc)-(mX + 1, ucSH - 1 - mYInc), mBDRC
        UserControl.Circle (mX + cSize3, ucSH - 1 - mYInc - cSize3), cSize3, mBDRC, 3.14, 1.5 * 3.14
        '
        If mInd <> selIndex + 1 Or drawAll = True Then
            UserControl.Line (mX, ucSH - 1 - mYInc - cSize3)-(mX, ucSH - 1 - mYInc - mTabHeig), mBDRC
        End If
        'close tab area to fill
        If mInd <> selIndex And drawAll <> True Then
            UserControl.Line (mX + bWid - getWidInc(mInd, cSize2) - cSize1, ucSH - 1 - mYInc - cSize2 + 1)-(mX + bWid - getWidInc(mInd, cSize2) - cSize1, ucSH - 1 - mYInc - mTabHeig), mBDRC
        Else
            'before fill hide all what dont need be visible
            For i = 0 To mTabHeig - cSize2
                UserControl.Line (mX + bWid - cSize1 - 1 - i, ucSH - 1 - mYInc - mTabHeig + 1 + i)-(mX + 10, ucSH - 1 - mYInc - mTabHeig + 1 + i), mBC
            Next i
        End If
        'fill
        ExtFloodFill UserControl.hdc, mX + 2, ucSH - 1 - mYInc - 2, UserControl.Point(mX + 2, ucSH - 1 - mYInc - 2), 1
        'draw shadow
        If mePTAppearance = ptAppearance3D Then
            'right
            If mInd = selIndex Or drawAll = True Then
                UserControl.Line (mX + bWid - 1, ucSH - 1 - mYInc - mTabHeig + 1)-(mX + bWid - cSize1 - 1, ucSH - 1 - mYInc - mTabHeig + 1), controlShadC
                UserControl.Line (mX + bWid - cSize1 + 2 - 1, ucSH - 1 - mYInc - mTabHeig)-(mX + bWid - cSize1 - getWidInc(mInd, cSize2 / 3) + 3, ucSH - 1 - mYInc - cSize2 / 3 + 1 - 3), controlShadC
                If mInd <> selIndex Then UserControl.Line (mX + bWid, mTabHeig + mYInc)-(mX, mTabHeig + mYInc), controlBorderC
            End If
            UserControl.PSet (mX + bWid - cSize1 - cSize2 - getWidInc(mInd, cSize2) + 1, ucSH - 1 - mYInc - 2), controlShadC
            UserControl.PSet (mX + bWid - cSize1 - cSize2 - getWidInc(mInd, cSize2) + 2, ucSH - 1 - mYInc - 2), controlShadC
            UserControl.PSet (mX + bWid - cSize1 - cSize2 - getWidInc(mInd, cSize2) + 3, ucSH - 1 - mYInc - 3), controlShadC
            UserControl.PSet (mX + bWid - cSize1 - cSize2 - getWidInc(mInd, cSize2) + 4, ucSH - 1 - mYInc - 3), controlShadC
            'center
            UserControl.Line (mX + bWid - cSize1 - getWidInc(mInd, cSize2) - cSize2, ucSH - 1 - mYInc - 1)-(mX + 1, ucSH - 1 - mYInc - 1), controlShadC
            'left
            If mInd <> selIndex - 1 Or drawAll = True Then
                UserControl.Line (mX + 1, ucSH - 1 - mYInc - cSize3 + 1)-(mX + 1, ucSH - 1 - mYInc - mTabHeig), controlHLightC
            End If
        End If
        'find font y-pos
        If mePTVertTextAlg = ptBottom Then
            UserControl.CurrentY = ucSH - mTabHeig + 2 - mYInc '+ UserControl.TextHeight(lstCaptions.Item(mInd + 1)) + 5
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = ucSH - mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 - mYInc
        Else
            UserControl.CurrentY = ucSH - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 - mYInc
        End If
    End If
    'draw picture - if exist
    If mePTPicAlig = ptPicCenter Then
        putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig, mX + getWidInc(mInd) + (bWid - getWidInc(mInd)) / 2 - icWid / 2
        fntMovPC = 0
    ElseIf mePTPicAlig = ptPicLeftEdge Then
        putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig, mX + getWidInc(mInd)
        fntMovPC = 0
    ElseIf mePTPicAlig = ptPicLeftOfCapton Then
        If mePTHorTextAlg = ptLeft Then
            putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig, mX + getWidInc(mInd)
            fntMovPC = icWid + mSpacing
        ElseIf mePTHorTextAlg = ptCenter Then
            putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig, mX + getWidInc(mInd) + (bWid - getWidInc(mInd)) / 2 - UserControl.TextWidth(lstCaptions.Item(mInd + 1)) / 2 - icWid - mSpacing
            fntMovPC = 0 '(icWid + mSpacing) / 2
        ElseIf mePTHorTextAlg = ptRight Then
            putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig, mX + bWid - icWid - mSpacing * 2 - UserControl.TextWidth(lstCaptions.Item(mInd + 1))
            fntMovPC = 0 'icWid + mSpacing
        End If
    ElseIf mePTPicAlig = ptPicRightEdge Then
        putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig, mX + bWid - icWid - 5
        fntMovPC = 0
    ElseIf mePTPicAlig = ptPicRightOfCaption Then
        If mePTHorTextAlg = ptLeft Then
            putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig, mX + getWidInc(mInd) + UserControl.TextWidth(lstCaptions.Item(mInd + 1)) + mSpacing
            fntMovPC = 0 'icWid + mSpacing
        ElseIf mePTHorTextAlg = ptCenter Then
            putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig, mX + getWidInc(mInd) / 2 + (bWid - getWidInc(mInd)) / 2 + UserControl.TextWidth(lstCaptions.Item(mInd + 1)) / 2 + mSpacing '+ icWid / 2
            fntMovPC = -(icWid + mSpacing) / 2
        ElseIf mePTHorTextAlg = ptRight Then
            putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig, mX + bWid - icWid - mSpacing
            fntMovPC = icWid + mSpacing
        End If
    End If
    If tabIcons(mInd) Is Nothing Then fntMovPC = 0
    'find font x-pos
    If mePTHorTextAlg = ptCenter Then
        UserControl.CurrentX = mX + (bWid - getWidInc(mInd)) / 2 - UserControl.TextWidth(lstCaptions.Item(mInd + 1)) / 2 + fntMovPC
    ElseIf mePTHorTextAlg = ptLeft Then
        UserControl.CurrentX = mX + fntMovPC
    Else
        UserControl.CurrentX = mX + bWid - UserControl.TextWidth(lstCaptions.Item(mInd + 1)) - fntMovPC - mSpacing - getWidInc(mInd)
    End If
    'make som efect if this is active tab
    If mInd = selIndex And efectSel = True Then
        If mePTOrientation = ptOrientationTop Then
            UserControl.CurrentY = UserControl.CurrentY - 1
        Else
            UserControl.CurrentY = UserControl.CurrentY + 1
        End If
        UserControl.CurrentX = UserControl.CurrentX - 1
    End If
    'print caption
    UserControl.Print lstCaptions.Item(mInd + 1)
    'hide line at bottom of active tab
    hideBottom mInd, mX, mYInc, bWid, mTabHeig, mBC, controlHLightC, controlShadC
    'find position of next tab
    
    mX = mX + bWid
    If mInd <> tabCnt - 1 Then mX = mX - getWidInc(mInd)  '+ cSize2 - cSize1 + mWINC
    If drawAll = False Then mX = mX + cSize2 - cSize1 + mWINC
End Sub
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Sub drawSSTab(ByVal mInd As Integer)
'''''    Dim ucSW As Integer, ucSH As Integer, bWid As Integer
    Dim cutWid As Integer, cutHeig As Integer, mTabHeig As Integer, mYInc As Integer
'''''    Dim mBDRC As OLE_COLOR, mBC As OLE_COLOR, mShadC As OLE_COLOR, mFC As OLE_COLOR
    Dim fntMovPC As Byte, pcMov As Byte
    '
    Dim havePIC As Boolean
    'setValues mInd, mBDRC, mBC, mShadC
    'set control font
   ' Set UserControl.Font = mFNT
   'If UserControl.ForeColor <> mFC Then UserControl.ForeColor = mFC
    'ucSW = UserControl.ScaleWidth
    'ucSH = UserControl.ScaleHeight
    'find tab height
    mTabHeig = tabHeig
    mYInc = Abs(tabHeigActive - tabHeig)
    If mInd = selIndex Then
        mTabHeig = tabHeigActive
        If tabHeigActive > tabHeig Then mYInc = 0
    Else
        If tabHeig > tabHeigActive Then mYInc = 0
    End If
    '
    cutWid = 5 'mTabHeig / 3 / Sin(90 * 3.14 / 180)
    cutHeig = 5 'mTabHeig / 3
    'find tab width
    bWid = UserControl.TextWidth(lstCaptions.Item(mInd + 1)) + mSpacing * 2 + 10 + getWidInc(mInd)
    'if there is picture then increase tab wid
    havePIC = False
    If Not tabIcons(mInd) Is Nothing Then
        If mePTPicAlig <> ptPicCenter Then bWid = bWid + icWid + mSpacing '* 2
        havePIC = True
    End If
    lstSizes.Add bWid
    '
    If bWid + mX < 0 Or mX > ucSW Then
        mX = mX + bWid
        Exit Sub
    End If
    'draw tab
    If mePTOrientation = ptOrientationTop Then
        UserControl.Line (mX, cutHeig + mYInc)-(mX, mTabHeig + mYInc), mBDRC
        UserControl.Line (mX, cutHeig + mYInc)-(mX + cutWid, mYInc), mBDRC
        UserControl.Line (mX + cutWid, mYInc)-(mX + bWid - cutWid, mYInc), mBDRC
        UserControl.Line (mX + bWid - cutWid, mYInc)-(mX + bWid, mYInc + cutHeig), mBDRC
        UserControl.Line (mX + bWid, mYInc + cutHeig)-(mX + bWid, mTabHeig + mYInc), mBDRC
        'close tab area for fill
        UserControl.Line (mX, mTabHeig + mYInc)-(mX + bWid, mTabHeig + mYInc), mBDRC
        'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
        'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
        If mX < 0 And mX + bWid > 0 Then
            ExtFloodFill UserControl.hdc, mX + bWid - 2, mTabHeig - 2 + mYInc, UserControl.Point(mX + bWid - 2, mTabHeig - 2 + mYInc), 1
        ElseIf mX < ucSW Then
            ExtFloodFill UserControl.hdc, mX + 2, mTabHeig - 2 + mYInc, UserControl.Point(mX + 2, mTabHeig - 2 + mYInc), 1
        End If
        'draw shadow
        If mePTAppearance = ptAppearance3D Then
            'left
            UserControl.Line (mX + 1, cutHeig + mYInc)-(mX + 1, mTabHeig + mYInc + 1), controlHLightC
            UserControl.Line (mX + 1, cutHeig + mYInc)-(mX + cutWid + 1, mYInc), controlHLightC
            If mInd = selIndex Then
                UserControl.Line (mX + 2, cutHeig + mYInc)-(mX + 2, mTabHeig + mYInc + 1), controlHLightC
                UserControl.Line (mX + 2, cutHeig + mYInc)-(mX + cutWid + 2, mYInc), controlHLightC
                UserControl.Line (mX + 3, cutHeig + mYInc)-(mX + cutWid + 3, mYInc), controlHLightC
            End If
            'right
            UserControl.Line (mX + bWid - 1, cutHeig + mYInc)-(mX + bWid - 1, mTabHeig + mYInc + 1), controlShadC
            UserControl.Line (mX + bWid - 1, cutHeig + mYInc)-(mX + bWid - cutWid - 1, mYInc), controlShadC
            If mInd = selIndex Then
                UserControl.Line (mX + bWid - 2, cutHeig + mYInc)-(mX + bWid - 2, mTabHeig + mYInc + 2), controlShadC
                UserControl.Line (mX + bWid - 2, cutHeig + mYInc)-(mX + bWid - cutWid - 2, mYInc), controlShadC
                UserControl.Line (mX + bWid - 3, cutHeig + mYInc)-(mX + bWid - cutWid - 3, mYInc), controlShadC
            End If
            'center
            UserControl.Line (mX + cutWid, 1 + mYInc)-(mX + bWid - cutWid + 1, 1 + mYInc), controlHLightC
            If mInd = selIndex Then UserControl.Line (mX + cutWid, 2 + mYInc)-(mX + bWid - cutWid, 2 + mYInc), controlHLightC
        End If
        'redraw line for close tab area
        If mInd <> selIndex Then UserControl.Line (mX, mTabHeig + mYInc)-(mX + bWid, mTabHeig + mYInc), mBDRC
        '
        If mePTVertTextAlg = ptTop Then
            UserControl.CurrentY = 2 + mYInc 'UserControl.TextHeight(lstCaptions.Item(mInd + 1)) + 5
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 + mYInc
        Else
            UserControl.CurrentY = mTabHeig - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 + mYInc
        End If
    '>>bottom
    Else
        UserControl.Line (mX, ucSH - 1 - cutHeig - mYInc)-(mX, ucSH - 1 - mTabHeig - mYInc), mBDRC
        UserControl.Line (mX, ucSH - 1 - cutHeig - mYInc)-(mX + cutWid, ucSH - 1 - mYInc), mBDRC
        UserControl.Line (mX + cutWid, ucSH - 1 - mYInc)-(mX + bWid - cutWid, ucSH - 1 - mYInc), mBDRC
        UserControl.Line (mX + bWid - cutWid, ucSH - 1 - mYInc)-(mX + bWid, ucSH - 1 - mYInc - cutHeig), mBDRC
        UserControl.Line (mX + bWid, ucSH - 1 - mYInc - cutHeig)-(mX + bWid, ucSH - 1 - mTabHeig - mYInc), mBDRC
        'close tab area for fill
        UserControl.Line (mX, ucSH - 1 - mTabHeig - mYInc)-(mX + bWid, ucSH - 1 - mTabHeig - mYInc), mBDRC
        'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
        'If UserControl.FillColor <> mBC Then UserControl.FillColor = mBC
        If mX < 0 And mX + bWid > 0 Then
            ExtFloodFill UserControl.hdc, mX + bWid - 2, ucSH - 1 - mTabHeig + 2 - mYInc, UserControl.Point(mX + bWid - 2, ucSH - 1 - mTabHeig + 2 - mYInc), 1
        ElseIf mX < ucSW Then
            ExtFloodFill UserControl.hdc, mX + 2, ucSH - 1 - mTabHeig + 2 - mYInc, UserControl.Point(mX + 2, ucSH - 1 - mTabHeig + 2 - mYInc), 1
        End If
        'draw shadow
        If mePTAppearance = ptAppearance3D Then
            'left
            UserControl.Line (mX + 1, ucSH - 1 - cutHeig - mYInc)-(mX + 1, ucSH - 1 - mTabHeig - mYInc - 1), controlHLightC
            UserControl.Line (mX + 1, ucSH - 1 - cutHeig - mYInc)-(mX + cutWid + 1, ucSH - 1 - mYInc), controlHLightC
            If mInd = selIndex Then
                UserControl.Line (mX + 2, ucSH - 1 - cutHeig - mYInc)-(mX + 2, ucSH - 1 - mTabHeig - mYInc - 1), controlHLightC
                UserControl.Line (mX + 2, ucSH - 1 - cutHeig - mYInc)-(mX + cutWid + 2, ucSH - 1 - mYInc), controlHLightC
                UserControl.Line (mX + 3, ucSH - 1 - cutHeig - mYInc)-(mX + cutWid + 3, ucSH - 1 - mYInc), controlHLightC
            End If
            'right
            UserControl.Line (mX + bWid - 1, ucSH - 1 - cutHeig - mYInc)-(mX + bWid - 1, ucSH - 1 - mTabHeig - mYInc - 1), controlShadC
            UserControl.Line (mX + bWid - 1, ucSH - 1 - cutHeig - mYInc)-(mX + bWid - cutWid - 1, ucSH - 1 - mYInc), controlShadC
            If mInd = selIndex Then
                UserControl.Line (mX + bWid - 2, ucSH - 1 - cutHeig - mYInc)-(mX + bWid - 2, ucSH - 1 - mTabHeig - mYInc - 2), controlShadC
                UserControl.Line (mX + bWid - 2, ucSH - 1 - cutHeig - mYInc)-(mX + bWid - cutWid - 2, ucSH - 1 - mYInc), controlShadC
                UserControl.Line (mX + bWid - 3, ucSH - 1 - cutHeig - mYInc)-(mX + bWid - cutWid - 3, ucSH - 1 - mYInc), controlShadC
            End If
            'center
            UserControl.Line (mX + cutWid, ucSH - 1 - 1 - mYInc)-(mX + bWid - cutWid + 1, ucSH - 1 - 1 - mYInc), controlShadC
            If mInd = selIndex Then UserControl.Line (mX + cutWid, ucSH - 1 - 2 - mYInc)-(mX + bWid - cutWid, ucSH - 1 - 2 - mYInc), controlShadC
        End If
        'redraw line for close tab area
        If mInd <> selIndex Then UserControl.Line (mX, ucSH - 1 - mTabHeig - mYInc)-(mX + bWid, ucSH - 1 - mTabHeig - mYInc), mBDRC
        '
        If mePTVertTextAlg = ptBottom Then
            UserControl.CurrentY = ucSH - mTabHeig + 2 - mYInc '+ UserControl.TextHeight(lstCaptions.Item(mInd + 1)) + 5
        ElseIf mePTVertTextAlg = ptMiddle Then
            UserControl.CurrentY = ucSH - mTabHeig / 2 - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) / 2 - mYInc
        Else
            UserControl.CurrentY = ucSH - UserControl.TextHeight(lstCaptions.Item(mInd + 1)) - 2 - mYInc
        End If
    End If
    'draw picture if exist
    putPicture mInd, bWid, fntMovPC, pcMov, mYInc, mTabHeig, mePTPicAlig
    'find font x-pos
    UserControl.CurrentX = getFontX(bWid, mInd, havePIC, mePTPicAlig) 'getFontX(bWid, mInd)
    'if this is active tab then make some effect
    If mInd = selIndex And efectSel = True Then
        If mePTOrientation = ptOrientationTop Then
            UserControl.CurrentY = UserControl.CurrentY - 1
        Else
            UserControl.CurrentY = UserControl.CurrentY + 1
        End If
        UserControl.CurrentX = UserControl.CurrentX - 1
    End If
    'print caption
    UserControl.Print lstCaptions.Item(mInd + 1)
    'hide line at bottom of active tab
    If mInd = selIndex Then
        Dim shINC As Byte
        If mePTAppearance = ptAppearance3D Then shINC = 3 Else shINC = 0
        If mePTOrientation = ptOrientationTop Then
            UserControl.Line (mX + shINC, mTabHeig + mYInc)-(mX + bWid - shINC + 1, mTabHeig + mYInc), mBC
            UserControl.Line (mX + shINC, mTabHeig + mYInc + 1)-(mX + bWid - shINC + 1, mTabHeig + mYInc + 1), mBC
            UserControl.Line (mX + shINC, mTabHeig + mYInc + 2)-(mX + bWid - shINC + 1, mTabHeig + mYInc + 2), mBC
        Else
            UserControl.Line (mX + shINC, ucSH - mTabHeig - 1 - mYInc)-(mX + bWid - shINC + 1, ucSH - mTabHeig - 1 - mYInc), mBC
            UserControl.Line (mX + shINC, ucSH - mTabHeig - 1 - mYInc - 1)-(mX + bWid - shINC + 1, ucSH - mTabHeig - 1 - mYInc - 1), mBC
            UserControl.Line (mX + shINC, ucSH - mTabHeig - 1 - mYInc - 2)-(mX + bWid - shINC + 1, ucSH - mTabHeig - 1 - mYInc - 2), mBC
        End If
    End If
    'find position ef next tab
    mX = mX + bWid
End Sub
'*************************************************************************************
'*************************************************************************************
'*************************************************************************************
Private Sub putPicture(ByVal mInd As Integer, ByVal bWid As Integer, ByVal fntMovPC As Integer, ByVal pcMov As Byte, ByVal mYInc As Integer, ByVal mTabHeig As Integer, ByVal mePTPicAligS As ePTPicAlig, Optional myX As Long = 0)
    'draw picture
    If tabIcons(mInd) Is Nothing Then Exit Sub
    If myX = 0 Then
        If mePTOrientation = ptOrientationTop Then
            'If Not tabIcons(mInd) Is Nothing Then
                If mInd = selIndex And efectSel = True Then pcMov = 1 Else pcMov = 0
                UserControl.PaintPicture tabIcons(mInd), mX + getPicX(bWid, icWid, mInd, fntMovPC, mePTPicAligS) - pcMov, mTabHeig / 2 + mYInc - icHeig / 2 + 1 - pcMov, icWid, icHeig
            'End If
        Else
            'If Not tabIcons(mInd) Is Nothing Then
                If mInd = selIndex And efectSel = True Then pcMov = 1 Else pcMov = 0
                UserControl.PaintPicture tabIcons(mInd), mX + getPicX(bWid, icWid, mInd, fntMovPC, mePTPicAligS) - pcMov, UserControl.ScaleHeight - mTabHeig / 2 - mYInc - icHeig / 2 - 1 + pcMov, icWid, icHeig
            'End If
        End If
    Else
        'If Not tabIcons(mInd) Is Nothing Then
            If mInd = selIndex And efectSel = True Then pcMov = 1 Else pcMov = 0
            If mePTOrientation = ptOrientationTop Then
                UserControl.PaintPicture tabIcons(mInd), myX, mTabHeig / 2 + mYInc - icHeig / 2 + 1 - pcMov, icWid, icHeig
            Else
                UserControl.PaintPicture tabIcons(mInd), myX, UserControl.ScaleHeight - mTabHeig / 2 - mYInc - icHeig / 2 - 1 - pcMov, icWid, icHeig
            End If
        'End If
    End If
End Sub
'
'-----------------------------------------------------------
'-------------------------------------------------------------
'hides bottom of selected tab
Private Sub hideBottom(ByVal mInd As Integer, ByVal mStX As Integer, ByVal mYInc As Integer, ByVal mBWid As Integer, ByVal mTabHeig As Integer, ByVal mBC As OLE_COLOR, ByVal mShadLeft As OLE_COLOR, ByVal mShadRight As OLE_COLOR)
    'if this is active tab then hide line at bottom of tab
    If mInd = selIndex Then
        If mePTOrientation = ptOrientationTop Then
            UserControl.Line (mStX + 1, mTabHeig + mYInc)-(mStX + mBWid - 1, mTabHeig + mYInc + 1), mBC, BF
            If mePTAppearance = ptAppearance3D Then
                UserControl.Line (mStX + 1, mTabHeig + mYInc)-(mStX + 1, mTabHeig + mYInc + 1), mShadLeft
                UserControl.Line (mStX + mBWid - 1, mTabHeig + mYInc)-(mStX + mBWid - 1, mTabHeig + mYInc + 1), mShadRight
            End If
        Else
            UserControl.Line (mStX + 1, UserControl.ScaleHeight - mTabHeig - 1 - mYInc)-(mStX + mBWid - 1, UserControl.ScaleHeight - mTabHeig - 2 - mYInc), mBC, BF
            If mePTAppearance = ptAppearance3D Then
                UserControl.Line (mStX + 1, UserControl.ScaleHeight - 1 - mTabHeig - mYInc)-(mStX + 1, UserControl.ScaleHeight - 1 - mTabHeig - mYInc - 1), mShadLeft
                UserControl.Line (mStX + mBWid - 1, UserControl.ScaleHeight - 1 - mTabHeig - mYInc)-(mStX + mBWid - 1, UserControl.ScaleHeight - 1 - mTabHeig - mYInc - 1), mShadRight
            End If
        End If
    End If
End Sub
'
'*************************************************************************************
'*************************************************************************************
Private Sub setValues(ByVal mInd As Integer, ByRef mBdrC1 As OLE_COLOR, ByRef mBC1 As OLE_COLOR, ByRef mShadC1 As OLE_COLOR)
    Dim mRET As Long
    If mePTColorScheme = ptColorUser Or getStyleByIndex(mInd) = ptVerticalLine Then
        mBdrC1 = controlBorderC
        mShadC1 = controlShadC
        If tabEnabl(mInd) = True Then
            If mInd = selIndex Then
                mBC1 = BCActive
                UserControl.FillColor = BCActive
                UserControl.ForeColor = FCActive
                'UserControl.Font = FCActive
                If Not UserControl.Font Is mFontActive Then Set UserControl.Font = mFontActive
            ElseIf mInd = hoverIndex Then
                mBC1 = BCHover
                UserControl.FillColor = BCHover
                UserControl.ForeColor = FCHover
                If Not UserControl.Font Is mFontHover Then Set UserControl.Font = mFontHover
            Else
                mBC1 = BC
                UserControl.FillColor = BC
                UserControl.ForeColor = FC
                If Not UserControl.Font Is mFont Then Set UserControl.Font = mFont
            End If
        Else
            mBC1 = BCDisabled
            UserControl.FillColor = BCDisabled
            UserControl.ForeColor = FCDisabled
            If Not UserControl.Font Is mFontDisabled Then Set UserControl.Font = mFontDisabled
        End If
    ElseIf mePTColorScheme = ptColorNoteOne Then
        'load color scheme
        If currColScheme <> 1 Then loadColors 1
        Dim mColorInd As Byte
        If mInd <= 7 Then
            mColorInd = mInd
        Else
            Dim tmpInd As Integer
            tmpInd = mInd
            Do While tmpInd > 7
                tmpInd = tmpInd - 7 - 1
            Loop
            mColorInd = tmpInd
        End If
        If tabEnabl(mInd) = True Then
            If mInd = selIndex Then
                mShadC1 = vbWhite
                mBC1 = mConColor(mColorInd)
                UserControl.FillColor = mConColor(mColorInd)
                If getStyleByIndex(mInd) = ptSSTab Then mShadC1 = controlShadC
                If Not UserControl.Font Is mFontActive Then Set UserControl.Font = mFontActive
                'mBC1 = mConColor(mColorInd)
                mBdrC1 = &HA19D9D
                UserControl.ForeColor = vbBlack
            ElseIf mInd = hoverIndex Then
                mShadC1 = vbWhite
                If getStyleByIndex(mInd) = ptSSTab Then mShadC1 = controlShadC
                If Not UserControl.Font Is mFontHover Then Set UserControl.Font = mFontHover
                mBC1 = getColorMix(mConColor(mColorInd), vbWhite)
                UserControl.FillColor = mBC1
                'mBdrC1 = &HA19D9D
                UserControl.ForeColor = &H808080
            Else
                mBC1 = mConColor(mColorInd)
                UserControl.FillColor = mBC1
                mBdrC1 = &HA19D9D
                If mInd = selIndex Then mShadC1 = vbWhite Else mShadC1 = mConColor(mColorInd)
                UserControl.ForeColor = vbBlack
                If getStyleByIndex(mInd) = ptSSTab Then mShadC1 = controlShadC
                If Not UserControl.Font Is mFont Then Set UserControl.Font = mFont
            End If
        Else
            mBC1 = mConColor(mColorInd)
            UserControl.FillColor = mBC1
            mBdrC1 = &HA19D9D
            mShadC1 = mConColor(mColorInd)
            UserControl.ForeColor = &H808080
    
            If Not UserControl.Font Is mFontDisabled Then Set UserControl.Font = mFontDisabled
            If getStyleByIndex(mInd) = ptSSTab Then mShadC1 = controlShadC
        End If
    End If


End Sub

Private Function getWidInc(ByVal mInd As Integer, Optional heigDec As Byte = 0) As Single
    getWidInc = 0
    If getStyleByIndex(mInd) = ptDistorted Then
        If mInd = selIndex Then
            getWidInc = tabHeigActive * Cos(60 * 3.14 / 180) * 2
        Else
            getWidInc = tabHeig * Cos(60 * 3.14 / 180) * 2
        End If
    ElseIf getStyleByIndex(mInd) = ptProTab Then
        If mInd = selIndex Then
            getWidInc = tabHeigActive * Cos(60 * 3.14 / 180) * 2
        Else
            getWidInc = tabHeig * Cos(60 * 3.14 / 180) * 2
        End If
    ElseIf getStyleByIndex(mInd) = ptCoolLeft Or getStyleByIndex(mInd) = ptCoolRight Then
        If mInd = selIndex Then
            getWidInc = (tabHeigActive - heigDec) * Cos(60 * 3.14 / 180) * 2
        Else
            getWidInc = (tabHeig - heigDec) * Cos(60 * 3.14 / 180) * 2
        End If
    End If
End Function

Private Function getFontX(ByVal bWid As Integer, ByVal mInd As Integer, Optional havePC As Boolean = False, Optional mPicAlign1 As ePTPicAlig = 0) As Single
    If mePTHorTextAlg = ptLeft Then
        getFontX = mX + mSpacing
        If getStyleByIndex(mInd) <> ptVerticalLine Then getFontX = getFontX + getWidInc(mInd) / 2 'getXSpacing
        If mPicAlign1 = ptPicLeftOfCapton Then 'Or mPicAlign1 = ptPicLeftEdge Then
            If havePC = True Then getFontX = getFontX + icWid + mSpacing
        End If
        If getStyleByIndex(mInd) = ptProTab Then getFontX = getFontX + 5
    ElseIf mePTHorTextAlg = ptCenter Then
        getFontX = mX + bWid / 2 - UserControl.TextWidth(lstCaptions.Item(mInd + 1)) / 2 '+ getWidInc(mind) / 8
        If mPicAlign1 = ptPicLeftOfCapton Then 'Or mPicAlign1 = ptPicLeftEdge Then
            If havePC = True Then getFontX = getFontX + (icWid + mSpacing) / 2
        ElseIf mPicAlign1 = ptPicRightOfCaption Then 'mPicAlign1 = ptPicRightEdge Or
            If havePC = True Then getFontX = getFontX - (icWid + mSpacing) / 2
        End If
    Else
        getFontX = mX + bWid - UserControl.TextWidth(lstCaptions.Item(mInd + 1)) - 5 ' - getWidInc(mInd) / 2
        If getStyleByIndex(mInd) <> ptVerticalLine Then getFontX = getFontX - getWidInc(mInd) / 2
        If mPicAlign1 = ptPicRightOfCaption Then 'mPicAlign1 = ptPicRightEdge Or
            If havePC = True Then getFontX = getFontX - (icWid + mSpacing) '- getWidInc(mInd) / 2
        End If
        'If getStyleByIndex(mInd) = ptDistorted Then getFontX = getFontX - 5
    End If
End Function

Private Function getPicX(ByVal bWid As Integer, ByVal mIcWid As Integer, ByVal mInd As Integer, Optional fntMov As Integer = 0, Optional mPicAlign1 As ePTPicAlig = 0) As Integer
    If mPicAlign1 = ptPicLeftEdge Then
        getPicX = mSpacing + 5 + getWidInc(mInd) / 2
        'If getStyleByIndex(mInd) = ptDistorted Then getPicX = getPicX + 5
    ElseIf mPicAlign1 = ptPicRightEdge Then
        getPicX = bWid - mSpacing - mIcWid - 5 - getWidInc(mInd) / 2
    ElseIf mPicAlign1 = ptPicLeftOfCapton Then
        If mePTHorTextAlg = ptLeft Then
            getPicX = mSpacing + 5 + getWidInc(mInd) / 2 'getFontX(bWid, mInd, 0, mPicAlign1) - mX + getWidInc(mInd) / 2 '- mIcWid - mSpacing 'mSpacing + 5
        ElseIf mePTHorTextAlg = ptCenter Then
            getPicX = bWid / 2 - (UserControl.TextWidth(lstCaptions.Item(mInd + 1)) + mIcWid + mSpacing) / 2
        Else
            getPicX = getFontX(bWid, mInd, 0, mPicAlign1) - mX - mSpacing - icWid
        End If
    ElseIf mPicAlign1 = ptPicRightOfCaption Then
        If mePTHorTextAlg = ptLeft Then
            getPicX = getFontX(bWid, mInd, 0, mPicAlign1) - mX + mSpacing + UserControl.TextWidth(lstCaptions.Item(mInd + 1))
        ElseIf mePTHorTextAlg = ptCenter Then
            getPicX = bWid / 2 + (UserControl.TextWidth(lstCaptions.Item(mInd + 1)) - mIcWid + mSpacing) / 2
        Else
            getPicX = bWid - mSpacing - mIcWid - 5 - getWidInc(mInd) / 2 'bWid - mSpacing - mIcWid - 5 - getWidInc(mInd) / 2
        End If
        'If getStyleByIndex(mInd) = ptDistorted And mePTHorTextAlg = ptRight Then getPicX = getPicX - 5
    ElseIf mPicAlign1 = ptPicCenter Then
        getPicX = bWid / 2 - mIcWid / 2
    End If
End Function

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'isMDown = True
    RaiseEvent MouseDown(Button, Shift, X, Y)
    If Button = 1 Then
        Dim mTabHeig1 As Integer, tmpTab1 As Integer
        If tabHeig > tabHeigActive Then mTabHeig1 = tabHeig Else mTabHeig1 = tabHeigActive
        If getTabIndexByPos(X, Y) <> selIndex And getTabIndexByPos(X, Y) > -1 Then
            If getTabIndexByPos(X, Y) >= 0 Then
                If tabEnabl(getTabIndexByPos(X, Y)) = True Then
                    tmpTab1 = getTabIndexByPos(X, Y)
                    RaiseEvent BeforeTabChange(selIndex, tmpTab1)
                    handleControls selIndex, tmpTab1
                    tmpTab1 = selIndex
                    selIndex = getTabIndexByPos(X, Y)
                    'set tab position to see it whole
                    'MsgBox selIndex & ">=" & firstItemL & "+" & lstSizes.Count & "-" & 2
                    If firstItemL > selIndex Then
                        firstItemL = selIndex + 1
                        If firstItemL < 1 Then firstItemL = 1
                    ElseIf selIndex >= firstItemL + lstSizes.Count - 2 Or (getStyleByIndex(selIndex) = ptCoolRight And selIndex >= firstItemL + lstSizes.Count - 3) Then  'Or mX > UserControl.ScaleWidth Then
                        Dim isMov As Boolean
                        isMov = False
                        If selIndex <> tabCnt - 1 Then
                            Do While selIndex >= firstItemL + lstSizes.Count - 2
                                firstItemL = firstItemL + 1
                                isMov = True
                            Loop
                        End If
                        If mX > UserControl.ScaleWidth And isMov = False Then
                            firstItemL = firstItemL + 1
                        End If
                    End If
                    
                    If firstItemL < 1 Then firstItemL = 1
                    If mStartX > 0 Then mStartX = 0
                    'raiseevent
                    RaiseEvent TabClick(selIndex)
                    reDraw
                    RaiseEvent TabChange(tmpTab1, selIndex)
                End If
            End If
        ElseIf getTabIndexByPos(X, Y) = -2 Or getTabIndexByPos(X, Y) = -3 Then
            If downSlide <> Abs(getTabIndexByPos(X, Y)) - 1 Then
                downSlide = Abs(getTabIndexByPos(X, Y)) - 1
                tmrSlide.Interval = 50
                If downSlide = 1 Then
                    If scrLEnb = True Then
                        'mStartX = mStartX + 10
                        scrlSide = 1
                        reDraw
                        tmrSlide.Enabled = True
                    End If
                Else
                    If scrREnb = True Then
                        'mStartX = mStartX - 10
                        scrlSide = -1
                        reDraw
                        tmrSlide.Enabled = True
                    End If
                End If
            End If
        End If
    End If
    'isMDown = False
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Ambient.UserMode <> True Then Exit Sub
    RaiseEvent MouseMove(Button, Shift, X, Y)
    hoverSlide = 0
    'toolTip
    If getTabIndexByPos(X, Y) <> hoverIndex And getTabIndexByPos(X, Y) >= 0 Then
        mTT.Style = meTTStyle
        mTT.DelayTime = 300
        mTT.VisibleTime = 2000
        mTT.Icon = TTIcon_Info
        mTT.Title = lstCaptions.Item(getTabIndexByPos(X, Y) + 1)
        mTT.TipText = strToolTips(getTabIndexByPos(X, Y))
        mTT.PopupOnDemand = False
        mTT.CreateToolTip UserControl.hwnd
        'MsgBox strToolTips(getTabIndexByPos(X, Y))
    End If
    
    If getTabIndexByPos(X, Y) <> selIndex And getTabIndexByPos(X, Y) <> hoverIndex Then
        hoverIndex = getTabIndexByPos(X, Y)
        If hoverIndex >= 0 Then
            If tabEnabl(hoverIndex) = True Then reDraw
        ElseIf hoverIndex < 0 Then
            reDraw
        End If
    ElseIf getTabIndexByPos(X, Y) = selIndex And hoverIndex <> selIndex Then
        hoverIndex = selIndex
        reDraw
    ElseIf getTabIndexByPos(X, Y) = -2 Or getTabIndexByPos(X, Y) = -3 Then
        If hoverSlide <> Abs(getTabIndexByPos(X, Y)) - 1 Then
            hoverSlide = Abs(getTabIndexByPos(X, Y)) - 1
            reDraw 'True
        End If
    End If
    If hoverIndex = -1 Then mTT.Destroy
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
    If downSlide <> 0 Then
        RaiseEvent AfterScroll
    End If
    downSlide = 0
    tmrSlide.Enabled = False
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'MsgBox "READ"
    meTTStyle = PropBag.ReadProperty("ToolTipStyle", TTStyle_Balloon)
    
    mePTStylesNormal = PropBag.ReadProperty("StyleNormal", ptVerticalLine)
    mePTStylesActive = PropBag.ReadProperty("StyleActive", 0)
    mePTOrientation = PropBag.ReadProperty("Orientation", 0)
    mePTHorTextAlg = PropBag.ReadProperty("TextAlignHorizontal", ptCenter)
    mePTVertTextAlg = PropBag.ReadProperty("TextAlignVerticall", ptMiddle)
    mePTSlideButtStyle = PropBag.ReadProperty("ScrollStyle", ptFilledArrow)
    '
    mePTColorScheme = PropBag.ReadProperty("ColorScheme", 0)
    '
    mePTPicAlig = PropBag.ReadProperty("IconAlign", 0)
    '
    controlBC = PropBag.ReadProperty("BackColor", vbButtonFace)
    controlBorderC = PropBag.ReadProperty("BorderColor", &H808080)
    tabAreaBC = PropBag.ReadProperty("TabAreaBackColor", vbHighlight)
    controlHLightC = PropBag.ReadProperty("HighlightColor", vbWhite)
    controlShadC = PropBag.ReadProperty("ShadowColor", &H808080)
    '
    BC = PropBag.ReadProperty("TabColor", vbButtonFace)
    BCActive = PropBag.ReadProperty("TabColorActive", vbButtonFace)
    BCHover = PropBag.ReadProperty("TabColorHover", vbButtonFace)
    BCDisabled = PropBag.ReadProperty("TabColorDisabled", vbButtonFace)
    '
    FC = PropBag.ReadProperty("ForeColor", vbBlack)
    FCActive = PropBag.ReadProperty("ForeColorActive", vbBlack)
    FCHover = PropBag.ReadProperty("ForeColorHover", vbBlack)
    FCDisabled = PropBag.ReadProperty("ForeColorDisabled", vbBlack)
    '
    shadC = PropBag.ReadProperty("TabShadowColor", vbWhite)
    '
    Set mFont = PropBag.ReadProperty("Font", Ambient.Font)
    Set UserControl.Font = mFont
    Set mFontActive = PropBag.ReadProperty("FontActive", Ambient.Font)
    Set mFontHover = PropBag.ReadProperty("FontHover", Ambient.Font)
    Set mFontDisabled = PropBag.ReadProperty("FontDisabled", Ambient.Font)
    '
    icWid = PropBag.ReadProperty("IconWidth", 16)
    icHeig = PropBag.ReadProperty("IconHeight", 16)
    tabHeig = PropBag.ReadProperty("TabHeight", 18)
    tabHeigActive = PropBag.ReadProperty("TabHeightActive", 20)
    startX = PropBag.ReadProperty("LeftSpacing", 10)
    buttSpacing = PropBag.ReadProperty("TabSpacing", 0)
    mSpacing = PropBag.ReadProperty("FontSpacing", 5)
    selIndex = PropBag.ReadProperty("ActiveTab", 0)
    '
    'If selIndex > tabCnt - 1 Then selIndex = tabCnt - 1
    '
    enbl = PropBag.ReadProperty("Enabled", True)
    efectSel = PropBag.ReadProperty("EnableFontMoving", True)
    shSlideButtons = PropBag.ReadProperty("ShowScroll", False)
    '
    scrHover = PropBag.ReadProperty("ScrollHoverButton", True)
    mePTAppearance = PropBag.ReadProperty("Appearance", 1)
    '
    drawClArea = PropBag.ReadProperty("DrawClientArea", True)
    aSize = PropBag.ReadProperty("AutoSize", False)
    '
    tabCnt = PropBag.ReadProperty("TabCount", 4)
    '
    Set lstCaptions = Nothing
    ReDim tabIcons(tabCnt - 1)
    ReDim tabEnabl(tabCnt - 1)
    ReDim strToolTips(tabCnt - 1)
    ReDim tabTags(tabCnt - 1)
    ReDim ctlLst(tabCnt - 1)
    ReDim tabVis(tabCnt - 1)
    
    Dim i As Integer, z As Integer
    For i = 0 To tabCnt - 1
        lstCaptions.Add PropBag.ReadProperty("TabCaption" & i, "Tab " & i + 1)
        Set tabIcons(i) = PropBag.ReadProperty("TabIcon" & i, Nothing)
        tabEnabl(i) = PropBag.ReadProperty("TabEnabled" & i, True)
        strToolTips(i) = PropBag.ReadProperty("TabToolTip" & i, "")
        tabTags(i) = PropBag.ReadProperty("TabTag" & i, "")
        tabVis(i) = PropBag.ReadProperty("TabVisible" & i, True)
    Next i
    
    Dim mCCount As Integer
    
    For i = 0 To tabCnt - 1
        
        mCCount = PropBag.ReadProperty("containdecControlsCount" & i, 0)
        For z = 1 To mCCount
            If PropBag.ReadProperty("CtlID" & i & "*" & z, "") <> "" Then
                ctlLst(i).Add PropBag.ReadProperty("CtlID" & i & "*" & z, "")
                'If i = 0 Then MsgBox PropBag.ReadProperty("CtlID" & i & "*" & z, "")
            End If
        Next z
    Next i
    '
    'slide propertyes
    '   back
    hSlideBack = PropBag.ReadProperty("ScrollAreaTransparent", True)
    '   back color
    slideBC = PropBag.ReadProperty("ScrollBackColor", vbHighlight)
    slideBCHover = PropBag.ReadProperty("ScrollBackColorHover", vbHighlight)
    slideBCDown = PropBag.ReadProperty("ScrollBackColorDown", vbHighlight)
    slideBCDisabled = PropBag.ReadProperty("ScrollBackColorDisabled", vbHighlight)
    '   fill color
    slideFillC = PropBag.ReadProperty("ScrollFillColor", vbBlack)
    slideFillCHover = PropBag.ReadProperty("ScrollFillColorHover", vbBlack)
    slideFillCDown = PropBag.ReadProperty("ScrollFillColorDown", vbBlack)
    slideFillCDisabled = PropBag.ReadProperty("ScrollFillColorDisabled", &H808080)
    '
    slideBDRC = PropBag.ReadProperty("ScrollBorderColor", &H404040)
    slideBDRCDisabled = PropBag.ReadProperty("ScrollBorderColorDisabled", &H404040)
    slideHLC = PropBag.ReadProperty("ScrollHighlightColor", vbWhite)
    slideShadC = PropBag.ReadProperty("ScrollShadowColor", &H404040)
    'align
    mePTSlideAlign = PropBag.ReadProperty("ScrollAlign", 1)
    '
   ' setTabCount
    handleControls 0, selIndex
    'reDraw
    ActiveTab = selIndex
End Sub

Private Sub UserControl_Resize()
    ucSW = UserControl.ScaleWidth
    ucSH = UserControl.ScaleHeight
    RaiseEvent Resize
    reDraw
    'MsgBox "RES"
End Sub

Private Sub UserControl_Terminate()
    mTT.Destroy
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "ToolTipStyle", meTTStyle, TTStyle_Balloon
    
    PropBag.WriteProperty "StyleNormal", mePTStylesNormal, ptVerticalLine
    PropBag.WriteProperty "StyleActive", mePTStylesActive, 0
    PropBag.WriteProperty "Orientation", mePTOrientation, 0
    PropBag.WriteProperty "TextAlignHorizontal", mePTHorTextAlg, ptCenter
    PropBag.WriteProperty "TextAlignVerticall", mePTVertTextAlg, ptMiddle
    PropBag.WriteProperty "ScrollStyle", mePTSlideButtStyle, ptFilledArrow
    PropBag.WriteProperty "IconAlign", mePTPicAlig, 0
    '
    PropBag.WriteProperty "ColorScheme", mePTColorScheme, 0
    '
    PropBag.WriteProperty "BackColor", controlBC, vbButtonFace
    PropBag.WriteProperty "BorderColor", controlBorderC, &H808080
    PropBag.WriteProperty "TabAreaBackColor", tabAreaBC, vbHighlight
    
    PropBag.WriteProperty "HighlightColor", controlHLightC, vbWhite
    PropBag.WriteProperty "ShadowColor", controlShadC, &H808080
    '
    PropBag.WriteProperty "TabColor", BC, vbButtonFace
    PropBag.WriteProperty "TabColorActive", BCActive, vbButtonFace
    PropBag.WriteProperty "TabColorHover", BCHover, vbButtonFace
    PropBag.WriteProperty "TabColorDisabled", BCDisabled, vbButtonFace
    '
    PropBag.WriteProperty "ForeColor", FC, vbBlack
    PropBag.WriteProperty "ForeColorActive", FCActive, vbBlack
    PropBag.WriteProperty "ForeColorHover", FCHover, vbBlack
    PropBag.WriteProperty "ForeColorDisabled", FCDisabled, vbBlack
    '
'    '
    PropBag.WriteProperty "TabShadowColor", shadC, vbWhite
    '
    PropBag.WriteProperty "Font", mFont, Ambient.Font
    PropBag.WriteProperty "FontActive", mFontActive, Ambient.Font
    PropBag.WriteProperty "FontHover", mFontHover, Ambient.Font
    PropBag.WriteProperty "FontDisabled", mFontDisabled, Ambient.Font
    '
    PropBag.WriteProperty "IconWidth", icWid, 16
    PropBag.WriteProperty "IconHeight", icHeig, 16
    PropBag.WriteProperty "TabHeight", tabHeig, 18
    PropBag.WriteProperty "TabHeightActive", tabHeigActive, 20
    PropBag.WriteProperty "LeftSpacing", startX, 10
    PropBag.WriteProperty "TabSpacing", buttSpacing, 0
    PropBag.WriteProperty "FontSpacing", mSpacing, 5
    PropBag.WriteProperty "ActiveTab", selIndex, 0
    '
    
    PropBag.WriteProperty "Enabled", enbl, True
    PropBag.WriteProperty "EnableFontMoving", efectSel, True
    PropBag.WriteProperty "ShowScroll", shSlideButtons, False
    '
    PropBag.WriteProperty "ScrollHoverButton", scrHover, True
    PropBag.WriteProperty "Appearance", mePTAppearance, 1
    '
    PropBag.WriteProperty "DrawClientArea", drawClArea, True
    PropBag.WriteProperty "AutoSize", aSize, False
    '
    'slide propertyes
    '   back
    PropBag.WriteProperty "ScrollAreaTransparent", hSlideBack, True
    '   back color
    PropBag.WriteProperty "ScrollBackColor", slideBC, vbHighlight
    PropBag.WriteProperty "ScrollBackColorHover", slideBCHover, vbHighlight
    PropBag.WriteProperty "ScrollBackColorDown", slideBCDown, vbHighlight
    PropBag.WriteProperty "ScrollBackColorDisabled", slideBCDisabled, vbHighlight
    '   fill color
    PropBag.WriteProperty "ScrollFillColor", slideFillC, vbBlack
    PropBag.WriteProperty "ScrollFillColorHover", slideFillCHover, vbBlack
    PropBag.WriteProperty "ScrollFillColorDown", slideFillCDown, vbBlack
    PropBag.WriteProperty "ScrollFillColorDisabled", slideFillCDisabled, &H808080
    '   border
    PropBag.WriteProperty "ScrollBorderColor", slideBDRC, &H404040
    PropBag.WriteProperty "ScrollBorderColorDisabled", slideBDRCDisabled, &H404040
    PropBag.WriteProperty "ScrollHighlightColor", slideHLC, vbWhite
    PropBag.WriteProperty "ScrollShadowColor", slideShadC, &H404040
    'align
    PropBag.WriteProperty "ScrollAlign", mePTSlideAlign, 1
    '
    PropBag.WriteProperty "TabCount", tabCnt, 4
    '
    Dim i As Integer, z As Integer
    For i = 0 To tabCnt - 1
        PropBag.WriteProperty "TabCaption" & i, lstCaptions.Item(i + 1), "Tab " & i + 1
        PropBag.WriteProperty "TabIcon" & i, tabIcons(i), Nothing
        PropBag.WriteProperty "TabEnabled" & i, tabEnabl(i), True
        PropBag.WriteProperty "TabToolTip" & i, strToolTips(i), ""
        PropBag.WriteProperty "TabTag" & i, tabTags(i), ""
        PropBag.WriteProperty "TabVisible" & i, tabVis(i), True
    Next i
    '
    For i = 0 To tabCnt - 1
        PropBag.WriteProperty "containdecControlsCount" & i, ctlLst(i).Count
        For z = 1 To ctlLst(i).Count
            PropBag.WriteProperty "CtlID" & i & "*" & z, ctlLst(i).Item(z), ""
        Next z
    Next i
End Sub
'function return tab index by mouse position
Private Function getTabIndexByPos(ByVal mx1 As Long, ByVal mY As Integer) As Integer
    Dim i As Integer, mTabHeig As Integer, mYInc As Integer, mTabHeig1 As Integer, slideHeig As Integer
    Dim mStartX1 As Integer
    Dim startITM As Integer, endITM As Integer, stepITM As Integer
    getTabIndexByPos = -1
    On Error Resume Next
    'first check is slide buttons visible,
    '   maybe mouse is over one of slide buttons
    If shSlideButtons = True Then
        If tabHeig > tabHeigActive Then mTabHeig1 = tabHeig Else mTabHeig1 = tabHeigActive
        slideHeig = Int(mTabHeig1 / 1.5)
        If slideHeig Mod 2 <> 0 Then slideHeig = slideHeig + 1
        '
        'if mouse y is in range slide_button_Y - slide_button_Y+ slide_button_Height then
        '   check is x position inside slide button range
        If (mY >= mTabHeig1 / 2 - slideHeig / 2 And mY <= mTabHeig1 / 2 + slideHeig / 2 And mePTOrientation = ptOrientationTop) Or (mY <= UserControl.ScaleHeight - mTabHeig1 / 2 + slideHeig / 2 And mY >= UserControl.ScaleHeight - mTabHeig1 / 2 - slideHeig / 2 And mePTOrientation = ptOrientationBottom) Then
            'if align is left
            If mePTSlideAlign = ptScrollLeft Then
                If mx1 >= 2 And mx1 <= 2 + slideHeig Then
                    getTabIndexByPos = -2
                    Exit Function
                ElseIf mx1 >= 3 + slideHeig And mx1 <= 3 + slideHeig * 2 Then
                    getTabIndexByPos = -3
                    Exit Function
                End If
                
            'or right
            Else
                'Dim mStartX1 As Integer
                mStartX1 = UserControl.ScaleWidth - 2 - slideHeig * 2
                If mx1 >= mStartX1 And mx1 <= mStartX1 + slideHeig Then
                    getTabIndexByPos = -2
                    Exit Function
                ElseIf mx1 >= mStartX1 + slideHeig And mx1 >= mStartX1 - slideHeig * 2 Then
                    getTabIndexByPos = -3
                    Exit Function
                End If
            End If
        End If
    End If
    '
    'check ordering:
    '   for some styles order is from last because this function 'getTabIndexByPos' works
    '   with tab coordinates and sizes and some tabs is over/under other, so we must
    '   select correct index. I didn't explane it so good. Try comment this if and
    '   leave only block after else and select ptCoolRight style and you will understand
    If mePTStylesNormal = ptCoolRight Or mePTStylesActive = ptCoolRight Then
        startITM = lstPositions.Count 'UBound(tabVis) + 1 '
        endITM = 1
        stepITM = -1
    Else
        startITM = 1
        endITM = lstPositions.Count
        stepITM = 1
    End If
    'find item (tab that is under pointer)
    For i = startITM To endITM Step stepITM
        'check is position in current tab area
        'If tabVis(i - 1) <> False Then
            If mx1 >= lstPositions.Item(i) And mx1 <= lstPositions.Item(i) + lstSizes.Item(i) Then
                'find tab height
                mTabHeig = tabHeig
                mYInc = Abs(tabHeigActive - tabHeig)
                If i - 1 = selIndex Then
                    mTabHeig = tabHeigActive
                    If tabHeigActive > tabHeig Then mYInc = 0
                Else
                    If tabHeig > tabHeigActive Then mYInc = 0
                End If
                '
                If mePTOrientation = ptOrientationBottom Then
                    If mY > UserControl.ScaleHeight - mYInc Or mY < UserControl.ScaleHeight - mYInc - mTabHeig Then
                        Exit Function
                    Else
                        getTabIndexByPos = firstItemL + i - 1 - 1
                    End If
                Else
                    If mY < mYInc Or mY > mYInc + mTabHeig + 1 Then
                        Exit Function
                    Else
                        getTabIndexByPos = firstItemL + i - 1 - 1
                    End If
                End If
                Exit Function
            End If
        'End If
    Next i
End Function
'
Private Sub setTabCount()
    On Error Resume Next
    Dim i  As Integer, z As Integer
    'declare 'dinamic field' variables
    Dim tmpIcs() As StdPicture
    Dim tmpEnbl() As Boolean
    Dim tmpVis() As Boolean
    Dim tmpColl As New Collection
    Dim tmpToolTip() As String
    Dim tmptabTags() As Variant
    Dim tmpCTL() As New Collection
    
    'create temp data
    If lstCaptions.Count > 0 Then
        ReDim tmpIcs(lstCaptions.Count - 1)
        ReDim tmpEnbl(lstCaptions.Count - 1)
        ReDim tmpVis(lstCaptions.Count - 1)
        ReDim tmpToolTip(lstCaptions.Count - 1)
        ReDim tmptabTags(lstCaptions.Count - 1)
        ReDim tmpCTL(lstCaptions.Count - 1)
        
        For i = 0 To lstCaptions.Count - 1
            tmpColl.Add lstCaptions.Item(i + 1)
            Set tmpIcs(i) = tabIcons(i)
            tmpEnbl(i) = tabEnabl(i)
            tmpToolTip(i) = strToolTips(i)
            Set tabIcons(i) = Nothing
            tmptabTags(i) = tabTags(i)
            tmpVis(i) = tabVis(i)
            '
            For z = 1 To ctlLst(i).Count
                tmpCTL(i).Add ctlLst(i).Item(z)
            Next z
        Next i
    End If
    '
    'redefine original variables (set new field/matrix/vector/whatever lenght)
    Set lstCaptions = Nothing
    ReDim tabIcons(tabCnt - 1)
    ReDim tabEnabl(tabCnt - 1)
    ReDim strToolTips(tabCnt - 1)
    ReDim tabTags(tabCnt - 1)
    ReDim ctlLst(tabCnt - 1)
    ReDim tabVis(tabCnt - 1)
    
    For i = 0 To tabCnt - 1
        'first copy data in tmp variables
        If tmpColl.Count >= i + 1 Then
            lstCaptions.Add tmpColl.Item(i + 1)
            Set tabIcons(i) = tmpIcs(i)
            tabEnabl(i) = tmpEnbl(i)
            strToolTips(i) = tmpToolTip(i)
            tabTags(i) = tmptabTags(i)
            tabVis(i) = tmpVis(i)
            For z = 1 To tmpCTL(i).Count
                ctlLst(i).Add tmpCTL(i).Item(z)
            Next z
        'then create new tabls (is there is any)
        Else
            lstCaptions.Add "Tab " & i + 1
            'Set tabIcons(i) = Nothing
            tabEnabl(i) = True
            tabVis(i) = True
            strToolTips(i) = ""
            tabTags(i) = ""
        End If
    Next i
End Sub
'
'**********************************************************************
'********  C  O  L  E  C  T  I  O  N  S  ******************************
'**********************************************************************
'
'Private Sub clearCollection(ByRef lstCollection As Collection)
'    Do While lstCollection.Count > 0
'        lstCollection.remove 1
'    Loop
'End Sub
'
Private Function isInCollection(ByRef lstCollection As Collection, ByVal vData As Variant) As Boolean
    isInCollection = False
    If lstCollection.Count = 0 Then Exit Function
    Dim i As Integer
    For i = 1 To lstCollection.Count
        If lstCollection.Item(i) = vData Then
            isInCollection = True
            Exit Function
        End If
    Next i
End Function
'
Private Sub replaceDataInCollection(ByRef lstCollection As Collection, ByVal dataIndex As Integer, ByVal nData As Variant)
    If lstCollection.Count = 0 Then Exit Sub
    Dim i As Integer, tmpColl As New Collection
    For i = 1 To lstCollection.Count
        If i <> dataIndex Then
            tmpColl.Add lstCollection.Item(i)
        Else
            tmpColl.Add nData
        End If
    Next i
    '
    'MsgBox lstCollection.Count
    Set lstCollection = Nothing
    Set lstCollection = New Collection
    '
    For i = 1 To tmpColl.Count
        lstCollection.Add tmpColl.Item(i)
    Next i
End Sub
'
Private Sub removeDataFromCollection(ByRef lstCollection As Collection, ByVal vData As Variant)
    If lstCollection.Count = 0 Then Exit Sub
    Dim i As Integer
    For i = 1 To lstCollection.Count
        If lstCollection.Item(i) = vData Then
            lstCollection.Remove (i)
            Exit Sub
        End If
    Next i
End Sub
'
'get mix between two colors
Private Function getColorMix(ByVal mColor1 As Long, ByVal mColor2 As Long, Optional cDist As Single = 2) As Long
    On Error Resume Next
    Err.Clear
    Dim cR1 As Long, cG1 As Long, cB1 As Long
    Dim cR2 As Long, cG2 As Long, cB2 As Long

    cB1 = mColor1 \ 65536
    cG1 = (mColor1 - cB1 * 65536) \ 256
    cR1 = mColor1 - cB1 * 65536 - cG1 * 256
    '
    cB2 = mColor2 \ 65536
    cG2 = (mColor2 - cB2 * 65536) \ 256
    cR2 = mColor2 - cB2 * 65536 - cG2 * 256
    '
    getColorMix = RGB((cR1 + cR2) / cDist, (cG1 + cG2) / cDist, (cB1 + cB2) / cDist)
    If Err.Number <> 0 Then
        getColorMix = mColor1
    End If
End Function
'load color schemes
Private Sub loadColors(ByVal ColorScheme As Byte)
    'colors for NoteOne scheme
    If ColorScheme = 1 Then
        ReDim mConColor(7)
        mConColor(0) = &HE4A88A '
        mConColor(1) = &H75DBFF '
        mConColor(2) = &H9FCDBD '
        mConColor(3) = &H9F9EF0 '
        mConColor(4) = &HE1A6BA '
        mConColor(5) = &HB4BF9A '
        mConColor(6) = &H83B6F7 '
        mConColor(7) = &HC0ABD8 '
        currColScheme = 1
    'elseif...  you can add new color schemes
    End If
End Sub
'
Private Function getStyleByIndex(ByVal mInd As Integer) As ePTStyles
    If mInd = selIndex Then
        getStyleByIndex = mePTStylesActive
    Else
        getStyleByIndex = mePTStylesNormal
    End If
End Function
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'--------------------- P U B L I C ----------------------------
'-------------------------------------------------------------------
'-------------------------------------------------------------------
'add new tab
Public Function addTab(ByVal strTabCaption As String, ByVal strTabToolTip As String, _
                    Optional bTabEnabled As Boolean = True, Optional bTabVisible As Boolean = True, Optional imgTabIcon As StdPicture = Nothing)
    'increase tab count
    tabCnt = tabCnt + 1
    'set new tab count
    setTabCount
    'set values for new tab
    replaceDataInCollection lstCaptions, lstCaptions.Count, strTabCaption
    Set tabIcons(tabCnt - 1) = imgTabIcon
    tabEnabl(tabCnt - 1) = bTabEnabled
    strToolTips(tabCnt - 1) = strTabToolTip
    tabVis(tabCnt - 1) = bTabVisible
    '
    reDraw
End Function
'remove tab
Public Sub RemoveTab(ByVal mTabIndex As Integer)
    'On Error Resume Next
    If mTabIndex < 0 Or mTabIndex > tabCnt Then Exit Sub
    
    Dim i  As Integer, z As Integer, k As Integer
    Dim tmpIcs() As StdPicture
    Dim tmpEnbl() As Boolean
    Dim tmpVis() As Boolean
    Dim tmpColl As New Collection
    Dim tmpToolTip() As String
    Dim tmptabTags() As Variant
    Dim tmpCTL() As New Collection
    
    'first do something with controls for this tab
    If mTabIndex = selIndex Then
        Dim mCTL As Control
        For Each mCTL In UserControl.ContainedControls
            'if control is not container
            If Not TypeOf mCTL Is ContainerCTL Then
                If mCTL.top > -50000 Then
                    mCTL.top = mCTL.top - 10000
                    mCTL.Visible = False
                End If
            Else
                'if is then check is for all tabs
                If mCTL.ForAllTabs = False Then
                    If mCTL.top > -50000 Then
                        mCTL.top = mCTL.top - 10000
                        mCTL.Visible = False
                    End If
                End If
            End If
        Next
    End If
    '
    If tabCnt <= 1 Then Exit Sub
    If lstCaptions.Count > 0 Then
        ReDim tmpIcs(lstCaptions.Count - 1)
        ReDim tmpEnbl(lstCaptions.Count - 1)
        ReDim tmpVis(lstCaptions.Count - 1)
        ReDim tmpToolTip(lstCaptions.Count - 1)
        ReDim tmptabTags(lstCaptions.Count - 1)
        ReDim tmpCTL(lstCaptions.Count - 1)
        
        z = 0
        For i = 0 To lstCaptions.Count - 1
            If i <> mTabIndex Then
                'MsgBox
                tmpColl.Add lstCaptions.Item(i + 1)
                Set tmpIcs(z) = tabIcons(i)
                tmpEnbl(z) = tabEnabl(i)
                tmpVis(z) = tabVis(i)
                tmpToolTip(z) = strToolTips(i)
                tmptabTags(z) = tabTags(i)
                
                For k = 1 To ctlLst(i).Count
                    tmpCTL(z).Add ctlLst(i).Item(k)
                Next k
                z = z + 1
            End If
            Set tabIcons(i) = Nothing
        Next i
    End If
    '
    tabCnt = z
    Set lstCaptions = Nothing
    ReDim tabIcons(z - 1)
    ReDim tabEnabl(z - 1)
    ReDim tabVis(z - 1)
    ReDim strToolTips(z - 1)
    ReDim tabTags(z - 1)
    ReDim ctlLst(z - 1)
    
    For i = 0 To z - 1 'tabCnt - 1
        lstCaptions.Add tmpColl.Item(i + 1)
        Set tabIcons(i) = tmpIcs(i)
        tabEnabl(i) = tmpEnbl(i)
        tabVis(i) = tmpVis(i)
        strToolTips(i) = tmpToolTip(i)
        tabTags(i) = tmptabTags(i)
        '
        For k = 1 To tmpCTL(i).Count
            ctlLst(i).Add tmpCTL(i).Item(k)
            'MsgBox i & vbCrLf & tmpCTL(i).Item(k)
        Next k
    Next i
    If selIndex > tabCnt - 1 Then selIndex = tabCnt - 1
    handleControls selIndex, selIndex
    reDraw
End Sub
'swap two tabs
Public Sub SwapTabs(ByVal tabIndex1 As Integer, ByVal tabIndex2 As Integer)
    If tabIndex1 > tabCnt - 1 Or tabIndex1 < 0 Then
        Err.Raise 380
        Exit Sub
    End If
    If tabIndex2 > tabCnt - 1 Or tabIndex2 < 0 Then
        Err.Raise 380
        Exit Sub
    End If
    'tmp variables
    Dim caption1 As String, enbl1 As Boolean, vis1 As Boolean, tollTip1 As String, tag1 As Variant, ic1 As StdPicture
    Dim lstTMP1 As New Collection, lstTMP2 As New Collection, i As Integer
    caption1 = lstCaptions.Item(tabIndex1 + 1)
    enbl1 = tabEnabl(tabIndex1)
    vis1 = tabVis(tabIndex1)
    tag1 = tabTags(tabIndex1)
    tollTip1 = strToolTips(tabIndex1)
    If Not tabIcons(tabIndex1) Is Nothing Then Set ic1 = tabIcons(tabIndex1)
    'replace1
    replaceDataInCollection lstCaptions, tabIndex1 + 1, lstCaptions.Item(tabIndex2 + 1)
    tabEnabl(tabIndex1) = tabEnabl(tabIndex2)
    tabVis(tabIndex1) = tabVis(tabIndex2)
    tabTags(tabIndex1) = tabTags(tabIndex2)
    strToolTips(tabIndex1) = strToolTips(tabIndex2)
    If Not tabIcons(tabIndex2) Is Nothing Then Set tabIcons(tabIndex1) = tabIcons(tabIndex2)
    'replace2
    replaceDataInCollection lstCaptions, tabIndex2 + 1, caption1
    tabEnabl(tabIndex2) = enbl1
    tabVis(tabIndex2) = vis1
    tabTags(tabIndex2) = tag1
    strToolTips(tabIndex2) = tollTip1
    If Not ic1 Is Nothing Then Set tabIcons(tabIndex2) = ic1
    '
    For i = 1 To ctlLst(tabIndex1).Count
        lstTMP1.Add ctlLst(tabIndex1).Item(i)
    Next i
    For i = 1 To ctlLst(tabIndex2).Count
        lstTMP2.Add ctlLst(tabIndex2).Item(i)
    Next i
    
    Set ctlLst(tabIndex1) = Nothing
    Set ctlLst(tabIndex2) = Nothing
    
    For i = 1 To lstTMP1.Count
        ctlLst(tabIndex2).Add lstTMP1.Item(i)
    Next i
    For i = 1 To lstTMP2.Count
        ctlLst(tabIndex1).Add lstTMP2.Item(i)
    Next i
    reDraw
    handleControls selIndex, selIndex
End Sub
                    
'handle contained controls
Private Sub handleControls(ByVal lastIndex As Integer, ByVal nIndex As Integer)
    Dim i As Integer, z As Integer, mStr1 As String
    Dim mCTL As Control
    Dim lstRemove As New Collection, haveCTL As Boolean, tmpVis As Boolean
    '
    On Error Resume Next
'    clearCollection ctlLst(lastIndex)
    'hide controls
    For Each mCTL In UserControl.ContainedControls
        If mCTL.top > -50000 Then
            If isInCollection(ctlLst(lastIndex), pGetControlId(mCTL)) <> True Then
                'If Not TypeOf mCTL Is ContainerCTL Then
                    ctlLst(lastIndex).Add pGetControlId(mCTL)
                'Else
                    'If mCTL.ForAllTabs = False Then
                        ctlLst(lastIndex).Add pGetControlId(mCTL)
                    'End If
               ' End If
            End If
        End If
    Next
        
    haveCTL = False
    'find controls that we need to show
    For Each mCTL In UserControl.ContainedControls
        Err.Clear
        tmpVis = mCTL.Visible
       '
        If Err.Number = 0 Then
            If isInCollection(ctlLst(nIndex), pGetControlId(mCTL)) Then ' Or isInCollection(ctlLst, pGetControlId(mCTL) & "-1") Then
                mCTL.Visible = True
                If mCTL.top < -50000 Then mCTL.top = mCTL.top + 100000
            Else
                If mCTL.top > -50000 Then
                    'If Not TypeOf mCTL Is ContainerCTL Then
                        mCTL.Visible = False
                        mCTL.top = mCTL.top - 100000
'                    Else
'                        If mCTL.ForAllTabs = False Then
'                            mCTL.Visible = False
'                            mCTL.top = mCTL.top - 100000
'                        End If
'                    End If
                End If
            End If
        End If
        Err.Clear
    Next
    
    Set ctlLst(nIndex) = Nothing
    For Each mCTL In UserControl.ContainedControls
        If mCTL.top > -50000 Then ctlLst(nIndex).Add pGetControlId(mCTL)
    Next
End Sub

' Function returns control's name & control's index combination
Private Function pGetControlId(ByRef oCtl As Control) As String
  On Error Resume Next
  
  Static sCtlName As String
  Static iCtlIndex As Integer
  
  iCtlIndex = -1
  
  sCtlName = oCtl.Name
  iCtlIndex = oCtl.Index
  pGetControlId = sCtlName & IIf(iCtlIndex <> -1, iCtlIndex, "")
End Function

'hide all controls
'Private Sub setCTLsVisible(ByVal nVis As Boolean)
'    Dim mCTL As Control, tmpVis As Boolean
'    On Error Resume Next
'    For Each mCTL In UserControl.ContainedControls
'        Err.Clear
'        tmpVis = mCTL.Visible
'        If isInCollection(ctlLst(selIndex), pGetControlId(mCTL)) Then
'            mCTL.Visible = nVis
'        End If
'        Err.Clear
'    Next
'End Sub


Private Sub setContainerBackColor(ByVal nColor As OLE_COLOR)
    On Error GoTo errH
    Dim mCTL1 As Control, isContainer As Boolean
    For Each mCTL1 In UserControl.ContainedControls
        If TypeOf mCTL1 Is ContainerCTL Then
            If mCTL1.AutoBackColor = True Then
                'MsgBox "IDE"
                mCTL1.BackColor = nColor
               ' MsgBox mCTL1.ctlCOM
            End If
        End If
    Next
errH:
End Sub


Private Sub checkProperties()
    If mePTStylesNormal = ptDistorted Or mePTStylesActive = ptDistorted Or _
                mePTStylesNormal = ptDistortedMenu Or mePTStylesActive = ptDistortedMenu Then
        If tabHeig <> tabHeigActive Then tabHeig = tabHeigActive
    ElseIf mePTStylesNormal = ptCoolLeft Or mePTStylesNormal = ptCoolRight Then
        If mePTStylesNormal <> mePTStylesActive Then mePTStylesNormal = mePTStylesActive
        
    End If
End Sub

