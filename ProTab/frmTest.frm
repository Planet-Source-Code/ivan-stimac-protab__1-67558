VERSION 5.00
Begin VB.Form frmTest 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   6435
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   ScaleHeight     =   429
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   796
   StartUpPosition =   3  'Windows Default
   Begin ProTab.ProTabCTL ProTabCTL1 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   10610
      StyleNormal     =   5
      StyleActive     =   5
      ScrollStyle     =   0
      ColorScheme     =   1
      ShadowColor     =   526344
      ForeColorActive =   8421504
      ForeColorHover  =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontActive {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabHeightActive =   18
      FontSpacing     =   0
      ShowScroll      =   -1  'True
      TabCaption0     =   "Tab styles"
      TabCaption1     =   "Color schemes"
      TabCaption2     =   "Scroll styles"
      TabCaption3     =   "About"
      containdecControlsCount0=   0
      containdecControlsCount1=   5
      CtlID1*1        =   "cmbColor"
      CtlID1*2        =   "cmbAppearance"
      CtlID1*3        =   "Label13"
      CtlID1*4        =   "Label12"
      CtlID1*5        =   "lblColorSchemes"
      containdecControlsCount2=   5
      CtlID2*1        =   "cmbScrAlig"
      CtlID2*2        =   "cmbScrollStyle"
      CtlID2*3        =   "Label15"
      CtlID2*4        =   "Label14"
      CtlID2*5        =   "lblScrollStyles"
      containdecControlsCount3=   1
      CtlID3*1        =   "lblAbout"
      Begin ProTab.ProTabCTL ProTabCTL2 
         Height          =   3375
         Left            =   240
         TabIndex        =   17
         Top             =   2280
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   5953
         StyleNormal     =   11
         StyleActive     =   11
         ScrollStyle     =   0
         IconAlign       =   3
         BackColor       =   16777215
         ShadowColor     =   16777215
         TabColor        =   16777215
         TabColorActive  =   16777215
         TabColorHover   =   16777215
         TabColorDisabled=   16777215
         ForeColorActive =   8421504
         ForeColorHover  =   4210752
         TabShadowColor  =   16576
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontActive {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabSpacing      =   3
         ActiveTab       =   1
         ShowScroll      =   -1  'True
         ScrollHoverButton=   0   'False
         TabCount        =   20
         TabCaption0     =   "Properties"
         TabIcon0        =   "frmTest.frx":0000
         TabCaption1     =   "Examples"
         TabIcon1        =   "frmTest.frx":031A
         TabIcon2        =   "frmTest.frx":0634
         containdecControlsCount0=   1
         CtlID0*1        =   "ContainerCTL1"
         containdecControlsCount1=   0
         containdecControlsCount2=   0
         containdecControlsCount3=   0
         containdecControlsCount4=   0
         containdecControlsCount5=   0
         containdecControlsCount6=   0
         containdecControlsCount7=   0
         containdecControlsCount8=   0
         containdecControlsCount9=   0
         containdecControlsCount10=   0
         containdecControlsCount11=   0
         containdecControlsCount12=   0
         containdecControlsCount13=   0
         containdecControlsCount14=   0
         containdecControlsCount15=   0
         containdecControlsCount16=   0
         containdecControlsCount17=   0
         containdecControlsCount18=   0
         containdecControlsCount19=   0
         Begin ProTab.ProTabCTL ProTabCTL5 
            Height          =   1215
            Index           =   1
            Left            =   7440
            TabIndex        =   26
            Top             =   1920
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   2143
            StyleNormal     =   6
            StyleActive     =   6
            ScrollStyle     =   0
            ShadowColor     =   526344
            ForeColorActive =   8421504
            ForeColorHover  =   4210752
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontActive {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabHeightActive =   18
            FontSpacing     =   0
            containdecControlsCount0=   0
            containdecControlsCount1=   0
            containdecControlsCount2=   0
            containdecControlsCount3=   0
         End
         Begin ProTab.ProTabCTL ProTabCTL4 
            Height          =   1215
            Index           =   1
            Left            =   3960
            TabIndex        =   25
            Top             =   1920
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   2143
            StyleNormal     =   9
            StyleActive     =   9
            ScrollStyle     =   0
            ShadowColor     =   526344
            TabColor        =   -2147483638
            ForeColorActive =   8421504
            ForeColorHover  =   4210752
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontActive {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabHeight       =   20
            FontSpacing     =   0
            containdecControlsCount0=   0
            containdecControlsCount1=   0
            containdecControlsCount2=   0
            containdecControlsCount3=   0
            Begin ProTab.ProTabCTL ProTabCTL6 
               Height          =   285
               Left            =   360
               TabIndex        =   28
               Top             =   840
               Width           =   2505
               _ExtentX        =   4419
               _ExtentY        =   503
               StyleNormal     =   9
               StyleActive     =   9
               ScrollStyle     =   0
               TabAreaBackColor=   -2147483633
               ShadowColor     =   526344
               TabColor        =   8438015
               TabColorActive  =   33023
               TabColorHover   =   12640511
               ForeColorActive =   8388608
               ForeColorHover  =   4210752
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   238
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty FontActive {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   238
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   238
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   238
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TabHeightActive =   18
               LeftSpacing     =   0
               FontSpacing     =   0
               EnableFontMoving=   0   'False
               DrawClientArea  =   0   'False
               AutoSize        =   -1  'True
               containdecControlsCount0=   0
               containdecControlsCount1=   0
               containdecControlsCount2=   0
               containdecControlsCount3=   0
            End
            Begin VB.Label Label6 
               Caption         =   "With autosize = true and drawClientArea = false"
               Height          =   495
               Left            =   360
               TabIndex        =   29
               Top             =   360
               Width           =   2655
            End
         End
         Begin ProTab.ProTabCTL ProTabCTL3 
            Height          =   1215
            Index           =   1
            Left            =   480
            TabIndex        =   24
            Top             =   1920
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   2143
            StyleNormal     =   3
            StyleActive     =   3
            ScrollStyle     =   0
            ShadowColor     =   526344
            ForeColorActive =   8421504
            ForeColorHover  =   4210752
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontActive {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontSpacing     =   0
            containdecControlsCount0=   0
            containdecControlsCount1=   0
            containdecControlsCount2=   0
            containdecControlsCount3=   0
         End
         Begin ProTab.ProTabCTL ProTabCTL5 
            Height          =   1215
            Index           =   0
            Left            =   7440
            TabIndex        =   23
            Top             =   600
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   2143
            StyleNormal     =   1
            StyleActive     =   1
            ScrollStyle     =   0
            ShadowColor     =   526344
            ForeColorActive =   8421504
            ForeColorHover  =   4210752
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontActive {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontSpacing     =   0
            containdecControlsCount0=   0
            containdecControlsCount1=   0
            containdecControlsCount2=   0
            containdecControlsCount3=   0
         End
         Begin ProTab.ProTabCTL ProTabCTL4 
            Height          =   1215
            Index           =   0
            Left            =   3960
            TabIndex        =   22
            Top             =   600
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   2143
            StyleNormal     =   14
            StyleActive     =   14
            ScrollStyle     =   0
            ShadowColor     =   526344
            TabColor        =   -2147483638
            ForeColorActive =   8421504
            ForeColorHover  =   4210752
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontActive {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabHeight       =   20
            FontSpacing     =   0
            containdecControlsCount0=   0
            containdecControlsCount1=   0
            containdecControlsCount2=   0
            containdecControlsCount3=   0
            Begin VB.Label Label4 
               Caption         =   "ProTab Style - dont use different tab height  propertie values"
               Height          =   495
               Left            =   240
               TabIndex        =   27
               Top             =   600
               Width           =   2895
            End
         End
         Begin ProTab.ProTabCTL ProTabCTL3 
            Height          =   1215
            Index           =   0
            Left            =   480
            TabIndex        =   21
            Top             =   600
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   2143
            ScrollStyle     =   0
            ColorScheme     =   1
            ShadowColor     =   526344
            ForeColorActive =   8421504
            ForeColorHover  =   4210752
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontActive {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontHover {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontDisabled {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontSpacing     =   0
            containdecControlsCount0=   0
            containdecControlsCount1=   0
            containdecControlsCount2=   0
            containdecControlsCount3=   0
         End
         Begin ProTab.ContainerCTL ContainerCTL1 
            Height          =   2535
            Left            =   120
            TabIndex        =   18
            Top             =   -99400
            Visible         =   0   'False
            Width           =   10935
            _ExtentX        =   19288
            _ExtentY        =   4471
            BackColor       =   16777215
            Begin VB.Label Label5 
               Caption         =   "Label5"
               Height          =   1815
               Left            =   480
               TabIndex        =   20
               Top             =   360
               Width           =   10215
            End
            Begin VB.Image Image3 
               Height          =   240
               Left            =   120
               Picture         =   "frmTest.frx":094E
               Top             =   120
               Width           =   240
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Understand some properties:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   238
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00996646&
               Height          =   255
               Left            =   420
               TabIndex        =   19
               Top             =   120
               Width           =   2715
            End
         End
      End
      Begin VB.ComboBox cmbScrAlig 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   -98440
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.ComboBox cmbScrollStyle 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   -99160
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.ComboBox cmbStyleActive 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   840
         Width           =   3735
      End
      Begin VB.ComboBox cmbStyleNormal 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   1560
         Width           =   3735
      End
      Begin VB.ComboBox cmbColor 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   -99160
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.ComboBox cmbAppearance 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   -98440
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Label lblAbout 
         BackStyle       =   0  'Transparent
         Caption         =   "##"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   4320
         TabIndex        =   16
         Top             =   -99160
         Visible         =   0   'False
         Width           =   6765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Scroll align:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   240
         TabIndex        =   15
         Top             =   -98680
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Scroll style:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   14
         Top             =   -99400
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblScrollStyles 
         BackStyle       =   0  'Transparent
         Caption         =   "##"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   4320
         TabIndex        =   13
         Top             =   -99160
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Appearance:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   -98680
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Color scheme:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   -99400
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Style active:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Style normal:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "##"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   4320
         TabIndex        =   6
         Top             =   840
         Width           =   6615
      End
      Begin VB.Label lblColorSchemes 
         BackStyle       =   0  'Transparent
         Caption         =   "##"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   4320
         TabIndex        =   3
         Top             =   -99160
         Visible         =   0   'False
         Width           =   6615
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private isLoaded As Boolean

Private Sub cmbAppearance_Click()
    Me.ProTabCTL1.Appearance = Me.cmbAppearance.ListIndex
End Sub

Private Sub cmbColor_Click()
    Me.ProTabCTL1.ColorScheme = Me.cmbColor.ListIndex
End Sub

Private Sub cmbScrAlig_Click()
    Me.ProTabCTL1.ScrollAlign = Me.cmbScrAlig.ListIndex
End Sub

Private Sub cmbScrollStyle_Click()
    Me.ProTabCTL1.ScrollStyle = Me.cmbScrollStyle.ListIndex
End Sub

Private Sub cmbStyleActive_Click()
    If isLoaded Then Me.ProTabCTL1.StyleActive = Me.cmbStyleActive.ListIndex
End Sub

Private Sub cmbStyleNormal_Click()
    If isLoaded Then Me.ProTabCTL1.StyleNormal = Me.cmbStyleNormal.ListIndex
End Sub

Private Sub Form_Load()
    Label3.Caption = "There is 15 professional styles. And that is not all. " & _
                    "You can choose one style for active and one for normal tab." & _
                    "I know what you think - crazy full spend days for develop " & _
                    "this control. And i can tell you I realy spend a days and weeks " & _
                    "if includes testing and debugs. " & vbCrLf & vbCrLf & _
                    "SO PLEASE IF YOU LIKE IT VISIT AGAIN OUR PSC AND RATE. " & _
                    "I DON 'T ASK YOU 5, JUST YOUR OPINION. " & vbCrLf & vbCrLf & _
                    "I don 't know englesh so good so sorry if this text is not correct."
                    
    Me.lblColorSchemes.Caption = "There is 2 color schemes. Note one and user. Note one colors " & _
                                 "each tab with different color. User color all tabs with user defined colors"
                                 
    Me.lblScrollStyles.Caption = "You can choose one of 3 scroll styles and by changing scroll color propertyes " & _
                                 "create cool scroll design"
    
    Me.lblAbout.Caption = "Name: ProTab" & vbCrLf & _
                        "Author: Ivan Stimac, Croatia" & vbCrLf & _
                        "Mail: ivan.stimac@po.htnet.hr" & vbCrLf & vbCrLf & _
                        "What more to say? Nothing, I just need to thank you for " & _
                        "downloading this code, learn more and be better programmer, " & _
                        "or learn english... heh. I don't think so..."
                        
    
    'fill text box
    
    Me.Label5.Caption = "AutoSize - if true set control width as width of all tabs" & vbCrLf & _
                        "FontSpacing - increase tab width for this walue at booth sides" & vbCrLf & _
                        "LeftSpacing - distance of left corner (or slide buttons if align is left) from first tab" & vbCrLf & _
                        "ScrollAlign - left or right - where scroll buttons will be" & vbCrLf & _
                        "ScrollAreaTransparent - if true then tab under scroll button will be visible" & vbCrLf & _
                        "ScrollHoverButton - scroll border will be draw only when hover and press scroll button" & vbCrLf & _
                        "TabSpacing - distance from tabs" & vbCrLf & _
                        "EnableFontMoving - if true then move tab caption to left and top of tab when it's active"
                        
                        
                        
                        
    cmbStyleNormal.Clear
    cmbStyleNormal.AddItem "ptRectangle"
    cmbStyleNormal.AddItem "ptRoundedRectangle"
    cmbStyleNormal.AddItem "ptRoundedRectangle2"
    cmbStyleNormal.AddItem "ptCornerCutLeft"
    cmbStyleNormal.AddItem "ptCornerCutRight"
    cmbStyleNormal.AddItem "ptCoolLeft"
    cmbStyleNormal.AddItem "ptCoolRight"
    cmbStyleNormal.AddItem "ptDistorted"
    cmbStyleNormal.AddItem "ptVerticalLine"
    cmbStyleNormal.AddItem "ptRoundMenu"
    cmbStyleNormal.AddItem "ptDistortedMenu"
    cmbStyleNormal.AddItem "ptXP"
    cmbStyleNormal.AddItem "ptSSTab"
    cmbStyleNormal.AddItem "ptFlatButton"
    cmbStyleNormal.AddItem "ptProTab"
    cmbStyleNormal.ListIndex = Me.ProTabCTL1.StyleNormal
    '
    cmbStyleActive.Clear
    cmbStyleActive.AddItem "ptRectangle"
    cmbStyleActive.AddItem "ptRoundedRectangle"
    cmbStyleActive.AddItem "ptRoundedRectangle2"
    cmbStyleActive.AddItem "ptCornerCutLeft"
    cmbStyleActive.AddItem "ptCornerCutRight"
    cmbStyleActive.AddItem "ptCoolLeft"
    cmbStyleActive.AddItem "ptCoolRight"
    cmbStyleActive.AddItem "ptDistorted"
    cmbStyleActive.AddItem "ptVerticalLine"
    cmbStyleActive.AddItem "ptRoundMenu"
    cmbStyleActive.AddItem "ptDistortedMenu"
    cmbStyleActive.AddItem "ptXP"
    cmbStyleActive.AddItem "ptSSTab"
    cmbStyleActive.AddItem "ptFlatButton"
    cmbStyleActive.AddItem "ptProTab"
    cmbStyleActive.ListIndex = Me.ProTabCTL1.StyleActive
    
    '
    cmbAppearance.AddItem "Flat"
    cmbAppearance.AddItem "3D"
    cmbAppearance.ListIndex = Me.ProTabCTL1.Appearance
    
    cmbColor.Clear
    cmbColor.AddItem "ptColorUser"
    cmbColor.AddItem "ptColorNoteOne"
    cmbColor.ListIndex = Me.ProTabCTL1.ColorScheme
    
    '
    cmbScrAlig.Clear
    cmbScrAlig.AddItem "ptScrollLeft"
    cmbScrAlig.AddItem "ptScrollRight"
    cmbScrAlig.ListIndex = Me.ProTabCTL1.ScrollAlign
    
    cmbScrollStyle.Clear
    cmbScrollStyle.AddItem "ptArrow"
    cmbScrollStyle.AddItem "ptTrinangle"
    cmbScrollStyle.AddItem "ptFilledArrow"
    cmbScrollStyle.ListIndex = Me.ProTabCTL1.ScrollStyle
    
    isLoaded = True
    
   ' me.ProTabCTL1.
    'Me.ProTabCTL2.ActiveTab = 3
    
End Sub

Private Sub ProTabCTL2_Click()
    'ProTabCTL2.SwapTabs
    'ProTabCTL2.TabVisible
End Sub

