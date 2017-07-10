VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{0C8DE9F2-EAFC-44DF-A13F-B5A9B36ED780}#2.0#0"; "lvButton.ocx"
Object = "{E0D30636-0F87-47D5-B501-08A4FFAC604E}#1.0#0"; "osenxpsuite2005.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "La Esperanza Hotel Booking And Reservation System"
   ClientHeight    =   9915
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   17100
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   661
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1140
   StartUpPosition =   2  'CenterScreen
   Begin La_Esperanza_Hotel_Reservation_And_Booki.b8Line b8Line2 
      Height          =   60
      Left            =   0
      TabIndex        =   17
      Top             =   9480
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   106
      BorderColor1    =   16761024
      BorderColor2    =   16744576
      BorderColor3    =   16744576
   End
   Begin La_Esperanza_Hotel_Reservation_And_Booki.b8SideBar b8SideBar1 
      Height          =   9135
      Left            =   0
      TabIndex        =   15
      Top             =   360
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   16113
      BackColor       =   16761024
      BackColor       =   16761024
      BorderColor1    =   16777152
      Begin La_Esperanza_Hotel_Reservation_And_Booki.b8Line b8Line8 
         Height          =   60
         Left            =   -1560
         TabIndex        =   27
         Top             =   6960
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   106
         BorderColor1    =   16744576
         BorderColor2    =   16777215
         BorderColor3    =   16777215
      End
      Begin lvButton.lvButtons_H lvButtons_H4 
         Height          =   615
         Left            =   0
         TabIndex        =   26
         Top             =   6000
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "Change Password"
         CapAlign        =   2
         BackStyle       =   5
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777152
         LockHover       =   1
         cGradient       =   0
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":0000
         ImgSize         =   32
         cBack           =   12648447
      End
      Begin La_Esperanza_Hotel_Reservation_And_Booki.b8Line b8Line6 
         Height          =   60
         Left            =   -1560
         TabIndex        =   25
         Top             =   5640
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   106
         BorderColor1    =   16744576
         BorderColor2    =   16777215
         BorderColor3    =   16777215
      End
      Begin lvButton.lvButtons_H lvButtons_H3 
         Height          =   615
         Left            =   0
         TabIndex        =   24
         Top             =   4680
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "Change Password"
         CapAlign        =   2
         BackStyle       =   5
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777152
         LockHover       =   1
         cGradient       =   0
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":041D
         ImgSize         =   32
         cBack           =   12648447
      End
      Begin La_Esperanza_Hotel_Reservation_And_Booki.b8Line b8Line5 
         Height          =   60
         Left            =   -1560
         TabIndex        =   23
         Top             =   4320
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   106
         BorderColor1    =   16744576
         BorderColor2    =   16777215
         BorderColor3    =   16777215
      End
      Begin La_Esperanza_Hotel_Reservation_And_Booki.b8Line b8Line4 
         Height          =   60
         Left            =   -1560
         TabIndex        =   22
         Top             =   3000
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   106
         BorderColor1    =   16744576
         BorderColor2    =   16777215
         BorderColor3    =   16777215
      End
      Begin La_Esperanza_Hotel_Reservation_And_Booki.b8Line b8Line3 
         Height          =   60
         Left            =   -720
         TabIndex        =   21
         Top             =   1680
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   106
         BorderColor1    =   16744576
         BorderColor2    =   16777215
         BorderColor3    =   16777215
      End
      Begin La_Esperanza_Hotel_Reservation_And_Booki.b8SideTab b8SideTab2 
         Height          =   375
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         Caption         =   "Quick Tools"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   16777215
      End
      Begin lvButton.lvButtons_H cmdChangePass 
         Height          =   615
         Left            =   0
         TabIndex        =   19
         Top             =   3360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "Change Password"
         CapAlign        =   2
         BackStyle       =   5
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777152
         LockHover       =   1
         cGradient       =   0
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":083A
         ImgSize         =   32
         cBack           =   12648447
      End
      Begin lvButton.lvButtons_H lvButtons_H2 
         Height          =   615
         Left            =   0
         TabIndex        =   18
         Top             =   2040
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "User List"
         CapAlign        =   2
         BackStyle       =   5
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777152
         LockHover       =   1
         cGradient       =   0
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":0C57
         ImgSize         =   32
         cBack           =   12648447
      End
      Begin lvButton.lvButtons_H lvButtons_H1 
         Height          =   615
         Left            =   0
         TabIndex        =   16
         Top             =   840
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "Guest List"
         CapAlign        =   2
         BackStyle       =   5
         Shape           =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         cBhover         =   16777152
         LockHover       =   1
         cGradient       =   0
         CapStyle        =   2
         Mode            =   0
         Value           =   0   'False
         Image           =   "frmMain.frx":10A8
         ImgSize         =   32
         cBack           =   12648447
      End
   End
   Begin XPControls.XPFrame XPFrame1 
      Height          =   2295
      Left            =   2160
      TabIndex        =   8
      Top             =   360
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   4048
      BackColor       =   16761024
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin La_Esperanza_Hotel_Reservation_And_Booki.b8SideBar b8SideBar2 
         Height          =   8775
         Left            =   0
         TabIndex        =   11
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   15478
         BackColor       =   16761024
         BackColor       =   16761024
         Begin XPControls.XPButton cmdView 
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   3120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            Caption         =   "View All"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin La_Esperanza_Hotel_Reservation_And_Booki.b8SideTab b8SideTab1 
            Height          =   2655
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   4683
            Caption         =   "Legend"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontBold        =   -1  'True
            FontName        =   "Courier New"
            FontSize        =   9.75
            ForeColor       =   16777215
            Begin MSComctlLib.ListView ListView2 
               Height          =   2055
               Left            =   120
               TabIndex        =   13
               Top             =   480
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   3625
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               _Version        =   393217
               Icons           =   "imglstIcons1"
               SmallIcons      =   "imglstIcons1"
               ColHdrIcons     =   "imglstIcons1"
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   0
            End
         End
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   8535
         Left            =   2040
         TabIndex        =   10
         Top             =   480
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   15055
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         Icons           =   "imglstIcons1"
         SmallIcons      =   "imglstIcons1"
         ColHdrIcons     =   "imglstIcons1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin La_Esperanza_Hotel_Reservation_And_Booki.b8TitleBar b8TitleBar1 
         Height          =   345
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   14920
         _ExtentX        =   26326
         _ExtentY        =   609
         Caption         =   "Room Status"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontBold        =   -1  'True
         FontName        =   "Courier New"
         FontSize        =   9.75
         ForeColor       =   0
         Icon            =   "frmMain.frx":1557
      End
   End
   Begin La_Esperanza_Hotel_Reservation_And_Booki.b8Line b8Line1 
      Height          =   60
      Left            =   0
      TabIndex        =   7
      Top             =   360
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   106
      BorderColor1    =   16761024
      BorderColor2    =   16744576
      BorderColor3    =   16744576
   End
   Begin osenxpsuite2005.OsenXPStatusBar XPStatBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   9165
      Width           =   17100
      _ExtentX        =   30163
      _ExtentY        =   661
      BackColor       =   14936810
      ForeColor       =   -2147483630
      ForeColorDissabled=   16777215
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowGripper     =   -1  'True
      ShowSeperators  =   -1  'True
      NumberOfPanels  =   9
      PWidth1         =   100
      PMinWidth1      =   0
      pTTText1        =   ""
      pType1          =   0
      pText1          =   "User Name:"
      pTextAlignment1 =   0
      PanelPicture1   =   "frmMain.frx":1573
      PanelPicAlignment1=   0
      pBckgColor1     =   0
      pGradient1      =   0
      pEdgeSpacing1   =   0
      pEdgeInner1     =   0
      pEdgeOuter1     =   0
      PWidth2         =   70
      PMinWidth2      =   0
      pTTText2        =   "Current User"
      pType2          =   0
      pText2          =   ""
      pTextAlignment2 =   0
      PanelPicture2   =   "frmMain.frx":18C5
      PanelPicAlignment2=   0
      pBckgColor2     =   0
      pGradient2      =   0
      pEdgeSpacing2   =   0
      pEdgeInner2     =   0
      pEdgeOuter2     =   0
      PWidth3         =   120
      PMinWidth3      =   0
      pTTText3        =   ""
      pType3          =   0
      pText3          =   "Account Type:"
      pTextAlignment3 =   0
      PanelPicture3   =   "frmMain.frx":18E1
      PanelPicAlignment3=   0
      pBckgColor3     =   0
      pGradient3      =   0
      pEdgeSpacing3   =   0
      pEdgeInner3     =   0
      pEdgeOuter3     =   0
      PWidth4         =   100
      PMinWidth4      =   0
      pTTText4        =   "Account Type"
      pType4          =   0
      pText4          =   ""
      pTextAlignment4 =   0
      PanelPicture4   =   "frmMain.frx":1C33
      PanelPicAlignment4=   0
      pBckgColor4     =   0
      pGradient4      =   0
      pEdgeSpacing4   =   0
      pEdgeInner4     =   0
      pEdgeOuter4     =   0
      PWidth5         =   400
      PMinWidth5      =   0
      pTTText5        =   ""
      pType5          =   0
      pText5          =   ""
      pTextAlignment5 =   0
      PanelPicture5   =   "frmMain.frx":1C4F
      PanelPicAlignment5=   0
      pBckgColor5     =   0
      pGradient5      =   0
      pEdgeSpacing5   =   0
      pEdgeInner5     =   0
      pEdgeOuter5     =   0
      PWidth6         =   70
      PMinWidth6      =   0
      pTTText6        =   "Capslock Indicator"
      pType6          =   5
      pText6          =   "CAPS"
      pTextAlignment6 =   0
      PanelPicture6   =   "frmMain.frx":1C6B
      PanelPicAlignment6=   0
      pBckgColor6     =   0
      pGradient6      =   0
      pEdgeSpacing6   =   0
      pEdgeInner6     =   0
      pEdgeOuter6     =   0
      PWidth7         =   50
      PMinWidth7      =   0
      pTTText7        =   "Numlock Indicator"
      pType7          =   6
      pText7          =   "NUM"
      pTextAlignment7 =   0
      PanelPicture7   =   "frmMain.frx":1C87
      PanelPicAlignment7=   0
      pBckgColor7     =   0
      pGradient7      =   0
      pEdgeSpacing7   =   0
      pEdgeInner7     =   0
      pEdgeOuter7     =   0
      PWidth8         =   100
      PMinWidth8      =   0
      pTTText8        =   "Current Time"
      pType8          =   0
      pText8          =   ""
      pTextAlignment8 =   0
      PanelPicture8   =   "frmMain.frx":1CA3
      PanelPicAlignment8=   0
      pBckgColor8     =   0
      pGradient8      =   0
      pEdgeSpacing8   =   0
      pEdgeInner8     =   0
      pEdgeOuter8     =   0
      PWidth9         =   100
      PMinWidth9      =   0
      pTTText9        =   "Current Date"
      pType9          =   0
      pText9          =   ""
      pTextAlignment9 =   0
      PanelPicture9   =   "frmMain.frx":1FF5
      PanelPicAlignment9=   0
      pBckgColor9     =   0
      pGradient9      =   0
      pEdgeSpacing9   =   0
      pEdgeInner9     =   0
      pEdgeOuter9     =   0
      DrawMode        =   1
      HaveXPForm      =   -1  'True
   End
   Begin osenxpsuite2005.OsenXPToolBar OsenXPToolBar1 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   17100
      _ExtentX        =   30163
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowEndPanel    =   -1  'True
      XPBlend         =   0   'False
      TotalButton     =   9
      Bpic1           =   "frmMain.frx":2347
      Bname1          =   "Room Status"
      BSCap1          =   -1  'True
      Btype1          =   0
      Bwidth1         =   0
      Bchecked1       =   0   'False
      Bvalue1         =   0   'False
      Bname2          =   "Button2"
      BSCap2          =   -1  'True
      Btype2          =   2
      Bwidth2         =   0
      Bchecked2       =   0   'False
      Bvalue2         =   0   'False
      Bpic3           =   "frmMain.frx":2699
      Bname3          =   "Check-In"
      BSCap3          =   -1  'True
      Btype3          =   0
      Bwidth3         =   0
      Bchecked3       =   0   'False
      Bvalue3         =   0   'False
      Bname4          =   "Button4"
      Btype4          =   2
      Bwidth4         =   0
      Bchecked4       =   0   'False
      Bvalue4         =   0   'False
      Bpic5           =   "frmMain.frx":29EB
      Bname5          =   "Check-Out"
      BSCap5          =   -1  'True
      Btype5          =   0
      Bwidth5         =   0
      Bchecked5       =   0   'False
      Bvalue5         =   0   'False
      Bname6          =   "Button6"
      Btype6          =   2
      Bwidth6         =   0
      Bchecked6       =   0   'False
      Bvalue6         =   0   'False
      Bpic7           =   "frmMain.frx":2D3D
      Bname7          =   "Bookings and Reservations"
      BSCap7          =   -1  'True
      Btype7          =   0
      Bwidth7         =   0
      Bchecked7       =   0   'False
      Bvalue7         =   0   'False
      Bname8          =   "Button8"
      Btype8          =   2
      Bwidth8         =   0
      Bchecked8       =   0   'False
      Bvalue8         =   0   'False
      Bpic9           =   "frmMain.frx":308F
      Bname9          =   "Cancel Booking or Reservation"
      BSCap9          =   -1  'True
      Btype9          =   0
      Bwidth9         =   0
      Bchecked9       =   0   'False
      Bvalue9         =   0   'False
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   840
      Left            =   10560
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1482
      ButtonWidth     =   1799
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "imglstIcons1"
      HotImageList    =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Room Status"
            Key             =   "RoomStatus"
            Object.ToolTipText     =   "View All Rooms Status"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Check-In"
            Key             =   "CheckIn"
            Object.ToolTipText     =   "Check-In Guest"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Check-Out"
            Key             =   "CheckOut"
            Object.ToolTipText     =   "Check-Out Guest"
            ImageIndex      =   13
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtUserID 
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   7560
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   7560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   7560
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imgListMain 
      Left            =   -120
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":33E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3833
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C85
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":40D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4529
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   7680
   End
   Begin MSComctlLib.StatusBar statBarMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9540
      Visible         =   0   'False
      Width           =   17100
      _ExtentX        =   30163
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin MSComctlLib.ImageList imglRunSearch 
      Left            =   720
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   61
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":497B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4C95
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4FAF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7761
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":946B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9785
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":98DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9A39
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9B93
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9EAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":ACFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BB51
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C9FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D84D
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FFFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10319
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1116B
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11485
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12ADF
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12DF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13113
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1342D
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13747
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14021
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":148FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1574D
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1659F
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":173F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17CCB
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18B1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A827
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AB41
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B993
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E145
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":208F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22601
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24DB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26ABD
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27397
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":281E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29843
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A11D
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A9F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B849
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CEA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DCF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E00F
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E8E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2EC03
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2EF1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F237
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F391
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F6AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F9C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2FCDF
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2FFF9
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30313
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3062D
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30787
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":308E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30A3B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTBList 
      Left            =   1560
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30B95
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3146F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31D49
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":32623
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":32EFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":337D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":340B1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstIcons1 
      Left            =   2400
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3498B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":351DF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5EE01
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":88A23
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":88FBD
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8EBDF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B8801
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BFD03
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EA9DD
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EB2B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EBB91
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EC46B
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EC5F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":116219
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3240
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13FE3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":140717
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":140FF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1418CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1421AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":142A87
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":143363
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":143C3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14451B
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1453F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1462CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1471A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":148083
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":148F5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":149E37
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   4080
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14AD11
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14B723
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14C135
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14C4CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14C869
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14CC03
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14CF9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14D9AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14E3C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14EDD3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14F7E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1501F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":150C09
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15161B
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":151BB7
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16g 
      Left            =   4920
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":152153
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1526ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":152C87
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":153021
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1533BB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":153755
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList itb32x32 
      Left            =   5760
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":153AEF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":155481
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":156E13
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1587A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15A137
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15BAC9
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15D45B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15EDED
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16077F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":162113
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":162DEF
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1636CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1643AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":165087
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":165D63
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":166A3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16771B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   1800
      Left            =   6000
      Picture         =   "frmMain.frx":167FF7
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   1320
   End
   Begin VB.Menu mnuUserManage 
      Caption         =   "&User Management"
      Begin VB.Menu mnuNewUser 
         Caption         =   "New &User"
         Shortcut        =   ^U
      End
      Begin VB.Menu space1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChgePass 
         Caption         =   "Change &Password"
      End
   End
   Begin VB.Menu mnuGuestMnage 
      Caption         =   "&Guest Management"
      Begin VB.Menu mnuGuest 
         Caption         =   "New &Guest"
         Shortcut        =   ^G
      End
      Begin VB.Menu space2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUpdateGuest 
         Caption         =   "&Update Preexisting Guest Info"
      End
   End
   Begin VB.Menu mnuBookWalk 
      Caption         =   "&Booking"
      Begin VB.Menu mnuNewBookWalk 
         Caption         =   "&New Booking"
         Shortcut        =   ^B
      End
      Begin VB.Menu space3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBookRecords 
         Caption         =   "Bo&oking Records"
      End
      Begin VB.Menu space4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancelRebook 
         Caption         =   "&Cancel/Rebook"
      End
   End
   Begin VB.Menu mnuReservation 
      Caption         =   "&Reservation"
      Begin VB.Menu mnuNewReservation 
         Caption         =   "New &Reservation"
         Shortcut        =   ^R
      End
      Begin VB.Menu space5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReservationRecords 
         Caption         =   "R&eservation Records"
      End
      Begin VB.Menu space6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCancelReservation 
         Caption         =   "C&ancel Reservation"
      End
   End
   Begin VB.Menu mnuMaintenance 
      Caption         =   "&Maintenance"
      Begin VB.Menu mnuRoomInfo 
         Caption         =   "Room &Info"
      End
      Begin VB.Menu space7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAmenities 
         Caption         =   "&Amenities"
      End
   End
   Begin VB.Menu mnuReports 
      Caption         =   "R&eports"
      Begin VB.Menu mnuTransactions 
         Caption         =   "&Transactions"
      End
      Begin VB.Menu space8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWalkHistoty 
         Caption         =   "&Walk-In History"
      End
   End
   Begin VB.Menu mnuLogOut 
      Caption         =   "&Log-Out"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "PopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuCheckOut 
         Caption         =   "Check-Out"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub b8TitleBar1_CloseMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
XPFrame1.Visible = False
End If
End Sub

Private Sub cmdChangePass_Click()
frmChangePass.Show vbModal, frmMain
End Sub

Private Sub cmdView_Click()
ListView1.ListItems.Clear
ListView1Prop
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
mnuLogOut_Click
End If
End Sub

Private Sub ListView1Prop()

Dim objList1 As ListItem

    Set ListView1.Icons = imglstIcons1
    
    With ListView1
    
        .View = lvwIcon
    
        .Arrange = lvwAutoTop
        
        .HotTracking = True
        
        .HoverSelection = True
        
        .LabelEdit = lvwManual
        
        .BorderStyle = ccFixedSingle
        
        .Appearance = ccFlat
        
        .OLEDragMode = ccOLEDragAutomatic
        
            For l = 1 To 2
            
            .ColumnHeaders.Add l
            
            Next l
             
    End With

strSQL = "SELECT tblRoom_Status.room_num, tblRoom_Status.room_type, tblRoom_Info.room_name, " & _
         "tblRoom_Info.room_rate, tblRoom_Info.room_rate_per_hour, tblRoom_Info.room_rate_on_season, " & _
         "tblRoom_Info.room_description, tblRoom_Status.room_status " & _
         "FROM tblRoom_Info INNER JOIN tblRoom_Status ON " & _
         "tblRoom_Info.room_type_num = tblRoom_Status.room_type;"
         
         
            Set recSet = New ADODB.Recordset
            
                With recSet
                
                    .Open strSQL, Conn, 3, 3
                    
                    u = 1
                    
                        Do Until .EOF
                        
                            If !room_status = 0 Then
                        
                                Set objList1 = ListView1.ListItems.Add(u, , "Room Number:" & !room_num & "  Room Type:" & !room_name & " Room Status:Available", 6)
                            
                            ElseIf !room_status = 1 Then
                            
                                Set objList1 = ListView1.ListItems.Add(u, , "Room Number:" & !room_num & " Room Type:" & !room_name & " Room Status:Occupied", 8)
                                            
                            End If
                        
                            .MoveNext
                            
                    u = u + 1
                        
                        Loop
                
                End With

End Sub

Private Sub ListView2prop()

Dim objList2 As ListItem

Set ListView2.Icons = imglstIcons1

With ListView2
.Arrange = lvwAutoTop
.HotTracking = True
.HoverSelection = True
.LabelEdit = lvwManual
.BorderStyle = ccFixedSingle
.Appearance = ccFlat
.OLEDragMode = ccOLEDragAutomatic

.ListItems.Add 1, , "Available", 6
.ListItems.Add 2, , "Occupied", 8

For a = 1 To 7
.ColumnHeaders.Add a
Next a
End With

End Sub

Private Sub Form_Load()
DBCON

Text1.Visible = False 'user first name
Text2.Visible = False 'user account type(for restrictions)
txtUserID.Visible = False
userFname = Text1.Text 'user first name
userAcctType = Text2.Text 'user account type(for restrictions)
statBarprop 'status bar properties

b8Line1.Width = frmMain.Width
b8Line2.Width = frmMain.Width

With XPFrame1

.Visible = False
.Top = 24
.Left = 144
.Width = 1001
.Height = 609

End With

ListView2prop
ListView1Prop

Image1.Top = Me.Top
Image1.Left = Me.Left
Image1.Width = Me.Width / 15
Image1.Height = Me.Height / 15


End Sub

Private Sub ListView2_Click()

    Select Case ListView2.SelectedItem.Index
    
    Case 1
    
        ListView1.ListItems.Clear
    
        SelectAvailableRoom
    
    Case 2
    
        ListView1.ListItems.Clear
    
        SelectOccupiedRoom
    
    End Select
End Sub


Private Sub SelectAvailableRoom()

strSQL = "SELECT tblRoom_Status.room_num, tblRoom_Status.room_type, tblRoom_Info.room_name, " & _
         "tblRoom_Info.room_rate, tblRoom_Info.room_rate_per_hour, tblRoom_Info.room_rate_on_season, " & _
         "tblRoom_Info.room_description, tblRoom_Status.room_status " & _
         "FROM tblRoom_Info INNER JOIN tblRoom_Status ON " & _
         "tblRoom_Info.room_type_num = tblRoom_Status.room_type " & _
         "WHERE (((tblRoom_Status.room_status)=0));"

    Set recSet = New ADODB.Recordset
    
        With recSet
        
            .Open strSQL, Conn, 3, 2
            
                    j = 1
                    
                        Do Until .EOF
                        
                            If !room_status = 0 Then
                        
                                Set objList1 = ListView1.ListItems.Add(j, , "Room Number:" & !room_num & "  Room Type:" & !room_name & " Room Status:Available", 6)
                            
                            ElseIf !room_status = 1 Then
                            
                                Set objList1 = ListView1.ListItems.Add(j, , "Room Number:" & !room_num & " Room Type:" & !room_name & " Room Status:Occupied", 8)
                                            
                            End If
                        
                            .MoveNext
                            
                    j = j + 1
                        
                        Loop
        
        End With
        
End Sub

Private Sub SelectOccupiedRoom()

strSQL = "SELECT tblRoom_Status.room_num, tblRoom_Status.room_type, tblRoom_Info.room_name, " & _
         "tblRoom_Info.room_rate, tblRoom_Info.room_rate_per_hour, tblRoom_Info.room_rate_on_season, " & _
         "tblRoom_Info.room_description, tblRoom_Status.room_status " & _
         "FROM tblRoom_Info INNER JOIN tblRoom_Status ON " & _
         "tblRoom_Info.room_type_num = tblRoom_Status.room_type " & _
         "WHERE (((tblRoom_Status.room_status)=1));"

    Set recSet = New ADODB.Recordset
    
        With recSet
        
            .Open strSQL, Conn, 3, 2
            
                    h = 1
                    
                        Do Until .EOF
                        
                            If !room_status = 0 Then
                        
                                Set objList1 = ListView1.ListItems.Add(h, , "Room Number:" & !room_num & "  Room Type:" & !room_name & " Room Status:Available", 6)
                            
                            ElseIf !room_status = 1 Then
                            
                                Set objList1 = ListView1.ListItems.Add(h, , "Room Number:" & !room_num & " Room Type:" & !room_name & " Room Status:Occupied", 8)
                                            
                            End If
                        
                            .MoveNext
                            
                    h = h + 1
                        
                        Loop
        
        End With
        
End Sub


Private Sub mnuAmenities_Click()
frmAmenities.Show vbModal, frmMain
End Sub

Private Sub mnuChgePass_Click()
frmChangePass.Show vbModal, frmMain
End Sub

Sub mnuGuest_Click()
frmNewGuest.Show vbModal, frmMain
End Sub

Private Sub mnuLogOut_Click()
If MsgBox("Are you sure you want to log-out?", vbInformation + vbYesNo, "Confirm Logout") = vbYes Then

 MsgBox "Logging Out Successful! Goodbye " & Text1.Text & "!", vbInformation + vbOKOnly, "Log-Out"

frmLogIn.Show

Unload Me

Else

End If
End Sub

Sub mnuNewBookWalk_Click()
frmCheckIn.lblControl.Caption = "book"
frmCheckIn.dtpickCheckIn.Locked = False
frmCheckIn.xpbtConfirm.Tag = "2"
frmCheckIn.Show vbModal, frmMain
End Sub

Sub mnuNewReservation_Click()
frmCheckIn.lblControl.Caption = "reserve"
frmCheckIn.dtpickCheckIn.Locked = False
frmCheckIn.xpbtConfirm.Tag = "3"
frmCheckIn.txtItemNum.Locked = True
frmCheckIn.Show vbModal, frmMain
End Sub

Private Sub mnuNewUser_Click()
frmNewUser.Show vbModal, frmMain
End Sub

Private Sub mnuRoomInfo_Click()
frmRoomMaintenance.Show vbModal, frmMain

End Sub




Private Sub OsenXPToolBar1_ButtonClick(Index As Integer, sText As String)
Select Case Index

Case 1

XPFrame1.Visible = True

Case 3

ForWalkIn

Case 5

frmCheckOut.Show vbModal, frmMain

Case 7

frmCheckInBooked.Show vbModal


End Select
End Sub

Private Sub statBarMain_PanelClick(ByVal Panel As Panel)
 'if clicked somewhere on the status bar
    Select Case Panel.Key
    
        Case "sbrDate"
    
            Call Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl") ' opens systems time/date changer
    
        Case "sbrTime"
        
            Call Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl") ' opens systems time/date changer
        
                         
        Case Else
        
    
    End Select
End Sub
Sub statBarprop()


        With statBarMain
        
            .Panels.Add 1, "sbrUser", "User:"
            .Panels.Add 2, "sbrUserName"
            .Panels.Add 3, "sbrAccount", "Account Type:"
            .Panels.Add 4, "sbrAccType"
            .Panels.Add 5, "space"
            .Panels.Add 6, , , 2 'numlock
            .Panels.Add 7, , , 1 'capslock
            .Panels.Add 8, "sbrDate" 'date
            .Panels.Add 9, "sbrTime" 'time
            

            
        .Font = "arial"

        .Font.Size = 10

        .Font.Bold = True

        .Panels(1).AutoSize = 2
        .Panels(2).AutoSize = 2
        .Panels(3).AutoSize = 2
        .Panels(4).AutoSize = 2
        .Panels(6).AutoSize = 2
        .Panels(7).AutoSize = 2
        .Panels(8).AutoSize = 2
        .Panels(9).AutoSize = 2

        
        .Panels(1).Picture = imglRunSearch.ListImages(45).Picture 'user
        .Panels(3).Picture = imgListMain.ListImages(5).Picture 'account type
        .Panels(8).Picture = imgListMain.ListImages(4).Picture 'date
        .Panels(9).Picture = imgListMain.ListImages(3).Picture 'time

        .Panels(1).Bevel = 0
        .Panels(3).Bevel = 0
        .Panels(5).Bevel = 0

        .Panels(1).Width = 50
        .Panels(3).Width = 50
        .Panels(8).Width = 50
            
 
        End With
        
End Sub

Private Sub Timer1_Timer()
statBarMain.Panels(9) = Time
statBarMain.Panels(8) = Date
statBarMain.Panels(2) = Text1.Text
statBarMain.Panels(4) = Text2.Text
Xpbar
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "RoomStatus"
MsgBox "room status here"

Case "CheckIn"

ForWalkIn

Case "CheckOut"

frmCheckOut.Show vbModal, frmMain

End Select
End Sub

Sub ForWalkIn()

frmCheckIn.lblControl.Caption = "checkin"
frmCheckIn.dtpickCheckIn.Locked = True
frmCheckIn.xpbtConfirm.Tag = "1"
frmCheckIn.Show vbModal, frmMain

End Sub

Private Sub Xpbar()
Dim UserName1 As String

With XPStatBar
.PanelCaption(2) = statBarMain.Panels(2).Text
.PanelCaption(4) = statBarMain.Panels(4).Text
.PanelCaption(9) = Format$(Now, "mm/dd/yyyy")
.PanelCaption(8) = Format(Now, "hh:mm:ss AM/PM")
End With
End Sub

