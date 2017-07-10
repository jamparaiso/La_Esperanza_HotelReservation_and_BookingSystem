VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmRoomMaintenance 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10485
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   13005
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmRoomMaintenance.frx":0000
   ScaleHeight     =   10485
   ScaleWidth      =   13005
   StartUpPosition =   2  'CenterScreen
   Begin XPControls.XPFrame XPFrame3 
      Height          =   3615
      Left            =   6240
      TabIndex        =   18
      Top             =   6600
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   6376
      Caption         =   "Change Room Rates And Description"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtRateHour 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4920
         TabIndex        =   21
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtSeasonRate 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4920
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtNormalRate 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         TabIndex        =   19
         Top             =   975
         Width           =   1575
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   1215
         Left            =   240
         TabIndex        =   22
         Top             =   1680
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2143
         _Version        =   393217
         TextRTF         =   $"frmRoomMaintenance.frx":455B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPButton cmdSaveRate 
         Height          =   495
         Left            =   3840
         TabIndex        =   28
         Top             =   3000
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "&Save"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SoundFile       =   "C:\WINDOWS\Media\Windows XP Notify.wav"
      End
      Begin XPControls.XPButton cmdCancelRate 
         Height          =   495
         Left            =   5280
         TabIndex        =   29
         Top             =   3000
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "&Cancel"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SoundFile       =   "C:\WINDOWS\Media\tada.wav"
      End
      Begin XPControls.XPButton cmdEditRate 
         Height          =   495
         Left            =   240
         TabIndex        =   30
         Top             =   3000
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "&Edit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPCombo cboRoomType2 
         Height          =   315
         Left            =   1320
         TabIndex        =   32
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         Text            =   "Choose Here"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label10 
         Caption         =   "Room Type:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Rate Per Hour:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   25
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Seasonal Rate:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   24
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Normal Rate:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Width           =   1215
      End
   End
   Begin XPControls.XPButton cmdBAck 
      Height          =   495
      Left            =   4200
      TabIndex        =   10
      Top             =   9600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Back"
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
   Begin XPControls.XPFrame XPFrame2 
      Height          =   1575
      Left            =   6240
      TabIndex        =   6
      Top             =   5040
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   2778
      Caption         =   "Add Room/Change Room Type"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin XPControls.XPCombo cboRoomType 
         Height          =   315
         Left            =   4200
         TabIndex        =   14
         Top             =   360
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         Text            =   "Choose Here"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtRoomNum 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin XPControls.XPButton cmdSave 
         Height          =   495
         Left            =   3840
         TabIndex        =   15
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "&Save"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPButton cmdCancel 
         Height          =   495
         Left            =   5280
         TabIndex        =   16
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "&Cancel"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPButton cmdAdd 
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "&Add Room"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPButton cmdEdit 
         Height          =   495
         Left            =   1680
         TabIndex        =   27
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "&Edit"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Room Number:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Room Type: "
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
   End
   Begin XPControls.XPFrame XPFrame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3836
      Caption         =   "Tools"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtRoomCount 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   13
         Top             =   1560
         Width           =   1455
      End
      Begin XPControls.XPButton cmdOk 
         Height          =   495
         Left            =   1560
         TabIndex        =   5
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "&OK"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPCombo cboViewBy 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         Text            =   "Choose Here"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPControls.XPButton cmdViewAll 
         Height          =   495
         Left            =   3000
         TabIndex        =   11
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "&View All"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label9 
         Caption         =   "Room Count:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "View By Room Type:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11760
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoomMaintenance.frx":45DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoomMaintenance.frx":49EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoomMaintenance.frx":4E18
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoomMaintenance.frx":5202
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoomMaintenance.frx":56C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoomMaintenance.frx":60BF
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRoomMaintenance.frx":6F99
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   7011
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Room Maintenance Section"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   4
      Top             =   240
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   240
      Picture         =   "frmRoomMaintenance.frx":73D2
      Stretch         =   -1  'True
      Top             =   120
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   945
      Left            =   0
      Top             =   0
      Width           =   12975
   End
End
Attribute VB_Name = "frmRoomMaintenance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RoomTypeNum As Integer
Dim RoomTypeNum1 As Integer
Dim RoomTypeNum2 As Integer


Private Sub cboRoomType_Change()

    Select Case cboRoomType.listIndex
    
        Case 0
        
            RoomTypeNum1 = 5
        
        Case 1
        
            RoomTypeNum1 = 4
        
        Case 2
        
            RoomTypeNum1 = 3
        
        Case 3
        
            RoomTypeNum1 = 2
        
        Case 4
        
            RoomTypeNum1 = 1
    
    End Select
    
End Sub

Private Sub cboRoomType2_Change()

    Select Case cboRoomType2.listIndex
    
        Case 0
        
            RoomTypeNum2 = 5
        
        Case 1
        
            RoomTypeNum2 = 4
        
        Case 2
        
            RoomTypeNum2 = 3
        
        Case 3
        
            RoomTypeNum2 = 2
        
        Case 4
        
            RoomTypeNum2 = 1
    
    End Select
    
strSQL = "select * from tblroom_info where room_type_num= " & RoomTypeNum2 & ""

    Set recSet = New ADODB.Recordset
    
    With recSet
    
        .Open strSQL, Conn, 3, 2
        
            txtNormalRate.Text = !room_rate
            
            txtSeasonRate.Text = !room_rate_on_season
            
            txtRateHour.Text = !room_rate_per_hour
            
            RichTextBox1.Text = !room_description
        
        .Close
        
    End With

End Sub

Private Sub cboViewBy_Change()

    Select Case cboViewBy.listIndex
    
        Case 0
        
            RoomTypeNum = 5
        
        Case 1
        
            RoomTypeNum = 4
        
        Case 2
        
            RoomTypeNum = 3
        
        Case 3
        
            RoomTypeNum = 2
        
        Case 4
        
            RoomTypeNum = 1
    
    End Select

End Sub

Private Sub cmdAdd_Click()

    MsgBox "The system doesn't automate room number, Please make sure that you supply the correct room number.", vbInformation + vbOKOnly, "Notice"
    
    cmdSave.Tag = 2
    
    EnableText
    
    ClearFields
    
    ListView1.Enabled = False
    
End Sub

Private Sub cmdBAck_Click()
End
End Sub

Private Sub cmdCancel_Click()

DisableText

cmdEdit.Enabled = True

cmdAdd.Enabled = True

cmdSave.Tag = 1

ListView1.Enabled = True

End Sub

Private Sub cmdCancelRate_Click()
DisableTextRate
End Sub

Private Sub cmdEdit_Click()

MsgBox "This will only change the room type of a certain room.", vbOKOnly, "Attention"

txtRoomNum.Enabled = False

cmdSave.Tag = 3

cmdAdd.Enabled = False

EnableText

End Sub

Private Sub cmdEditRate_Click()
MsgBox "Please choose the room type you want to edit it's information", vbOKOnly, "Edit Room Info"
EnableTextRate

cboRoomType2.SetFocus

End Sub

Private Sub cmdOk_Click()

    ListView1.ListItems.Clear
    
        strSQL = "SELECT tblRoom_Status.room_num, tblRoom_Status.room_type, tblRoom_Info.room_name, " & _
                 "tblRoom_Info.room_rate, tblRoom_Info.room_rate_per_hour, tblRoom_Info.room_rate_on_season, " & _
                 "tblRoom_Info.room_description, tblRoom_Status.room_status " & _
                 "FROM tblRoom_Info INNER JOIN tblRoom_Status ON " & _
                 "tblRoom_Info.room_type_num = tblRoom_Status.room_type " & _
                 "WHERE (((tblRoom_Status.room_type)=" & RoomTypeNum & "));"
                 
            
    Set recSet = New ADODB.Recordset
    
        With recSet
        
            .Open strSQL, Conn, 3, 3
        
                a = 1
        
            Do Until .EOF
        
                Set objList = ListView1.ListItems.Add(a, , !room_num, , 4)
        
                    If !room_status = 1 Then
                    
                        objList.ListSubItems.Add 1, , "Occupied", 2
                    
                    ElseIf !room_status = 0 Then
                    
                        objList.ListSubItems.Add 1, , "Available", 3
                    
                    End If
            
                    objList.ListSubItems.Add 2, , !room_name, 1
                    objList.ListSubItems.Add 3, , !room_rate, 7
                    objList.ListSubItems.Add 4, , !room_rate_on_season, 7
                    objList.ListSubItems.Add 5, , !room_rate_per_hour, 7
                    objList.ListSubItems.Add 6, , !room_description, 6
            
        
        
                 a = a + 1
        
            .MoveNext
        
            Loop
        
        .Close
        
        
    End With
        
        txtRoomCount.Text = a - 1

End Sub

Private Sub cmdSave_Click()

If cmdSave.Tag = 2 Then 'for new room

        AddNewRoom
    
ElseIf cmdSave.Tag = 3 Then 'for changing room type

    If ListView1.SelectedItem.SubItems(1) = "Available" Then

        ChangeRoomType
    
    Else
    
        MsgBox "You can't edit room that has been occupied. Please wait till the guest check before editing.", vbCritical + vbOKOnly, "Error"
    
        Exit Sub
    
    End If

End If

End Sub

Private Sub ChangeRoomType()

strSQL = "select * from tblroom_status where room_num= " & txtRoomNum.Text & ""

    Set recSet = New ADODB.Recordset
    
        With recSet
            .Open strSQL, Conn, 3, 2
            
            
            If Not .EOF Then
            
                !room_type = RoomTypeNum1
                
                .Update
                
                .Close
                
                cmdViewAll_Click
                    
                ListView1.Enabled = True
                    
                DisableText
                
                MsgBox "Room type succesfully change.", vbOKOnly, "Success"
                
                Exit Sub
                
            Else
            
                MsgBox "No room found. Please try again.", vbCritical + vbOKOnly, "Error"
                
                txtRoomNum.Text = ""
                
                txtRoomNum.SetFocus
                
                Exit Sub
            
            End If
        End With

End Sub

Private Sub cmdSaveRate_Click()

    If MsgBox("Are you sure that you want to edit the information of this room?", vbInformation + vbYesNo, "Confirm Edit") = vbYes Then
    
        strSQL = "select * from tblroom_info where room_type_num=" & RoomTypeNum2 & ""
        
            Set recSet = New ADODB.Recordset
            
                With recSet
                .Open strSQL, Conn, 3, 2
                
                    !room_rate = txtNormalRate.Text
                    
                    !room_rate_per_hour = txtRateHour.Text
                    
                    !room_rate_on_season = txtSeasonRate.Text
                    
                    !room_description = RichTextBox1.Text
                    
                .Update
                
                .Close
                
                MsgBox "Room information has been successfully changed.", vbInformation + vbOKOnly, "Success"
                
                DisableTextRate
                
                cmdViewAll_Click
                    
                ListView1.Enabled = True
                
                End With
            
    End If

End Sub

Private Sub cmdViewAll_Click()

ListView1.ColumnHeaders.Clear

ListView1.ListItems.Clear

ListViewProp

End Sub

Private Sub Form_Load()
DBCON

DisableText

DisableTextRate

FormPos frmRoomMaintenance

ListViewProp

cboViewByItems

cmdSave.Tag = 1

End Sub

Private Sub DisableTextRate()

    txtNormalRate.Enabled = False
    
    txtSeasonRate.Enabled = False
    
    txtRateHour.Enabled = False
    
    RichTextBox1.Enabled = False
    
    cmdSaveRate.Enabled = False
    
    cmdCancelRate.Enabled = False

End Sub

Private Sub EnableTextRate()

    txtNormalRate.Enabled = True
    
    txtSeasonRate.Enabled = True
    
    txtRateHour.Enabled = True
    
    RichTextBox1.Enabled = True
    
    cmdSaveRate.Enabled = True
    
    cmdCancelRate.Enabled = True

End Sub

Private Sub AddNewRoom()

strSQL = "select * from tblroom_status where room_num= " & txtRoomNum.Text & ""

Set recSet = New ADODB.Recordset

    With recSet
    
        .Open strSQL, Conn, 3, 2
        
            If Not .EOF Then
            
                MsgBox "This room number is already taken. Please choose another one.", vbCritical + vbOKOnly, "Error"
                
                txtRoomNum.Text = ""
                
                txtRoomNum.SetFocus
                
                Exit Sub
                
            ElseIf .EOF Then
            
                If MsgBox("Are you sure you want to add this room?", vbYesNo, "Confirm Save") = vbYes Then
                
                    .AddNew
                    
                    !room_num = txtRoomNum.Text
                    
                    !room_type = RoomTypeNum1
                    
                    !room_status = 0
                    
                    .Update
                    
                    '.Close
                    
                    MsgBox "Adding new room successfull.You may perform transaction using the new added room.", vbOKOnly, "Success"
                    
                    cmdViewAll_Click
                    
                    ListView1.Enabled = True
                    
                    DisableText
                
                Exit Sub
                
                Else
                
                    Exit Sub
                
                End If
                            
            End If
    
    End With
    
End Sub

Private Sub ClearFields()

txtRoomNum.Text = ""

txtNormalRate.Text = ""

txtSeasonRate.Text = ""

txtRateHour.Text = ""

RichTextBox1.Text = ""

End Sub

Private Sub DisableText()

txtRoomNum.Enabled = False

cmdSave.Enabled = False

cmdCancel.Enabled = False

End Sub

Private Sub EnableText()

txtRoomNum.Enabled = True

cmdSave.Enabled = True

cmdCancel.Enabled = True

End Sub

Private Sub ListViewProp()

Dim objList As ListItem

        With ListView1
        
            .View = lvwReport
            
            .ColumnHeaders.Add 1, , "Room Number", ListView1.Width * 0.15, , 4
            
            .ColumnHeaders.Add 2, , "Status", ListView1.Width * 0.15, , 5
            
            .ColumnHeaders.Add 3, , "Room Type", ListView1.Width * 0.16, , 1
            
            .ColumnHeaders.Add 4, , "Room Rate", ListView1.Width * 0.15, , 7
            
            .ColumnHeaders.Add 5, , "Seasonal Rate", ListView1.Width * 0.18, , 7
            
            .ColumnHeaders.Add 6, , "Rate Per Hour", ListView1.Width * 0.18, , 7
            
            .ColumnHeaders.Add 7, , "Room Description", ListView1.Width * 1, , 6
            
            .FullRowSelect = True
            
            .GridLines = True
        
        End With
        

    strSQL = "SELECT tblRoom_Status.room_num, tblRoom_Info.room_name, tblRoom_Info.room_rate, " & _
             "tblRoom_Info.room_rate_per_hour, tblRoom_Info.room_rate_on_season, " & _
             "tblRoom_Info.room_description, tblRoom_Status.room_status " & _
             "FROM tblRoom_Info INNER JOIN tblRoom_Status ON " & _
             "tblRoom_Info.room_type_num = tblRoom_Status.room_type;"



    Set recSet = New ADODB.Recordset
            
            With recSet
            
            .Open strSQL, Conn, 3, 3
            
                    i = 1
                    
                    Do Until .EOF
                    
                    Set objList = ListView1.ListItems.Add(i, , !room_num, , 4)
                    
                            If !room_status = 1 Then
                            
                                objList.ListSubItems.Add 1, , "Occupied", 2
                            
                            ElseIf !room_status = 0 Then
                            
                                objList.ListSubItems.Add 1, , "Available", 3
                            
                            End If
                        
                        objList.ListSubItems.Add 2, , !room_name, 1
                        
                        objList.ListSubItems.Add 3, , !room_rate, 7
                        
                        objList.ListSubItems.Add 4, , !room_rate_on_season, 7
                        
                        objList.ListSubItems.Add 5, , !room_rate_per_hour, 7
                        
                        objList.ListSubItems.Add 6, , !room_description, 6
                        
                    
                    
                    i = i + 1
                    
                    .MoveNext
                    
                    Loop
            
            End With

txtRoomCount.Text = i - 1
End Sub

Private Sub cboViewByItems()

    strSQL = "select * from tblroom_info"

        Set recSet = New ADODB.Recordset
        
                With recSet
                    .Open strSQL, Conn, 3, 3
                    
                        Do Until .EOF
                        
                        a = 0
                        
                        cboViewBy.AddItem !room_name, a
                        
                        cboRoomType.AddItem !room_name, a
                        
                        cboRoomType2.AddItem !room_name, a
                        
                        a = a + 1
                        
                        .MoveNext
                        
                        Loop
                        
                    .Close
                    
                End With


End Sub

Private Sub ListView1_Click()

txtRoomNum.Text = ListView1.SelectedItem

End Sub

Private Sub txtNormalRate_KeyPress(KeyAscii As Integer)
KeyAscii = OnlyNumericKeys(KeyAscii, txtNormalRate)
End Sub

Private Sub txtRateHour_KeyPress(KeyAscii As Integer)
KeyAscii = OnlyNumericKeys(KeyAscii, txtNormalRate)
End Sub

Private Sub txtRoomNum_KeyPress(KeyAscii As Integer)

KeyAscii = OnlyNumericKeys(KeyAscii, txtRoomNum)

End Sub

Private Sub txtSeasonRate_KeyPress(KeyAscii As Integer)
KeyAscii = OnlyNumericKeys(KeyAscii, txtNormalRate)
End Sub

