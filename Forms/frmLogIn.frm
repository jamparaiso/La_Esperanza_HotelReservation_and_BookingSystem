VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Object = "{E0D30636-0F87-47D5-B501-08A4FFAC604E}#1.0#0"; "osenxpsuite2005.OCX"
Begin VB.Form frmLogIn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Log-In"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4965
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogIn.frx":0000
   ScaleHeight     =   3690
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3
      Left            =   0
      Top             =   2880
   End
   Begin osenxpsuite2005.OsenXPStatusBar OsenXPStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   3315
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   661
      BackColor       =   14936810
      ForeColor       =   -2147483630
      ForeColorDissabled=   16777215
      MaskColor       =   16711935
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowGripper     =   -1  'True
      ShowSeperators  =   -1  'True
      NumberOfPanels  =   2
      PWidth1         =   130
      PMinWidth1      =   0
      pTTText1        =   ""
      pType1          =   0
      pText1          =   "Log-In Attempt Left:"
      pTextAlignment1 =   0
      pTextBold1      =   -1  'True
      PanelPicture1   =   "frmLogIn.frx":455B
      PanelPicAlignment1=   0
      pBckgColor1     =   0
      pGradient1      =   0
      pEdgeSpacing1   =   0
      pEdgeInner1     =   0
      pEdgeOuter1     =   0
      PWidth2         =   20
      PMinWidth2      =   0
      pTTText2        =   ""
      pType2          =   0
      pText2          =   ""
      pTextAlignment2 =   0
      pTextBold2      =   -1  'True
      PanelPicture2   =   "frmLogIn.frx":4577
      PanelPicAlignment2=   0
      pBckgColor2     =   0
      pGradient2      =   0
      pEdgeSpacing2   =   0
      pEdgeInner2     =   0
      pEdgeOuter2     =   0
      DrawMode        =   1
      HaveXPForm      =   -1  'True
   End
   Begin XPControls.XPButton cmdExit 
      Height          =   495
      Left            =   3480
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Exit"
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
   Begin XPControls.XPButton cmdClear 
      Height          =   495
      Left            =   1920
      TabIndex        =   6
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Clear"
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
   Begin XPControls.XPButton cmdLogIn 
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Log-In"
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
   Begin VB.TextBox txtRestrict 
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox txtPassWord 
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
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox txtUserName 
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
      Left            =   1440
      TabIndex        =   2
      Top             =   960
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1440
      Top             =   4200
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Log-In"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2280
      TabIndex        =   9
      Top             =   480
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "La Esperanza Hotel Booking And Reservation System"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   960
      TabIndex        =   8
      Top             =   120
      Width           =   3870
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmLogIn.frx":4593
      Top             =   120
      Width           =   720
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Height          =   855
      Left            =   0
      Top             =   0
      Width           =   4935
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      Height          =   2175
      Left            =   0
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FailCount As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then 'if the enter key is pressed
cmdLogIn_Click
ElseIf KeyCode = vbKeyEscape Then
cmdExit_Click
End If
End Sub

Private Sub Form_Load()
DBCON

OsenXPStatusBar1.PanelCaption(2) = Timer1.Interval
End Sub

Sub clear_Fields()
txtUsername.Text = ""
txtPassword.Text = ""
txtUsername.SetFocus
End Sub

Private Sub cmdClear_Click()
clear_Fields
End Sub

Private Sub cmdExit_Click()
If MsgBox("Are You Sure You Want To Exit?", vbYesNo, "Exit") = vbYes Then

    End
    
End If
End Sub

Private Sub cmdLogIn_Click()

Dim CaseUname As Integer
Dim CasePass As Integer

strSQL = "SELECT tbluser_info.user_id, tblUser_Info.uname, tblUser_Info.pword, tblUser_Info.first_name, " & _
            "tblAccount_Types.acct_description, tblUser_Info.acc_type " & _
            "FROM tblAccount_Types INNER JOIN tblUser_Info ON " & _
            "tblAccount_Types.account_id = tblUser_Info.[acc_type] " & _
            "WHERE (((tblUser_Info.uname)='" & txtUsername.Text & "') AND ((tblUser_Info.pword)='" & txtPassword.Text & "'));"




Set recSet = New ADODB.Recordset

 With recSet
 
  .Open strSQL, Conn, 3, 3
  
  If .BOF Or .EOF Then '2 if no user found,null
  
    If MsgBox("Invalid Username Or Password! Try Again.", vbCritical + vbYesNo, "Error") = vbYes Then '1
    
    Timer1.Interval = Timer1.Interval - 1
    
    OsenXPStatusBar1.PanelCaption(2) = Timer1.Interval
    
     'if the user press YES above
    clear_Fields
    
    Else
    
       If MsgBox("The System Will be Closed. Are You Sure?", vbInformation + vbYesNo, _
                "Close System") = vbYes Then '0 if the user press YES above
    
            End
        
        Else ' if the user press NO,refresh
        
            Unload Me
        
            frmLogIn.Show
    
        End If '0
  
    End If '1
  
 Else 'tbl is not empty and found a user refer to if 2
 CaseUname = InStr(txtUsername.Text, !uname)
 CasePass = InStr(txtPassword.Text, !pword)
 
        If CaseUname And CasePass = 1 Then
  
  MsgBox "Welcome" & " " & !first_name & "!", vbInformation + vbOKOnly, "Logged-In as" & " " & !acct_description
  
  frmMain.Text1.Text = !first_name 'transfer to frmmain
  frmMain.Text2.Text = !acct_description 'transfer to frmmain
  frmMain.txtUserID.Text = !user_id 'transfer to frmmain
  
  txtRestrict.Text = !acc_type 'restriction number
  
  user_restriction 'function
  
  frmMain.Show
   
  Unload Me
  
        Else
        
        MsgBox "Invalid Username Or Password! Try Again.", vbCritical + vbYesNo, "Error"
        
        clear_Fields
  
        End If
  
  End If '3
  
  .Close

 End With

End Sub

Sub user_restriction()
'this affects the main form
Select Case txtRestrict.Text

Case 2

frmMain.mnuUserManage.Visible = False

Case 3

frmMain.mnuUserManage.Visible = False
frmMain.mnuReports.Visible = False
frmMain.mnuMaintenance.Visible = False

End Select

End Sub


Private Sub Timer1_Timer()

    If Timer1.Interval = 0 Then
    
        MsgBox "You have reach the limit of log-in attempt. The System will be locked. Please notify the Owner to unlock.", vbCritical + vbOKOnly, "System Lock"
        
        frmLock.Show
        
        Unload Me
    
    End If

End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Or KeyAscii = 39 Then
KeyAscii = 0

txtUsername.ToolTipText = "These Characters are not allowed:' , """
End If
End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
If KeyAscii = 34 Or KeyAscii = 39 Then
KeyAscii = 0

txtUsername.ToolTipText = "These Characters are not allowed:' , """
End If
End Sub
