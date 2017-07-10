VERSION 5.00
Begin VB.Form frmLock 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   ClientHeight    =   4965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6465
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   ScaleHeight     =   4965
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox imgBg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4425
      Left            =   0
      ScaleHeight     =   4425
      ScaleWidth      =   6435
      TabIndex        =   3
      Top             =   0
      Width           =   6435
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   -360
         ScaleHeight     =   645
         ScaleWidth      =   7005
         TabIndex        =   4
         Top             =   0
         Width           =   7035
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Locked"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   26.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   735
            Left            =   450
            TabIndex        =   5
            Top             =   -90
            Width           =   2145
         End
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmLock.frx":0000
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   5295
      End
   End
   Begin VB.PictureBox bgPassword 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   429
      TabIndex        =   0
      Top             =   4470
      Width           =   6435
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2880
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   60
         Width           =   3495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Password"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1230
         TabIndex        =   2
         Top             =   90
         Width           =   1605
      End
   End
End
Attribute VB_Name = "frmLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SystemPass As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        
        CheckSystemPass
    
    End If

End Sub

Private Sub CheckSystemPass()


strSQL = "select * from tblsystempassword"

    Set recSet = New ADODB.Recordset
    
        With recSet
            .Open strSQL, Conn, 3, 2
            
            SystemPass = InStr(txtPassword.Text, !system_password)
            
                If SystemPass = 1 Then
                
                    MsgBox "The system has been unlocked! Proceeding into log-in form", vbInformation + vbOKOnly, "System Unlocked"
                    
                    frmLogIn.Show
                    
                    Unload Me
                    
                Else
                
                    MsgBox "Wrong System Password! Please Try Again.", vbCritical + vbOKOnly, "Wrong Password"
                    
                    txtPassword.Text = ""
                    
                    txtPassword.SetFocus
                
                End If
            
        End With


End Sub

Private Sub Form_Load()
DBCON
FormPos frmLock
End Sub


