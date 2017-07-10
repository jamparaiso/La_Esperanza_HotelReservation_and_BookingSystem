VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BB5807FE-DBD2-11D3-87C1-4C980CC10374}#1.0#0"; "MyHover.ocx"
Object = "{2210EC79-A724-4033-AAF4-790E2467C0E8}#1.0#0"; "vbalCmdBar6.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin vbalCmdBar6.vbalCommandBar vbalCommandBar1 
      Height          =   495
      Left            =   0
      Top             =   600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   0
   End
   Begin vbalCmdBar6.vbalCommandBarDock vbalCommandBarDock1 
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   873
   End
   Begin MyHoverButton.Button Button1 
      Height          =   975
      Left            =   3840
      TabIndex        =   6
      Top             =   2400
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1720
      HoverBackColor  =   12632319
      DownBackColor   =   12648447
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Form1.frx":0000
      HoverPicture    =   "Form1.frx":0EDA
      DisabledPicture =   "Form1.frx":0EF6
      DownPicture     =   "Form1.frx":0F12
      MouseIcon       =   "Form1.frx":0F2E
      HoverCaption    =   ""
      DownCaption     =   ""
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   720
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   720
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   58785793
      CurrentDate     =   41164
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   960
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   58785793
      CurrentDate     =   41164
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
Text2.Text = Format(Combo1.Text, "hh:mm AM/PM")
End Sub

Sub VerifyRooms()

strSQL = "SELECT tblBooking_Info.booking_date, tblBooking_Info.out_date, " & _
         "tblBooking_Info.expected_time_in, tblBooking_Info.expected_time_out, " & _
         "tblBooking_Info.room_num " & _
         "From tblBooking_Info " & _
         "WHERE (((tblBooking_Info.room_num)=302));"
         
    Set recSet = New ADODB.Recordset
    
     With recSet
     .Open strSQL, Conn, 3, 3
     
     Do While Not .EOF
     
     If DTPicker1.Value >= !booking_date And _
        DTPicker1.Value <= !out_date And _
        TimeNo(Text2.Text) >= TimeNo(!expected_time_in) Then
            
            If DTPicker1.Value = !out_date And TimeNo(Text1.Text) > TimeNo(!expected_time_out) Then
            
            MsgBox "booked1"
            Exit Sub
            
            Else
            
            MsgBox "occupied1"
            Exit Sub
            
            End If
        
        
     ElseIf DTPicker1.Value >= !booking_date And _
        DTPicker1.Value <= !out_date And _
        TimeNo(Text2.Text) < TimeNo(!expected_time_in) Then
        
            If DTPicker1.Value = !out_date And TimeNo(Text1.Text) > TimeNo(!expected_time_out) Then
            
            MsgBox "booked1"
            Exit Sub
            
            Else
            
            MsgBox "occupied1"
            Exit Sub
            
            End If
            
     ElseIf DTPicker2.Value >= !booking_date And _
            DTPicker2.Value <= !out_date And _
            TimeNo(Text2.Text) >= TimeNo(!expected_time_in) Then
            
           MsgBox "occupied2"
           Exit Sub
           
     ElseIf DTPicker2.Value >= !booking_date And _
            DTPicker2.Value <= !out_date And _
            TimeNo(Text2.Text) < TimeNo(!expected_time_in) Then
            
            If DTPicker2.Value = !booking_date Then
            
            MsgBox "Booked2"
            Exit Sub
            
            Else
            MsgBox "occupied2"
            Exit Sub
            
            End If
     ElseIf DTPicker1.Value <= !booking_date And _
            DTPicker2.Value >= !out_date Then
            
            MsgBox "occupied3"
            Exit Sub
            

            
     End If
     .MoveNext
     Loop
     
     End With
     
'reservation-----------------------------------------------------
strSQL = "SELECT tblReservation_Info.reservation_date, tblReservation_Info.out_date, " & _
         "tblReservation_Info.expected_time_in, tblReservation_Info.expected_time_out, " & _
         "tblReservation_Info.room_num " & _
         "From tblReservation_Info " & _
         "WHERE (((tblReservation_Info.room_num)=302));"

         
    Set recSet = New ADODB.Recordset
    
     With recSet
     .Open strSQL, Conn, 3, 3
     
     Do While Not .EOF
     
     If DTPicker1.Value >= !reservation_date And _
        DTPicker1.Value <= !out_date And _
        TimeNo(Text2.Text) >= TimeNo(!expected_time_in) Then
            
            If DTPicker1.Value = !out_date And TimeNo(Text1.Text) > TimeNo(!expected_time_out) Then
            
            MsgBox "booked1rs"
            Exit Sub
            
            Else
            
            MsgBox "occupied1rs"
            Exit Sub
            
            End If
        
        
     ElseIf DTPicker1.Value >= !reservation_date And _
        DTPicker1.Value <= !out_date And _
        TimeNo(Text2.Text) < TimeNo(!expected_time_in) Then
        
            If DTPicker1.Value = !out_date And TimeNo(Text1.Text) > TimeNo(!expected_time_out) Then
            
            MsgBox "booked1rs"
            Exit Sub
            
            Else
            
            MsgBox "occupied1rs"
            Exit Sub
            
            End If
            
     ElseIf DTPicker2.Value >= !reservation_date And _
            DTPicker2.Value <= !out_date And _
            TimeNo(Text2.Text) >= TimeNo(!expected_time_in) Then
            
           MsgBox "occupied2rs"
           Exit Sub
           
     ElseIf DTPicker2.Value >= !reservation_date And _
            DTPicker2.Value <= !out_date And _
            TimeNo(Text2.Text) < TimeNo(!expected_time_in) Then
            
            If DTPicker2.Value = !reservation_date Then
            
            MsgBox "Booked2rs"
            Exit Sub
            
            Else
            MsgBox "occupied2rs"
            Exit Sub
            
            End If
     ElseIf DTPicker1.Value <= !reservation_date And _
            DTPicker2.Value >= !out_date Then
            
            MsgBox "occupied3"
            Exit Sub
            

            
     End If
     .MoveNext
     Loop
     
     End With
     
    MsgBox "you can book"

End Sub

Private Sub Command1_Click()
VerifyRooms
End Sub

Private Sub Form_Load()
DBCON

DTPicker1.Format = dtpShortDate
DTPicker2.Format = dtpShortDate

Text1.Text = Format("2:00 PM", "hh:mm AM/PM")

With Combo1
.AddItem "12:00 PM"
.AddItem "1:00 PM"
.AddItem "2:00 PM"
.AddItem "3:00 PM"
.AddItem "4:00 PM"
.AddItem "5:00 PM"
.AddItem "6:00 PM"


End With

End Sub

