VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDF6FCF6-E2A0-4DA6-8DF8-FA27594705C8}#26.1#0"; "XPControls.ocx"
Begin VB.Form frmAmenities 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   ClientHeight    =   5295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9030
   Icon            =   "frmAmenities.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAmenities.frx":0442
   ScaleHeight     =   5295
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XPControls.XPButton cmdAdd 
      Height          =   495
      Left            =   1080
      TabIndex        =   7
      Top             =   3840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Add"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtItemPrice 
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
      Left            =   7200
      TabIndex        =   3
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox txtItemName 
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
      Left            =   7200
      TabIndex        =   2
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox txtItemNum 
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
      Left            =   7200
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4471
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin XPControls.XPButton cmdEdit 
      Height          =   495
      Left            =   2640
      TabIndex        =   8
      Top             =   3840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Edit"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton cmdDelete 
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      Top             =   3840
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Delete"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton cmdSave 
      Height          =   495
      Left            =   1080
      TabIndex        =   10
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Save"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton cmdCancel 
      Height          =   495
      Left            =   2640
      TabIndex        =   11
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPControls.XPButton cmdBack 
      Height          =   495
      Left            =   4080
      TabIndex        =   12
      Top             =   4560
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "&Back"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amenity Section"
      BeginProperty Font 
         Name            =   "Poor Richard"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   13
      Top             =   240
      Width           =   2985
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   120
      Picture         =   "frmAmenities.frx":499D
      Stretch         =   -1  'True
      Top             =   120
      Width           =   705
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   9015
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Price:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6360
      TabIndex        =   6
      Top             =   3480
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6360
      TabIndex        =   5
      Top             =   2520
      Width           =   1200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Item Num:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6360
      TabIndex        =   4
      Top             =   1560
      Width           =   1080
   End
End
Attribute VB_Name = "frmAmenities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const ItemNameIndex As Integer = 1 'index for column 2 for listview1
Private Const ItemPriceIndex As Integer = 2 'index for column 3 for listview 1

Private Sub cmdAdd_Click()

ListView1.Enabled = False

disable_AddDelEdit

enable_TextFields

clear_Fields

txtItemNum.Enabled = False

txtItemName.SetFocus

strSQL = "select max(item_num) as maxItemnum from tblamenities"

    Set recSet = New ADODB.Recordset
    
        With recSet
            .Open strSQL, Conn, 3, 3
            
             txtItemNum.Text = !maxitemnum + 1
            
            .Close
        
        End With
        
 cmdSave.Tag = 1

End Sub



Private Sub cmdBack_Click()
Unload Me
End Sub

Private Sub cmdCancel_Click()
ListView1.Enabled = True
enable_AddDelEdit
disable_TextFields
cmdSave.Tag = ""
End Sub

Private Sub cmdDelete_Click()
If txtItemNum.Text <> "" Then '0

    If MsgBox("Are you sure you want to delete this file?", vbInformation + vbYesNo, "Delete") = vbYes Then '1
    
        If MsgBox("Deleted records cannot recovered.Do you still want to continue?", vbInformation + vbYesNo, "Confirm Delete") = vbYes Then '3
        
        strSQL = "select * from tblamenities where item_num= " & txtItemNum.Text & ""
        
            Set recSet = New ADODB.Recordset
            
             With recSet
             
              .Open strSQL, Conn, 3, 2
              
              .Delete
              
              .Update
              
              .Close
             
             End With
             
          If MsgBox("Deleting Successful, Do you want to stay in this form?", vbInformation + vbYesNo, "Delete success") = vbYes Then '4
          
            Form_Load
            
         Else '4
         
         End If '4
        
        Else '3
        
         If MsgBox("Delete Cancelled,do yo want to stay in this Form?", vbInformation + vbYesNo, "Delete Cancelled") = vbYes Then '5
        
         Else '5
         
         Unload Me
        
         End If '5
         
        End If '3
    
    Else '1
    
     If MsgBox("Delete Cancalled, do you want to stay in this Form?", vbInformation + vbYesNo, "Delete Cancelled") = vbYes Then '2
        
     Else '2
     
     Unload Me
             
     End If '2
    
    End If '1
    
Else '0

MsgBox "No Record selected! Please Choose a record.", vbInformation + vbOKOnly, "Error"

End If '0
End Sub

Private Sub cmdEdit_Click()
disable_AddDelEdit

enable_TextFields

txtItemNum.Enabled = False

txtItemName.SetFocus

cmdSave.Tag = 2
End Sub

Sub clear_Fields()
txtItemName.Text = ""
txtItemNum.Text = ""
txtItemPrice.Text = ""
End Sub

Private Sub cmdSave_Click()
On Error GoTo ErrCatch

If cmdSave.Tag = 1 Then '0
    If MsgBox("Are You sure you want to save this record?", vbInformation + vbYesNo, "Confirm Save") = vbYes Then '1
    
        strSQL = "select * from tblamenities"
        
            Set recSet = New ADODB.Recordset
            
                With recSet
                
                    .Open strSQL, Conn, 3, 3
                    
                        .AddNew
                        
                        !item_name = txtItemName.Text
                        !item_price = txtItemPrice.Text
                        
                        .Update
                        
                        
                       If MsgBox("Successfully saved!Do You want to stay in this Form?", vbInformation + vbYesNo, "Save Success") = vbYes Then '2
                       
                       cmdSave.Tag = ""
                       
                        ListView1.Enabled = True
                       
                        ListView1.ListItems.Clear
                        
                        List_ViewProp
                        
                        List_ViewList
                        
                        enable_AddDelEdit

                            clear_Fields
                            
                            disable_TextFields
                       
                       Else '2
                       
                        MsgBox "Returning to Main Form...", vbOKOnly, "Returning"
                        
                        cmdSave.Tag = ""
                        
                        Unload Me
                       
                       End If '2
                    
                    .Close
                
                End With
    Else '1
    
        MsgBox "Save cancelled", vbOKOnly, "Cancelled"
        
        enable_AddDelEdit
        
        clear_Fields
        
        disable_TextFields
        
        cmdSave.Tag = ""
    
    End If '1
    
ElseIf cmdSave.Tag = 2 Then '0
    If txtItemName.Text <> "" And txtItemPrice.Text <> "" Then '3
    
        If MsgBox("Are You sure you want to edit this record?", vbInformation + vbYesNo, "Confirm Edit") = vbYes Then '4
    
            strSQL = "select * from tblamenities where item_num= " & txtItemNum.Text & ""
        
                Set recSet = New ADODB.Recordset
        
                     With recSet
                        .Open strSQL, Conn, 3, 2
                
                            !item_name = txtItemName.Text
                            !item_price = txtItemPrice.Text
                    
                        .Update
                        
                        .Close
                        
                     If MsgBox("Edit Successfull!Do you want to stay in this form?", vbInformation + vbYesNo, "Edit Success") = vbYes Then '5
                     
                            cmdSave.Tag = ""
                     
                            Unload Me
                     
                            frmAmenities.Show vbModal, frmMain
                     
                     Else '5
                     
                            MsgBox "Returning to Main Form...", vbInformation + vbOKOnly, "Returning"
                            
                            Unload Me
                     
                     End If '5
            
                     End With
         Else '4
         
            MsgBox "Edit Cancelled.", vbInformation + vbOKOnly, "Edit Cancelled"
            
            enable_AddDelEdit
        
            clear_Fields
        
            disable_TextFields
        
            cmdSave.Tag = ""
         
         End If '4
    
    Else '3
    
        MsgBox "Please Choose the record you want to edit.", vbExclamation + vbOKOnly, "No Record to Edit"
    
    End If '3
End If '0

ErrCatch:
Select Case Err.Number

Case 2147252571

MsgBox "Invalid Input! Try again!", vbCritical + vbOKOnly, "Error"

txtItemPrice.SetFocus

Exit Sub

End Select

End Sub

Private Sub Form_Load()
FormPos frmAmenities
DBCON

cmdSave.Tag = ""

List_ViewProp

List_ViewList

enable_AddDelEdit

clear_Fields

disable_TextFields
End Sub

Sub enable_TextFields()
txtItemNum.Enabled = True
txtItemName.Enabled = True
txtItemPrice.Enabled = True
End Sub

Sub disable_TextFields()
txtItemNum.Enabled = False
txtItemName.Enabled = False
txtItemPrice.Enabled = False
End Sub

Sub enable_AddDelEdit()
cmdAdd.Enabled = True
cmdDelete.Enabled = True
cmdEdit.Enabled = True
cmdSave.Enabled = False
cmdCancel.Enabled = False
End Sub

Sub disable_AddDelEdit()
cmdAdd.Enabled = False
cmdDelete.Enabled = False
cmdEdit.Enabled = False
cmdSave.Enabled = True
cmdCancel.Enabled = True
End Sub

Sub List_ViewProp()

    With ListView1
    
        .View = lvwReport
        .FullRowSelect = True
        .GridLines = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Item Num", .Width * 0.15
        .ColumnHeaders.Add 2, , "Item Name", .Width * 0.5
        .ColumnHeaders.Add 3, , "Item Price", .Width * 0.35
        
    End With
End Sub


Sub List_ViewList()

Dim objList As ListItem

strSQL = "select * from tblamenities"

 Set recSet = New ADODB.Recordset
 
  With recSet
  
   .Open strSQL, Conn, 3, 3
   
   Do Until .EOF
   
    Set objList = ListView1.ListItems.Add(, , !item_num) 'must set this first so that the subitems can follow
        objList.SubItems(ItemNameIndex) = !item_name
        objList.SubItems(ItemPriceIndex) = !item_price
   
    .MoveNext
    
   Loop
   
   .Close
  
  End With
End Sub


Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)

    With Item
    
    txtItemNum.Text = .Text
    txtItemName.Text = .SubItems(ItemNameIndex)
    txtItemPrice.Text = .SubItems(ItemPriceIndex)

    End With
    
End Sub

Private Sub txtItemPrice_KeyPress(KeyAscii As Integer)
KeyAscii = OnlyNumericKeys(KeyAscii, txtItemPrice)
End Sub
