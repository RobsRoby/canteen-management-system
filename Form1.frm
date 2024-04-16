VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Canteen - Today"
   ClientHeight    =   11985
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   14385
   DrawStyle       =   1  'Dash
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11985
   ScaleWidth      =   14385
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7260
      Left            =   2520
      TabIndex        =   26
      Top             =   2520
      Width           =   7335
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   735
      Left            =   3120
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
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
      Caption         =   "Adodc3"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   735
      Left            =   1680
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
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
      Caption         =   "Adodc2"
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
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   14325
      TabIndex        =   18
      Top             =   615
      Visible         =   0   'False
      Width           =   14385
      Begin VB.CommandButton Command9 
         Caption         =   "Exit Search"
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   525
         Left            =   4800
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   120
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Search"
         Height          =   525
         Left            =   9240
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Remove Quantity (1)"
      Height          =   495
      Left            =   12360
      TabIndex        =   17
      Top             =   3240
      Width           =   1815
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
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
   Begin VB.ListBox List2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7260
      Left            =   9840
      TabIndex        =   15
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Delete an Item"
      Height          =   495
      Left            =   360
      TabIndex        =   12
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Items"
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Debt List"
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   3360
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   14325
      TabIndex        =   5
      Top             =   0
      Width           =   14385
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Today's Items"
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5265
         TabIndex        =   6
         Top             =   120
         Width           =   3855
      End
   End
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00404040&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   14325
      TabIndex        =   1
      Top             =   10515
      Width           =   14385
      Begin VB.CommandButton Command1 
         Caption         =   "Sell"
         Height          =   495
         Left            =   6360
         TabIndex        =   7
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Sell as Debt"
         Height          =   495
         Left            =   4200
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add Item"
         Height          =   495
         Left            =   8400
         TabIndex        =   3
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Log Out "
         Height          =   495
         Left            =   12480
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   11610
      Width           =   14385
      _ExtentX        =   25374
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14076
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Picture         =   "Form1.frx":0000
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "3/4/2018"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "9:51 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Emoji"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7680
      TabIndex        =   24
      Top             =   9960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Delete"
      Height          =   375
      Left            =   5760
      TabIndex        =   25
      Top             =   9960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   2520
      TabIndex        =   29
      Top             =   1440
      Width           =   9615
   End
   Begin VB.Label Label12 
      Height          =   615
      Left            =   10080
      TabIndex        =   28
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   12480
      TabIndex        =   27
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label10 
      Caption         =   "Click"
      Height          =   495
      Left            =   9960
      TabIndex        =   23
      Top             =   6240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Height          =   615
      Left            =   10200
      TabIndex        =   22
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity Left"
      Height          =   495
      Left            =   12240
      TabIndex        =   16
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   495
      Left            =   10440
      TabIndex        =   14
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Items"
      Height          =   495
      Left            =   5640
      TabIndex        =   13
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Items Left"
      Height          =   495
      Left            =   0
      TabIndex        =   9
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Menu a 
      Caption         =   "Menu"
      Begin VB.Menu b 
         Caption         =   "Sort by Alphabetical"
      End
      Begin VB.Menu c 
         Caption         =   "Remove all Items"
      End
   End
   Begin VB.Menu d 
      Caption         =   "Print"
      Begin VB.Menu e 
         Caption         =   "Total Sales"
      End
      Begin VB.Menu f 
         Caption         =   "All Items for Today"
      End
      Begin VB.Menu z 
         Caption         =   "Students in Debt"
      End
   End
   Begin VB.Menu g 
      Caption         =   "Tools"
      Begin VB.Menu h 
         Caption         =   "Calculator"
      End
      Begin VB.Menu i 
         Caption         =   "Notepad"
      End
   End
   Begin VB.Menu u 
      Caption         =   "Search"
   End
   Begin VB.Menu v 
      Caption         =   "Log Out"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ans3 As String
Dim HALO As String
Dim ans4 As String
Dim m_NoScroll As Boolean
Dim ans5 As String
Dim ans6 As String
Dim Quant As Integer
Dim Minus As Integer
Dim btnClick As String

Private Sub c_Click()
Adodc1.RecordSource = "SELECT * FROM tblItem"
Adodc1.Refresh
    Do While Not (Adodc1.Recordset.EOF)
    Adodc1.Recordset.Delete
    Adodc1.Recordset.MoveNext
Loop
Call Form_Load
End Sub

Private Sub Command1_Click()
If List1.ListIndex = -1 Then
    MsgBox "Select an item to delete from the list box.", vbExclamation, "Error"
    Exit Sub
Else

ans3 = InputBox("Quantity", "Item", O)
If Not IsNumeric(ans3) Then
MsgBox "Your Input is not a number.", vbCritical, "Input Failed"
Exit Sub
End If

If ans3 = "" Then
Exit Sub
End If

ans3 = Label12.Caption
Dim X As Byte
X = MsgBox("Are you sure to sell this?", vbYesNo + vbQuestion, "Confirm")
    
If X = vbYes Then
    If Label6.Caption = "2" Then
    Minus = Label6.Caption - 1
    Label6.Caption = Minus
    Label6.Caption = "1"
    Call Form_Load
    MsgBox "Marked as Sold!", vbInformation, "Succesful"
    Adodc1.RecordSource = "SELECT * FROM tblItem"
    Adodc1.Refresh
    With Adodc1.Recordset
    .Fields("Quantity") = Label6.Caption
    .Update
    Call Form_Load
    End With
    Exit Sub
    
    Else
     
     If Label6.Caption = "1" Then
     Adodc1.Recordset.Delete
     MsgBox "Marked as Sold!", vbInformation, "Succesful"
     Label6.Caption = ""
     Adodc3.RecordSource = "SELECT * FROM tblSelled"
     Adodc3.Refresh
     Adodc3.Recordset.AddNew
     Adodc3.Recordset(0) = Label11.Caption
     Call Form_Load
     Exit Sub
     Else
     Minus = Label6.Caption - 1
     Label6.Caption = Minus
     End If
    
    End If
End If

If Label6.Caption = "0" Then
    Label6.Caption = ""
    Call Form_Load
    Exit Sub
End If

Adodc3.RecordSource = "SELECT * FROM tblSelled"
Adodc3.Refresh
Adodc3.Recordset.AddNew
Adodc3.Recordset(0) = Label11.Caption
MsgBox "Marked as Sold!", vbInformation, "Succesful"
Adodc1.RecordSource = "SELECT * FROM tblItem"
Adodc1.Refresh
    With Adodc1.Recordset
    .Fields("Quantity") = Label6.Caption
    .Update
    Call Form_Load
    End With

End If
End Sub

Private Sub Command10_Click()
If List1.ListIndex = -1 Then
    MsgBox "Select an item to delete from the list box.", vbExclamation, "Error"
    Exit Sub
Else

X = MsgBox("Remove (1) in this? ", vbYesNo + vbQuestion, "Confirm")
    
If X = vbYes Then
    If Label6.Caption = "2" Then
    Minus = Label6.Caption - 1
    Label6.Caption = Minus
    Label6.Caption = "1"
    Call Form_Load
    MsgBox "Marked as Sold!", vbInformation, "Succesful"
    Adodc1.RecordSource = "SELECT * FROM tblItem"
    Adodc1.Refresh
    With Adodc1.Recordset
    .Fields("Quantity") = Label6.Caption
    .Update
    Call Form_Load
    End With
    Exit Sub
    
    Else
     
     If Label6.Caption = "1" Then
     Adodc1.Recordset.Delete
     MsgBox "The Item was deleted!", vbInformation, "Succesful"
     Label6.Caption = ""
     Adodc3.RecordSource = "SELECT * FROM tblSelled"
     Adodc3.Refresh
     Adodc3.Recordset.AddNew
     Adodc3.Recordset(0) = Label11.Caption
     Call Form_Load
     Exit Sub
     Else
     Minus = Label6.Caption - 1
     Label6.Caption = Minus
     End If
    
    End If
End If

If Label6.Caption = "0" Then
    Label6.Caption = ""
    Call Form_Load
    Exit Sub
End If

MsgBox "Removed Successfully", vbInformation, "Succesful"
Adodc1.RecordSource = "SELECT * FROM tblItem"
Adodc1.Refresh
    With Adodc1.Recordset
    .Fields("Quantity") = Label6.Caption
    .Update
    Call Form_Load
    End With

End If

End Sub

Private Sub Command11_Click()
Label6.Caption = ""
List1.Enabled = True
Label10.Caption = "Click"
Command10.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command6.Enabled = True
Command5.Enabled = True
a.Enabled = True
d.Enabled = True
Command12.Visible = False
Command11.Visible = False
Command11.Left = 6840
End Sub

Private Sub Command12_Click()
If List1.ListIndex = -1 Then
    MsgBox "Select an item to delete from the list box.", vbExclamation, "Error"
    Exit Sub
Else
    Adodc1.Recordset.Delete
End If
End Sub

Private Sub Command2_Click()
Dim ans As String
ans = InputBox("Enter the name of the Student", "Name")

If ans = "" Then
Exit Sub

Else

Dim ans2 As String
ans2 = InputBox("Enter the Quantity", "Quantity", 0)
If Not IsNumeric(ans2) Then
Exit Sub
Else
Minus = Label6.Caption - 1
Label6.Caption = Minus
Adodc1.Refresh
    With Adodc1.Recordset
    .Fields("Quantity") = Label6.Caption
    .Update
    Call Form_Load
    End With
    
Adodc2.RecordSource = "SELECT * FROM tbldebt"
Adodc2.Refresh
Adodc2.Recordset.AddNew
Adodc2.Recordset(3) = Label11.Caption
Adodc2.Recordset(1) = Label9.Caption
Adodc2.Recordset(0) = ans
Adodc2.Recordset(2) = ans2
Adodc2.Recordset.Update

End If
End If
MsgBox "Listed Successfuly!", vbInformation, "Succesful"

If List1.ListIndex = -1 Then
    MsgBox "Select an item to delete from the list box.", vbExclamation, "Error"
    Exit Sub
Else
Adodc1.RecordSource = "SELECT * FROM tblItem"
Adodc1.Refresh
    Dim X As Byte
    X = MsgBox("Are you sure to sell this?", vbYesNo + vbQuestion, "Confirm")
    If X = vbYes Then
    Adodc3.RecordSource = "SELECT * FROM tblSelled"
    Adodc3.Refresh
    Adodc3.Recordset.AddNew
    Adodc3.Recordset(0) = Label11.Caption
    Adodc3.Recordset(1) = Label9.Caption
    Adodc3.Recordset(2) = Label12.Caption
    Adodc3.Recordset.Update
    MsgBox "Marked as Sold!", vbInformation, "Succesful"
        Adodc1.Recordset.Delete
    End If
    Label10.Caption = "Your Items."
    Call Form_Load

End If
Label6.Caption = ""
Label8.Caption = ""
Call Form_Load
End Sub

Private Sub Command3_Click()
ans3 = InputBox("Name of the Item?", "Item")
If ans3 = "" Then
Exit Sub

Else

ans4 = InputBox("How much is this Item?", "Price", 0)
If ans4 = "0" Then
Exit Sub
End If
If Not IsNumeric(ans4) Then
Exit Sub
End If


If IsNumeric(ans4) Then
Hey = InputBox("Enter the Quantity", "Quantity", 0)
    
    If IsNumeric(Hey) Then
     Adodc1.RecordSource = "SELECT * FROM tblItem"
        Adodc1.Refresh
        Adodc1.Recordset.AddNew
        Adodc1.Recordset(0) = ans3
        Adodc1.Recordset(1) = ans4
        Adodc1.Recordset(2) = Hey
        Adodc1.Recordset.Update
        Else
        MsgBox "Your Input is not a number.", vbCritical, "Input Failed"
        Exit Sub
    End If
Else
MsgBox "Your Input is not a number.", vbCritical, "Input Failed"
Exit Sub
End If

End If

MsgBox "Listed Successfuly!", vbInformation, "Succesful"
Call Form_Load

End Sub

Private Sub Command4_Click()
Dim X As Byte
    X = MsgBox("Do you want to DELETE all Items before you log out?", vbQuestion + vbYesNo, "Delete")
If X = vbYes Then
    Adodc1.RecordSource = "SELECT * FROM tblItem"
    Adodc1.Refresh
Do While Not (Adodc1.Recordset.EOF)
    Adodc1.Recordset.Delete
    Adodc1.Recordset.MoveNext
Loop
    frmLogin.Show
    Form1.Visible = False
    frmLogin.txtusername.Text = ""
    frmLogin.txtpassword.Text = ""
End If
If X = vbNo Then
X = MsgBox("Do you want to logout?", vbQuestion + vbYesNo, "Message")
If X = vbYes Then
    frmLogin.Show
    Form1.Visible = False
    frmLogin.txtusername.Text = ""
    frmLogin.txtpassword.Text = ""
End If
End If

End Sub

Private Sub Command5_Click()
Form3.Show
Form1.Visible = False
Adodc1.Enabled = False

End Sub

Private Sub Command6_Click()
Form2.Show
Form1.Visible = False
End Sub

Private Sub Command7_Click()
If Label10.Caption = "Click" Then
Command11.Left = 7680
Command11.Visible = True
Command12.Visible = True
Label10.Caption = "Delete"
Command10.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command6.Enabled = False
Command5.Enabled = False
a.Enabled = False
d.Enabled = False
End If

End Sub

Private Sub Command8_Click()
Adodc1.RecordSource = "SELECT * FROM tblItem WHERE Item = '" & List1.Text & "'"
Adodc1.Refresh
List1.Clear
Do While Not (Adodc1.Recordset.EOF)
    List1.Clear
    List1.AddItem (Adodc1.Recordset.Fields(0))
    Adodc1.Recordset.MoveNext
Loop
End Sub

Private Sub Command9_Click()
Picture3.Visible = False
Command9.Visible = False
Text1.Text = ""
Text1.Visible = False
Command8.Visible = False
Text1.Visible = False
Call Form_Load
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim X As Byte
    X = MsgBox("Do you want to DELETE all Items before you log out?", vbQuestion + vbYesNo, "Delete")
If X = vbYes Then
    Adodc1.RecordSource = "SELECT * FROM tblItem"
    Adodc1.Refresh
Do While Not (Adodc1.Recordset.EOF)
    Adodc1.Recordset.Delete
    Adodc1.Recordset.MoveNext
Loop
    frmLogin.Show
    Form1.Visible = False
    frmLogin.txtusername.Text = ""
    frmLogin.txtpassword.Text = ""
End If
If X = vbNo Then
X = MsgBox("Do you want to logout?", vbQuestion + vbYesNo, "Message")
If X = vbYes Then
    frmLogin.Show
    Form1.Visible = False
    frmLogin.txtusername.Text = ""
    frmLogin.txtpassword.Text = ""
End If
End If

End Sub
Private Sub Form_Load()
Form1.Adodc3.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Selled.mdb;Persist Security Info=False"
Form1.Adodc2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Debt.mdb;Persist Security Info=False"
Form1.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Items.mdb;Persist Security Info=False"
Adodc1.RecordSource = "SELECT * FROM tblItem"
Adodc1.Refresh
List1.Clear
List2.Clear
Do While Not (Adodc1.Recordset.EOF)
    List1.AddItem (Adodc1.Recordset.Fields(0))
    List2.AddItem (Adodc1.Recordset.Fields(1))
    Adodc1.Recordset.MoveNext
Loop
 
Form1.Height = 12720
Command11.Left = 6840
Command9.Visible = False
Text1.Text = ""
Text1.Visible = False
Command8.Visible = False
Text1.Visible = False
StatusBar1.Panels(3).Text = frmLogin.txtusername.Text

End Sub

Private Sub Label8_Click()

End Sub

Private Sub List1_Click()
Adodc1.RecordSource = "SELECT * FROM tblItem WHERE Item = '" & List1.Text & "'"
Adodc1.Refresh
With Adodc1.Recordset
            Label11.Caption = .Fields(0)
            Label6.Caption = .Fields("Quantity")
            Label9.Caption = .Fields("Price")
End With
End Sub
Private Sub List1_Scroll()
    If Not m_NoScroll Then
        m_NoScroll = True
        List2.TopIndex = List1.TopIndex
        m_NoScroll = False
    End If
End Sub

Private Sub List2_Scroll()
    If Not m_NoScroll Then
        m_NoScroll = True
        List1.TopIndex = List2.TopIndex
        m_NoScroll = False
    End If
End Sub
Private Sub u_Click()
Picture3.Visible = True
Command9.Visible = True
Text1.Visible = True
Command8.Visible = True
Text1.Visible = True
End Sub

Private Sub v_Click()
Dim X As Byte
    X = MsgBox("Do you want to DELETE all Items before you log out?", vbQuestion + vbYesNo, "Delete")
If X = vbYes Then
    Adodc1.RecordSource = "SELECT * FROM tblItem"
    Adodc1.Refresh
Do While Not (Adodc1.Recordset.EOF)
    Adodc1.Recordset.Delete
    Adodc1.Recordset.MoveNext
Loop
    frmLogin.Show
    Form1.Visible = False
    frmLogin.txtusername.Text = ""
    frmLogin.txtpassword.Text = ""
End If
If X = vbNo Then
X = MsgBox("Do you want to logout?", vbQuestion + vbYesNo, "Message")
If X = vbYes Then
    frmLogin.Show
    Form1.Visible = False
    frmLogin.txtusername.Text = ""
    frmLogin.txtpassword.Text = ""
End If
End If
End Sub
