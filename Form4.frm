VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item For Today"
   ClientHeight    =   10845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13425
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10845
   ScaleWidth      =   13425
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00404040&
      Height          =   735
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   13365
      TabIndex        =   26
      Top             =   10110
      Width           =   13425
      Begin VB.CommandButton Command6 
         Caption         =   "Confirm"
         Height          =   495
         Left            =   5880
         TabIndex        =   27
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.ListBox List2 
      Height          =   5130
      Left            =   7800
      TabIndex        =   22
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ListBox List3 
      Height          =   5130
      Left            =   9000
      TabIndex        =   21
      Top             =   1080
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   14640
      Top             =   1680
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
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   9720
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Edit"
      Enabled         =   0   'False
      Height          =   2895
      Left            =   240
      TabIndex        =   10
      Top             =   6360
      Width           =   6495
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   405
         Left            =   5040
         TabIndex        =   20
         Top             =   1800
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Ok"
         Height          =   495
         Left            =   1560
         TabIndex        =   18
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   3240
         TabIndex        =   17
         Top             =   2160
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   480
         Width           =   5535
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1200
         TabIndex        =   11
         Top             =   1200
         Width           =   4095
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Quantity"
         Height          =   495
         Left            =   4920
         TabIndex        =   16
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Price"
         Height          =   495
         Left            =   2640
         TabIndex        =   14
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Item Name"
         Height          =   255
         Left            =   2640
         TabIndex        =   13
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete"
      Height          =   495
      Left            =   10200
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   9600
      Width           =   6255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit"
      Height          =   495
      Left            =   10200
      TabIndex        =   6
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add to List"
      Height          =   495
      Left            =   9360
      TabIndex        =   5
      Top             =   9240
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   7200
      TabIndex        =   3
      Top             =   7680
      Width           =   5775
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   7200
      TabIndex        =   1
      Top             =   6840
      Width           =   5775
   End
   Begin VB.ListBox List1 
      Height          =   5130
      Left            =   2520
      TabIndex        =   0
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Items"
      Height          =   495
      Left            =   4560
      TabIndex        =   25
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   495
      Left            =   7680
      TabIndex        =   24
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      Height          =   495
      Left            =   9000
      TabIndex        =   23
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Add Items"
      Height          =   495
      Left            =   4800
      TabIndex        =   15
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Search:"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   9360
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      Height          =   495
      Left            =   9600
      TabIndex        =   4
      Top             =   8160
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
      Height          =   255
      Left            =   9600
      TabIndex        =   2
      Top             =   7440
      Width           =   1215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_NoScroll As Boolean

Private Sub Command1_Click()
If Text1.Text = "" Then
        MsgBox "You must never leave a blank.", vbExclamation, "Error"
        Exit Sub
End If
    
If Text2.Text = "" Then
        MsgBox "You must never leave a blank.", vbExclamation, "Error"
        Exit Sub
End If
If Combo1.Text = "" Then
        MsgBox "Select an item to delete from the box.", vbExclamation, "Error"
        Exit Sub
End If
        Adodc1.RecordSource = "SELECT * FROM tblItem"
        Adodc1.Refresh
        Adodc1.Recordset.AddNew
        Adodc1.Recordset(0) = Text1.Text
        Adodc1.Recordset(1) = Text2.Text
        Adodc1.Recordset(2) = Combo1.Text
        Adodc1.Recordset.Update
        Text1.Text = ""
        Text2.Text = ""
Call Form_Load
End Sub

Private Sub Command2_Click()
Frame1.Enabled = True
Text1.Enabled = False
Text2.Enabled = False
Combo1.Enabled = False
Command1.Enabled = False
End Sub

Private Sub Command3_Click()
If List1.ListIndex = -1 Then
    MsgBox "Select an item to delete from the list box.", vbExclamation, "Error"
    Exit Sub
Else
    Adodc1.RecordSource = "SELECT * FROM tblItem WHERE Item = '" & List1.Text & "'"
    Adodc1.Recordset.Delete
End If
Text1.Text = ""
Text2.Text = ""
Label4.Caption = ""
Text4.Text = ""
Text6.Text = ""
Text5.Text = ""
Call Form_Load
End Sub

Private Sub Command4_Click()
Adodc1.RecordSource = "SELECT * FROM tblItem"
Adodc1.Refresh
    With Adodc1.Recordset
    .Fields("Item") = Text4.Text
     .Fields("Quantity") = Text6.Text
     .Fields("Price") = Text5.Text
     .Update
    Call Form_Load
    End With
End Sub

Private Sub Command5_Click()
Frame1.Enabled = False
Text1.Text = ""
Text2.Text = ""
Text6.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Frame1.Enabled = True
Text1.Enabled = True
Text2.Enabled = True
Combo1.Enabled = True
Command1.Enabled = True
Call Form_Load
End Sub

Private Sub Command6_Click()
Form1.Show
Form4.Visible = False
End Sub
Private Sub Form_Unload(Cancel As Integer)
X = MsgBox("Do you want to logout?", vbQuestion + vbYesNo, "Message")
If X = vbYes Then
    frmLogin.Show
    Form1.Visible = False
    frmLogin.txtusername.Text = ""
    frmLogin.txtpassword.Text = ""
End If
End Sub

Private Sub Form_Load()
Form4.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Items.mdb;Persist Security Info=False"
Adodc1.RecordSource = "SELECT * FROM tblItem"
Adodc1.Refresh
List1.Clear
List2.Clear
List3.Clear
Do While Not (Adodc1.Recordset.EOF)
    List1.AddItem (Adodc1.Recordset.Fields(0))
    List2.AddItem (Adodc1.Recordset.Fields(1))
    List3.AddItem (Adodc1.Recordset.Fields(2))
    Adodc1.Recordset.MoveNext
Loop


Dim quantity As Byte
quantity = 1

Do While quantity <= 50
Combo1.AddItem (quantity)
quantity = quantity + 1
Loop
End Sub

Private Sub List1_Click()
Adodc1.RecordSource = "SELECT * FROM tblItem WHERE Item = '" & List1.Text & "'"
Adodc1.Refresh
With Adodc1.Recordset
            Label4.Caption = .Fields(0)
            Text4.Text = .Fields(0)
            Text6.Text = .Fields("Quantity")
            Text5.Text = .Fields("Price")
End With
Text3.Text = ""
End Sub

Private Sub Text3_Change()
Adodc1.RecordSource = "SELECT * FROM tblItem WHERE Item = '" & List1.Text & "'"
Adodc1.Refresh
List1.Clear
Do While Not (Adodc1.Recordset.EOF)
    List1.Clear
    List1.AddItem (Adodc1.Recordset.Fields(0))
    Adodc1.Recordset.MoveNext
Loop
End Sub
