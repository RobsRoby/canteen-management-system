VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debt List"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14085
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   14085
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   495
      Left            =   5880
      TabIndex        =   20
      Top             =   6360
      Width           =   1215
   End
   Begin VB.ListBox List3 
      Height          =   5130
      Left            =   6600
      TabIndex        =   19
      Top             =   960
      Width           =   1095
   End
   Begin VB.ListBox List2 
      Height          =   5130
      Left            =   5400
      TabIndex        =   18
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Add"
      Height          =   495
      Left            =   4320
      TabIndex        =   15
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Enabled         =   0   'False
      Height          =   3615
      Left            =   7920
      TabIndex        =   8
      Top             =   2520
      Width           =   5895
      Begin VB.CommandButton Command6 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   3240
         TabIndex        =   23
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Height          =   405
         Left            =   2640
         TabIndex        =   16
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Ok"
         Height          =   495
         Left            =   1800
         TabIndex        =   14
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1200
         TabIndex        =   10
         Top             =   1200
         Width           =   3975
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   5655
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Quantity"
         Height          =   495
         Left            =   2520
         TabIndex        =   17
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Item Name"
         Height          =   255
         Left            =   2520
         TabIndex        =   12
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Price"
         Height          =   495
         Left            =   2520
         TabIndex        =   11
         Top             =   1560
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit"
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   6360
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   12480
      Top             =   6360
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
   Begin VB.CommandButton Command3 
      Caption         =   "Paid"
      Height          =   495
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6360
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   5130
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   4935
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   10680
      TabIndex        =   22
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   8400
      TabIndex        =   21
      Top             =   6480
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Debt List"
      Height          =   735
      Left            =   9360
      TabIndex        =   13
      Top             =   360
      Width           =   3495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Quantity"
      Height          =   495
      Left            =   6600
      TabIndex        =   6
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Price"
      Height          =   495
      Left            =   5400
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Student Name"
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Label2"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10560
      TabIndex        =   3
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Need to Pay"
      Height          =   495
      Left            =   10560
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command2_Click()
Frame1.Enabled = True
Command1.Enabled = False
Command5.Enabled = False
Command3.Enabled = False

End Sub

Private Sub Command3_Click()
If List1.ListIndex = -1 Then
    MsgBox "Select an item to delete from the list box.", vbExclamation, "Error"
    Exit Sub
Else
    Adodc1.RecordSource = "SELECT * FROM tbldebt WHERE Debt = '" & List1.Text & "'"
    Adodc1.Recordset.Delete
End If
Call Form_Load
End Sub

Private Sub Command4_Click()
If Text4.Text = "" Then
MsgBox "Enter the Item!", vbCritical, "Edit/Add failed"
Exit Sub
Else

End If
If Text5.Text = "" Then
MsgBox "Enter the Price", vbCritical, "Edit/Add failed"
Exit Sub
Else

End If
If Combo1.Text = "Quantity" Then
MsgBox "Choose the quantity you want.", vbCritical, "Edit/Add failed"
Exit Sub
Else

End If

End Sub

Private Sub Command5_Click()
Frame1.Enabled = True
Text4.Text = ""
Text5.Text = ""
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
Dim ans3 As String
ans3 = InputBox("Enter the Price", "Price", 0)
If Not IsNumeric(ans3) Then
Exit Sub
Else
Adodc1.RecordSource = "SELECT * FROM tbldebt"
Adodc1.Refresh
Adodc1.Recordset.AddNew
Adodc1.Recordset(1) = ans3
Adodc1.Recordset(0) = ans
Adodc1.Recordset(2) = ans2
Adodc1.Recordset.Update
End If
End If
End If

Call Form_Load
End Sub

Private Sub Command6_Click()
Command1.Enabled = True
Command5.Enabled = True
Command3.Enabled = True
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
    Form1.Visible = True
    Form1.Adodc1.Enabled = True
    Unload Me
End Sub

Private Sub List3_Scroll()
    If Not m_NoScroll Then
        m_NoScroll = True
        List1.TopIndex = List3.TopIndex
        List2.TopIndex = List3.TopIndex
        m_NoScroll = False
    End If
End Sub


Private Sub List1_Scroll()
    If Not m_NoScroll Then
        m_NoScroll = True
        List2.TopIndex = List1.TopIndex
        List3.TopIndex = List1.TopIndex
        m_NoScroll = False
    End If
End Sub

Private Sub List2_Scroll()
    If Not m_NoScroll Then
        m_NoScroll = True
        List1.TopIndex = List2.TopIndex
        List3.TopIndex = List2.TopIndex
        m_NoScroll = False
    End If
End Sub

Private Sub Form_Load()
Form3.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Debt.mdb;Persist Security Info=False"
Adodc1.RecordSource = "SELECT * FROM tblDebt"
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

End Sub

Private Sub List1_Click()
Adodc1.RecordSource = "SELECT * FROM tbldebt WHERE Debt = '" & List1.Text & "'"
Adodc1.Refresh
With Adodc1.Recordset
            Text4.Text = .Fields(0)
            Text5.Text = .Fields(1)
            Text6.Text = .Fields(2)
            Label10.Caption = .Fields(1)
            Label11.Caption = .Fields(2)
End With
Val (Label0.Caption) * Val(Labe11.Caption)
End Sub

