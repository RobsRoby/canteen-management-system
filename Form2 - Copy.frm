VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8355
   FillColor       =   &H00400000&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1991.337
   ScaleMode       =   0  'User
   ScaleWidth      =   8355
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   9360
      Top             =   1920
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
   Begin VB.CommandButton cmdreg 
      Caption         =   "Register"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   17
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   4680
      TabIndex        =   14
      Top             =   8760
      Width           =   1455
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   2040
      TabIndex        =   13
      Top             =   8760
      Width           =   1455
   End
   Begin VB.TextBox txtconfirm 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1253
      TabIndex        =   12
      Top             =   7560
      Width           =   5895
   End
   Begin VB.TextBox txtRegPassword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1253
      TabIndex        =   11
      Top             =   6480
      Width           =   5895
   End
   Begin VB.TextBox txtregusername 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1253
      TabIndex        =   10
      Top             =   5400
      Width           =   5895
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5640
      TabIndex        =   5
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3480
      TabIndex        =   4
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox txtpassword 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Franklin Gothic Demi"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   2040
      Width           =   5895
   End
   Begin VB.TextBox txtusername 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Franklin Gothic Demi"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   1080
      Width           =   5895
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   0
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   480
      Top             =   120
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Franklin Gothic Demi"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   1
      Left            =   1898
      TabIndex        =   16
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Canteen Management"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Index           =   0
      Left            =   1560
      TabIndex        =   15
      Top             =   240
      Width           =   4815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2933
      TabIndex        =   9
      Top             =   8160
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3593
      TabIndex        =   8
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3593
      TabIndex        =   7
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Registration"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Index           =   1
      Left            =   2633
      TabIndex        =   6
      Top             =   4440
      Width           =   3135
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   8400
      Y1              =   2070.415
      Y2              =   2070.415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3465
      TabIndex        =   3
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3465
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 
Private Sub cmdCancel_Click(Index As Integer)
    Timer2.Enabled = True
End Sub

Private Sub cmdexit_Click()
End
End Sub



Private Sub cmdlogin_Click()
If txtusername.Text = "" Then
 MsgBox "Enter username/password!", vbCritical, "Login Failed"
Exit Sub
End If

If txtpassword.Text = "" Then
 MsgBox "Enter username/password!", vbCritical, "Login Failed"
Exit Sub
End If

Dim X As Byte
    X = MsgBox("Start the Day?", vbQuestion + vbYesNo, "New Day")
If X = vbYes Then
    Adodc1.RecordSource = "SELECT * FROM tblLogin WHERE Username = '" & txtusername.Text & "' AND Password = '" & txtpassword.Text & "'"
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
        MsgBox "Welcome", vbInformation, "Login Successful"
        frmSplash2.Show
        frmLogin.Visible = False
        
        Else
        MsgBox "Invalid username/password!", vbCritical, "Login Failed"
        txtusername.Text = ""
        txtpassword.Text = ""
    End If
End If
If X = vbNo Then
    Adodc1.RecordSource = "SELECT * FROM tblLogin WHERE Username = '" & txtusername.Text & "' AND Password = '" & txtpassword.Text & "'"
    Adodc1.Refresh
    If Adodc1.Recordset.RecordCount > 0 Then
        MsgBox "Welcome", vbInformation, "Login Successful"
        frmSplash.Show
        frmLogin.Visible = False
        
        Else
        MsgBox "Invalid username/password!", vbCritical, "Login Failed"
        txtusername.Text = ""
        txtpassword.Text = ""
    End If
End If
End Sub

Private Sub cmdreg_Click()
    frmRegPass.Show
    txtregusername.Text = ""
    txtRegPassword.Text = ""
    txtconfirm.Text = ""
    Unload Me
End Sub

Private Sub cmdSubmit_Click(Index As Integer)
        Adodc1.RecordSource = "SELECT * FROM tblLogin WHERE Username = '" & txtregusername.Text & "'"
Adodc1.Refresh

If Adodc1.Recordset.RecordCount > 0 Then
    MsgBox "Username is already in use!", vbCritical, "Registration failed"
    Exit Sub
Else
    Dim ques As Byte
    ques = MsgBox("Are you sure you want to save?", vbQuestion + vbYesNo, "Message")
    
    If ques = vbYes Then
        Adodc1.Recordset.AddNew
        Adodc1.Recordset(0) = txtregusername.Text
        Adodc1.Recordset(1) = txtRegPassword.Text
        Adodc1.Recordset.Update
            MsgBox "New account has been saved.", vbInformation, "Succesful"
            
            Timer2.Enabled = True
    End If
End If
End Sub

Private Sub Form_Load()
    txtpassword.PasswordChar = "*"
    frmLogin.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\StudentDB.mdb;Persist Security Info=False"
End Sub

Private Sub Timer1_Timer()
frmLogin.Height = frmLogin.Height + 40
If frmLogin.Height > 10290 Then
    Timer1.Enabled = False
    frmLogin.Enabled = True
    txtusername.Enabled = False
    txtpassword.Enabled = False
    cmdlogin.Enabled = False
    cmdexit.Enabled = False
    cmdreg.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
frmLogin.Height = frmLogin.Height - 40

If frmLogin.Height < 4680 Then
    Timer2.Enabled = False
    frmLogin.Enabled = True
    txtusername.Enabled = True
    txtpassword.Enabled = True
    cmdlogin.Enabled = True
    cmdexit.Enabled = True
    cmdreg.Enabled = True
    
End If
End Sub

