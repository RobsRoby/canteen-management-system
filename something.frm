VERSION 5.00
Begin VB.Form frmRegPass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Register Login"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   7845
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrloading 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdexit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdlogin 
      BackColor       =   &H00FFFFFF&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Script MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CheckBox chkshowpass 
      Caption         =   "Show Password"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox txtpass 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2400
      Width           =   3735
   End
   Begin VB.TextBox txtuser 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Register Login"
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
      Height          =   855
      Left            =   1560
      TabIndex        =   7
      Top             =   600
      Width           =   4815
   End
   Begin VB.Label lblpass 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3240
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label lbluser 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "USERNAME"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   3240
      TabIndex        =   0
      Top             =   2040
      Width           =   1335
   End
End
Attribute VB_Name = "frmRegPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Const userN As String = "admin"
    Const passW As String = "password"
    

Private Sub chkshowpass_Click()
    If chkshowpass.Value = 1 Then
        txtpass.PasswordChar = ""
    Else
        txtpass.PasswordChar = "*"
    End If
End Sub

Private Sub cmdexit_Click()
    frmLogin.Show
    Unload Me
End Sub

Private Sub cmdlogin_Click()
    Dim loginSuccessful As Boolean
    
    If txtuser.Text = userN Then
        If txtpass.Text = passW Then
            loginSuccessful = True
            chkshowpass.Enabled = False
            tmrloading.Enabled = True
            frmLogin.Show
            frmLogin.Timer1 = True
            Unload Me
        Else
        
        End If
    Else
    
    End If
    
    If loginSuccessful = False Then MsgBox "Your username or password is incorrect.", vbCritical + vbOKOnly, "Login Error"
End Sub

Private Sub Form_Activate()
    txtuser.SetFocus
    
End Sub

Private Sub Form_Load()
    Call initializeSettings
End Sub

Private Sub initializeSettings()
    txtuser.Text = ""
    txtpass.Text = ""
    chkshowpass.Enabled = True
    chkshowpass.Value = 0
End Sub


