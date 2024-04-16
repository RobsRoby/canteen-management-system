VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "comctl32.ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4440
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   9345
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   240
      Negotiate       =   -1  'True
      TabIndex        =   2
      Top             =   4680
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   8880
      Top             =   120
   End
   Begin VB.Label lbladmin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome!"
      BeginProperty Font 
         Name            =   "Forte"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   2760
      TabIndex        =   3
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   3788
      TabIndex        =   0
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   2280
      TabIndex        =   1
      Top             =   3600
      Width           =   4455
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    lbladmin.Caption = frmLogin.txtusername.Text
    frmLogin.Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\StudentDB.mdb;Persist Security Info=False"
End Sub

Private Sub Label4_Click()
Label4.Caption = frmLogin.txtusername.Text
End Sub

Private Sub Timer1_Timer()
     ProgressBar1.Value = ProgressBar1.Value + 1
Select Case ProgressBar1.Value

Case 5

Label2.Caption = "Loading"

Case 10

Label2.Caption = "Loading."

Case 15

Label2.Caption = "Loading.."

Case 20

Label2.Caption = "Loading..."

Case 25

Label2.Caption = "Loading...."

Case 30

Label2.Caption = "Loading....."

Case 40

Label2.Caption = "Loading Forms..."

Case 50

Label2.Caption = "Loading Database..."

Case 70

Label2.Caption = "Establishing Connection..."

Case 80

Label2.Caption = "Connection Established..."

Case 100

Label2.Caption = "Finish..."

End Select

Label3.Caption = ProgressBar1.Value & "%"

If ProgressBar1.Value = 100 Then
Form1.Show
ProgressBar1.Value = 0
Timer1.Enabled = False
    

Unload Me

End If
End Sub
