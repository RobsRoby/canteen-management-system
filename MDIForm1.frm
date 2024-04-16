VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Canteen"
   ClientHeight    =   10170
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   15960
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu a 
      Caption         =   "Menu"
      Begin VB.Menu o 
         Caption         =   "Sort by Alphabetical"
      End
      Begin VB.Menu p 
         Caption         =   "Add Item"
      End
      Begin VB.Menu l 
         Caption         =   "Remove Items"
      End
   End
   Begin VB.Menu c 
      Caption         =   "Print "
      Begin VB.Menu z 
         Caption         =   "Total Sales"
      End
      Begin VB.Menu m 
         Caption         =   "All Items for today"
      End
      Begin VB.Menu n 
         Caption         =   "Students  in Debt"
      End
   End
   Begin VB.Menu b 
      Caption         =   "Search"
   End
   Begin VB.Menu d 
      Caption         =   "Log Out"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
    Form1.Show
    Form2.Visible = False
    Form1.Visible = True
End Sub

Private Sub Command4_Click()
Dim X As Byte
    X = MsgBox("Are you sure you want to logout?", vbQuestion + vbYesNo, "Message")
If X = vbYes Then
    frmLogin.Show
    frmLogin.txtusername.Text = ""
    frmLogin.txtpassword.Text = ""
    MDIForm1.Visible = False
    Form2.Visible = False
End If

End Sub

Private Sub d_Click()
Dim X As Byte
    X = MsgBox("Are you sure you want to logout?", vbQuestion + vbYesNo, "Message")
If X = vbYes Then
    frmLogin.Show
    frmLogin.txtusername.Text = ""
    frmLogin.txtpassword.Text = ""
    MDIForm1.Visible = False
    Form2.Visible = False
End If

End Sub

Private Sub MDIForm_Load()
If WindowState = vbMaximized Then
      WindowState = vbNormal
   ElseIf WindowState = vbNormal Then
      WindowState = vbMaximized
End If

StatusBar1.Panels(3).Text = frmLogin.txtusername.Text
Form2.Show
Form2.SetFocus
Form2.Visible = True
Form1.Visible = False
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Dim X As Integer
    X = MsgBox("Do you want to logout?", vbQuestion + vbYesNo, "Message")
If X = vbYes Then
    frmLogin.Show
    frmLogin.txtusername.Text = ""
    frmLogin.txtpassword.Text = ""
    MDIForm1.Visible = False
    Form2.Visible = False
Else
    Cancel = 1
End If
End Sub
