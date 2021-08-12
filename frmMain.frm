VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "ImageCombo Demonstration"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   4350
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageCombo ImageCombo1 
      Height          =   330
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "ImageCombo1"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1080
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()
    Dim ciCurrent As ComboItem
    Dim x As Long
    
    
    Set ImageCombo1.ImageList = ImageList1
    
    
    For x = 1 To 3
        Set ciCurrent = ImageCombo1.ComboItems.Add
        ciCurrent.Text = "Item " & x
        ciCurrent.Image = x
        ciCurrent.Key = "Item " & x
    Next
    
End Sub

Private Sub ImageCombo1_Click()
    MsgBox "You chose the item with key '" & _
           ImageCombo1.SelectedItem.Key & "'", _
           vbExclamation
End Sub
