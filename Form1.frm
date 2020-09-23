VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\ApIMenu.vbp"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5865
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   391
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   457
   StartUpPosition =   3  'Windows Default
   Begin pIMenu.CIMenu IMenu1 
      Height          =   5550
      Left            =   225
      TabIndex        =   0
      Top             =   135
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   9790
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6030
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":27B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4F68
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":771C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9ED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C684
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuA 
      Caption         =   "Edit"
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "Rename"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

Dim Menu As CMenuItem
Dim I As Integer

With IMenu1
    .ItemHeight = 55
    .ButtonHeight = 20
    Set .ImageList = ImageList1
    
For I = 1 To 6
    .Add "Menu" + CStr(I)
    Set Menu = .Item(I)
    Menu.Add "Item1", Rnd * 5 + 1
    Menu.Add "Item2", Rnd * 5 + 1
    Menu.Add "Item3", Rnd * 5 + 1
    Menu.Add "Item4", Rnd * 5 + 1
    Menu.Add "Item5", Rnd * 5 + 1
    Menu.Add "Item6", Rnd * 5 + 1
    
Next
        
    .Redraw
End With


End Sub

Private Sub IMenu1_ItemClick(id As Long, Button As Integer, Shift As Integer)
With IMenu1
Caption = "Menu = " + CStr(.Selected) + "  " + _
          "Item = " + CStr(id)


End With
If Button = 2 Then
   ' PopupMenu mnuA
End If
End Sub

Private Sub mnuRemove_Click()
With IMenu1
    If .ItemCount(.Selected) > 0 Then
    .SubItem.Remove .SelectedItem
    .Redraw
    End If
End With
End Sub

Private Sub mnuRename_Click()
Dim a As String
Dim x As String
x = IMenu1.SubItem.Item(IMenu1.SelectedItem).Caption
a = InputBox("Enter new caption", Rename, x)

IMenu1.SetItemCaption IMenu1.SelectedItem, a

End Sub
