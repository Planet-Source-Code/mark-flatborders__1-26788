VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   7065
   StartUpPosition =   3  'Windows Default
   Begin Progetto1.FlatApply FlatApply14 
      Height          =   255
      Left            =   2160
      TabIndex        =   35
      Top             =   3720
      Width           =   495
      _extentx        =   873
      _extenty        =   450
      line1left       =   16777215
      line1top        =   16777215
      line2right      =   8421504
      line2botton     =   8421504
      effect          =   1
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   120
      TabIndex        =   34
      Text            =   "Combo2"
      Top             =   3720
      Width           =   2535
   End
   Begin Progetto1.FlatApply FlatApply13 
      Height          =   255
      Left            =   2040
      TabIndex        =   33
      Top             =   5280
      Width           =   615
      _extentx        =   1085
      _extenty        =   450
      line1left       =   16777215
      line1top        =   16777215
      line2right      =   8421504
      line2botton     =   8421504
      effect          =   1
   End
   Begin Progetto1.FlatApply FlatApply12 
      Height          =   255
      Left            =   2040
      TabIndex        =   32
      Top             =   4680
      Width           =   615
      _extentx        =   1085
      _extenty        =   450
      line1left       =   8421504
      line1top        =   8421504
      line2right      =   16777215
      line2botton     =   16777215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   120
      TabIndex        =   30
      Top             =   5280
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "InvFlat"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   6480
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   120
      TabIndex        =   29
      Top             =   4680
      Width           =   2535
   End
   Begin Progetto1.FlatApply FlatApply11 
      Height          =   255
      Left            =   2160
      TabIndex        =   28
      Top             =   6000
      Width           =   495
      _extentx        =   873
      _extenty        =   450
      line1left       =   8421504
      line1top        =   8421504
      line2right      =   16777215
      line2botton     =   16777215
   End
   Begin Progetto1.FlatApply FlatApply10 
      Height          =   255
      Left            =   2160
      TabIndex        =   27
      Top             =   3120
      Width           =   495
      _extentx        =   873
      _extenty        =   450
      line1left       =   8421504
      line1top        =   8421504
      line2right      =   16777215
      line2botton     =   16777215
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3240
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0353
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":06A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":07B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0B0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0E5D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1502
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1614
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1967
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Progetto1.FlatApply FlatApply9 
      Height          =   255
      Left            =   6120
      TabIndex        =   25
      Top             =   6000
      Width           =   735
      _extentx        =   1296
      _extenty        =   450
      line1left       =   8421504
      line1top        =   8421504
      line2right      =   16777215
      line2botton     =   16777215
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   975
      Left            =   3120
      TabIndex        =   24
      Top             =   5880
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   1720
      View            =   2
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   8819
      EndProperty
   End
   Begin Progetto1.FlatApply FlatApply8 
      Height          =   255
      Left            =   6120
      TabIndex        =   22
      Top             =   4680
      Width           =   615
      _extentx        =   1085
      _extenty        =   450
      line1left       =   8421504
      line1top        =   8421504
      line2right      =   16777215
      line2botton     =   16777215
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   1215
      Left            =   3120
      TabIndex        =   21
      Top             =   4560
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2143
      _Version        =   393217
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin Progetto1.FlatApply FlatApply7 
      Height          =   255
      Left            =   1560
      TabIndex        =   19
      Top             =   2040
      Width           =   975
      _extentx        =   1720
      _extenty        =   450
      line1left       =   16777215
      line1top        =   16777215
      line2right      =   8421504
      line2botton     =   8421504
      effect          =   1
   End
   Begin Progetto1.FlatApply FlatApply6 
      Height          =   255
      Left            =   1560
      TabIndex        =   17
      Top             =   1560
      Width           =   975
      _extentx        =   1720
      _extenty        =   450
      line1left       =   8421504
      line1top        =   8421504
      line2right      =   16777215
      line2botton     =   16777215
   End
   Begin Progetto1.FlatApply FlatApply5 
      Height          =   255
      Left            =   6120
      TabIndex        =   15
      Top             =   3240
      Width           =   615
      _extentx        =   1085
      _extenty        =   450
      line1left       =   8421504
      line1top        =   8421504
      line2right      =   16777215
      line2botton     =   16777215
   End
   Begin Progetto1.FlatApply FlatApply4 
      Height          =   255
      Left            =   6000
      TabIndex        =   14
      Top             =   2160
      Width           =   615
      _extentx        =   1085
      _extenty        =   450
      line1left       =   8421504
      line1top        =   8421504
      line2right      =   16777215
      line2botton     =   16777215
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   3120
      TabIndex        =   12
      Top             =   1920
      Width           =   3735
   End
   Begin Progetto1.FlatApply FlatApply3 
      Height          =   255
      Left            =   6120
      TabIndex        =   11
      Top             =   360
      Width           =   615
      _extentx        =   1085
      _extenty        =   450
      line1left       =   8421504
      line1top        =   8421504
      line2right      =   16777215
      line2botton     =   16777215
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   3120
      TabIndex        =   9
      Top             =   240
      Width           =   3735
   End
   Begin Progetto1.FlatApply FlatApply2 
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   840
      Width           =   495
      _extentx        =   873
      _extenty        =   450
      line1left       =   16777215
      line1top        =   16777215
      line2right      =   8421504
      line2botton     =   8421504
      effect          =   1
   End
   Begin Progetto1.FlatApply FlatApply1 
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   360
      Width           =   495
      _extentx        =   873
      _extenty        =   450
      line1left       =   8421504
      line1top        =   8421504
      line2right      =   16777215
      line2botton     =   16777215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   360
      Width           =   2295
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   3120
      TabIndex        =   3
      Top             =   3000
      Width           =   3735
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   3240
      Width           =   2535
   End
   Begin VB.HScrollBar Scroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   6000
      Width           =   2535
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Command/Scroll"
      Height          =   195
      Left            =   120
      TabIndex        =   31
      Top             =   4320
      Width           =   1170
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H80000003&
      Height          =   1935
      Left            =   0
      Top             =   4440
      Width           =   2775
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Combo"
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   2880
      Width           =   495
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Treeview/ListView"
      Height          =   195
      Left            =   3120
      TabIndex        =   23
      Top             =   4320
      Width           =   1320
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H80000003&
      Height          =   2535
      Left            =   3000
      Top             =   4440
      Width           =   3975
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H80000003&
      Height          =   1215
      Left            =   0
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Label"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   1320
      Width           =   390
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000003&
      Height          =   1215
      Left            =   0
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "DriveFileBox"
      Height          =   195
      Left            =   3120
      TabIndex        =   13
      Top             =   1680
      Width           =   885
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H80000003&
      Height          =   2415
      Left            =   3000
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "ListBox"
      Height          =   195
      Left            =   3120
      TabIndex        =   10
      Top             =   0
      Width           =   510
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000003&
      Height          =   1455
      Left            =   3000
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TextBox"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   585
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000003&
      Height          =   1095
      Left            =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
  FlatApply1.Effect = IIf((FlatApply1.Effect) = inset, raised, inset)
  FlatApply10.Effect = IIf((FlatApply10.Effect) = inset, raised, inset)
  FlatApply14.Effect = raised
End Sub

Private Sub Form_Load()
 FlatApply1.app3D Text1
 FlatApply2.app3D Text2
 FlatApply3.app3D List1
 FlatApply4.app3D Dir1
 FlatApply5.app3D File1
 FlatApply6.app3D Label6
 FlatApply7.app3D Label7
 FlatApply8.app3D TreeView1
 FlatApply9.app3D ListView1
 FlatApply10.app3D Combo1
 FlatApply11.app3D Scroll1
 FlatApply12.app3D Command2
 FlatApply13.app3D Command3
 FlatApply14.app3D Combo2
 FlatApply10.Effect = raised
 FlatApply14.Effect = inset
 fillTreeList
End Sub

Sub fillTreeList()
 Dim i As Integer
 For i = 1 To 10
   TreeView1.Nodes.Add , , , "Nodo prova..." & i, i, i
   ListView1.ListItems.Add , , "Nodo prova..." & i, i, i
   List1.AddItem "Nodo di prova..." & i
 Next
End Sub
