VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008000&
   Caption         =   "Form2"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5925
   LinkTopic       =   "Form2"
   ScaleHeight     =   6045
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CheckBox Check2 
      DownPicture     =   "Form2.frx":0000
      Height          =   975
      Left            =   1920
      Picture         =   "Form2.frx":0208
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2040
      Width           =   735
   End
   Begin VB.OptionButton Option2 
      DownPicture     =   "Form2.frx":04DA
      Height          =   975
      Left            =   4200
      Picture         =   "Form2.frx":06E2
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2040
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      DownPicture     =   "Form2.frx":09B4
      Height          =   975
      Left            =   3360
      Picture         =   "Form2.frx":0BBC
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2040
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      DownPicture     =   "Form2.frx":0E8E
      Height          =   975
      Left            =   480
      Picture         =   "Form2.frx":1096
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2040
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   100
      ImageHeight     =   100
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":1368
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":1592
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      DownPicture     =   "Form2.frx":17C7
      Height          =   975
      Left            =   4440
      Picture         =   "Form2.frx":19CF
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DownPicture     =   "Form2.frx":1CA1
      Height          =   975
      Index           =   1
      Left            =   3360
      Picture         =   "Form2.frx":1EA9
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DownPicture     =   "Form2.frx":217B
      Height          =   975
      Index           =   0
      Left            =   2520
      Picture         =   "Form2.frx":2383
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
   Begin VB.PictureBox picButton 
      AutoRedraw      =   -1  'True
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1455
      Left            =   720
      Picture         =   "Form2.frx":2655
      ScaleHeight     =   1455
      ScaleWidth      =   1455
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   735
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "Form2.frx":286F
         Top             =   728
         Width           =   855
      End
   End
   Begin VB.Label Label2 
      Caption         =   "CheckBox control"
      Height          =   255
      Index           =   3
      Left            =   1560
      TabIndex        =   15
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Two OptionButtons"
      Height          =   255
      Index           =   5
      Left            =   3480
      TabIndex        =   11
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "CheckBox control - simulating a Command Button"
      Height          =   615
      Index           =   4
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "CommandButton by itself."
      Height          =   495
      Index           =   2
      Left            =   4080
      TabIndex        =   8
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "CommandButton ctrls. in a ctrl array."
      Height          =   495
      Index           =   1
      Left            =   2520
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "PictureBox control - click the red diamond"
      Height          =   495
      Index           =   0
      Left            =   600
      TabIndex        =   6
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   $"Form2.frx":288D
      Height          =   2175
      Left            =   1320
      TabIndex        =   1
      Top             =   3720
      Width           =   3255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Doug Gaede
'Example version 3.2.0
'January 17, 2001

Dim ShapeTheControls As clsTransForm 'make a reference to the class

Private Sub Check1_Click()
'This line forces a CheckBox control to simulate a
'CommandButton control, if you have a need.  It is no longer
'necessary since I fixed real CommandButton drawing in this
'version.
Check1.Value = 0

End Sub

Private Sub Form_Activate()
'In your own application, when you first create the region
'data files, you may need to move some ShapeMe calls
'(currently in Form_Load) into Form_Activate for the regions
'to be created and saved correctly.  You may also need to add
'a temporary DoEvents at the end of Form_Activate.  Then
'you can move all the calls to Form_Load and eliminate the
'DoEvents if you added one.

End Sub

Private Sub Form_Load()
Set ShapeTheControls = New clsTransForm 'instantiate the object from the class

'shape the picturebox
ShapeTheControls.ShapeMe picButton, RGB(255, 255, 255), True, App.Path & "\Form2RegionData2.dat"

'Turns out you CAN shape all the types of buttons in
'Form_Load, but only if you are loading the region from a
'file.  And when you first make the data file (by changing
'True to False) during development, it must be done in
'Form_Activate.  Then you can move it here.  Remember
'to change False back to True.

ShapeTheControls.ShapeMe Command1(0), RGB(255, 255, 255), True, App.Path & "\Form2RegionData1.dat"
ShapeTheControls.ShapeMe Command1(1), RGB(255, 255, 255), True, App.Path & "\Form2RegionData1.dat"
ShapeTheControls.ShapeMe Command2, RGB(255, 255, 255), True, App.Path & "\Form2RegionData1.dat"
ShapeTheControls.ShapeMe Check1, RGB(255, 255, 255), True, App.Path & "\Form2RegionData1.dat"
ShapeTheControls.ShapeMe Check2, RGB(255, 255, 255), True, App.Path & "\Form2RegionData1.dat"
ShapeTheControls.ShapeMe Option1, RGB(255, 255, 255), True, App.Path & "\Form2RegionData1.dat"
ShapeTheControls.ShapeMe Option2, RGB(255, 255, 255), True, App.Path & "\Form2RegionData1.dat"

Set ShapeTheControls = Nothing 'destroy the object
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set Form2 = Nothing 'good practice to free resources VB doesn't normally free when you unload a form!

End Sub

Private Sub picButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Set picButton.Picture = ImageList1.ListImages(2).Picture
Text1 = "In"

End Sub

Private Sub picButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Set picButton.Picture = ImageList1.ListImages(1).Picture
Text1 = "Out"

End Sub
