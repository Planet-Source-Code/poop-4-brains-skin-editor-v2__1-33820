VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Skin Designer 2"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   13305
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   448
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   887
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar Status 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   38
      Top             =   6465
      Width           =   13305
      _ExtentX        =   23469
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.HScrollBar Hor 
      Height          =   255
      LargeChange     =   20
      Left            =   0
      SmallChange     =   5
      TabIndex        =   19
      Top             =   6120
      Width           =   9255
   End
   Begin VB.VScrollBar Ver 
      Height          =   6135
      LargeChange     =   20
      Left            =   9240
      SmallChange     =   5
      TabIndex        =   18
      Top             =   0
      Width           =   255
   End
   Begin VB.Timer tmrBack 
      Interval        =   500
      Left            =   2880
      Top             =   2160
   End
   Begin TabDlg.SSTab tbProps 
      Height          =   6375
      Left            =   9600
      TabIndex        =   2
      Top             =   0
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   11245
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   529
      TabCaption(0)   =   "Window"
      TabPicture(0)   =   "frmMain.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "clrWLight"
      Tab(0).Control(1)=   "clrWDark"
      Tab(0).Control(2)=   "clrWBack"
      Tab(0).Control(3)=   "sldHeight"
      Tab(0).Control(4)=   "sldWidth"
      Tab(0).Control(5)=   "Label11"
      Tab(0).Control(6)=   "Label10"
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(8)=   "Label8"
      Tab(0).Control(9)=   "Label7"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Title Bar"
      TabPicture(1)   =   "frmMain.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(3)=   "Label4"
      Tab(1).Control(4)=   "Label6"
      Tab(1).Control(5)=   "txtCaption"
      Tab(1).Control(6)=   "clrTCap"
      Tab(1).Control(7)=   "clrTDark"
      Tab(1).Control(8)=   "clrTLight"
      Tab(1).Control(9)=   "clrTBack"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Control Buttons"
      TabPicture(2)   =   "frmMain.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label12"
      Tab(2).Control(1)=   "lstButtons"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Button"
      TabPicture(3)   =   "frmMain.frx":091E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label5"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label13"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label14"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label15"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label16"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "txtBCaption"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "tmrControls"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "clrBBack"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "clrBDark"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "clrBLight"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "clrBCap"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "cmdSetupButton"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "cmdDeleteButton"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "cmdResizeButton"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).ControlCount=   14
      Begin VB.CommandButton cmdResizeButton 
         Caption         =   "Resize Button..."
         Height          =   255
         Left            =   480
         TabIndex        =   39
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton cmdDeleteButton 
         Caption         =   "Delete Button..."
         Height          =   255
         Left            =   480
         TabIndex        =   40
         Top             =   960
         Width           =   1695
      End
      Begin VB.CommandButton cmdSetupButton 
         Caption         =   "Make New Button..."
         Height          =   255
         Left            =   480
         TabIndex        =   37
         Top             =   720
         Width           =   1695
      End
      Begin VB.PictureBox clrBCap 
         Height          =   255
         Left            =   1440
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   36
         Top             =   3840
         Width           =   735
      End
      Begin VB.PictureBox clrBLight 
         Height          =   255
         Left            =   1440
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   34
         Top             =   3360
         Width           =   735
      End
      Begin VB.PictureBox clrBDark 
         Height          =   255
         Left            =   1440
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   32
         Top             =   2880
         Width           =   735
      End
      Begin VB.PictureBox clrBBack 
         Height          =   255
         Left            =   1440
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   30
         Top             =   2400
         Width           =   735
      End
      Begin VB.Timer tmrControls 
         Interval        =   10
         Left            =   2160
         Top             =   840
      End
      Begin VB.TextBox txtBCaption 
         Height          =   285
         Left            =   480
         TabIndex        =   27
         Top             =   1800
         Width           =   1335
      End
      Begin VB.PictureBox clrWLight 
         Height          =   255
         Left            =   -73560
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   25
         Top             =   3260
         Width           =   735
      End
      Begin VB.PictureBox clrWDark 
         Height          =   255
         Left            =   -73560
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   24
         Top             =   2780
         Width           =   735
      End
      Begin VB.PictureBox clrWBack 
         Height          =   255
         Left            =   -73560
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   23
         Top             =   2300
         Width           =   735
      End
      Begin VB.ListBox lstButtons 
         Height          =   2985
         ItemData        =   "frmMain.frx":093A
         Left            =   -74640
         List            =   "frmMain.frx":0947
         Style           =   1  'Checkbox
         TabIndex        =   17
         Top             =   960
         Width           =   1935
      End
      Begin MSComctlLib.Slider sldHeight 
         Height          =   255
         Left            =   -74640
         TabIndex        =   16
         Top             =   1580
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   50
         SmallChange     =   10
         Min             =   1
         Max             =   500
         SelStart        =   1
         Value           =   1
      End
      Begin MSComctlLib.Slider sldWidth 
         Height          =   255
         Left            =   -74640
         TabIndex        =   14
         Top             =   860
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   50
         SmallChange     =   10
         Min             =   1
         Max             =   500
         SelStart        =   1
         Value           =   1
      End
      Begin VB.PictureBox clrTBack 
         BackColor       =   &H00C00000&
         Height          =   255
         Left            =   -73560
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   12
         Top             =   1680
         Width           =   735
      End
      Begin VB.PictureBox clrTLight 
         BackColor       =   &H00FF0000&
         Height          =   255
         Left            =   -73560
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   10
         Top             =   3120
         Width           =   735
      End
      Begin VB.PictureBox clrTDark 
         BackColor       =   &H00000000&
         Height          =   255
         Left            =   -73560
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   9
         Top             =   2640
         Width           =   735
      End
      Begin VB.PictureBox clrTCap 
         BackColor       =   &H00FF0000&
         Height          =   255
         Left            =   -73560
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   8
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtCaption 
         Height          =   285
         Left            =   -74520
         TabIndex        =   3
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label16 
         Caption         =   "Caption"
         Height          =   255
         Left            =   480
         TabIndex        =   35
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "Light Color"
         Height          =   255
         Left            =   480
         TabIndex        =   33
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label Label14 
         Caption         =   "Dark Color"
         Height          =   255
         Left            =   480
         TabIndex        =   31
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label Label13 
         Caption         =   "Back"
         Height          =   255
         Left            =   480
         TabIndex        =   29
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "Enabled Buttons"
         Height          =   255
         Left            =   -74640
         TabIndex        =   28
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label5 
         Caption         =   "Caption"
         Height          =   255
         Left            =   480
         TabIndex        =   26
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Light Color"
         Height          =   255
         Left            =   -74640
         TabIndex        =   22
         Top             =   3260
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Dark Color"
         Height          =   255
         Left            =   -74640
         TabIndex        =   21
         Top             =   2780
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Back"
         Height          =   255
         Left            =   -74640
         TabIndex        =   20
         Top             =   2300
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Height"
         Height          =   255
         Left            =   -74640
         TabIndex        =   15
         Top             =   1340
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Width"
         Height          =   255
         Left            =   -74640
         TabIndex        =   13
         Top             =   620
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Backcolor"
         Height          =   255
         Left            =   -74520
         TabIndex        =   11
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Light Color"
         Height          =   255
         Left            =   -74520
         TabIndex        =   7
         Top             =   3120
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Dark Color"
         Height          =   255
         Left            =   -74520
         TabIndex        =   6
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Caption"
         Height          =   255
         Left            =   -74520
         TabIndex        =   5
         Top             =   2160
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Caption"
         Height          =   255
         Left            =   -74520
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.PictureBox Container 
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      Height          =   6135
      Left            =   0
      ScaleHeight     =   405
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   613
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin VB.Timer tmrScroll 
         Interval        =   10
         Left            =   2400
         Top             =   2040
      End
      Begin VB.PictureBox Board 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   0
         ScaleHeight     =   121
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   121
         TabIndex        =   1
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSkin 
         Caption         =   "New Skin"
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "Load Skin"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Skin"
      End
      Begin VB.Menu mnuBMP 
         Caption         =   "Convert to .BMP"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuClearButtons 
         Caption         =   "Clear Buttons"
      End
      Begin VB.Menu mnuFileAss 
         Caption         =   "Create File Association"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Down As Long, MakeButton As Long, ResizeButton As Long

Function NewSkin()
Skin.Title = "New Skin"
txtCaption.Text = "New Skin"

Skin.bntClose = True
Skin.bntMax = False
Skin.bntMin = False

lstButtons.Selected(0) = False
lstButtons.Selected(1) = False
lstButtons.Selected(2) = True

Skin.Height = 200
Skin.Width = 300

sldHeight.Value = 200
sldWidth.Value = 300

Skin.TColors.Back = RGB(150, 0, 0)
Skin.TColors.Caption = vbRed
Skin.TColors.DarkColor = vbBlack
Skin.TColors.LightColor = vbRed

clrTBack.BackColor = RGB(150, 0, 0)
clrTCap.BackColor = vbRed
clrTDark.BackColor = vbBlack
clrTLight.BackColor = vbRed

Skin.WColors.Back = RGB(0, 0, 150)
Skin.WColors.DarkColor = vbBlack
Skin.WColors.LightColor = RGB(0, 0, 255)

clrWBack.BackColor = RGB(0, 0, 150)
clrWDark.BackColor = vbBlack
clrWLight.BackColor = RGB(0, 0, 255)

clrTBack.BackColor = RGB(150, 0, 0)
clrTDark.BackColor = vbBlack
clrTLight.BackColor = RGB(255, 0, 0)

ClearButtons
End Function

Function UpdateSkinView(Optional ShowSelect As Boolean = True)
Dim I As Long, XS, YS

Board.Cls
Board.Width = Skin.Width
Board.Height = Skin.Height

Board.Line (0, 0)-(Board.Width - 2, Board.Height - 2), Skin.WColors.LightColor, BF
Board.Line (1, 1)-(Board.Width - 1, Board.Height - 1), Skin.WColors.DarkColor, BF
Board.Line (1, 1)-(Board.Width - 2, Board.Height - 2), Skin.WColors.Back, BF

Board.Line (0, 0)-(Board.Width - 2, 15 - 2), Skin.TColors.LightColor, BF
Board.Line (1, 1)-(Board.Width - 1, 15 - 1), Skin.TColors.DarkColor, BF
Board.Line (1, 1)-(Board.Width - 2, 15 - 2), Skin.TColors.Back, BF

Board.ForeColor = Skin.TColors.DarkColor
Board.CurrentX = 3
Board.CurrentY = 2
Board.Print Skin.Title
Board.ForeColor = Skin.TColors.Caption
Board.CurrentX = 4
Board.CurrentY = 2
Board.Print Skin.Title

If Skin.bntClose Then
Board.ForeColor = Skin.TColors.DarkColor
Board.CurrentX = Board.ScaleWidth - 13
Board.CurrentY = 2
Board.Print "X"
Board.ForeColor = Skin.TColors.Caption
Board.CurrentX = Board.ScaleWidth - 12
Board.CurrentY = 2
Board.Print "X"
End If

If Skin.bntMax Then
Board.ForeColor = Skin.TColors.DarkColor
Board.CurrentX = Board.ScaleWidth - 28
Board.CurrentY = 3
Board.Line (Board.CurrentX, Board.CurrentY)-(Board.CurrentX + 10, Board.CurrentY + 10), Board.ForeColor, B
Board.ForeColor = Skin.TColors.LightColor
Board.CurrentX = Board.ScaleWidth - 27
Board.CurrentY = 3
Board.Line (Board.CurrentX, Board.CurrentY)-(Board.CurrentX + 10, Board.CurrentY + 10), Board.ForeColor, B
End If

If Skin.bntMin Then
Board.ForeColor = Skin.TColors.DarkColor
Board.CurrentX = Board.ScaleWidth - 43
Board.CurrentY = 3
Board.Line (Board.CurrentX, Board.CurrentY + 10)-(Board.CurrentX + 10, Board.CurrentY + 10), Board.ForeColor, B
Board.ForeColor = Skin.TColors.LightColor
Board.CurrentX = Board.ScaleWidth - 42
Board.CurrentY = 3
Board.Line (Board.CurrentX, Board.CurrentY + 10)-(Board.CurrentX + 10, Board.CurrentY + 10), Board.ForeColor, B
End If

SelRGB = SelRGB + SelDir
If SelRGB < 120 Then SelDir = 20
If SelRGB > 220 Then SelDir = -20

For I = 1 To 20
If B(I).Act = True Then
Board.DrawWidth = 1
 
If SelectedButton = I And ShowSelect = True Then Board.Line (B(I).X - 5, B(I).Y - 5)-(B(I).X2 + 5, B(I).Y2 + 5), RGB(0, 0, SelRGB), BF

Board.Line (B(I).X, B(I).Y)-(B(I).X2 - 2, B(I).Y2 - 2), B(I).Color.LightColor, BF
Board.Line (B(I).X + 1, B(I).Y + 1)-(B(I).X2 - 1, B(I).Y2 - 1), B(I).Color.DarkColor, BF
Board.Line (B(I).X + 1, B(I).Y + 1)-(B(I).X2 - 2, B(I).Y2 - 2), B(I).Color.Back, BF

Board.ForeColor = B(I).Color.DarkColor
XS = (B(I).X2 - B(I).X) \ 2 - Board.TextWidth(B(I).Caption) \ 2
YS = (B(I).Y2 - B(I).Y) \ 2 - Board.TextHeight(B(I).Caption) \ 2
Board.CurrentX = B(I).X + XS
Board.CurrentY = B(I).Y + YS
Board.Print B(I).Caption

Board.ForeColor = B(I).Color.Caption
XS = ((B(I).X2 - B(I).X) \ 2 - Board.TextWidth(B(I).Caption) \ 2) - 1
YS = ((B(I).Y2 - B(I).Y) \ 2 - Board.TextHeight(B(I).Caption) \ 2) - 1
Board.CurrentX = B(I).X + XS
Board.CurrentY = B(I).Y + YS
Board.Print B(I).Caption
End If
Next I
End Function

Function ChangeColor(pic As PictureBox)
On Error GoTo NoColor:
Dim cColor
'32755

Set cColor = CreateObject("MSCOMDLG.CommonDialog")

cColor.CancelError = True
cColor.ShowColor

pic.BackColor = cColor.Color

Set cColor = Nothing 'remove the object from mem

Exit Function
NoColor:

Set cColor = Nothing
End Function

Private Sub Board_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Down = True

Select Case MakeButton
Case True
MakeNewButton X, Y, X, Y
Case False
If ResizeButton = False Then
If SelectButton(X, Y) = 0 Then Exit Sub
B(SelectedButton).X = X
B(SelectedButton).Y = Y
Else
B(SelectedButton).X = X
B(SelectedButton).Y = Y
End If

End Select
End Sub

Private Sub Board_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Down = False Then Exit Sub

MakeButton = False
B(SelectedButton).X2 = X
B(SelectedButton).Y2 = Y
End Sub

Private Sub Board_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Down = False
MakeButton = False
End Sub

Private Sub clrBBack_Click()
ChangeColor clrBBack
B(SelectedButton).Color.Back = clrBBack.BackColor
End Sub

Private Sub clrBCap_Click()
ChangeColor clrBCap
B(SelectedButton).Color.Caption = clrBCap.BackColor
End Sub

Private Sub clrBDark_Click()
ChangeColor clrBDark
B(SelectedButton).Color.DarkColor = clrBDark.BackColor
End Sub

Private Sub clrBLight_Click()
ChangeColor clrBLight
B(SelectedButton).Color.LightColor = clrBLight.BackColor
End Sub

Private Sub clrTBack_Click()
ChangeColor clrTBack
Skin.TColors.Back = clrTBack.BackColor
End Sub

Private Sub clrTCap_Click()
ChangeColor clrTCap
Skin.TColors.Caption = clrTCap.BackColor
End Sub

Private Sub clrTDark_Click()
ChangeColor clrTDark
Skin.TColors.DarkColor = clrTDark.BackColor
End Sub

Private Sub clrTLight_Click()
ChangeColor clrTLight
Skin.TColors.LightColor = clrTLight.BackColor
End Sub

Private Sub clrWBack_Click()
ChangeColor clrWBack
Skin.WColors.Back = clrWBack.BackColor
End Sub

Private Sub clrWDark_Click()
ChangeColor clrWDark
Skin.WColors.DarkColor = clrWDark.BackColor
End Sub

Private Sub clrWLight_Click()
ChangeColor clrWLight
Skin.WColors.LightColor = clrWLight.BackColor
End Sub

Private Sub cmdDeleteButton_Click()
B(SelectedButton).Act = False
End Sub

Private Sub cmdResizeButton_Click()
ResizeButton = True
MakeButton = False
Stat "Click on the board to resize button..."
End Sub

Private Sub cmdSetupButton_Click()
MakeButton = True
ResizeButton = False
Stat "Click on board to create new button..."
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Container_Click()
SelectedButton = 0
End Sub

Private Sub Form_Load()
mnuSkin_Click
Down = False

If Len(Command) > 2 Then 'if the program has a path to load then load it
LoadSkin Command
Stat "Skin loaded from: " & Command, "Skin Designer 2 - " & Skin.Title
End If

Me.Visible = True
Hor.Enabled = False
Ver.Enabled = False
tmrBack_Timer
Stat "Welcome to Skin Designer 2"
End Sub

Function Stat(str As String, Optional fcap As String = "")   'just a function to quickly set the status bar
If Len(fcap) > 0 Then Me.Caption = fcap
Status.SimpleText = str
End Function

Private Sub Hor_Change()
Board.Left = Hor.Value
End Sub

Private Sub lstButtons_Click()
Select Case lstButtons.ListIndex
Case 0 'min button
Skin.bntMin = lstButtons.Selected(lstButtons.ListIndex)
Case 1 'max button
Skin.bntMax = lstButtons.Selected(lstButtons.ListIndex)
Case 2 'close button
Skin.bntClose = lstButtons.Selected(lstButtons.ListIndex)
End Select
End Sub

Private Sub mnuAbout_Click()
Load frmAbout
frmAbout.Visible = True
End Sub

Private Sub mnuBMP_Click()
On Error GoTo NoFile:
Dim cPath
'32755

Set cPath = CreateObject("MSCOMDLG.CommonDialog")

cPath.Filter = "Bitmap Files (*.bmp*)|*.bmp*"
cPath.CancelError = True
cPath.ShowSave

UpdateSkinView False  'update the skin
SavePicture Board.Image, Replace(cPath.FileName, ".bmp", "") & ".bmp" 'save the skin
Stat "Converted to bitmap... Saved to: " & Replace(cPath.FileName, ".bmp", "") & ".bmp"

Set cPath = Nothing 'remove the object from mem

Exit Sub
NoFile:

Set cPath = Nothing
End Sub

Private Sub mnuClearButtons_Click()
If MsgBox("Are you sure you want to clear all buttons from the skin?", vbYesNo, "Buttons") = vbNo Then Exit Sub
ClearButtons
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuFileAss_Click()
If FileExist(App.path & "\" & App.EXEName & ".exe") = False Then MsgBox ("Cannot create association because project is not compiled into an executable"), vbCritical, "Error": Exit Sub

CreateAssociation "SKN", "Skin Designer 2", App.path & "\" & App.EXEName & ".exe"
End Sub

Function FileExist(path As String) As Boolean
On Error GoTo Oops

FileExist = True
Open path For Input As #1
Close path

Exit Function
Oops:
If Err.Number = 53 Then FileExist = False
End Function

Private Sub mnuLoad_Click()
On Error GoTo NoFile:
Dim cPath
'32755

Set cPath = CreateObject("MSCOMDLG.CommonDialog")

cPath.Filter = "Skin Files (*.skn*)|*.skn*"
cPath.CancelError = True
cPath.ShowOpen

LoadSkin cPath.FileName
Stat "Skin loaded from: " & cPath.FileName, "Skin Designer 2 - " & Skin.Title

Set cPath = Nothing 'remove the object from mem

Exit Sub
NoFile:

Set cPath = Nothing
End Sub

Private Sub mnuSave_Click()
On Error GoTo NoFile:
Dim cPath
'32755

Set cPath = CreateObject("MSCOMDLG.CommonDialog")

cPath.Filter = "Skin Files (*.skn*)|*.skn*"
cPath.CancelError = True
cPath.ShowSave
SaveSkin Replace(cPath.FileName, ".skn", "") & ".skn"
Stat "Skin saved to: " & Replace(cPath.FileName, ".skn", "") & ".skn", "Skin Designer 2 - " & Skin.Title

Set cPath = Nothing 'remove the object from mem

Exit Sub
NoFile:

Set cPath = Nothing
End Sub

Private Sub mnuSkin_Click()
Stat "New skin started", "Skin Designer 2 - New Skin"
Container.Enabled = True
NewSkin
End Sub

Private Sub sldHeight_Click()
Skin.Height = sldHeight.Value
End Sub



Private Sub sldWidth_Click()
Skin.Width = sldWidth.Value
End Sub

Private Sub tmrBack_Timer()
UpdateSkinView
End Sub

Private Sub tmrControls_Timer()
Dim Bl As Boolean
Bl = IIf(SelectedButton = 0, False, True)
If Bl = True Then Bl = B(SelectedButton).Act

txtBCaption.Enabled = Bl
clrBBack.Enabled = Bl
clrBCap.Enabled = Bl
clrBDark.Enabled = Bl
clrBLight.Enabled = Bl
cmdResizeButton.Enabled = Bl
cmdDeleteButton.Enabled = Bl
End Sub

Private Sub tmrScroll_Timer()
If Board.Height > Container.ScaleHeight Then
Ver.Enabled = True
Ver.Max = (Container.ScaleHeight - Board.Height)
Else
Ver.Enabled = False
Board.Top = 0
End If

If Board.Width > Container.ScaleWidth Then
Hor.Enabled = True
Hor.Max = (Container.ScaleWidth - Board.Width)
Else
Hor.Enabled = False
Board.Left = 0
End If
End Sub

Private Sub txtBCaption_Change()
B(SelectedButton).Caption = txtBCaption.Text
End Sub

Private Sub txtCaption_Change()
Skin.Title = txtCaption.Text
End Sub

Private Sub Ver_Change()
Board.Top = Ver.Value
End Sub
