Attribute VB_Name = "modObjects"
Option Explicit

Type Button
Act As Long

X As Long
Y As Long
X2 As Long
Y2 As Long

Color As Colors
Caption As String

Selected As Long
End Type

Public B(1 To 20) As Button 'i dont think youll be adding many buttons
Public SelectedButton As Long 'the index of the selected button
Public SelRGB As Long, SelDir As Long

Function isOver(X, Y, X1, Y1, X2, Y2) As Boolean
isOver = False
If X > X1 And X < X2 And Y > Y1 And Y < Y2 Then isOver = True
End Function

Function isOverButton(X, Y) As Long
Dim I As Long
isOverButton = 0
For I = 1 To 20
If isOver(X, Y, B(I).X, B(I).Y, B(I).X2, B(I).Y2) = True Then isOverButton = I: Exit For
Next I
End Function

Function SelectButton(X, Y)
If isOverButton(X, Y) = 0 Then Exit Function
Dim I As Long

For I = 1 To 20
B(I).Selected = False
Next I

I = isOverButton(X, Y)
SelectedButton = I

frmMain.txtBCaption.Text = B(I).Caption
frmMain.clrBBack.BackColor = B(I).Color.Back
frmMain.clrBDark.BackColor = B(I).Color.DarkColor
frmMain.clrBLight.BackColor = B(I).Color.LightColor
frmMain.clrBCap.BackColor = B(I).Color.Caption
End Function

Function ClearButtons()
Dim I As Long
For I = 1 To 20
B(I).Act = False
Next I
End Function

Function GetNewest() As Long
Dim I As Long
For I = 1 To 20
If B(I).Act = True Then GetNewest = I
Next I
End Function

Function MakeNewButton(X1, Y1, X2, Y2)
Dim I As Long
For I = 1 To 20
If B(I).Act = False Then
B(I).Act = True
B(I).X = X1
B(I).Y = Y1
B(I).Y2 = Y2
B(I).X2 = X2
B(I).Color.Back = Skin.TColors.Back
B(I).Color.Caption = Skin.TColors.Caption
B(I).Color.DarkColor = Skin.TColors.DarkColor
B(I).Color.LightColor = Skin.TColors.LightColor

frmMain.txtBCaption.Text = B(I).Caption
frmMain.clrBBack.BackColor = B(I).Color.Back
frmMain.clrBDark.BackColor = B(I).Color.DarkColor
frmMain.clrBLight.BackColor = B(I).Color.LightColor
frmMain.clrBCap.BackColor = B(I).Color.Caption

SelectedButton = I
B(I).Selected = True
Exit For
End If
Next I
End Function
