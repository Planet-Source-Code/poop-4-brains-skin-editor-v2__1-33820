Attribute VB_Name = "modFile"
Option Explicit

Type Colors
LightColor As Long
DarkColor As Long
Back As Long
Caption As Long
End Type

Public Type SkinInfo
Title As String
bntMax As Boolean
bntClose As Boolean
bntMin As Boolean

TitleHeight As Long

TColors As Colors
WColors As Colors

Width As Long
Height As Long
End Type

Public Skin As SkinInfo

Function SaveSkin(path As String)
Dim I As Long

Open path For Output As #1
Write #1, Skin.Title, Skin.Width, Skin.Height, Skin.TitleHeight
Write #1, Skin.bntClose, Skin.bntMin, Skin.bntMax
Write #1, Skin.TColors.Back, Skin.TColors.Caption, Skin.TColors.DarkColor, Skin.TColors.LightColor
Write #1, Skin.WColors.Back, Skin.WColors.Caption, Skin.WColors.DarkColor, Skin.WColors.LightColor

For I = 1 To 20
If B(I).Act = True Then
Write #1, B(I).X, B(I).Y, B(I).X2, B(I).Y2, B(I).Caption, B(I).Caption, B(I).Color.Back, B(I).Color.Caption, B(I).Color.DarkColor, B(I).Color.LightColor
End If
Next I

Close #1
End Function

Function LoadSkin(path As String)
Dim I As Long

Open path For Input As #1
Input #1, Skin.Title, Skin.Width, Skin.Height, Skin.TitleHeight
Input #1, Skin.bntClose, Skin.bntMin, Skin.bntMax
Input #1, Skin.TColors.Back, Skin.TColors.Caption, Skin.TColors.DarkColor, Skin.TColors.LightColor
Input #1, Skin.WColors.Back, Skin.WColors.Caption, Skin.WColors.DarkColor, Skin.WColors.LightColor

Do Until EOF(1)
I = I + 1
Input #1, B(I).X, B(I).Y, B(I).X2, B(I).Y2, B(I).Caption, B(I).Caption, B(I).Color.Back, B(I).Color.Caption, B(I).Color.DarkColor, B(I).Color.LightColor
B(I).Act = True
DoEvents
If I > 20 Then Exit Do
Loop

Close #1
End Function
