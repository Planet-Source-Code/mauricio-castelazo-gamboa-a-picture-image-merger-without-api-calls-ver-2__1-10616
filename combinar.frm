VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "If you like this code...Vote for it !!!"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   Icon            =   "combinar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   1230
      ItemData        =   "combinar.frx":0E42
      Left            =   3180
      List            =   "combinar.frx":0E58
      TabIndex        =   4
      Top             =   2640
      Width           =   3105
   End
   Begin VB.TextBox Text1 
      Height          =   765
      Left            =   3180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "combinar.frx":0F02
      Top             =   3930
      Width           =   3135
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   2280
      Left            =   30
      ScaleHeight     =   148
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   2
      Top             =   2430
      Width           =   3060
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2310
      Left            =   3210
      Picture         =   "combinar.frx":1056
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   1
      Top             =   60
      Width           =   3060
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   2310
      Left            =   30
      Picture         =   "combinar.frx":348A
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   0
      Top             =   60
      Width           =   3060
   End
   Begin VB.Label Label1 
      Caption         =   "Double click the option to see how it works."
      Height          =   225
      Left            =   3180
      TabIndex        =   5
      Top             =   2400
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code only wants to show you how to use the pset and point functions
'of a picture box.
'
'With this code, you can invert an image, merge it with a second one and more.
'This code is free, use it everywhere you want, modify it as you need it.
'
'Comments at "castelazo@cyberlatino.com.mx"
'Author: Mauricio Castelazo Gamboa
'HomePage: http://www.cyberlatino.com.mx
'
'If you like the code, go back to Planet Source Code, and vote for it!!!

Private Sub MERGE(INVERTED1 As Boolean, INVERTED2 As Boolean)
    On Error GoTo fin:
    Dim X As Long
    Dim y As Long
    Dim R1 As Integer
    Dim G1 As Integer
    Dim B1 As Integer
    Dim R2 As Integer
    Dim G2 As Integer
    Dim B2 As Integer

    Picture3.Cls
    Picture3.Height = Picture1.Height
    Picture3.Width = Picture1.Width
    For X = 0 To Picture1.ScaleWidth
        DoEvents
        For y = 0 To Picture1.ScaleHeight
            'HERE WE GET THE RGB VALUES OF THE FIRST PICTURE
                GET_COLORS Picture1.Point(X, y), R1, G1, B1, INVERTED1
            'HERE WE GET THE RGB VALUES OF THE SECOND PICTURE
                GET_COLORS Picture2.Point(X, y), R2, G2, B2, INVERTED2
            'WE PUT AN AVERAGE OF BOTH PIXELS IN THE THIRD PICTURE BOX
            'pset puts a pixel on a specified x,y point, with a RGB color.
                Picture3.PSet (X, y), RGB((R1 + R2) / 2, (G2 + G1) / 2, (B2 + B1) / 2)
        Next y
    Next X
Beep
Exit Sub
fin:
TEMP = MsgBox(Err.Description, vbExclamation, "Picture Mixer")
End Sub

Private Sub NEGATIVE_IMAGE(PICTURE As PictureBox)
    On Error GoTo fin:
    Dim X As Long
    Dim y As Long
    Dim R1 As Integer
    Dim G1 As Integer
    Dim B1 As Integer
    Dim R2 As Integer
    Dim G2 As Integer
    Dim B2 As Integer

    Picture3.Cls
    Picture3.Height = Picture1.Height
    Picture3.Width = Picture1.Width
    For X = 0 To Picture1.ScaleWidth
        DoEvents
        For y = 0 To Picture1.ScaleHeight
            'HERE WE GET THE RGB VALUES OF THE PICTURE
                GET_COLORS PICTURE.Point(X, y), R1, G1, B1, True
            'pset puts a pixel on a specified x,y point, with a RGB color.
                Picture3.PSet (X, y), RGB(R1, G1, B1)
        Next y
    Next X
Beep
Exit Sub
fin:
TEMP = MsgBox(Err.Description, vbExclamation, "Picture Mixer")
End Sub


Private Sub GET_COLORS(COLOR As Long, ByRef R As Integer, ByRef G As Integer, ByRef B As Integer, INVERTED As Boolean)
    'here you get the RGB values
    Dim TEMP As Long
    TEMP = (COLOR And 255)
    R = TEMP And 255
    TEMP = Int(COLOR / 256)
    G = TEMP And 255
    TEMP = Int(COLOR / 65536)
    B = TEMP And 255
    'Now, we are going to check if we need to invert an image...
    If INVERTED = True Then
        R = Abs(R - 255)
        G = Abs(G - 255)
        B = Abs(B - 255)
    End If
End Sub

Private Sub Form_Load()
    Dim y As Byte
    y = MsgBox("Thanks for downloading this code, I hope it to be useful for you." & vbCrLf & "If you like this code, please, VOTE FOR IT!!!" & vbCrLf & vbCrLf & "Else, visit my homepage www.cyberlatino.com.mx", vbInformation, "Image merger")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim y As Byte
    y = MsgBox("Remember, if you like this code, please, VOTE FOR IT!!!" & vbCrLf & "at http://www.planet-source-code.com/vb" & vbCrLf & vbCrLf & "Also, visit my homepage www.cyberlatino.com.mx" & vbCrLf & "Or email me at castelazo@cyberlatino.com.mx", vbInformation, "Image merger")
End Sub

Private Sub List1_DblClick()
    On Error GoTo fin:
    Select Case List1.ListIndex
        Case 0:
            MERGE False, False
        Case 1:
            MERGE True, False
        Case 2:
            MERGE False, True
        Case 3:
            MERGE True, True
        Case 4:
            NEGATIVE_IMAGE Picture1
        Case 5:
            NEGATIVE_IMAGE Picture2
    End Select

Exit Sub
fin:
TEMP = MsgBox(Err.Description, vbExclamation, "Picture Mixer")
End Sub
