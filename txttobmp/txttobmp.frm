VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   Caption         =   "TXT TO PICTURE"
   ClientHeight    =   5160
   ClientLeft      =   4980
   ClientTop       =   5010
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prog2 
      Height          =   135
      Left            =   120
      TabIndex        =   6
      Top             =   4920
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar prog 
      Height          =   135
      Left            =   4200
      TabIndex        =   5
      Top             =   4920
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3. Draw From Text"
      Height          =   375
      Left            =   6000
      TabIndex        =   4
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox pb 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawWidth       =   20
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   4200
      ScaleHeight     =   4215
      ScaleWidth      =   3975
      TabIndex        =   3
      Top             =   600
      Width           =   3975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2. create Pic.TXT file"
      Height          =   435
      Left            =   3360
      TabIndex        =   2
      Top             =   90
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1. load picture"
      Height          =   435
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4140
      Left            =   90
      ScaleHeight     =   4140
      ScaleWidth      =   3840
      TabIndex        =   0
      Top             =   630
      Width           =   3840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Path As String

Private Sub Command1_Click()
    'load 16 shades of gray picture to picturebox1
    Dim PictureFile As String
    PictureFile = Path + "pic.bmp"
    Picture1.Picture = LoadPicture(PictureFile)
End Sub

Private Sub Command2_Click()
    prog2.Value = 0
    Dim GLine As String
    Dim PicX As Single
    Dim PicY As Single
    Dim PicColor As Double
    Dim OutFile As String
    
    OutFile = Path + "pic.txt"
    Open OutFile For Output As #1
   
    For PicY = 0 To Picture1.Height Step Screen.TwipsPerPixelY * 2
    '<-- every second line is used because... _
    pixle.width     / pixle.height     = 1 / 1   => TRUE _
    character.width / character.height = 1 / 2   => TRUE
        prog2.Max = 4141
        prog2.Value = prog2.Value + Screen.TwipsPerPixelX * 2
        Debug.Print prog2.Value
        GLine = ""
        For PicX = 0 To Picture1.Width Step Screen.TwipsPerPixelX
            PicColor = Picture1.Point(PicX, PicY) 'get pixle color
            
            Select Case PicColor
            Case &HFFFFFF           'white
                GLine = GLine + " "
            Case &HEEEEEE
                GLine = GLine + "."
            Case &HDDDDDD
                GLine = GLine + ","
            Case &HCCCCCC
                GLine = GLine + ":"
            Case &HBBBBBB
                GLine = GLine + "÷"
            Case &HAAAAAA
                GLine = GLine + "¤"
            Case &H999999
                GLine = GLine + "l"
            Case &H888888
                GLine = GLine + "J"
            Case &H777777
                GLine = GLine + "F"
            Case &H666666
                GLine = GLine + "E"
            Case &H555555
                GLine = GLine + "9"
            Case &H444444
                GLine = GLine + "$"
            Case &H333333
                GLine = GLine + "€"
            Case &H222222
                GLine = GLine + "8"
            Case &H111111
                GLine = GLine + "@"
            Case &H0                'Black
                GLine = GLine + "#"
            End Select
        Next
        Print #1, GLine
    Next
    GLine = GLine + vbCrLf + "end"
    Print #1, GLine
    Close
    x = MsgBox("Done. Now open file " + OutFile + ". To see entire picture...: Select all and reduce font size (to about '4') then come back and redraw the picture using the text ")
    
End Sub

Private Sub Command3_Click()
 prog.Value = 0
 Dim color As String
 Dim piclen As Integer
 Dim spot As String
 txt = App.Path & "\pic.txt"
      Open txt For Input As #1
      
      For up = 0 To 150
      prog.Max = 143
      prog.Value = prog.Value + 1
      y = y + 15
      x = 0
      If tline = "end" Then Exit Sub
      Line Input #1, tline
      piclen = 0
      
      For across = 0 To Len(tline)
      piclen = piclen + 1
      
      On Error GoTo done:
      spot = Mid(tline, piclen, (1))
      
      Select Case spot
            Case " "
                color = "&HFFFFFF"
            Case "."
                color = "&HEEEEEE"
            Case ","
               color = "&HDDDDDD"
            Case ":"
                color = "&HCCCCCC"
            Case "÷"
               color = "&HBBBBBB"
            Case "¤"
                color = "&HAAAAAA"
            Case "l"
               color = "&H999999"
            Case "J"
                color = "&H888888"
            Case "F"
                color = "&H777777"
            Case "E"
                color = "&H666666"
            Case "9"
               color = "&H555555"
            Case "$"
                color = "&H444444"
            Case "€"
                color = "&H333333"
            Case "8"
                color = "&H222222"
            Case "@"
                 color = "&H111111"
            Case "#"                'Black
                color = "&H0"
                End Select
x = x + 10

pb.Line -(x, y), color
Next across
done:
Next up
partyout:
Close #1
End Sub

Private Sub Form_Load()
    Path = App.Path
    If Right$(Path, 1) <> "\" Then Path = Path + "\"
End Sub
