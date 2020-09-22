VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   2475
      TabIndex        =   6
      Top             =   1680
      Width           =   2535
   End
   Begin VB.PictureBox Picture3 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   1515
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2160
      Top             =   720
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   3075
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   300
      ScaleWidth      =   3780
      TabIndex        =   0
      Top             =   2760
      Width           =   3780
   End
   Begin VB.Label Label4 
      Caption         =   "Random crap:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Date:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Time:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Source Picture:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const PIC_WIDTH = 210
Private Const PIC_HEIGHT = 300

Public Sub DisplayDigital(DisplayPic As PictureBox, Strin As String, SourcePic As PictureBox)
    Dim SPLT() As String
    Dim i As Integer
    
    ReDim SPLT(1 To Len(Strin))
    For i = 1 To Len(Strin)
        SPLT(i) = Mid(Strin, i, 1)
    Next i
    
    DisplayPic.Cls
    DisplayPic.BorderStyle = 0
    DisplayPic.AutoRedraw = True
    
    For i = 1 To UBound(SPLT())
        Select Case UCase(SPLT(i))
            Case "0"
                DisplayPic.PaintPicture SourcePic, CLng((i - 1) * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT, CLng(0 * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT
            Case "1"
                DisplayPic.PaintPicture SourcePic, CLng((i - 1) * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT, CLng(1 * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT
            Case "2"
                DisplayPic.PaintPicture SourcePic, CLng((i - 1) * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT, CLng(2 * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT
            Case "3"
                DisplayPic.PaintPicture SourcePic, CLng((i - 1) * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT, CLng(3 * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT
            Case "4"
                DisplayPic.PaintPicture SourcePic, CLng((i - 1) * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT, CLng(4 * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT
            Case "5"
                DisplayPic.PaintPicture SourcePic, CLng((i - 1) * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT, CLng(5 * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT
            Case "6"
                DisplayPic.PaintPicture SourcePic, CLng((i - 1) * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT, CLng(6 * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT
            Case "7"
                DisplayPic.PaintPicture SourcePic, CLng((i - 1) * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT, CLng(7 * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT
            Case "8"
                DisplayPic.PaintPicture SourcePic, CLng((i - 1) * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT, CLng(8 * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT
            Case "9"
                DisplayPic.PaintPicture SourcePic, CLng((i - 1) * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT, CLng(9 * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT
            Case " "
                DisplayPic.PaintPicture SourcePic, CLng((i - 1) * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT, CLng(10 * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT
            Case ":"
                DisplayPic.PaintPicture SourcePic, CLng((i - 1) * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT, CLng(11 * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT
            Case ";"
                DisplayPic.PaintPicture SourcePic, CLng((i - 1) * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT, CLng(12 * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT
            Case "-"
                DisplayPic.PaintPicture SourcePic, CLng((i - 1) * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT, CLng(13 * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT
            Case "'"
                DisplayPic.PaintPicture SourcePic, CLng((i - 1) * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT, CLng(14 * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT
            Case "."
                DisplayPic.PaintPicture SourcePic, CLng((i - 1) * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT, CLng(15 * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT
            Case ","
                DisplayPic.PaintPicture SourcePic, CLng((i - 1) * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT, CLng(16 * PIC_WIDTH), 0, PIC_WIDTH, PIC_HEIGHT
        End Select
        
    Next i
    DisplayPic.Width = (i - 1) * PIC_WIDTH
    DisplayPic.Height = PIC_HEIGHT
End Sub

Private Sub Form_Load()
    Timer1_Timer
End Sub

Private Sub Timer1_Timer()
    Dim sTIME As String, sSeconds As String, sDate As String

    sSeconds = Mid(Format(Now, "long time"), InStrRev(Format(Now, "long time"), ":"), 3)
    sTIME = Format(Now, "short time") & sSeconds
    sDate = Replace(Format(Now, "short date"), "/", "-")
    
    DisplayDigital Picture2, sTIME, Picture1
    DisplayDigital Picture3, sDate, Picture1
    DisplayDigital Picture4, "8'34;03,22 15:678;", Picture1
End Sub
