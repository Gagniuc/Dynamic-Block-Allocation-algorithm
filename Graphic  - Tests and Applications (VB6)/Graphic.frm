VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Dynamic data block allocation for DNA sequences"
   ClientHeight    =   11775
   ClientLeft      =   675
   ClientTop       =   1080
   ClientWidth     =   16695
   LinkTopic       =   "Form1"
   ScaleHeight     =   11775
   ScaleWidth      =   16695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      Caption         =   "Processing state"
      Height          =   975
      Left            =   240
      TabIndex        =   42
      Top             =   6000
      Width           =   5055
      Begin VB.Label PS 
         Caption         =   "Processing state: 100 %"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   43
         Top             =   360
         Width           =   4575
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Block size/hits"
      Height          =   11535
      Left            =   14280
      TabIndex        =   40
      Top             =   120
      Width           =   2295
      Begin VB.ListBox hits_rem 
         Appearance      =   0  'Flat
         Height          =   3540
         Left            =   240
         TabIndex        =   44
         Top             =   7560
         Width           =   1815
      End
      Begin VB.ListBox Hit 
         Appearance      =   0  'Flat
         Height          =   6270
         Left            =   240
         TabIndex        =   41
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "All possible blocks:"
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label9 
         Caption         =   "Filtered blocks (different from zero):"
         Height          =   495
         Left            =   240
         TabIndex        =   46
         Top             =   7080
         Width           =   1815
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Method used"
      Height          =   1215
      Left            =   240
      TabIndex        =   31
      Top             =   2280
      Width           =   5055
      Begin VB.OptionButton Option1 
         Caption         =   "Multi Brute Force Method"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   33
         Top             =   720
         Value           =   -1  'True
         Width           =   4335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Double Brute Force Method"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Block distribution"
      Height          =   5655
      Left            =   5400
      TabIndex        =   24
      Top             =   6000
      Width           =   8655
      Begin VB.CheckBox Use_Circle 
         Caption         =   "Use circles for values"
         Height          =   255
         Left            =   3720
         TabIndex        =   45
         Top             =   5160
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CommandButton EO1 
         Caption         =   "Erase output"
         Height          =   255
         Left            =   6840
         TabIndex        =   26
         Top             =   360
         Width           =   1455
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4335
         Left            =   600
         ScaleHeight     =   287
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   511
         TabIndex        =   25
         Top             =   720
         Width           =   7695
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   360
         TabIndex        =   30
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Max_ScaleY2 
         BackStyle       =   0  'Transparent
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   4920
         Width           =   615
      End
      Begin VB.Label Max_len2 
         Caption         =   "1000b"
         Height          =   255
         Left            =   8040
         TabIndex        =   28
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label Min_len2 
         Caption         =   "10b"
         Height          =   255
         Left            =   480
         TabIndex        =   27
         Top             =   5280
         Width           =   375
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00000000&
         X1              =   8280
         X2              =   8280
         Y1              =   5160
         Y2              =   5040
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         X1              =   600
         X2              =   600
         Y1              =   5160
         Y2              =   5040
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00000000&
         X1              =   600
         X2              =   480
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00000000&
         X1              =   480
         X2              =   600
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Preset (shows different allocation patterns of data blocks)"
      Height          =   4575
      Left            =   240
      TabIndex        =   23
      Top             =   7080
      Width           =   5055
      Begin VB.CommandButton Preset_6 
         Caption         =   "Preset 6"
         Height          =   495
         Left            =   240
         TabIndex        =   48
         Top             =   3840
         Width           =   4575
      End
      Begin VB.CommandButton Preset_5 
         Caption         =   "Preset 5"
         Height          =   495
         Left            =   240
         TabIndex        =   38
         Top             =   3240
         Width           =   4575
      End
      Begin VB.CommandButton Preset_4 
         Caption         =   "Preset 4"
         Height          =   495
         Left            =   240
         TabIndex        =   37
         Top             =   2640
         Width           =   4575
      End
      Begin VB.CommandButton Preset_3 
         Caption         =   "Preset 3"
         Height          =   495
         Left            =   240
         TabIndex        =   36
         Top             =   2040
         Width           =   4575
      End
      Begin VB.CommandButton Preset_2 
         Caption         =   "Preset 2"
         Height          =   495
         Left            =   240
         TabIndex        =   35
         Top             =   1440
         Width           =   4575
      End
      Begin VB.CommandButton Preset_1 
         Caption         =   "Preset 1"
         Height          =   495
         Left            =   240
         TabIndex        =   34
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label preset_msg 
         Caption         =   "No preset !"
         Height          =   615
         Left            =   360
         TabIndex        =   39
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Run"
      Height          =   2295
      Left            =   240
      TabIndex        =   19
      Top             =   3600
      Width           =   5055
      Begin VB.OptionButton PointLine 
         Caption         =   "Plot points"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   22
         Top             =   720
         Value           =   -1  'True
         Width           =   3255
      End
      Begin VB.OptionButton PointLine 
         Caption         =   "Plot lines"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   21
         Top             =   360
         Width           =   3255
      End
      Begin VB.CommandButton Process 
         Caption         =   "Allocate data blocks for random DNA sequences !"
         Height          =   735
         Left            =   360
         TabIndex        =   20
         Top             =   1200
         Width           =   4455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Parameters"
      Height          =   2055
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   5055
      Begin VB.HScrollBar Min_Block 
         Height          =   255
         Left            =   1080
         Max             =   100
         Min             =   1
         TabIndex        =   13
         Top             =   480
         Value           =   9
         Width           =   3135
      End
      Begin VB.HScrollBar Last_Block 
         Height          =   255
         Left            =   1080
         Max             =   100
         Min             =   1
         TabIndex        =   12
         Top             =   1440
         Value           =   3
         Width           =   3135
      End
      Begin VB.HScrollBar Max_Block 
         Height          =   255
         Left            =   1080
         Max             =   1000
         Min             =   100
         TabIndex        =   11
         Top             =   840
         Value           =   1000
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Min Block"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Last Block"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Max Block"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   855
      End
      Begin VB.Label MinB 
         Caption         =   "-"
         Height          =   255
         Left            =   4440
         TabIndex        =   16
         Top             =   480
         Width           =   375
      End
      Begin VB.Label MaxB 
         Caption         =   "-"
         Height          =   255
         Left            =   4440
         TabIndex        =   15
         Top             =   840
         Width           =   375
      End
      Begin VB.Label LB 
         Caption         =   "-"
         Height          =   255
         Left            =   4440
         TabIndex        =   14
         Top             =   1440
         Width           =   375
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "X = block size, Y = Sequence length"
      Height          =   5775
      Left            =   5400
      TabIndex        =   0
      Top             =   120
      Width           =   8655
      Begin VB.VScrollBar MaxScale 
         Height          =   3735
         Left            =   240
         Max             =   1000
         Min             =   4
         TabIndex        =   8
         Top             =   960
         Value           =   200
         Width           =   255
      End
      Begin VB.HScrollBar LenSeq 
         Height          =   255
         Left            =   960
         Max             =   30000
         Min             =   20
         TabIndex        =   7
         Top             =   5280
         Value           =   30000
         Width           =   6735
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4335
         Left            =   600
         ScaleHeight     =   287
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   503
         TabIndex        =   2
         Top             =   720
         Width           =   7575
      End
      Begin VB.CommandButton EO 
         Caption         =   "Erase output"
         Height          =   255
         Left            =   6720
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00000000&
         X1              =   480
         X2              =   600
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         X1              =   600
         X2              =   480
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         X1              =   600
         X2              =   600
         Y1              =   5160
         Y2              =   5040
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00000000&
         X1              =   8160
         X2              =   8160
         Y1              =   5160
         Y2              =   5040
      End
      Begin VB.Label Min_len 
         Caption         =   "10b"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   5280
         Width           =   375
      End
      Begin VB.Label Max_len 
         Caption         =   "1000b"
         Height          =   255
         Left            =   7920
         TabIndex        =   5
         Top             =   5280
         Width           =   495
      End
      Begin VB.Label Max_ScaleY 
         BackStyle       =   0  'Transparent
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   4920
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   255
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   __________________________________                          _________
'  /   Dynamic data block allocation  \________________________/         |
' |                                                                      |
' |     Description:  Dynamic Data Block Allocation                      |
' |        Category:  Open source                                        |
' |          Author:  Paul Gagniuc                                       |
' |                                                                      |
' |    Date Created:  July 2010                                          |
' |          Update:  July 2021                                          |
' |       Tested On:  Win98, WinXP, WinVista, Win7, Win8, Win10          |
' |           Email:  paul_gagniuc@acad.ro                               |
' |                                                                      |
' |           Notes:  Graphic  - Tests and Applications                  |
' |                  _____________________________                       |
' |_________________/                             \______________________|
'

Option Explicit

Dim tmpx As Variant
Dim tmpy As Variant
Dim tmpxx As Variant
Dim tmpyy As Variant


Private Sub Form_Load()
    Max_ScaleY.Caption = MaxScale.Value
    Max_len.Caption = LenSeq.Value
    Min_len.Caption = LenSeq.Min

    LB.Caption = Last_Block.Value
    MinB.Caption = Min_Block.Value
    MaxB.Caption = Max_Block.Value
End Sub



Private Sub Process_Click()
Dim yy, xx, yy2, xx2, mai_mare, CountAloc, max_use As Variant
Dim o, L, m, w As Variant

    Preset_1.Enabled = False
    Preset_2.Enabled = False
    Preset_3.Enabled = False
    Preset_4.Enabled = False
    Preset_5.Enabled = False
    Preset_6.Enabled = False

    yy = Picture1.ScaleHeight / MaxScale.Value
    xx = Picture1.ScaleWidth / LenSeq.Value
    tmpx = 0
    tmpy = 0

    Hit.Clear
    hits_rem.Clear

    '-------------------------------
    For o = 1 To MaxScale.Value
        Hit.AddItem 0
    Next o
    '-------------------------------

    For L = LenSeq.Min To LenSeq.Value
        '-------------------------------
        DoEvents
        PS.Caption = "Processing state: " & Int((100 / LenSeq.Value) * L) & " %"
        '-------------------------------
        If Option1(0).Value = True Then
            m = Block_Alocation1(L, Last_Block.Value, Min_Block.Value, Max_Block.Value)
        Else
            m = Block_Alocation2(L, Last_Block.Value, Min_Block.Value, Max_Block.Value)
        End If
        '-------------------------------

        If m > MaxScale.Value Then
            mai_mare = mai_mare + 1
            GoTo 3
        End If
        
        '-------------------------------
        If m <= Hit.ListCount - 1 Then Hit.List(m) = Hit.List(m) + 1
3:
            
            If PointLine(0).Value Then
                Picture1.Line (tmpx, tmpy)-(xx * L, Val(yy * m)), vbRed
                tmpx = (xx * L)
                tmpy = (Val(yy * m))
            Else
                Picture1.PSet (xx * L, yy * m), vbRed
            End If
        '-------------------------------
    Next L

    '-------------------------------
    For w = 1 To Hit.ListCount
        If Val(Hit.List(w)) > 0 Then
            CountAloc = CountAloc + 1
            hits_rem.AddItem Hit.List(w)
            If max_use < Val(Hit.List(w)) Then max_use = Val(Hit.List(w))
        End If
    Next w
    '-------------------------------
    yy2 = Picture2.ScaleHeight / max_use
    xx2 = Picture2.ScaleWidth / CountAloc

    tmpxx = 0
    tmpyy = 0

    For L = 0 To CountAloc '- 1

        Picture2.Line (tmpxx, tmpyy)-(xx2 * L, Val(yy2 * Val(hits_rem.List(L)))), vbBlue

        If Use_Circle.Value = 1 Then Picture2.Circle (tmpxx, tmpyy), 2

        tmpxx = xx2 * L
        tmpyy = Val(yy2 * Val(hits_rem.List(L)))

    Next L
    '-------------------------------

    Preset_1.Enabled = True
    Preset_2.Enabled = True
    Preset_3.Enabled = True
    Preset_4.Enabled = True
    Preset_5.Enabled = True
    Preset_6.Enabled = True

End Sub

'Multi Brute Force Algorithm (MBFA)
Function Block_Alocation2(ByVal L As Variant, ByVal Last_Block As Variant, ByVal MinBlock As Variant, ByVal MaxBlock As Variant) As Variant
    Dim RestetBlock As Variant
    Dim q, t As Variant

    t = 1
    RestetBlock = MinBlock

    q = L - Last_Block

1:
    Do Until t = 0 Or MinBlock > MaxBlock
        MinBlock = MinBlock + 1
        t = q Mod MinBlock
    Loop


    If MinBlock > MaxBlock Or q < RestetBlock Then
        q = q - 1
        MinBlock = RestetBlock
        If q < RestetBlock Then GoTo 2
        GoTo 1
    End If

2:
    Block_Alocation2 = MinBlock
End Function

'Double Brute Force Algorithm (DBFA)
Function Block_Alocation1(ByVal L As Variant, ByVal Last_Block As Variant, ByVal MinBlock As Variant, ByVal MaxBlock As Variant) As Variant
    Dim a, t, b, m, v, r As Integer

    a = 1
    t = 1
    b = 1
    m = MinBlock

    Do Until t > Last_Block And v = 0
        a = a + 1
        t = (L Mod a)
        r = (L - t)
        v = r Mod 2
    Loop

    Do Until b = 0 Or m >= MaxBlock
        m = m + 1
        b = r Mod m
    Loop

    Block_Alocation1 = m
End Function


Private Sub MaxScale_Change()
    Max_ScaleY.Caption = MaxScale.Value
End Sub

Private Sub Last_Block_Change()
    LB.Caption = Last_Block.Value
End Sub

Private Sub Max_Block_Change()
    MaxB.Caption = Max_Block.Value
End Sub

Private Sub Min_Block_Change()
    MinB.Caption = Min_Block.Value
End Sub

Private Sub EO_Click()
    Picture1.Cls
End Sub

Private Sub EO1_Click()
    Picture2.Cls
End Sub

Private Sub LenSeq_Change()
    Max_len.Caption = LenSeq.Value
    Min_len.Caption = LenSeq.Min
End Sub

Private Sub Preset_1_Click()
'Param
    Min_Block.Value = 5
    Max_Block.Value = 1000
    Last_Block.Value = 3
'Methods
    Option1(0).Value = True
    Option1(1).Value = False
'Run
    PointLine(0).Value = False
    PointLine(1).Value = True
'Graph
    MaxScale.Value = 20
    LenSeq.Value = 30000

    Max_len2.Caption = LenSeq.Value
    Max_ScaleY2.Caption = MaxScale.Value
'Msg
    preset_msg.Caption = "Preset_1"
End Sub

Private Sub Preset_2_Click()
'Param
    Min_Block.Value = 5
    Max_Block.Value = 1000
    Last_Block.Value = 3
'Methods
    Option1(0).Value = True
    Option1(1).Value = False
'Run
    PointLine(0).Value = True
    PointLine(1).Value = False
'Graph
    MaxScale.Value = 20
    LenSeq.Value = 300

    Max_len2.Caption = LenSeq.Value
    Max_ScaleY2.Caption = MaxScale.Value
'Msg
    preset_msg.Caption = "Preset_2"
End Sub

Private Sub Preset_3_Click()
'Param
    Min_Block.Value = 9
    Max_Block.Value = 1000
    Last_Block.Value = 3
'Methods
    Option1(0).Value = False
    Option1(1).Value = True
'Run
    PointLine(0).Value = False
    PointLine(1).Value = True
'Graph
    MaxScale.Value = 1000
    LenSeq.Value = 30000

    Max_len2.Caption = LenSeq.Value
    Max_ScaleY2.Caption = MaxScale.Value
'Msg
    preset_msg.Caption = "Preset_3"
End Sub

Private Sub Preset_4_Click()
'Param
    Min_Block.Value = 65
    Max_Block.Value = 1000
    Last_Block.Value = 3
'Methods
    Option1(0).Value = True
    Option1(1).Value = False
'Run
    PointLine(0).Value = False
    PointLine(1).Value = True
'Graph
    MaxScale.Value = 1000
    LenSeq.Value = 30000

    Max_len2.Caption = LenSeq.Value
    Max_ScaleY2.Caption = MaxScale.Value
'Msg
    preset_msg.Caption = "Preset_4"
End Sub

Private Sub Preset_5_Click()
'Param
    Min_Block.Value = 65
    Max_Block.Value = 1000
    Last_Block.Value = 3
'Methods
    Option1(0).Value = False
    Option1(1).Value = True
'Run
    PointLine(0).Value = False
    PointLine(1).Value = True
'Graph
    MaxScale.Value = 1000
    LenSeq.Value = 30000

    Max_len2.Caption = LenSeq.Value
    Max_ScaleY2.Caption = MaxScale.Value
'Msg
    preset_msg.Caption = "Preset_5"
End Sub

Private Sub Preset_6_Click()
'Param
    Min_Block.Value = 65
    Max_Block.Value = 1000
    Last_Block.Value = 14
'Methods
    Option1(0).Value = True
    Option1(1).Value = False
'Run
    PointLine(0).Value = False
    PointLine(1).Value = True
'Graph
    MaxScale.Value = 1000
    LenSeq.Value = 30000

    Max_len2.Caption = LenSeq.Value
    Max_ScaleY2.Caption = MaxScale.Value
'Msg
    preset_msg.Caption = "Preset_6"
End Sub

