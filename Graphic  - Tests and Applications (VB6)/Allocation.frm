VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Dynamic data block allocation for DNA sequences"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12810
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   12810
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Parameters"
      Height          =   7575
      Left            =   9240
      TabIndex        =   3
      Top             =   120
      Width           =   3375
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         FillColor       =   &H000000FF&
         ForeColor       =   &H000000FF&
         Height          =   2055
         Left            =   240
         ScaleHeight     =   133
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   189
         TabIndex        =   23
         Top             =   3240
         Width           =   2895
      End
      Begin VB.CommandButton EOP 
         Caption         =   "Erase OutPut"
         Height          =   495
         Left            =   240
         TabIndex        =   20
         Top             =   6840
         Width           =   2895
      End
      Begin VB.CheckBox IH 
         Caption         =   "Info header"
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   5520
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CommandButton Alocate 
         Caption         =   "Dynamic data block allocation"
         Height          =   615
         Left            =   240
         TabIndex        =   17
         Top             =   6120
         Width           =   2895
      End
      Begin VB.HScrollBar MB 
         Height          =   255
         Left            =   960
         Max             =   10
         Min             =   1
         TabIndex        =   15
         Top             =   2280
         Value           =   3
         Width           =   1815
      End
      Begin VB.HScrollBar Max_Block 
         Height          =   255
         Left            =   960
         Max             =   1000
         Min             =   100
         TabIndex        =   8
         Top             =   840
         Value           =   1000
         Width           =   1815
      End
      Begin VB.HScrollBar Last_Block 
         Height          =   255
         Left            =   960
         Max             =   100
         Min             =   1
         TabIndex        =   5
         Top             =   1320
         Value           =   3
         Width           =   1815
      End
      Begin VB.HScrollBar Min_Block 
         Height          =   255
         Left            =   960
         Max             =   100
         Min             =   1
         TabIndex        =   4
         Top             =   480
         Value           =   9
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Visual arrangement of the sequence:"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   120
         X2              =   3240
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label MC 
         Caption         =   "-"
         Height          =   255
         Left            =   2880
         TabIndex        =   19
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Matrix Col's"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label LB 
         Caption         =   "-"
         Height          =   255
         Left            =   2880
         TabIndex        =   12
         Top             =   1320
         Width           =   375
      End
      Begin VB.Label MaxB 
         Caption         =   "-"
         Height          =   255
         Left            =   2880
         TabIndex        =   11
         Top             =   840
         Width           =   375
      End
      Begin VB.Label MinB 
         Caption         =   "-"
         Height          =   255
         Left            =   2880
         TabIndex        =   10
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Max Block"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Last Block"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Min Block"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sequences"
      Height          =   7575
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8895
      Begin VB.TextBox Nucleotide_number 
         Height          =   285
         Left            =   7920
         TabIndex        =   22
         Text            =   "250"
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton GenSeqRND 
         Caption         =   "< Generate"
         Height          =   315
         Left            =   6840
         TabIndex        =   21
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox sequence_output 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6165
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   2
         Top             =   1200
         Width           =   8415
      End
      Begin VB.TextBox sequence_input 
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Text            =   "ACGTCGTTACATTATTCATTCGTACCGATCAGTATCGATCGTAGCTATACGATTCACGTCGTTACATTATTCATTCGTACCGATCAGTATCGATCGTAGCTATACGATTCACACACACACAC"
         Top             =   600
         Width           =   6495
      End
      Begin VB.Label Label5 
         Caption         =   "Output sequence:"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Input_seq 
         Caption         =   "Input sequence:"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   1935
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
' |       Tested On:  Win98, WinXP, WinVista, Win7, Win8                 |
' |           Email:  paul_gagniuc@acad.ro                               |
' |                                                                      |
' |           Notes:  Graphic  - Tests and Applications                  |
' |                  _____________________________                       |
' |_________________/                             \______________________|
'

Option Explicit

Dim tmpx As Variant

Private Sub Alocate_Click()
    Dim Total_Rows, b, i, r, Total_Blocks, Col, PlotBlock As Integer
    Dim BlockData, tmp_sequence As String
    Dim xx, yy As Variant

    Picture1.Cls

    yy = Picture1.ScaleHeight
    xx = Picture1.ScaleWidth / (Len(sequence_input.Text))

    tmpx = 1
    Total_Rows = 1
    '---------------------------------------------------
    b = Block_Alocation2(Len(sequence_input.Text), Last_Block.Value, Min_Block.Value, Max_Block.Value)
    '---------------------------------------------------
    For i = 1 To Len(sequence_input.Text) Step b
        Total_Blocks = Total_Blocks + 1
        Col = Col + 1

        BlockData = Mid(sequence_input.Text, i, b)

        tmp_sequence = tmp_sequence & " " & BlockData
        '---------------------------------------------------
        PlotBlock = Val(xx * Len(BlockData))

    For r = tmpx To tmpx + PlotBlock - 2
        Picture1.PSet (r, Val(2 * Total_Rows)), vbRed
    Next r

    tmpx = tmpx + PlotBlock
    '---------------------------------------------------
    If Col >= MB.Value Then
        Col = 0
        tmpx = 1
        Total_Rows = Total_Rows + 1
        tmp_sequence = tmp_sequence & vbCrLf
    End If

    Next i
    '---------------------------------------------------

    '---------------------------------------------------
    If IH.Value = 1 Then
        sequence_output.Text = sequence_output.Text & _
        vbCrLf & "-------------------------------------" & _
        vbCrLf & "Input Parameters:" & _
        vbCrLf & "Last_Block: " & Last_Block.Value & _
        vbCrLf & "Min_Block: " & Min_Block.Value & _
        vbCrLf & "Max_Block: " & Max_Block.Value & _
        vbCrLf & "-------------------------------------" & _
        vbCrLf & "Results:" & _
        vbCrLf & "Sequence length: " & Len(sequence_input.Text) & _
        vbCrLf & "Block size found: " & b & _
        vbCrLf & "Number of Blocks: " & Total_Blocks & _
        vbCrLf & "-------------------------------------" & _
        vbCrLf & tmp_sequence
    Else
        sequence_output.Text = sequence_output.Text & vbCrLf & tmp_sequence
    End If
    '---------------------------------------------------
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
Function Block_Alocation1(ByVal L As Variant) As Integer
    Dim a, t, b, m, v, r As Integer
    
    a = 1
    t = 1
    b = 1
    m = 10

    Do Until t > 3 And v = 0
        a = a + 1
        t = (L Mod a)
        r = (L - t)
        v = r Mod 2
    Loop

    Do Until b = 0 Or m > 1000
        m = m + 1
        b = r Mod m
    Loop

    Block_Alocation1 = m
End Function

Private Sub EOP_Click()
    sequence_output.Text = ""
    Picture1.Cls
End Sub

Private Sub Form_Load()
    LB.Caption = Last_Block.Value
    MinB.Caption = Min_Block.Value
    MaxB.Caption = Max_Block.Value
    MC.Caption = MB.Value

    Input_seq.Caption = "Input sequence: " & Len(sequence_input) & " b"
End Sub

Private Sub GenSeqRND_Click()
    Dim x As Integer
    x = Val(Nucleotide_number.Text)
    sequence_input.Text = Novo_Sequence(x, "ADN")
End Sub

Private Sub Last_Block_Change()
    LB.Caption = Last_Block.Value
End Sub

Private Sub Max_Block_Change()
    MaxB.Caption = Max_Block.Value
End Sub

Private Sub MB_Change()
    MC.Caption = MB.Value
End Sub

Private Sub Min_Block_Change()
    MinB.Caption = Min_Block.Value
End Sub

Function Novo_Sequence(ByVal nr As Variant, ByVal tip As String) As String
    Dim N, C As Integer
    Dim p As String

    Dim nucleo(1 To 5) As String
    nucleo(1) = "A"
    nucleo(2) = "T"
    nucleo(3) = "G"
    nucleo(4) = "C"
    nucleo(5) = "U"

    For N = 1 To nr

    If (tip = "ADN") Then
        C = Int(3 * Rnd(3))
        p = p & nucleo(C + 1)
    End If

    If (tip = "ARN") Then
        C = Int(4 * Rnd(4))
        If (C + 1 = 2) Then C = 4
        p = p & nucleo(C + 1)
    End If

    Next N

    Novo_Sequence = p
End Function

Private Sub sequence_input_Change()
    Input_seq.Caption = "Input sequence: " & Len(sequence_input) & " b"
End Sub
