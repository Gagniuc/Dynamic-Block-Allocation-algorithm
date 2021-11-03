Attribute VB_Name = "article_exemple"
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

Private Sub main()

    Dim Last_Block, Min_Block, Max_Block, Max_Rows, Max_Cols, b, i, Total_Blocks, Col As Integer
    Dim sequence_input, output, BlockData As String

    
    sequence_input = Novo_Sequence(300, "ADN")

    Last_Block = 3
    Min_Block = 9
    Max_Block = 1000

    Max_Rows = 1
    Max_Cols = 3

    b = MBFA(Len(sequence_input), Last_Block, Min_Block, Max_Block)

    For i = 1 To Len(sequence_input) Step b
        Total_Blocks = Total_Blocks + 1
        Col = Col + 1

        BlockData = Mid(sequence_input, i, b)

        output = output & "|" & BlockData

        If Col >= Max_Cols Then
            Col = 0

            Max_Rows = Max_Rows + 1
            output = output & vbCrLf

        End If

    Next i

    MsgBox output
End Sub

Function MBFA(L, Last_Block, MinBlock, MaxBlock) As Variant

    Dim RestetBlock As Variant
    Dim q, t As Variant

    RestetBlock = MinBlock

    t = 1
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
    MBFA = MinBlock
End Function


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
