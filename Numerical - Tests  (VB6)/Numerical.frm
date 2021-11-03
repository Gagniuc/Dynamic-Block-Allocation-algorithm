VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Dynamic data block allocation for DNA sequences"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Dynamic data block allocation for DNA sequences"
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.TextBox sequence_output 
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Top             =   1200
         Width           =   7095
      End
      Begin VB.CommandButton Alocate 
         Caption         =   "Allocate"
         Height          =   735
         Left            =   2640
         TabIndex        =   2
         Top             =   2040
         Width           =   2295
      End
      Begin VB.TextBox sequence_input 
         Height          =   285
         Left            =   360
         TabIndex        =   1
         Text            =   "2134567823"
         Top             =   720
         Width           =   7095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   ________________________________                          ___________
'  / Dynamic data block allocation  \________________________/           |
' |                                                                      |
' |     Description:  Dynamic Data Block Allocation                      |
' |        Category:  Open source                                        |
' |          Author:  Paul Gagniuc                                       |
' |                                                                      |
' |    Date Created:  July 2010                                          |
' |       Tested On:  WinXP, WinVista, Win7, Win8                        |
' |           Email:  paul_gagniuc@acad.ro                               |
' |                                                                      |
' |           Notes:  Numerical - Tests                                  |
' |                  _____________________________                       |
' |_________________/                             \______________________|
'

Option Explicit

Private Sub Alocate_Click()
'Multi Brute Force Algorithm (MBFA)

Dim q, t As Variant
Dim last_block, block As Integer

last_block = 3 'we declare the number of nucleotides in the last block of data
block = 9      'we declare the minimum length of a data block
t = 1          'we declare y, different from zero

q = Val(sequence_input) - last_block

1:
Do Until t = 0 Or block > 1000
    block = block + 1
    t = q Mod block
Loop

If block > 1000 Or q < 9 Then
    q = q - 1
    block = 9
    GoTo 1
End If

sequence_output.Text = "Total length of the sequence: " & q & ". Assigned block size is: " & _
block & ", and the last block: " & Val(sequence_input) - q

End Sub

