VERSION 5.00
Begin VB.Form frmSolitaire 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   1080
   ClientLeft      =   1875
   ClientTop       =   1830
   ClientWidth     =   6660
   FillColor       =   &H00FFFF80&
   ForeColor       =   &H00FFFF80&
   Icon            =   "SolitaireRound.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   72
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   444
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblMarblesLeft 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "32"
      DragMode        =   1  'Automatic
      Enabled         =   0   'False
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   255
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Height          =   495
      Left            =   960
      Shape           =   3  'Circle
      Top             =   120
      Width           =   495
   End
   Begin VB.Image imgHole 
      Height          =   480
      Index           =   0
      Left            =   2520
      Picture         =   "SolitaireRound.frx":08CA
      Top             =   120
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgMarble 
      DragMode        =   1  'Automatic
      Height          =   480
      Index           =   0
      Left            =   1680
      Top             =   120
      Width           =   480
   End
   Begin VB.Menu a 
      Caption         =   "                                                                      "
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnuInstructions 
         Caption         =   "&Instructions"
      End
      Begin VB.Menu mnuNewGame 
         Caption         =   "&New-Game"
      End
      Begin VB.Menu mnuRedo 
         Caption         =   "&Redo        (Shift+click)"
      End
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo        (Ctrl+click)"
      End
      Begin VB.Menu mnuSound 
         Caption         =   "&Sound"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu b 
      Caption         =   "                                                   "
   End
End
Attribute VB_Name = "frmSolitaire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------
'Author:    Anders Fransson
'Email:     anders.fransson@home.se
'Internet:  http://hem1.passagen.se/fylke
'Date:      97-07-30
'-------------------------------------------------------------------------

Option Explicit

Private m_iDragIndex As Integer
Private m_iMarblesLeft As Integer
Private m_iOldMovesIndex As Integer
Private m_vOldMoves(1 To 2, 1 To 35) As Integer

Private Const SIZE As Integer = 7
Private Const HOLE_WIDTH As Integer = 50
Private Const BORDER_TOP As Integer = -19
Private Const BORDER_LEFT As Integer = 17
Private Const BORDER_DIAMETER As Integer = 445
Private Const FORM_DIAMETER As Integer = 7000

Private Const TEXT_SOLITAIRE As String = "Solitaire"
Private Const TEXT_INSTRUCTIONS As String = "Solitaire playing instructions"
Private Const TEXT_INSTRUCTION_1 As String = "Remove as many marbles as possible."
Private Const TEXT_INSTRUCTION_2 As String = "Legal moves are horizontal or vertical jumps of one marble"
Private Const TEXT_INSTRUCTION_3 As String = "over another to an empty hole."
Private Const TEXT_INSTRUCTION_4 As String = "Just drag a marble and drop it in an empty hole."
Private Const TEXT_INSTRUCTION_5 As String = "The marble that has been jumped over is removed."
Private Const TEXT_INSTRUCTION_6 As String = "The ultimate challange is to get one marble left"
Private Const TEXT_INSTRUCTION_7 As String = "at the same position as the initial empty hole."

Private Declare Function CreateEllipticRgn Lib "gdi32" _
    (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function SetWindowRgn Lib "user32" _
    (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long

Private Static Sub Form_Load()

    Dim i%, j%
    Dim hr&, dl&
    Dim usew&, useh&
    
    'Make round form
    Me.Width = FORM_DIAMETER
    Me.Height = FORM_DIAMETER
    usew& = Me.Width / Screen.TwipsPerPixelX
    useh& = Me.Height / Screen.TwipsPerPixelY
    hr& = CreateEllipticRgn(20, 19, usew, useh)
    dl& = SetWindowRgn(Me.hwnd, hr, True)
    Me.Height = FORM_DIAMETER + 50
    Me.Width = FORM_DIAMETER + 50

    'Place border
    shpBorder.Move BORDER_LEFT, BORDER_TOP, BORDER_DIAMETER, BORDER_DIAMETER
    lblMarblesLeft.Move 100, 70
    
    'Load and place images
    imgMarble(0).Visible = False
    imgMarble(0).Picture = Me.Icon
    imgMarble(0).DragIcon = Me.Icon
    For i = 0 To SIZE - 1: For j = 0 To SIZE - 1
        If Not ((i < 2 And (j < 2 Or j > 4)) Or (i > 4 And (j < 2 Or j > 4))) Then
            Load imgMarble(i * SIZE + j)
            Load imgHole(i * SIZE + j)
            imgMarble(i * SIZE + j).Move 74 + HOLE_WIDTH * j, _
                44 + HOLE_WIDTH * i
            imgHole(i * SIZE + j).Move 74 + HOLE_WIDTH * j, _
                44 + HOLE_WIDTH * i
            imgHole(i * SIZE + j).Visible = True
        End If
    Next j: Next i
        
    NewGame
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)

    'Undo or redo move if Shift or Ctrl
    If Shift = 1 Then Redo
    If Shift = 2 Then Undo

End Sub

Private Sub imgHole_MouseDown(Index As Integer, Button As Integer, _
    Shift As Integer, X As Single, Y As Single)

    'Undo or redo move if Shift or Ctrl
    If Shift = 1 Then Redo
    If Shift = 2 Then Undo

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, _
    X As Single, Y As Single)
        
    'Show last draged marble if it has been dropped outside form
    If Not m_iDragIndex = 0 Then imgMarble(m_iDragIndex).Visible = True

End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)

    m_iDragIndex = 0
    Source.Visible = True

End Sub

Private Sub imgMarble_DragDrop(Index As Integer, Source As Control, _
    X As Single, Y As Single)

    m_iDragIndex = 0
    Source.Visible = True

End Sub

Private Sub imgMarble_DragOver(Index As Integer, Source As Control, _
    X As Single, Y As Single, State As Integer)

    Source.Visible = False
    m_iDragIndex = Source.Index
    
End Sub

Private Static Sub imgHole_DragDrop(Index As Integer, Source As Control, _
    X As Single, Y As Single)
    
    Dim xHole%, yHole%, xSource%, ySource%
    
    'Calculate coordinates for source-marble and drop-hole
    xHole = (Index) Mod SIZE
    yHole = (Index) \ SIZE
    xSource = Source.Index Mod SIZE
    ySource = Source.Index \ SIZE

    m_iDragIndex = 0

    'Show source-marble and exit sub if move isn't valid
    If Not ((Abs(yHole - ySource) = 2 And (xHole = xSource)) Or _
       (Abs(xHole - xSource) = 2 And (yHole = ySource))) Then
        Source.Visible = True
        Exit Sub
    End If
 
    'Show source-marble and exit sub if move isn't valid
    If Not imgMarble((Index + Source.Index) / 2).Visible Then
        Source.Visible = True
        Exit Sub
    End If
    
    If mnuSound.Checked Then PlaySound App.Path & "\Drop.wav"
    lblMarblesLeft.Move imgMarble((Index + Source.Index) / 2).Left + 7, _
        imgMarble((Index + Source.Index) / 2).Top + 9
    
    'Update move-menus
    mnuUndo.Enabled = True
    mnuRedo.Enabled = False
    
    'Update the old-moves variable
    m_iOldMovesIndex = m_iOldMovesIndex + 1
    m_vOldMoves(1, m_iOldMovesIndex) = Source.Index
    m_vOldMoves(2, m_iOldMovesIndex) = Index
    m_vOldMoves(1, m_iOldMovesIndex + 1) = 0
    m_vOldMoves(2, m_iOldMovesIndex + 1) = 0
    
    'Hide and show involved marbles
    Source.Visible = False
    imgMarble((Index + Source.Index) / 2).Visible = False
    imgMarble(Index).Visible = True
    
    'Update form caption
    m_iMarblesLeft = m_iMarblesLeft - 1
    lblMarblesLeft = m_iMarblesLeft

End Sub

Private Sub mnuInstructions_Click()
 
    MsgBox TEXT_INSTRUCTION_1 & vbNewLine & vbNewLine & _
        TEXT_INSTRUCTION_2 & vbNewLine & _
        TEXT_INSTRUCTION_3 & vbNewLine & vbNewLine & _
        TEXT_INSTRUCTION_4 & vbNewLine & _
        TEXT_INSTRUCTION_5 & vbNewLine & vbNewLine & _
        TEXT_INSTRUCTION_6 & vbNewLine & _
        TEXT_INSTRUCTION_7 & vbNewLine, _
        vbInformation, TEXT_INSTRUCTIONS
        
End Sub

Private Sub mnuNewGame_Click()

    NewGame

End Sub

Private Sub mnuRedo_Click()
    
    Redo

End Sub

Private Sub mnuSound_Click()
    
    mnuSound.Checked = Not mnuSound.Checked

End Sub

Private Sub mnuUndo_Click()

    Undo
    
End Sub

Private Sub mnuExit_Click()

    Unload Me
    
End Sub

Private Static Sub NewGame()

    Dim i%, j%
    
    'Show marbles
    For i = 0 To SIZE - 1: For j = 0 To SIZE - 1
        If Not ((i < 2 And (j < 2 Or j > 4)) Or (i > 4 And (j < 2 Or j > 4))) Then _
            imgMarble(i * SIZE + j).Visible = True
        If (i = 3 And j = 3) Then imgMarble(i * SIZE + j).Visible = False
    Next j: Next i

    'Reset old-moves
    For i = 1 To 2
        For j = LBound(m_vOldMoves, 1) To UBound(m_vOldMoves, 1)
            m_vOldMoves(i, j) = 0
        Next j
    Next i

    'Start values
    m_iDragIndex = 0
    m_iOldMovesIndex = 0
    m_iMarblesLeft = 32
    mnuUndo.Enabled = False
    mnuRedo.Enabled = False
    lblMarblesLeft = m_iMarblesLeft
    lblMarblesLeft.Move 231, 203
'    Me.Caption = TEXT_SOLITAIRE
    
End Sub

Private Sub Undo()
    
    'Exit if Undo-menu is disabled
    If Not mnuUndo.Enabled Then Exit Sub
    
    If mnuSound.Checked Then PlaySound App.Path & "\Drop.wav"
    
    'Update form caption
    m_iMarblesLeft = m_iMarblesLeft + 1
    lblMarblesLeft = m_iMarblesLeft
    
    'Update marbles visability and old-moves
    imgMarble(m_vOldMoves(1, m_iOldMovesIndex)).Visible = True
    imgMarble((m_vOldMoves(1, m_iOldMovesIndex) + _
        m_vOldMoves(2, m_iOldMovesIndex)) / 2).Visible = True
    'Plave label with marbles left
    lblMarblesLeft.Move imgMarble(m_vOldMoves(2, m_iOldMovesIndex)).Left + 7, _
        imgMarble(m_vOldMoves(2, m_iOldMovesIndex)).Top + 9
    imgMarble(m_vOldMoves(2, m_iOldMovesIndex)).Visible = False
    m_iOldMovesIndex = m_iOldMovesIndex - 1
    
    'Disable Undo-menu if there is no more move to undo
    If m_iOldMovesIndex = 0 Then mnuUndo.Enabled = False
        
    'Redo is now possible
    mnuRedo.Enabled = True

End Sub

Private Sub Redo()
    
    'Exit if Redo-menu is disabled
    If Not mnuRedo.Enabled Then Exit Sub
    
    If mnuSound.Checked Then PlaySound App.Path & "\Drop.wav"
    
    'Update form caption
    m_iMarblesLeft = m_iMarblesLeft - 1
    lblMarblesLeft = m_iMarblesLeft
    
    'Update old-moves and marbles visability
    m_iOldMovesIndex = m_iOldMovesIndex + 1
    imgMarble(m_vOldMoves(1, m_iOldMovesIndex)).Visible = False
    imgMarble((m_vOldMoves(1, m_iOldMovesIndex) + _
        m_vOldMoves(2, m_iOldMovesIndex)) / 2).Visible = False
    imgMarble(m_vOldMoves(2, m_iOldMovesIndex)).Visible = True
    
    'Plave label with marbles left
    lblMarblesLeft.Move imgMarble((m_vOldMoves(1, m_iOldMovesIndex) + _
        m_vOldMoves(2, m_iOldMovesIndex)) / 2).Left + 7, _
        imgMarble((m_vOldMoves(1, m_iOldMovesIndex) + _
        m_vOldMoves(2, m_iOldMovesIndex)) / 2).Top + 9
    
    'Disable Redo-menu if there is no more move to redo
    If m_vOldMoves(1, m_iOldMovesIndex + 1) = 0 Then mnuRedo.Enabled = False

    'Undo is now possible
    mnuUndo.Enabled = True

End Sub
