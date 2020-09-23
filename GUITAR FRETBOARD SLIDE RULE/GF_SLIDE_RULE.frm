VERSION 5.00
Begin VB.Form GF_SLIDE_RULE_2 
   Caption         =   "GF SLIDE RULE "
   ClientHeight    =   7245
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14025
   LinkTopic       =   "Form1"
   ScaleHeight     =   7245
   ScaleWidth      =   14025
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstDrawingOptions 
      Height          =   450
      ItemData        =   "GF_SLIDE_RULE.frx":0000
      Left            =   5040
      List            =   "GF_SLIDE_RULE.frx":0010
      TabIndex        =   56
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CheckBox chkShowOpenStrings 
      Caption         =   "view capo open strings"
      Height          =   255
      Left            =   7560
      TabIndex        =   55
      Top             =   1200
      Width           =   2055
   End
   Begin VB.HScrollBar hsbCapo 
      Height          =   200
      Left            =   120
      TabIndex        =   54
      Top             =   4680
      Width           =   500
   End
   Begin VB.TextBox txtCapo 
      BackColor       =   &H000000FF&
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   0
      TabIndex        =   53
      Top             =   4320
      Width           =   700
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   12240
      TabIndex        =   51
      Top             =   840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox lstInversions 
      Enabled         =   0   'False
      Height          =   255
      Left            =   12240
      TabIndex        =   42
      Top             =   480
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2415
      Left            =   0
      ScaleHeight     =   2355
      ScaleWidth      =   11355
      TabIndex        =   41
      Top             =   1800
      Width           =   11415
   End
   Begin VB.CheckBox chkReverseFretboard 
      Caption         =   "Right-Handed"
      Height          =   255
      Left            =   9960
      TabIndex        =   40
      Top             =   1200
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.TextBox txtLastFret 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   12240
      Locked          =   -1  'True
      TabIndex        =   39
      Text            =   "21"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtFirstFret 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   11760
      Locked          =   -1  'True
      TabIndex        =   38
      Text            =   "0"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtLastPosition 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   12720
      Locked          =   -1  'True
      TabIndex        =   37
      Text            =   "21"
      Top             =   0
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   495
      Left            =   11880
      Max             =   22
      TabIndex        =   36
      Top             =   480
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdDrawEntireScale 
      BackColor       =   &H00808080&
      Caption         =   "Draw Entire Scale"
      Height          =   255
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   1560
      Width           =   11415
   End
   Begin VB.CheckBox chkInstrument 
      Caption         =   "Guitar"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   34
      Top             =   1200
      Width           =   855
   End
   Begin VB.CheckBox chkInstrument 
      Caption         =   "4 String Bass"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   33
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CheckBox chkInstrument 
      Caption         =   "5 String Bass"
      Height          =   255
      Index           =   2
      Left            =   2280
      TabIndex        =   32
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CheckBox chkInstrument 
      Caption         =   "7 String Guitar"
      Height          =   255
      Index           =   3
      Left            =   3600
      TabIndex        =   31
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtPosition 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   30
      Top             =   0
      Width           =   11750
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   777
      Max             =   18
      TabIndex        =   29
      Top             =   4320
      Width           =   11415
   End
   Begin VB.CommandButton cmdDigit 
      Caption         =   "2"
      Height          =   375
      Index           =   2
      Left            =   1200
      TabIndex        =   28
      ToolTipText     =   "or 9"
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton cmdDigit 
      Caption         =   "b3"
      Height          =   375
      Index           =   3
      Left            =   1680
      TabIndex        =   27
      ToolTipText     =   "or #9"
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton cmdDigit 
      Caption         =   "3"
      Height          =   375
      Index           =   4
      Left            =   2160
      TabIndex        =   26
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton cmdDigit 
      Caption         =   "4"
      Height          =   375
      Index           =   5
      Left            =   2640
      TabIndex        =   25
      ToolTipText     =   "or 11"
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton cmdDigit 
      Caption         =   "b5"
      Height          =   375
      Index           =   6
      Left            =   3120
      TabIndex        =   24
      ToolTipText     =   "or +11"
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton cmdDigit 
      Caption         =   "5"
      Height          =   375
      Index           =   7
      Left            =   3600
      TabIndex        =   23
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton cmdDigit 
      Caption         =   "#5"
      Height          =   375
      Index           =   8
      Left            =   4080
      TabIndex        =   22
      ToolTipText     =   "or b6"
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton cmdDigit 
      Caption         =   "6"
      Height          =   375
      Index           =   9
      Left            =   4560
      TabIndex        =   21
      ToolTipText     =   "or 13"
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton cmdDigit 
      Caption         =   "b7"
      Height          =   375
      Index           =   10
      Left            =   5040
      TabIndex        =   20
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton cmdDigit 
      Caption         =   "7"
      Height          =   375
      Index           =   11
      Left            =   5520
      TabIndex        =   19
      Top             =   5640
      Width           =   375
   End
   Begin VB.CommandButton cmdDigit 
      Caption         =   "b2"
      Height          =   375
      Index           =   1
      Left            =   720
      TabIndex        =   18
      ToolTipText     =   "or b9"
      Top             =   5640
      Width           =   375
   End
   Begin VB.Frame Frame3 
      Caption         =   "Notes:"
      Height          =   3615
      Left            =   11760
      TabIndex        =   11
      Top             =   1200
      Width           =   1455
      Begin VB.ListBox lstPitch 
         Height          =   2760
         ItemData        =   "GF_SLIDE_RULE.frx":0044
         Left            =   120
         List            =   "GF_SLIDE_RULE.frx":006C
         Style           =   1  'Checkbox
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cmdClear2 
         Caption         =   "clear"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   3240
         Width           =   735
      End
   End
   Begin VB.OptionButton optNoteCount 
      Caption         =   "INTERVALS"
      Height          =   375
      Index           =   1
      Left            =   6480
      TabIndex        =   10
      Top             =   5280
      Width           =   1455
   End
   Begin VB.OptionButton optNoteCount 
      Caption         =   "TRIADS"
      Height          =   255
      Index           =   2
      Left            =   6480
      TabIndex        =   9
      Top             =   5640
      Width           =   1455
   End
   Begin VB.OptionButton optNoteCount 
      Caption         =   "TETRATONICS"
      Height          =   375
      Index           =   3
      Left            =   6480
      TabIndex        =   8
      Top             =   5880
      Width           =   1455
   End
   Begin VB.OptionButton optNoteCount 
      Caption         =   "PENTATONICS"
      Height          =   255
      Index           =   4
      Left            =   6480
      TabIndex        =   7
      Top             =   6240
      Width           =   1455
   End
   Begin VB.OptionButton optNoteCount 
      Caption         =   "HEXATONICS"
      Height          =   375
      Index           =   5
      Left            =   6480
      TabIndex        =   6
      Top             =   6480
      Width           =   1455
   End
   Begin VB.OptionButton optNoteCount 
      Caption         =   "HEPTATONICS"
      Height          =   255
      Index           =   6
      Left            =   11640
      TabIndex        =   5
      Top             =   5040
      Width           =   1455
   End
   Begin VB.OptionButton optNoteCount 
      Caption         =   "8 NOTE "
      Height          =   375
      Index           =   7
      Left            =   11640
      TabIndex        =   4
      Top             =   5280
      Width           =   1455
   End
   Begin VB.OptionButton optNoteCount 
      Caption         =   "9 NOTE"
      Height          =   255
      Index           =   8
      Left            =   11640
      TabIndex        =   3
      Top             =   5640
      Width           =   1455
   End
   Begin VB.OptionButton optNoteCount 
      Caption         =   "10 NOTE"
      Height          =   375
      Index           =   9
      Left            =   11640
      TabIndex        =   2
      Top             =   5880
      Width           =   1455
   End
   Begin VB.OptionButton optNoteCount 
      Caption         =   "11 NOTE"
      Height          =   255
      Index           =   10
      Left            =   11640
      TabIndex        =   1
      Top             =   6240
      Width           =   1455
   End
   Begin VB.OptionButton optNoteCount 
      Caption         =   "CHROMATIC "
      Height          =   255
      Index           =   11
      Left            =   11640
      TabIndex        =   0
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "----------------------------------FIND RELATED SCALE INVERSIONS"
      Height          =   2175
      Left            =   0
      TabIndex        =   43
      Top             =   4920
      Width           =   6135
      Begin VB.CommandButton cmdClear 
         Caption         =   "clear"
         Height          =   255
         Left            =   4920
         TabIndex        =   49
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton cmdEnter 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2520
         MaskColor       =   &H00FF0000&
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   1680
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox txtFormula 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   47
         Top             =   240
         Width           =   5655
      End
      Begin VB.TextBox txtInput 
         Height          =   285
         Left            =   240
         TabIndex        =   46
         Text            =   "100010010000"
         Top             =   1800
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdDigit 
         Caption         =   "R"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   45
         Top             =   720
         Width           =   375
      End
      Begin VB.ListBox lstRoot 
         Height          =   255
         ItemData        =   "GF_SLIDE_RULE.frx":00A8
         Left            =   2040
         List            =   "GF_SLIDE_RULE.frx":00D0
         TabIndex        =   44
         ToolTipText     =   "(highlight key before drawing...)"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "          KEY: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   960
         TabIndex        =   50
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "-------------------------------------------------------Scale Menu (click item)"
      Height          =   2295
      Left            =   6240
      TabIndex        =   14
      Top             =   4800
      Width           =   6975
      Begin VB.ListBox lstAppendix 
         Height          =   2010
         Left            =   1680
         TabIndex        =   16
         Top             =   240
         Width           =   3615
      End
      Begin VB.OptionButton optNoteCount 
         Caption         =   "OCTAVES"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblTotal 
         Height          =   255
         Left            =   5400
         TabIndex        =   17
         Top             =   1920
         Width           =   1455
      End
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   135
      Left            =   11640
      TabIndex        =   52
      Top             =   6600
      Width           =   1335
   End
End
Attribute VB_Name = "GF_SLIDE_RULE_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
' ================
' DRAWING sections
' ================
Dim FirstFret As Byte, LastFret As Byte, LastPosition As Byte
Dim highString As Byte, lowString As Byte
Dim twelfthFretMark1 As Single, twelfthFretMark2 As Single

Private Type Tablature
    E As Byte
    F As Byte
    Fs As Byte
    G As Byte
    Ab As Byte
    A As Byte
    Bb As Byte
    B As Byte
    C As Byte
    Cs As Byte
    D As Byte
    Eb As Byte
End Type

Dim StringOne() As Tablature
Dim StringTwo() As Tablature
Dim StringThree() As Tablature
Dim StringFour() As Tablature
Dim StringFive() As Tablature
Dim StringSix() As Tablature
Dim StringSeven() As Tablature

Dim storeNoteIndex As Integer

Dim FretNumber As Integer, StringNumber As Integer, SelectedNotes As Integer
Dim ArrayIndex As Integer
Dim DotWidth As Integer, CapoFret As Integer

' ==================================
' LISTING, NAMING & SORTING sections
' ==================================
Dim I As Variant, Isharp As Variant, II As Variant, IIIflat As Variant
Dim III As Variant, IV As Variant, IVsharp As Variant, V As Variant
Dim Vsharp As Variant, VI, VIIflat As Variant, VII As Variant

Dim parentScaleFormula As String, WholeandHalfStepEquivalent As String

Dim dynamicList() As String, dynamicindex As Integer
Dim Spelling As Integer
Dim Notecount As Integer

Dim Frame As Integer
Dim Length As Integer

Dim PitchNames As Variant

Dim Digits(11) As Integer, ctr2 As Integer
Dim CountNotes As Integer


Private Sub Form_Load()
   
   Me.Height = 7665
   Me.Width = 13700
   
   Picture1.AutoRedraw = True
   Picture1.BackColor = vbWhite
   Picture1.Height = 2415
   Picture1.Left = 120
   Picture1.Top = 1920
   Picture1.Width = 11415
   
   HScroll1.Height = 375
   HScroll1.Left = 1000
   HScroll1.Max = 18
   HScroll1.Min = 0
   HScroll1.SmallChange = 1
   HScroll1.Top = 4440
   HScroll1.Value = 0
   HScroll1.Width = 10595
   
   VScroll2.Height = 495
   VScroll2.Left = 12000
   VScroll2.Max = 22
   VScroll2.Min = 0
   VScroll2.Top = 600
   VScroll2.Value = 0
   VScroll2.Visible = False
   ''''''''''''''''''''''''''''''''''''''''''
   lstDrawingOptions.ListIndex = 1
   lstDrawingOptions.Left = 5222
   lstDrawingOptions.Top = 1200
   
   txtFirstFret.Height = 400
   txtFirstFret.Top = 120
   txtFirstFret.Left = 12000
   
   txtLastFret.Height = 400
   txtLastFret.Left = 12400
   txtLastFret.Text = 22
   txtLastFret.Top = 120
   
   txtLastPosition.Height = 400
   txtLastPosition.Left = 12800
   txtLastPosition.Text = 22
   txtLastPosition.Top = 120
   
   chkInstrument.Item(0).Caption = "Guitar"
   chkInstrument.Item(0).Left = 240
   chkInstrument.Item(0).Top = 1320
   chkInstrument.Item(0).Value = 1
   
   chkInstrument.Item(1).Caption = "4 String Bass"
   chkInstrument.Item(1).Left = 1080
   chkInstrument.Item(1).Top = 1320
   chkInstrument.Item(1).Width = 1335
   
   chkInstrument.Item(2).Caption = "5 String Bass"
   chkInstrument.Item(2).Top = 1320
   chkInstrument.Item(2).Left = 2400
   chkInstrument.Item(2).Width = 1335
   
   chkInstrument.Item(3).Caption = "7 String Guitar"
   chkInstrument.Item(3).Left = 3720
   chkInstrument.Item(3).Top = 1320
   chkInstrument.Item(3).Width = 1335
   
   chkReverseFretboard.Left = 9960
   chkReverseFretboard.Top = 1320
   chkReverseFretboard.Value = 0
   
   txtPosition.Height = 1095
   txtPosition.Left = 120
   'txtPosition.MultiLine = True (read only)
   'txtPosition.ScrollBars = 2 (read only)
   txtPosition.Top = 120
   txtPosition.Width = 11750
   
   txtCapo.Height = 285
   txtCapo.Left = 120
   txtCapo.Text = "Capo 0"
   txtCapo.Top = 4440
   txtCapo.Width = 700
   
   hsbCapo.Height = 200
   hsbCapo.Left = 250
   hsbCapo.Max = 18
   hsbCapo.Top = 4725
   hsbCapo.Width = 500

   cmdDrawEntireScale.BackColor = &H808080
   cmdDrawEntireScale.Left = 120
   cmdDrawEntireScale.Top = 1680
   cmdDrawEntireScale.Width = 11415
   
   lstAppendix.BackColor = &H80000005
   lstAppendix.Height = 2100
   lstAppendix.Left = 1680
   lstAppendix.Top = 240
   lstAppendix.Width = 3615
   
   lstInversions.Height = 255
   lstInversions.Left = 12240
   lstInversions.Top = 600
   lstInversions.Visible = False ' >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>  (may be replaced with an array)
                                                                        
   
   '======frmload Key of C - open position=
   optNoteCount.Item(6).Value = True
   lstAppendix.ListIndex = 0
   txtFirstFret.Text = 0
   txtLastPosition = 4
   Call RemoveUnison
   '=======================================
   
   chkShowOpenStrings.Caption = "view Open strings"
   chkShowOpenStrings.Top = 1320
   chkShowOpenStrings.Value = 0
   
End Sub

Private Sub chkInstrument_Click(Index As Integer)
   
   Select Case Index
      
      Case 0                  ' guitar
      highString = 1
      lowString = 6
      twelfthFretMark1 = 1.5
      twelfthFretMark2 = 5.5
      Frame = 0
      Case 1                  ' 4 string bass
      highString = 3
      lowString = 6
      twelfthFretMark1 = 3.5
      twelfthFretMark2 = 5.5
      Frame = 1
      Case 2                  ' 5 string bass
      highString = 3
      lowString = 7
      twelfthFretMark1 = 3.5
      twelfthFretMark2 = 6.5
      Frame = 2
      Case 3                  ' 7 string guitar
      highString = 1
      lowString = 7
      twelfthFretMark1 = 1.5
      twelfthFretMark2 = 6.5
      Frame = 3
   
   End Select
   
   Call DrawObjects
   Call RemoveUnison
   
End Sub

Private Sub chkInstrument_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   
   Select Case Index
      
      Case 0
      chkInstrument.Item(1).Value = 0   ' < (zeroes prevent simultaneous options from occuring)
      chkInstrument.Item(2).Value = 0
      chkInstrument.Item(3).Value = 0
      Case 1
      chkInstrument.Item(0).Value = 0
      chkInstrument.Item(2).Value = 0
      chkInstrument.Item(3).Value = 0
      Case 2
      chkInstrument.Item(0).Value = 0
      chkInstrument.Item(1).Value = 0
      chkInstrument.Item(3).Value = 0
      Frame = 2
      Case 3
      chkInstrument.Item(0).Value = 0
      chkInstrument.Item(1).Value = 0
      chkInstrument.Item(2).Value = 0
   
   End Select
   
End Sub

Private Sub txtFirstFret_Change()
   
   If IsNumeric(txtFirstFret.Text) = True And Val(txtFirstFret.Text) >= 0 And Val(txtFirstFret.Text) < Val(txtLastPosition.Text) Then
      
      FirstFret = Val(txtFirstFret.Text)

   End If
   
   Call DrawObjects

End Sub

Private Sub txtLastFret_Change()
   
   If IsNumeric(txtLastFret.Text) = True And Val(txtLastFret.Text) <= 24 Then
      LastFret = Val(txtLastFret.Text)
   Else
     txtLastFret.Text = 22
   End If
   
   Call DrawObjects

End Sub

Private Sub txtLastPosition_Change()
   
   If IsNumeric(txtLastPosition.Text) = True And Val(txtLastPosition.Text) <= 24 Then
      LastPosition = Val(txtLastPosition.Text)
   Else
     txtLastPosition.Text = 22
   End If
   
   Call DrawObjects

End Sub

Private Sub chkReverseFretboard_Click()
   
   Call DrawObjects
   Call RemoveUnison

End Sub

Private Sub DrawObjects()
   
   Call SelectFretboardOrientation
   Call Picture1.Cls
   Call DrawFretboard(Picture1)
   
   If CapoFret <> 0 Then
   Call DrawCapo(Picture1)
   End If
   
   If chkShowOpenStrings.Value = 1 Then
   Call DrawOpenCapoDots(Picture1) ' Additional
   End If
   
   Call DrawDots(Picture1)
   
End Sub

Private Sub SelectFretboardOrientation()
   
   If chkReverseFretboard = 0 Then
      
      Picture1.Scale (-1, 0)-(25, 8) ' > default Scale position
      chkReverseFretboard.Caption = "Left-Handed" ' Right-Handed   :> Bob
   
   ElseIf chkReverseFretboard = 1 Then
      
      Picture1.Scale (25, 0)-(-1, 8) ' > reverse Scale
      chkReverseFretboard.Caption = "Left-Handed" ' Left-Handed
   
   End If

End Sub

Private Sub DrawFretboard(Fretboard As PictureBox)
   
   Dim Frets As Single, Strings As Single
      
      Fretboard.ForeColor = vbBlack ' ADDITIONAL PROCEDURE
      
      For Frets = 0.5 To LastFret + 0.5 Step 1
         If Frets = 0.5 Then
            Fretboard.DrawWidth = 6             ' >  Nut
         Else
            Fretboard.DrawWidth = 3             ' >  frets
         End If
         Fretboard.Line (Frets, highString)-(Frets, lowString)
      Next Frets
   
      Fretboard.DrawWidth = 1
      
      For Strings = highString To lowString
         If chkInstrument.Item(0).Value = 1 And (Strings = 4 Or Strings = 5 Or Strings = 6) Then
            Fretboard.DrawWidth = 2.75          ' thicker gauge
         ElseIf chkInstrument.Item(0).Value = 0 Then
            Fretboard.DrawWidth = 1             ' thinner strings
         End If
         Fretboard.Line (0.5, Strings)-(LastFret + 0.5, Strings)
      Next Strings
   
End Sub

Private Sub DrawFretboardMarkers()
   
   ' * * * *  |  * * * * << FRETMARKERS
   
   Dim fretmarkcounter As Integer
   
   For fretmarkcounter = 3 To 9 Step 2
      
      Picture1.PSet (fretmarkcounter, ((lowString - highString) / 2 + highString)), &HC0C0C0         ' for frets 3,5,7,9
      Picture1.PSet (fretmarkcounter + 12, ((lowString - highString) / 2 + highString)), &HC0C0C0    ' for frets 15,17,19,21
   
   Next fretmarkcounter
   
   Picture1.PSet (12, twelfthFretMark1), &HC0C0C0    ' 12th fret FRETMARKER
   Picture1.PSet (12, twelfthFretMark2), &HC0C0C0    ' 12th fret FRETMARKER

End Sub

Private Sub DrawCapo(Capo As PictureBox)
   Dim capoHeight As Single
   
   Capo.DrawWidth = 17
   
   For capoHeight = highString To lowString Step 0.5
   Capo.PSet (CapoFret, capoHeight), vbRed
   Next capoHeight

End Sub

Private Sub DrawDots(Dots As PictureBox)
   
   Dots.DrawWidth = 7
   Call DrawFretboardMarkers
   Dots.DrawWidth = DotWidth
   
      For StringNumber = highString To lowString
         For FretNumber = FirstFret To LastPosition
            For SelectedNotes = 0 To lstPitch.ListCount - 1
               If lstPitch.Selected(SelectedNotes) = True Then
                  
                  Call StoreValues
                  
                  Picture1.CurrentX = FretNumber
                  Picture1.CurrentY = StringNumber
                  Picture1.ForeColor = vbBlack
                  
                  Select Case SelectedNotes
                        Case 0
                        Call CheckPitch0 ' > C
                        Case 1
                        Call CheckPitch1 ' > C#
                        Case 2
                        Call CheckPitch2 ' > D
                        Case 3
                        Call CheckPitch3 ' > Eb
                        Case 4
                        Call CheckPitch4 ' > E
                        Case 5
                        Call CheckPitch5 ' > F
                        Case 6
                        Call CheckPitch6 ' > F#
                        Case 7
                        Call CheckPitch7 ' > G
                        Case 8
                        Call CheckPitch8 ' > Ab
                        Case 9
                        Call CheckPitch9 ' > A
                        Case 10
                        Call CheckPitch10 '> Bb
                        Case 11
                        Call CheckPitch11 '> B
                        
                  End Select
               End If
            Next SelectedNotes
         Next FretNumber
      Next StringNumber
   
End Sub

Private Sub DrawOpenCapoDots(Dots As PictureBox)
   
   Dots.DrawWidth = 7
   Call DrawFretboardMarkers
   Dots.DrawWidth = DotWidth
   
      For StringNumber = highString To lowString
            
         FretNumber = CapoFret
            
            For SelectedNotes = 0 To lstPitch.ListCount - 1
               
               If lstPitch.Selected(SelectedNotes) = True Then
                  
                  Call StoreValues
                  
                  Picture1.CurrentX = FretNumber
                  Picture1.CurrentY = StringNumber
                  Picture1.ForeColor = vbBlack '&HC0& ' gray
                  Select Case SelectedNotes
                        Case 0
                        Call CheckPitch0 ' > C
                        Case 1
                        Call CheckPitch1 ' > C#
                        Case 2
                        Call CheckPitch2 ' > D
                        Case 3
                        Call CheckPitch3 ' > Eb
                        Case 4
                        Call CheckPitch4 ' > E
                        Case 5
                        Call CheckPitch5 ' > F
                        Case 6
                        Call CheckPitch6 ' > F#
                        Case 7
                        Call CheckPitch7 ' > G
                        Case 8
                        Call CheckPitch8 ' > Ab
                        Case 9
                        Call CheckPitch9 ' > A
                        Case 10
                        Call CheckPitch10 '> Bb
                        Case 11
                        Call CheckPitch11 '> B
                        
                  End Select
            
            End If
         
         Next SelectedNotes
      
      Next StringNumber
   
End Sub

Private Sub CheckPitch0()
' ====================================================================
   ' The next 11 procedures contain statements similar to this one...
' ====================================================================
   
   Select Case StringNumber
                  
   Case 1
      For ArrayIndex = LBound(StringOne) To UBound(StringOne)
         If StringOne(ArrayIndex).C = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                           
   Case 2
      For ArrayIndex = LBound(StringTwo) To UBound(StringTwo)
         If StringTwo(ArrayIndex).C = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 3
      For ArrayIndex = LBound(StringThree) To UBound(StringThree)
         If StringThree(ArrayIndex).C = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 4
      For ArrayIndex = LBound(StringFour) To UBound(StringFour)
         If StringFour(ArrayIndex).C = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 5
      For ArrayIndex = LBound(StringFive) To UBound(StringFive)
         If StringFive(ArrayIndex).C = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
   Case 6
      For ArrayIndex = LBound(StringSix) To UBound(StringSix)
         If StringSix(ArrayIndex).C = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
   Case 7
      For ArrayIndex = LBound(StringSeven) To UBound(StringSeven)
         If StringSeven(ArrayIndex).C = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   End Select
   
   ' ===========================================================
   ' ...hence CTRL C ... CTRL V ...CTRL F ...CTRL H ...REPLACE "extensions"
   ' ===========================================================

End Sub

' Simile
Private Sub CheckPitch1()

   Select Case StringNumber
                  
   Case 1
      For ArrayIndex = LBound(StringOne) To UBound(StringOne)
         If StringOne(ArrayIndex).Cs = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                           
   Case 2
      For ArrayIndex = LBound(StringTwo) To UBound(StringTwo)
         If StringTwo(ArrayIndex).Cs = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 3
      For ArrayIndex = LBound(StringThree) To UBound(StringThree)
         If StringThree(ArrayIndex).Cs = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 4
      For ArrayIndex = LBound(StringFour) To UBound(StringFour)
         If StringFour(ArrayIndex).Cs = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 5
      For ArrayIndex = LBound(StringFive) To UBound(StringFive)
         If StringFive(ArrayIndex).Cs = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
   Case 6
      For ArrayIndex = LBound(StringSix) To UBound(StringSix)
         If StringSix(ArrayIndex).Cs = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
   Case 7
      For ArrayIndex = LBound(StringSeven) To UBound(StringSeven)
         If StringSeven(ArrayIndex).Cs = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   End Select

End Sub

' Simile
Private Sub CheckPitch2()

   Select Case StringNumber
                  
   Case 1
      For ArrayIndex = LBound(StringOne) To UBound(StringOne)
         If StringOne(ArrayIndex).D = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                           
   Case 2
      For ArrayIndex = LBound(StringTwo) To UBound(StringTwo)
         If StringTwo(ArrayIndex).D = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 3
      For ArrayIndex = LBound(StringThree) To UBound(StringThree)
         If StringThree(ArrayIndex).D = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 4
      For ArrayIndex = LBound(StringFour) To UBound(StringFour)
         If StringFour(ArrayIndex).D = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 5
      For ArrayIndex = LBound(StringFive) To UBound(StringFive)
         If StringFive(ArrayIndex).D = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
   Case 6
      For ArrayIndex = LBound(StringSix) To UBound(StringSix)
         If StringSix(ArrayIndex).D = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
   Case 7
      For ArrayIndex = LBound(StringSeven) To UBound(StringSeven)
         If StringSeven(ArrayIndex).D = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   End Select

End Sub

' Simile
Private Sub CheckPitch3()

   Select Case StringNumber
                  
   Case 1
      For ArrayIndex = LBound(StringOne) To UBound(StringOne)
         If StringOne(ArrayIndex).Eb = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                           
   Case 2
      For ArrayIndex = LBound(StringTwo) To UBound(StringTwo)
         If StringTwo(ArrayIndex).Eb = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 3
      For ArrayIndex = LBound(StringThree) To UBound(StringThree)
         If StringThree(ArrayIndex).Eb = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 4
      For ArrayIndex = LBound(StringFour) To UBound(StringFour)
         If StringFour(ArrayIndex).Eb = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 5
      For ArrayIndex = LBound(StringFive) To UBound(StringFive)
         If StringFive(ArrayIndex).Eb = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
   Case 6
      For ArrayIndex = LBound(StringSix) To UBound(StringSix)
         If StringSix(ArrayIndex).Eb = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
   Case 7
      For ArrayIndex = LBound(StringSeven) To UBound(StringSeven)
         If StringSeven(ArrayIndex).Eb = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   End Select

End Sub

' Simile
Private Sub CheckPitch4()

   Select Case StringNumber
                  
   Case 1
      For ArrayIndex = LBound(StringOne) To UBound(StringOne)
         If StringOne(ArrayIndex).E = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                           
   Case 2
      For ArrayIndex = LBound(StringTwo) To UBound(StringTwo)
         If StringTwo(ArrayIndex).E = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 3
      For ArrayIndex = LBound(StringThree) To UBound(StringThree)
         If StringThree(ArrayIndex).E = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 4
      For ArrayIndex = LBound(StringFour) To UBound(StringFour)
         If StringFour(ArrayIndex).E = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 5
      For ArrayIndex = LBound(StringFive) To UBound(StringFive)
         If StringFive(ArrayIndex).E = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
   Case 6
      For ArrayIndex = LBound(StringSix) To UBound(StringSix)
         If StringSix(ArrayIndex).E = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
   Case 7
      For ArrayIndex = LBound(StringSeven) To UBound(StringSeven)
         If StringSeven(ArrayIndex).E = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   End Select

End Sub

' Simile
Private Sub CheckPitch5()

   Select Case StringNumber
                  
   Case 1
      For ArrayIndex = LBound(StringOne) To UBound(StringOne)
         If StringOne(ArrayIndex).F = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                           
   Case 2
      For ArrayIndex = LBound(StringTwo) To UBound(StringTwo)
         If StringTwo(ArrayIndex).F = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 3
      For ArrayIndex = LBound(StringThree) To UBound(StringThree)
         If StringThree(ArrayIndex).F = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 4
      For ArrayIndex = LBound(StringFour) To UBound(StringFour)
         If StringFour(ArrayIndex).F = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 5
      For ArrayIndex = LBound(StringFive) To UBound(StringFive)
         If StringFive(ArrayIndex).F = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
   Case 6
      For ArrayIndex = LBound(StringSix) To UBound(StringSix)
         If StringSix(ArrayIndex).F = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
   Case 7
      For ArrayIndex = LBound(StringSeven) To UBound(StringSeven)
         If StringSeven(ArrayIndex).F = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   End Select

End Sub

' Simile
Private Sub CheckPitch6()

   Select Case StringNumber
                  
   Case 1
      For ArrayIndex = LBound(StringOne) To UBound(StringOne)
         If StringOne(ArrayIndex).Fs = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                           
   Case 2
      For ArrayIndex = LBound(StringTwo) To UBound(StringTwo)
         If StringTwo(ArrayIndex).Fs = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 3
      For ArrayIndex = LBound(StringThree) To UBound(StringThree)
         If StringThree(ArrayIndex).Fs = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 4
      For ArrayIndex = LBound(StringFour) To UBound(StringFour)
         If StringFour(ArrayIndex).Fs = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 5
      For ArrayIndex = LBound(StringFive) To UBound(StringFive)
         If StringFive(ArrayIndex).Fs = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
   Case 6
      For ArrayIndex = LBound(StringSix) To UBound(StringSix)
         If StringSix(ArrayIndex).Fs = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
   Case 7
      For ArrayIndex = LBound(StringSeven) To UBound(StringSeven)
         If StringSeven(ArrayIndex).Fs = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   End Select

End Sub

' Simile
Private Sub CheckPitch7()

   Select Case StringNumber
                  
   Case 1
      For ArrayIndex = LBound(StringOne) To UBound(StringOne)
         If StringOne(ArrayIndex).G = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                           
   Case 2
      For ArrayIndex = LBound(StringTwo) To UBound(StringTwo)
         If StringTwo(ArrayIndex).G = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 3
      For ArrayIndex = LBound(StringThree) To UBound(StringThree)
         If StringThree(ArrayIndex).G = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 4
      For ArrayIndex = LBound(StringFour) To UBound(StringFour)
         If StringFour(ArrayIndex).G = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 5
      For ArrayIndex = LBound(StringFive) To UBound(StringFive)
         If StringFive(ArrayIndex).G = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
   Case 6
      For ArrayIndex = LBound(StringSix) To UBound(StringSix)
         If StringSix(ArrayIndex).G = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
   Case 7
      For ArrayIndex = LBound(StringSeven) To UBound(StringSeven)
         If StringSeven(ArrayIndex).G = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   End Select

End Sub

' Simile
Private Sub CheckPitch8()

   Select Case StringNumber
                  
   Case 1
      For ArrayIndex = LBound(StringOne) To UBound(StringOne)
         If StringOne(ArrayIndex).Ab = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                           
   Case 2
      For ArrayIndex = LBound(StringTwo) To UBound(StringTwo)
         If StringTwo(ArrayIndex).Ab = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 3
      For ArrayIndex = LBound(StringThree) To UBound(StringThree)
         If StringThree(ArrayIndex).Ab = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 4
      For ArrayIndex = LBound(StringFour) To UBound(StringFour)
         If StringFour(ArrayIndex).Ab = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 5
      For ArrayIndex = LBound(StringFive) To UBound(StringFive)
         If StringFive(ArrayIndex).Ab = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
   Case 6
      For ArrayIndex = LBound(StringSix) To UBound(StringSix)
         If StringSix(ArrayIndex).Ab = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
   Case 7
      For ArrayIndex = LBound(StringSeven) To UBound(StringSeven)
         If StringSeven(ArrayIndex).Ab = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   End Select

End Sub

' Simile
Private Sub CheckPitch9()

   Select Case StringNumber
                  
   Case 1
      For ArrayIndex = LBound(StringOne) To UBound(StringOne)
         If StringOne(ArrayIndex).A = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                           
   Case 2
      For ArrayIndex = LBound(StringTwo) To UBound(StringTwo)
         If StringTwo(ArrayIndex).A = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 3
      For ArrayIndex = LBound(StringThree) To UBound(StringThree)
         If StringThree(ArrayIndex).A = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 4
      For ArrayIndex = LBound(StringFour) To UBound(StringFour)
         If StringFour(ArrayIndex).A = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 5
      For ArrayIndex = LBound(StringFive) To UBound(StringFive)
         If StringFive(ArrayIndex).A = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
   Case 6
      For ArrayIndex = LBound(StringSix) To UBound(StringSix)
         If StringSix(ArrayIndex).A = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
   Case 7
      For ArrayIndex = LBound(StringSeven) To UBound(StringSeven)
         If StringSeven(ArrayIndex).A = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   End Select

End Sub

' Simile
Private Sub CheckPitch10()

   Select Case StringNumber
                  
   Case 1
      For ArrayIndex = LBound(StringOne) To UBound(StringOne)
         If StringOne(ArrayIndex).Bb = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                           
   Case 2
      For ArrayIndex = LBound(StringTwo) To UBound(StringTwo)
         If StringTwo(ArrayIndex).Bb = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 3
      For ArrayIndex = LBound(StringThree) To UBound(StringThree)
         If StringThree(ArrayIndex).Bb = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 4
      For ArrayIndex = LBound(StringFour) To UBound(StringFour)
         If StringFour(ArrayIndex).Bb = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 5
      For ArrayIndex = LBound(StringFive) To UBound(StringFive)
         If StringFive(ArrayIndex).Bb = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
   Case 6
      For ArrayIndex = LBound(StringSix) To UBound(StringSix)
         If StringSix(ArrayIndex).Bb = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
   Case 7
      For ArrayIndex = LBound(StringSeven) To UBound(StringSeven)
         If StringSeven(ArrayIndex).Bb = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   End Select

End Sub

' Simile
Private Sub CheckPitch11()

   Select Case StringNumber
                  
   Case 1
      For ArrayIndex = LBound(StringOne) To UBound(StringOne)
         If StringOne(ArrayIndex).B = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                           
   Case 2
      For ArrayIndex = LBound(StringTwo) To UBound(StringTwo)
         If StringTwo(ArrayIndex).B = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 3
      For ArrayIndex = LBound(StringThree) To UBound(StringThree)
         If StringThree(ArrayIndex).B = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 4
      For ArrayIndex = LBound(StringFour) To UBound(StringFour)
         If StringFour(ArrayIndex).B = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   Case 5
      For ArrayIndex = LBound(StringFive) To UBound(StringFive)
         If StringFive(ArrayIndex).B = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
   Case 6
      For ArrayIndex = LBound(StringSix) To UBound(StringSix)
         If StringSix(ArrayIndex).B = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
   Case 7
      For ArrayIndex = LBound(StringSeven) To UBound(StringSeven)
         If StringSeven(ArrayIndex).B = FretNumber Then
            Picture1.PSet (FretNumber, StringNumber)
            Call AddNoteCaptions
         End If
      Next ArrayIndex
                        
   End Select

End Sub

Private Sub AddNoteCaptions()
   
   Select Case chkReverseFretboard.Value
   
   Case 0
   Picture1.CurrentX = FretNumber - 0.24
   Picture1.CurrentY = StringNumber - 0.25
   Picture1.ForeColor = &HC0FFFF   ' pitch caption forecolor
   Picture1.FontSize = 7
   
   Case 1
   Picture1.CurrentX = FretNumber + 0.24
   Picture1.CurrentY = StringNumber - 0.25
   Picture1.ForeColor = &HC0FFFF   ' pitch caption forecolor
   Picture1.FontSize = 7
   
   End Select
   
   If lstDrawingOptions.ListIndex = 1 Then
      
      Picture1.Print lstPitch.List(SelectedNotes)
   
   ElseIf lstDrawingOptions.ListIndex = 2 Then
   
      If FretNumber - FirstFret > 0 And FirstFret > 0 Then
      Picture1.Print FretNumber - FirstFret
      ElseIf FretNumber - FirstFret = 0 And FirstFret <> 0 And (CapoFret <> FirstFret) Then
      Picture1.Print "1"
      
      ElseIf FretNumber - FirstFret = 0 And FirstFret <> 0 And (CapoFret = FirstFret) Then
      Picture1.Print "0"
      
      ElseIf FirstFret = 0 Then
      Picture1.Print FretNumber
      End If
   
   ElseIf lstDrawingOptions.ListIndex = 3 Then
      
      Picture1.Print FretNumber
   
   End If
   
End Sub

Private Sub lstDrawingOptions_Click()
   
   If lstDrawingOptions.ListIndex = 0 Then
      DotWidth = 12
   ElseIf lstDrawingOptions.ListIndex = 1 Then
      DotWidth = 16
   ElseIf lstDrawingOptions.ListIndex = 2 Or lstDrawingOptions.ListIndex = 3 Then
      DotWidth = 18
   End If
   
   Call DrawObjects
   Call RemoveUnison
   
End Sub

Private Sub chkShowOpenStrings_Click()
   Call DrawObjects
   Call RemoveUnison
End Sub


Private Sub StoreValues()

Dim GuitarTabNumber As Byte

   storeNoteIndex = -1
   
   For GuitarTabNumber = 0 To 12 Step 12
      
      storeNoteIndex = storeNoteIndex + 1
      
      ReDim Preserve StringOne(storeNoteIndex)
      ReDim Preserve StringTwo(storeNoteIndex)
      ReDim Preserve StringThree(storeNoteIndex)
      ReDim Preserve StringFour(storeNoteIndex)
      ReDim Preserve StringFive(storeNoteIndex)
      ReDim Preserve StringSix(storeNoteIndex)
      ReDim Preserve StringSeven(storeNoteIndex)
      
      StringOne(storeNoteIndex).E = GuitarTabNumber
      StringTwo(storeNoteIndex).B = GuitarTabNumber
      StringThree(storeNoteIndex).G = GuitarTabNumber
      StringFour(storeNoteIndex).D = GuitarTabNumber
      StringFive(storeNoteIndex).A = GuitarTabNumber
      StringSix(storeNoteIndex).E = GuitarTabNumber
      StringSeven(storeNoteIndex).B = GuitarTabNumber
      
   Next GuitarTabNumber

   storeNoteIndex = -1
   
   For GuitarTabNumber = 1 To 13 Step 12
      
      storeNoteIndex = storeNoteIndex + 1
      StringOne(storeNoteIndex).F = GuitarTabNumber
      StringTwo(storeNoteIndex).C = GuitarTabNumber
      StringThree(storeNoteIndex).Ab = GuitarTabNumber
      StringFour(storeNoteIndex).Eb = GuitarTabNumber
      StringFive(storeNoteIndex).Bb = GuitarTabNumber
      StringSix(storeNoteIndex).F = GuitarTabNumber
      StringSeven(storeNoteIndex).C = GuitarTabNumber
   Next GuitarTabNumber
      
   storeNoteIndex = -1
   
   For GuitarTabNumber = 2 To 14 Step 12
      storeNoteIndex = storeNoteIndex + 1
      StringOne(storeNoteIndex).Fs = GuitarTabNumber
      StringTwo(storeNoteIndex).Cs = GuitarTabNumber
      StringThree(storeNoteIndex).A = GuitarTabNumber
      StringFour(storeNoteIndex).E = GuitarTabNumber
      StringFive(storeNoteIndex).B = GuitarTabNumber
      StringSix(storeNoteIndex).Fs = GuitarTabNumber
      StringSeven(storeNoteIndex).Cs = GuitarTabNumber
   Next GuitarTabNumber
      
   storeNoteIndex = -1
   
   For GuitarTabNumber = 3 To 15 Step 12
      storeNoteIndex = storeNoteIndex + 1
      StringOne(storeNoteIndex).G = GuitarTabNumber
      StringTwo(storeNoteIndex).D = GuitarTabNumber
      StringThree(storeNoteIndex).Bb = GuitarTabNumber
      StringFour(storeNoteIndex).F = GuitarTabNumber
      StringFive(storeNoteIndex).C = GuitarTabNumber
      StringSix(storeNoteIndex).G = GuitarTabNumber
      StringSeven(storeNoteIndex).D = GuitarTabNumber
   Next GuitarTabNumber
   
   storeNoteIndex = -1
   
   For GuitarTabNumber = 4 To 16 Step 12
      
      storeNoteIndex = storeNoteIndex + 1
      StringOne(storeNoteIndex).Ab = GuitarTabNumber
      StringTwo(storeNoteIndex).Eb = GuitarTabNumber
      StringThree(storeNoteIndex).B = GuitarTabNumber
      StringFour(storeNoteIndex).Fs = GuitarTabNumber
      StringFive(storeNoteIndex).Cs = GuitarTabNumber
      StringSix(storeNoteIndex).Ab = GuitarTabNumber
      StringSeven(storeNoteIndex).Eb = GuitarTabNumber
   
   Next GuitarTabNumber
   
   storeNoteIndex = -1
   
   For GuitarTabNumber = 5 To 17 Step 12
      
      storeNoteIndex = storeNoteIndex + 1
      StringOne(storeNoteIndex).A = GuitarTabNumber
      StringTwo(storeNoteIndex).E = GuitarTabNumber
      StringThree(storeNoteIndex).C = GuitarTabNumber
      StringFour(storeNoteIndex).G = GuitarTabNumber
      StringFive(storeNoteIndex).D = GuitarTabNumber
      StringSix(storeNoteIndex).A = GuitarTabNumber
      StringSeven(storeNoteIndex).E = GuitarTabNumber
   
   Next GuitarTabNumber
   
   storeNoteIndex = -1
   
   For GuitarTabNumber = 6 To 18 Step 12
      
      storeNoteIndex = storeNoteIndex + 1
      StringOne(storeNoteIndex).Bb = GuitarTabNumber
      StringTwo(storeNoteIndex).F = GuitarTabNumber
      StringThree(storeNoteIndex).Cs = GuitarTabNumber
      StringFour(storeNoteIndex).Ab = GuitarTabNumber
      StringFive(storeNoteIndex).Eb = GuitarTabNumber
      StringSix(storeNoteIndex).Bb = GuitarTabNumber
      StringSeven(storeNoteIndex).F = GuitarTabNumber
   
   Next GuitarTabNumber
   
   storeNoteIndex = -1
   
   For GuitarTabNumber = 7 To 19 Step 12
      
      storeNoteIndex = storeNoteIndex + 1
      StringOne(storeNoteIndex).B = GuitarTabNumber
      StringTwo(storeNoteIndex).Fs = GuitarTabNumber
      StringThree(storeNoteIndex).D = GuitarTabNumber
      StringFour(storeNoteIndex).A = GuitarTabNumber
      StringFive(storeNoteIndex).E = GuitarTabNumber
      StringSix(storeNoteIndex).B = GuitarTabNumber
      StringSeven(storeNoteIndex).Fs = GuitarTabNumber
   
   Next GuitarTabNumber
   
   storeNoteIndex = -1
   
   For GuitarTabNumber = 8 To 20 Step 12
      
      storeNoteIndex = storeNoteIndex + 1
      StringOne(storeNoteIndex).C = GuitarTabNumber
      StringTwo(storeNoteIndex).G = GuitarTabNumber
      StringThree(storeNoteIndex).Eb = GuitarTabNumber
      StringFour(storeNoteIndex).Bb = GuitarTabNumber
      StringFive(storeNoteIndex).F = GuitarTabNumber
      StringSix(storeNoteIndex).C = GuitarTabNumber
      StringSeven(storeNoteIndex).G = GuitarTabNumber
   
   Next GuitarTabNumber
      
   storeNoteIndex = -1
   
   For GuitarTabNumber = 9 To 21 Step 12
      storeNoteIndex = storeNoteIndex + 1
      StringOne(storeNoteIndex).Cs = GuitarTabNumber
      StringTwo(storeNoteIndex).Ab = GuitarTabNumber
      StringThree(storeNoteIndex).E = GuitarTabNumber
      StringFour(storeNoteIndex).B = GuitarTabNumber
      StringFive(storeNoteIndex).Fs = GuitarTabNumber
      StringSix(storeNoteIndex).Cs = GuitarTabNumber
      StringSeven(storeNoteIndex).Ab = GuitarTabNumber
   Next GuitarTabNumber
   
   storeNoteIndex = -1
   
   For GuitarTabNumber = 10 To 22 Step 12
      
      storeNoteIndex = storeNoteIndex + 1
      StringOne(storeNoteIndex).D = GuitarTabNumber
      StringTwo(storeNoteIndex).A = GuitarTabNumber
      StringThree(storeNoteIndex).F = GuitarTabNumber
      StringFour(storeNoteIndex).C = GuitarTabNumber
      StringFive(storeNoteIndex).G = GuitarTabNumber
      StringSix(storeNoteIndex).D = GuitarTabNumber
      StringSeven(storeNoteIndex).A = GuitarTabNumber
   
   Next GuitarTabNumber
   
   storeNoteIndex = -1
   
   For GuitarTabNumber = 11 To 23 Step 12
      
      storeNoteIndex = storeNoteIndex + 1
      StringOne(storeNoteIndex).Eb = GuitarTabNumber
      StringTwo(storeNoteIndex).Bb = GuitarTabNumber
      StringThree(storeNoteIndex).Fs = GuitarTabNumber
      StringFour(storeNoteIndex).Cs = GuitarTabNumber
      StringFive(storeNoteIndex).Ab = GuitarTabNumber
      StringSix(storeNoteIndex).Eb = GuitarTabNumber
      StringSeven(storeNoteIndex).Bb = GuitarTabNumber
   
   Next GuitarTabNumber


End Sub

Private Sub HScroll1_Change()
   
   txtFirstFret.Text = HScroll1.Value
   txtLastPosition.Text = Val(txtFirstFret.Text + 4)
   Length = Len("FRET" & HScroll1.Value)
   txtPosition.Text = "FRET " & HScroll1.Value & Space(181 - Length) & lstAppendix.Text
   
   Call RemoveUnison
   
End Sub

Private Sub hsbCapo_Change()
   
   CapoFret = hsbCapo.Value
   txtCapo.Text = "Capo " & CapoFret
   HScroll1.Min = CapoFret
   
   
   Call DrawObjects
   Call RemoveUnison

End Sub

Private Sub VScroll2_Change()
   VScroll2.Max = HScroll1.Value + 4
   txtFirstFret.Text = VScroll2.Value
End Sub

Private Sub RemoveUnison()
   
   If txtLastPosition.Text = Val(txtFirstFret.Text + 4) Then
   If HScroll1.Value = 0 And (Frame = 0 Or Frame = 3) Then
      
      Picture1.PSet (4, 3), vbWhite               ' REMOVE 3RD STRING UNISON B(guitar open positions)
      Picture1.DrawWidth = 1
      Picture1.Line (4 - 0.5, 3)-(4 + 0.5, 3), vbBlack
      
   ElseIf HScroll1.Value <> 0 And (Frame = 0 Or Frame = 3) Then
   
   Dim erasure As Byte
   
      If HScroll1.Value <> CapoFret Then
      
      For erasure = 1 To 18
         If HScroll1.Value = erasure Then
         Picture1.PSet (erasure, 2), vbWhite      ' REMOVE 2ND STRING UNISON(for movable guitar patterns)
         Picture1.DrawWidth = 1
         Picture1.Line (erasure - 0.5, 2)-(erasure + 0.5, 2), vbBlack
         End If
      Next erasure
   
      ElseIf HScroll1.Value = CapoFret Then
      
      For erasure = 1 To 18
         If HScroll1.Value = erasure Then
         Picture1.PSet (erasure + 4, 3), vbWhite  ' REMOVE 3RD STRING UNISON (guitar CAPO positions)
         Picture1.DrawWidth = 1
         Picture1.Line (erasure + 4 - 0.5, 3)-(erasure + 4 + 0.5, 3), vbBlack
         End If
      Next erasure
      End If
   
   End If
   End If

End Sub

Private Sub cmdDrawEntireScale_Click()
   If CapoFret = 0 Then
   HScroll1.Value = 0
   txtFirstFret.Text = 0
   txtLastPosition = 22
   ElseIf CapoFret <> 0 Then
   HScroll1.Value = CapoFret
   txtFirstFret.Text = CapoFret
   txtLastPosition = 22
   End If
End Sub

Private Sub optNoteCount_Click(Index As Integer)
   Dim X As Integer
   Notecount = 0
   
   If Index <> 0 Then
   optNoteCount.Item(0).Value = False
   End If
   
   For X = 1 To optNoteCount.Count - 1
      If Index = 0 Then
      optNoteCount.Item(X).Value = False
      End If
   Next X
   
   Notecount = Index + 1
   Call List_Scales
   
   lblTotal.Caption = "Item No. : " & lstAppendix.ListIndex + 1 & " / " & lstAppendix.ListCount

End Sub

' Private Sub Optional()_Click

' ====================================================================================
' This list includes 351 "PARENT FORMULAS" from which other scales may be derived.
' One particular scale formula may be used to derive  several other related scales.
' For example, a Major Scale formula may be used to draw the Dorian mode, Phrygian mode,
' Lydian mode, Mixolydian mode, Natural minor scale and Locrian mode by locating the
' inversion/index number.

' The initial goal is to REDUCE THE INPUT DATA and the initial number of search comparisons
' and produce the data later only when a particular scale is drawn. The process of
' searching a list of 4095 names for every click can seemingly be avoided by reducing the
' list to 351 and further sorting the remainder based on the number of tones contained in
' the selected scale.  All scale combinations for 12 chromatic tones can be extracted by enumerating
' numbers from 000,000,000,000 up to 111,111,111,111 billion using only ones and zeroes
' or by converting decimal numbers  1 through 4095  to binary... (or  ... 2048 to 4095)...
' The total will boil down to 351 after "ReverseInStrduplicates" are removed.
' ====================================================================================

Private Sub List_Scales()
   
   Call LabelRoots
   lstAppendix.Clear
   
   Select Case Notecount
   Case 1
   Call Search_Octaves
   Case 2
   Call Search_Interval_Number
   Case 3
   Call Search_Triads
   Case 4
   Call Search_Tetratonics
   Case 5
   Call Search_Pentatonics
   Case 6
   Call Search_Hexatonics
   Case 7
   Call Search_Heptatonics
   Case 8
   Call Search_Eight
   Case 9
   Call Search_Nine
   Case 10
   Call Search_Ten
   Case 11
   Call Search_Eleven
   Case 12
   Call Search_Twelve
   End Select
 
   
   'Private Sub Search_X()
   
   'For Spelling = LBound(I) To UBound(I)
   '           OR
   'For Spelling = 0 To 11
                                                       ' =======================================================================================================================================
      'lstAppendix.AddItem I(Spelling) & " Major 7"    ' > This procedure appends the phrase " Major 7" to musical function Array I    >Private Sub LabelRoots().... I = Array("C","Db","D",etc...
                                                       '   This is a shorter  alternative to listing ...  .AddItem  "C Major 7"
   'Next Spelling                                      '                                                  .AddItem "Db Major 7"
                                                       '                                                  .AddItem  "D Major 7"...
                                                       '   ...which is done 12 times for each parent scale formula. There are  351 parent scale formulas.
                                                       '   Symmetrical scales or arpeggios are still listed 12 times since pitch names change when key signatures change.

                                                       ' =======================================================================================================================================
  
   
End Sub

Private Sub lstAppendix_Click()

   ' ==================================================================================
   ' In this section, a string of 12 ones and zeros is used to represent individual
   ' Music Scale intervals so that a significant mathematical relationship is
   ' immediately  established once the INPUT string data is entered. The ones
   ' are used to set control array lstPitch to a value of 1 (checked), while the
   ' zeroes are used to set lstPitch to a value of zero. The position of the characters
   ' in the string ... (>Mid<).. determines the index. Refer to procedure
   ' Private Sub CheckCheckboxes() which is called everytime an item is selected.
   ' ==================================================================================
   
   Me.MousePointer = 11
   
   lstInversions.Clear
   
   parentScaleFormula = dynamicList(lstAppendix.ListIndex \ 12)
   
   Call Extract12ParentScaleInversions
      
   Call CheckCheckboxes
   
   Call ChangeNames
   
   Call Rename
   
   Length = Len("FRET" & HScroll1.Value)
   txtPosition.Text = "FRET " & HScroll1.Value & Space(181 - Length) & lstAppendix.Text
   lblTotal.Caption = "Item No. : " & lstAppendix.ListIndex + 1 & " / " & lstAppendix.ListCount
      
   Call RemoveUnison
      
   Me.MousePointer = 0

End Sub

Private Sub lstAppendix_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   Dim Key As Integer
      
   Key = lstAppendix.ListIndex Mod 12
      
   lstRoot.ListIndex = Key
   
   If Notecount = 12 Then
   Call DrawObjects
   Call RemoveUnison
   End If
      
End Sub

Private Sub Extract12ParentScaleInversions()
   
'==============================================================================
' The procedure after the greens reduces the input formulas by 1 / 12TH.
' Only the parent scale formula "101011010101" is listed  ...
' followed by the names of the scale modes.
   

' dynamicList(0) = "101011010101"

' For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 216" & Space(3) & _
                                             'I(Spelling) & " Major Scale   or" & Space(3) & _
                                             'I(Spelling) & " Ionian Mode" & Space(3) & _
                                            'II(Spelling) & " Dorian Mode" & Space(3) & _
                                           'III(Spelling) & " Phrygian Mode" & Space(3) & _
                                            'IV(Spelling) & " Lydian Mode" & Space(3) & _
                                             'V(Spelling) & " Mixolydian Mode" & Space(3) & _
                                            'VI(Spelling) & " Aeolian Mode   or" & Space(3) & _
                                            'VI(Spelling) & " Natural Minor Scale" & Space(3) & _
                                           'VII(Spelling) & " Locrian Mode": Next Spelling
   
' The word (Spelling) is an integer variable representing the index of
' variant ArrayI, ArrayIsharp, ArrayII, ArrayIIIflat...
   
' Array names (II, III, IV, V, VI and VII) can be thought of as the DISTANCE
' or the INTERVAL of the root of a scale mode from the first note of the Major Scale
' or Key signature Tonic(I).


      
      Dim ctr As Integer
      Dim InversionIndex As Integer
      
      For ctr = Len(parentScaleFormula) To 1 Step -1
         lstInversions.AddItem Mid(parentScaleFormula, ctr + 1, Len(parentScaleFormula)) & Mid(parentScaleFormula, 1, ctr)  ' >>>>>>>>>>>>>>>>>>> lists all 12 inversions of the parent formula in an ASCENDING HALF STEP sequence ...
      Next ctr
         
      If lstAppendix.ListCount <> 0 Then
         InversionIndex = (lstAppendix.ListIndex) Mod 12  '  Mod > Make the first list index correspond with the second list index...
                                                           ' 0 Mod 12  = 0
                                                           ' 1 Mod 12  = 1
                                                           ' 2 Mod 12  = 2
                                                           ' 3 Mod 12  = 3
                                                           ' 4 Mod 12  = 4
                                                           ' 5 Mod 12  = 5
                                                           ' 6 Mod 12  = 6
                                                           ' 7 Mod 12  = 7
                                                           ' 8 Mod 12  = 8
                                                           ' 9 Mod 12  = 9
                                                          ' 10 Mod 12  = 10
                                                          ' 11 Mod 12  = 11
                                                          ' 12 Mod 12  = 0
                                                          ' 13 Mod 12  = 1
                                                          ' 14 Mod 12  = 2
                                                          ' 15 Mod 12  = 3
                                                          ' 16 Mod 12  = 4 etc...
         lstInversions.ListIndex = InversionIndex
         WholeandHalfStepEquivalent = lstInversions.Text
      End If

End Sub

Private Sub CheckCheckboxes()
      
      Dim characterposition As Byte
      
      For characterposition = 1 To Len(WholeandHalfStepEquivalent)
            If Mid(WholeandHalfStepEquivalent, characterposition, 1) = 1 Then
               lstPitch.Selected(characterposition - 1) = True
            ElseIf Mid(WholeandHalfStepEquivalent, characterposition, 1) = 0 Then
               lstPitch.Selected(characterposition - 1) = False
            End If
      Next characterposition

End Sub

Private Sub Rename()
   
   Dim commonIndex As Integer
   
   For commonIndex = 0 To 11
      If lstRoot.ListIndex <> -1 Then
         lstPitch.List(commonIndex) = PitchNames(commonIndex)
      End If
   Next commonIndex

End Sub

Private Sub ChangeNames()
   
   Select Case lstRoot.ListIndex
   
   ' This array is used to rename the checkboxes when the key modulates ...
   ' The Key Reference Functions are written horizontally.
   Case 0 ' Key of C
   PitchNames = Array("C", "C#", "D", "Eb", "E", "F", "F#", "G", "G#", "A", "Bb", "B")
   Case 1 ' Key of Db
   PitchNames = Array("C", "Db", "D", "Eb", "Fb", "F", "Gb", "G", "Ab", "A", "Bb", "Cb")
   Case 2 ' Key of D
   PitchNames = Array("C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B")
   Case 3 ' Key of Eb
   PitchNames = Array("C", "Db", "D", "Eb", "E", "F", "Gb", "G", "Ab", "A", "Bb", "B")
   Case 4 ' Key of E
   PitchNames = Array("B#", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B")
   Case 5 ' Key of F
   PitchNames = Array("C", "C#", "D", "Eb", "E", "F", "F#", "G", "Ab", "A", "Bb", "B")
   Case 6 ' Key of F#
   PitchNames = Array("B#", "C#", "Cx", "D#", "E", "E#", "F#", "Fx", "G#", "A", "A#", "B")
   Case 7 ' Key of G
   PitchNames = Array("C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "Bb", "B")
   Case 8 ' Key of Ab
   PitchNames = Array("C", "Db", "D", "Eb", "E", "F", "Gb", "G", "Ab", "A", "Bb", "Cb")
   Case 9 ' Key of A
   PitchNames = Array("C", "C#", "D", "D#", "E", "F", "F#", "G", "G#", "A", "A#", "B")
   Case 10 'Key of Bb
   PitchNames = Array("C", "Db", "D", "Eb", "E", "F", "F#", "G", "Ab", "A", "Bb", "B")
   Case 11 'Key of B
   PitchNames = Array("B#", "C#", "D", "D#", "E", "E#", "F#", "Fx", "G#", "A", "A#", "B")
         
   End Select
   
End Sub
   ' ============================================================================================================
   ' KEY REFERENCE FUNCTIONS = Array("I", "#I", "II", "bIII", "III", "IV", "#IV", "V", "#V", "VI", "bVII", "VII")
   ' "#I" "bIII" "IV" "#V" "bVII" represent notes not found in the key signature.
   ' ============================================================================================================

Private Sub LabelRoots()

   ' The names contained in these arrays are ATTACHED to the scale/chord names in the
   ' "Private Sub Search procedures" below...
   ' The Key Reference Functions  can be seen vertically and have identical indexes.
   
         I = Array("C", "Db", "D", "Eb", "E", "F", "F#", "G", "Ab", "A", "Bb", "B")
    Isharp = Array("C#", "D", "D#", "E", "F", "F#", "Fx", "G#", "A", "A#", "B", "B#")
        II = Array("D", "Eb", "E", "F", "F#", "G", "G#", "A", "Bb", "B", "C", "C#")
   IIIflat = Array("Eb", "Fb", "F", "Gb", "G", "Ab", "A", "Bb", "Cb", "C", "Db", "D")
       III = Array("E", "F", "F#", "G", "G#", "A", "A#", "B", "C", "C#", "D", "D#")
        IV = Array("F", "Gb", "G", "Ab", "A", "Bb", "B", "C", "Db", "D", "Eb", "E")
   IVsharp = Array("F#", "G", "G#", "A", "A#", "B", "B#", "C#", "D", "D#", "E", "E#")
         V = Array("G", "Ab", "A", "Bb", "B", "C", "C#", "D", "Eb", "E", "F", "F#")
    Vsharp = Array("G#", "A", "A#", "B", "B#", "C#", "Cx", "D#", "E", "F", "F#", "Fx")
        VI = Array("A", "Bb", "B", "C", "C#", "D", "D#", "E", "F", "F#", "G", "G#")
   VIIflat = Array("Bb", "Cb", "C", "Db", "D", "Eb", "E", "F", "Gb", "G", "Ab", "A")
       VII = Array("B", "C", "C#", "D", "D#", "E", "E#", "F#", "G", "G#", "A", "A#")

End Sub

Private Sub Search_Octaves()
ReDim dynamicList(0)

dynamicList(0) = "100000000000" ' R          [octaves]
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 001" & Space(3) & _
                                             I(Spelling) & " octaves": Next Spelling
 

End Sub


Private Sub Search_Interval_Number()

'======================
'INTERVALS (Total = 66)
'======================
ReDim dynamicList(5)

dynamicList(0) = "001000000100" ' 2 6        [4ths and 5ths]
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 002" & Space(3) & _
                                            II(Spelling) & "5   or   " & II(Spelling) & " perfect fifth" & Space(3) & _
                                            VI(Spelling) & " perfect fourth": Next Spelling

dynamicList(1) = "100010000000" ' R 3        [Maj 3rd and min6th]
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 003" & Space(3) & _
                                             I(Spelling) & " major third" & Space(3) & _
                                           III(Spelling) & " minor sixth": Next Spelling

dynamicList(2) = "001000000001" ' 2 7        [min 3rd and Maj 6th]
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 004" & Space(3) & _
                                            II(Spelling) & " major sixth" & Space(3) & _
                                           VII(Spelling) & " minor third ": Next Spelling

dynamicList(3) = "000000000101" ' 6 7        [whole step and min 7th]
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 005" & Space(3) & _
                                            VI(Spelling) & " major second" & Space(3) & _
                                           VII(Spelling) & " flatted seventh ": Next Spelling

dynamicList(4) = "100000000001" ' R 7        [half step and Maj 7th]
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 006" & Space(3) & _
                                           VII(Spelling) & " minor second" & Space(3) & _
                                             I(Spelling) & " major seventh ": Next Spelling

dynamicList(5) = "000001000001" ' 4 7        [b5]
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 007" & Space(3) & _
                                            IV(Spelling) & " augmented fourth" & Space(3) & _
                                           VII(Spelling) & " flatted fifth ": Next Spelling
End Sub

Private Sub Search_Triads()

'====================================
'TRIADS (3 NOTE SCALES) (Total = 220)
'====================================

ReDim dynamicList(18)
 
 dynamicList(0) = "100010010000" 'R 3 5      C
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 008" & Space(3) & _
                                             I(Spelling): Next Spelling
 
dynamicList(1) = "100010000100" 'R 3 6       Am, C6 voicing"
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 009" & Space(3) & _
                                            VI(Spelling) & "m" & Space(3) & _
                                             I(Spelling) & "6 voicing": Next Spelling
 
dynamicList(2) = "001001000001" '2 4 7       Dm6 voicing, Bdim
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 010" & Space(3) & _
                                             VII(Spelling) & "dim" & Space(3) & _
                                            II(Spelling) & "m6 voicing": Next Spelling

dynamicList(3) = "000001010001" '4 5 7       G7voicing
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 011" & Space(3) & _
                                             V(Spelling) & "7 voicing": Next Spelling

dynamicList(4) = "100010001000" 'R 3 #5      Caug, Eaug, G#aug '(symmetrical)
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 012" & Space(3) & _
                                             I(Spelling) & "aug" & Space(3) & _
                                           III(Spelling) & "aug" & Space(3) & _
                                        Vsharp(Spelling) & "aug": Next Spelling

dynamicList(5) = "101000010000" 'R 2 5       Csus2, Gsus4
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 013" & Space(3) & _
                                             I(Spelling) & "sus2" & Space(3) & _
                                             V(Spelling) & "sus4": Next Spelling
                                             
dynamicList(6) = "101001000000" 'R 2 4       Csus4/sus2, Dm7 voicing
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 014" & Space(3) & _
                                             I(Spelling) & "sus4/sus2" & Space(3) & _
                                            II(Spelling) & "m7 voicing": Next Spelling

dynamicList(7) = "101000000100" 'R 2 6       D7/5
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 015" & Space(3) & _
                                             II(Spelling) & "7/5": Next Spelling

dynamicList(8) = "100011000000" 'R 3 4       Cadd4 voicing (no 5th) PT
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 016" & Space(3) & _
                                             I(Spelling) & "add4 voicing": Next Spelling

dynamicList(9) = "100010000001" 'R 3 7       CM7 voicing
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 017" & Space(3) & _
                                             I(Spelling) & "M7 voicing": Next Spelling

dynamicList(10) = "000001000101" '4 6 7      F(b5)
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 018" & Space(3) & _
                                            IV(Spelling) & "(b5)": Next Spelling

dynamicList(11) = "000011000001" '3 4 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 019": Next Spelling

dynamicList(12) = "100001000001" 'R 4 7      CM7sus4 voicing (no 5th) PT
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 020" & Space(3) & _
                                             I(Spelling) & "M7sus4 voicing": Next Spelling

dynamicList(13) = "100000000101" 'R 6 7      Amadd9 voicing (no 5th) PT
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 021" & Space(3) & _
                                             VI(Spelling) & "madd9 voicing": Next Spelling

dynamicList(14) = "101000000001" 'R 2 7      (CM7sus2 voicing PT)
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 022": Next Spelling

dynamicList(15) = "101010000000" 'R 2 3      Cadd9 voicing (no 5th) PT
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 023" & Space(3) & _
                                             I(Spelling) & "add9 voicing": Next Spelling

dynamicList(16) = "100000001100" 'R #5 6     AmM7 voicing
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 024" & Space(3) & _
                                            VI(Spelling) & "mM7 voicing": Next Spelling

dynamicList(17) = "100000001001" 'R #5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 025": Next Spelling

dynamicList(18) = "000000001110" '#5 6 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 026": Next Spelling

End Sub

Private Sub Search_Tetratonics()

'=========================================
'TETRATONICS (4 NOTE SCALES) (Total = 495)
'=========================================

ReDim dynamicList(42)

dynamicList(0) = "100010010001" 'R 3 5 7     CM7, Emb6, GM13 voicing, (Am9 NR voicing)
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 027" & Space(3) & _
                                             I(Spelling) & "M7" & Space(3) & _
                                           III(Spelling) & "m(b6)" & Space(3) & _
                                            V(Spelling) & "M13 voicing" & Space(10) & _
                                            VI(Spelling) & "m9 NR voicing   ": Next Spelling
                                          
dynamicList(1) = "101001000100" 'R 2 4 6     Dm7, F6, (G9sus4 NR voicing), (G11 NR voicing), (BbM9 NR voicing)
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 028" & Space(3) & _
                                            II(Spelling) & "m7" & Space(3) & _
                                            IV(Spelling) & "6" & Space(10) & _
                                             V(Spelling) & "9sus4 NR voicing   or" & Space(3) & _
                                             V(Spelling) & "11 NR voicing" & Space(3) & _
                                       VIIflat(Spelling) & "M9 NR voicing": Next Spelling

dynamicList(2) = "001001010001" '2 4 5 7     G7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 029" & Space(3) & _
                                             V(Spelling) & "7": Next Spelling

dynamicList(3) = "001001000101" '2 4 6 7     Dm6, Bm7b5, (G9 NR voicing)
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 030" & Space(3) & _
                                            II(Spelling) & "m6" & Space(3) & _
                                           VII(Spelling) & "m7b5" & Space(10) & _
                                             V(Spelling) & "9 NR voicing ": Next Spelling

dynamicList(4) = "101010010000" 'R 2 3 5     Cadd2, Em7#5, G6sus4, D11voicing
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 031" & Space(3) & _
                                             I(Spelling) & "add9" & Space(3) & _
                                           III(Spelling) & "m7#5" & Space(3) & _
                                             V(Spelling) & "6sus4" & Space(3) & _
                                            II(Spelling) & "11 voicing": Next Spelling

dynamicList(5) = "100010000101" ' R 3 6 7    Amin(add2), C6M7, CM13voicing
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 032" & Space(3) & _
                                            VI(Spelling) & "m(add9)" & Space(3) & _
                                             I(Spelling) & "6M7   or" & Space(3) & _
                                             I(Spelling) & "M13 voicing": Next Spelling

dynamicList(6) = "100001010001" 'R 4 5 7     CM7sus4, CMll, Fsus2add#4
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 033" & Space(3) & _
                                             I(Spelling) & "M7sus4   or" & Space(3) & _
                                             I(Spelling) & "M11" & Space(3) & _
                                            IV(Spelling) & "sus2add#4": Next Spelling

dynamicList(7) = "101001010000" 'R 2 4 5     G7sus4, F9add6, Dm11 voicing, D11 voicing, (Bb 6/9 NR voicing)
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 034" & Space(3) & _
                                             V(Spelling) & "7sus4" & Space(3) & _
                                            IV(Spelling) & "9add6" & Space(3) & _
                                            II(Spelling) & "m11 voicing   or" & Space(3) & _
                                            II(Spelling) & "11 voicing" & Space(10) & _
                                            VIIflat(Spelling) & "6/9 NR voicing": Next Spelling

dynamicList(8) = "101000010001" 'R 2 5 7     Gadd4, CM9 voicing
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 035" & Space(3) & _
                                             V(Spelling) & "add4" & Space(3) & _
                                             I(Spelling) & "M9 voicing": Next Spelling

dynamicList(9) = "101010000001" 'R 2 3 7     CM9 voicing, D13voicing
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 036" & Space(3) & _
                                             I(Spelling) & "M9 voicing" & Space(3) & _
                                            II(Spelling) & "13 voicing": Next Spelling
 
dynamicList(10) = "101011000000" 'R 2 3 4    Dm9 voicing
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 037" & Space(3) & _
                                            II(Spelling) & "m9 voicing": Next Spelling
 
dynamicList(11) = "000001010101" '4 5 6 7    G9 voicing
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 038" & Space(3) & _
                                             V(Spelling) & "9 voicing": Next Spelling
 
dynamicList(12) = "001001010100" '2 4 5 6    G7sus2, G9 voicing, Dm11 voicing, (F6/9 voicing)
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 039" & Space(3) & _
                                             V(Spelling) & "7sus2   or" & Space(3) & _
                                             V(Spelling) & "9 voicing" & Space(3) & _
                                            II(Spelling) & "m11 voicing" & Space(3) & _
                                            IV(Spelling) & "6/9 voicing   ": Next Spelling

dynamicList(13) = "000011000101" '3 4 6 7    FM7b5, (G13 NR voicing)
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 040" & Space(3) & _
                                            IV(Spelling) & "M7b5" & Space(10) & _
                                             V(Spelling) & "13 NR voicing ": Next Spelling
 
dynamicList(14) = "100011000001" 'R 3 4 7    FM7(+11)
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 041" & Space(3) & _
                                            IV(Spelling) & "M7(+11)": Next Spelling
 
dynamicList(15) = "000011010001" '3 4 5 7    FM9(+11)voicing, G13voicing
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 042" & Space(3) & _
                                            IV(Spelling) & "M9(+11)" & Space(3) & _
                                             V(Spelling) & "13 voicing": Next Spelling
 
dynamicList(16) = "101001000001" 'R 2 4 7    D13#9 voicing
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 043" & Space(3) & _
                                            II(Spelling) & "13#9 voicing": Next Spelling
 
dynamicList(17) = "100001000101" 'R 4 6 7    B7b5b9voicing
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 044" & Space(3) & _
                                           VII(Spelling) & "7b5b9": Next Spelling
 
dynamicList(18) = "001011000001" '2 3 4 7    Dm6/9voicing, (G13 NR voicing)
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 045" & Space(3) & _
                                            II(Spelling) & "m6/9 voicing" & Space(10) & _
                                             V(Spelling) & "13 NR voicing": Next Spelling
 
dynamicList(19) = "101000000101" 'R 2 6 7    Amadd2add4 voicing (no 5th) PT
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 046" & Space(3) & _
                                            VI(Spelling) & "madd2add4 voicing": Next Spelling

dynamicList(20) = "100010001100" 'R 3 #5 6   AmM7, (D9(+11) NR voicing))
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 047" & Space(3) & _
                                            VI(Spelling) & "mM7" & Space(10) & _
                                            II(Spelling) & "9(+11) NR": Next Spelling
                                            
dynamicList(21) = "100001001100" 'R 4 #5 6   AmM7#5
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 048" & Space(3) & _
                                            VI(Spelling) & "mM7#5": Next Spelling
                                            
dynamicList(22) = "100000001101" 'R #5 6 7   AmM9 voicing
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 049" & Space(3) & _
                                            VI(Spelling) & "mM9 voicing": Next Spelling
                                            
dynamicList(23) = "100010001001" 'R 3 #5 7   CM7#5, Emajb6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 050" & Space(3) & _
                                             I(Spelling) & "M7#5" & Space(3) & _
                                           III(Spelling) & "majb6": Next Spelling

dynamicList(24) = "101010001000" 'R 2 3 #5   E7#5, D9(+11) voicing, (B9(+11)NR voicing)
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 051" & Space(3) & _
                                           III(Spelling) & "7#5" & Space(3) & _
                                           II(Spelling) & "9(+11) voicing" & Space(10) & _
                                           VIIflat(Spelling) & "9(+11) NR voicing": Next Spelling

dynamicList(25) = "001011001000" '2 3 4 #5   E7b9 voicing
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 052" & Space(3) & _
                                           III(Spelling) & "7b9 voicing": Next Spelling

dynamicList(26) = "000011001001" '3 4 #5 7   FmM7b5, (C#7#9 NR voicing), (G13b9 NR voicing)
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 053" & Space(3) & _
                                            IV(Spelling) & "mM7b5   " & _
                                        Isharp(Spelling) & "7#9 NR   " & _
                                             V(Spelling) & "13b9 NR": Next Spelling

dynamicList(27) = "001001001001" '2 4 #5 7   Ddim7, Fdim7, G#dim7, Bdim7 (symmetrical)(E7b9, G7b9, Bb7b9, C#7b9 NR voicings)
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 054" & Space(3) & _
                                            II(Spelling) & "dim7" & Space(3) & _
                                            IV(Spelling) & "dim7" & Space(3) & _
                                        Vsharp(Spelling) & "dim7" & Space(3) & _
                                           VII(Spelling) & "dim7" & Space(136) & _
                                           III(Spelling) & "7b9 NR voicing" & Space(3) & _
                                             V(Spelling) & "7b9 NR voicing" & Space(3) & _
                                       VIIflat(Spelling) & "7b9 NR voicing" & Space(3) & _
                                        Isharp(Spelling) & "7b9 NR voicing": Next Spelling

dynamicList(28) = "000011001100" '3 4 #5 6   FM7#9 voicing (no 5th) PT
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 055" & Space(3) & _
                                            IV(Spelling) & "M7#9 voicing": Next Spelling

dynamicList(29) = "100001001001" 'R 4 #5 7   CM7sus4#5 PT
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 056" & Space(3) & _
                                             I(Spelling) & "M7sus4#5": Next Spelling

dynamicList(30) = "101000001100" 'R 2 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 057": Next Spelling

dynamicList(31) = "101000001001" 'R 2 #5 7   (CM7sus2#5 PT)
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 058": Next Spelling

dynamicList(32) = "101000101000" 'R 2 #4 #5  D7b5"
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 059" & Space(3) & _
                                            II(Spelling) & "7b5": Next Spelling

dynamicList(33) = "100000101001" 'R #4 #5 7  G#7#9 voicing
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 060" & Space(3) & _
                                             Vsharp(Spelling) & "7#9 voicing": Next Spelling

dynamicList(34) = "000111001000" 'b3 3 4 #5  EM7b9 voicing (no 5th) PT
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 061" & Space(3) & _
                                           III(Spelling) & "M7b9": Next Spelling

dynamicList(35) = "000010001110" '3 #5 6 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 062": Next Spelling

dynamicList(36) = "000111000100" 'b3 3 4 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 063": Next Spelling

dynamicList(37) = "100111000000" 'R b3 3 4
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 064": Next Spelling

dynamicList(38) = "100000001110" 'R #5 6 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 065": Next Spelling

dynamicList(39) = "000000111100" '#4 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 066": Next Spelling

dynamicList(40) = "001011100000" '2 3 4 #4
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 067": Next Spelling

dynamicList(41) = "001000011100" '2 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 068": Next Spelling

dynamicList(42) = "100001100001" 'R 4 #4 7  (D13#9 NR voicing)
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 069" & Space(3) & _
                                            II(Spelling) & "13#9 NR": Next Spelling

End Sub


Private Sub Search_Pentatonics()

'=========================================
'PENTATONICS (5 NOTE SCALES) (Total = 792)
'=========================================
 
 ReDim dynamicList(65)
 
dynamicList(0) = "101010010100" 'R 2 3 5 6   C 6/9, C Pentatonic Major Scale, C Ryo, C Mongolian, C Chinese, C Pentatonic Majeur Mode, D Egyptian, D Yo, D9sus4,  Em11#5, G6sus4add9, G Ritsu, G Ritusen, G Hyogoi, Am7sus, Am11, A Pentatonic Minor Mode, A Pentatonic Minor Scale"
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 070" & Space(3) & _
                                            VI(Spelling) & " Pentatonic Minor Scale   or" & Space(3) & _
                                            VI(Spelling) & " Pentatonic Minor Mode   or" & Space(3) & _
                                            VI(Spelling) & "m7sus   or" & Space(3) & _
                                            VI(Spelling) & "m11" & Space(3) & _
                                             I(Spelling) & "6/9   or" & Space(3) & _
                                             I(Spelling) & " Pentatonic Major Scale   or" & Space(3) & _
                                             I(Spelling) & " Ryo   or" & Space(3) & _
                                             I(Spelling) & " Mongolian   or" & Space(3) & _
                                             I(Spelling) & " Chinese   or" & Space(3) & _
                                             I(Spelling) & " Pentatonic Majeur Mode" & Space(3) & _
                                            II(Spelling) & " Egyptian   or" & Space(3) & _
                                            II(Spelling) & " Yo   or" & Space(3) & _
                                            II(Spelling) & "9sus4   or" & Space(3) & _
                                           III(Spelling) & "m11#5" & Space(3) & _
                                             V(Spelling) & "6sus4add9   or" & Space(3) & _
                                             V(Spelling) & " Ritsu   or" & Space(3) & _
                                             V(Spelling) & " Hyogoi" & Space(3): Next Spelling

dynamicList(1) = "100011000101" 'R 3 4 6 7   E Japanese, (EgpKumoi), F Chinese,FM7(+11), A Hirajoshi, B Iwato Scale
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 071" & Space(3) & _
                                           III(Spelling) & " Japanese" & Space(3) & _
                                            IV(Spelling) & " Chinese   or" & Space(3) & _
                                            IV(Spelling) & "M7(+11)" & Space(3) & _
                                            VI(Spelling) & " Hirajoshi" & Space(3) & _
                                           VII(Spelling) & " Iwato": Next Spelling
   
dynamicList(2) = "001011000101" '2 3 4 6 7   D Kumoi, Dm6/9, Dmin6thadd9th, E Kokin Joshi, E Hybrid, B P5
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 072" & Space(3) & _
                                            II(Spelling) & " Kumoi   or" & Space(3) & _
                                            II(Spelling) & "m6/9   or" & Space(3) & _
                                            II(Spelling) & "m6add9   or" & Space(3) & _
                                           III(Spelling) & " Kokin Joshi" & Space(3) & _
                                           VII(Spelling) & " P5": Next Spelling
 
dynamicList(3) = "100011010001" 'R 3 4 5 7   E Balinese, (or E Pelog1)
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 073" & Space(3) & _
                                           III(Spelling) & " Balinese": Next Spelling
                                           
dynamicList(4) = "001010011001" '2 3 5 #5 7  G Scriabin, G Altered Pentatonic Mode, E7#9, E Hybrid, (Bb13b5b9 NR voicing)"
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 074" & Space(3) & _
                                             V(Spelling) & " Scriabin   or" & Space(3) & _
                                             V(Spelling) & " Altered Pentatonic Mode" & Space(3) & _
                                           III(Spelling) & "7#9" & Space(10) & _
                                       VIIflat(Spelling) & "13b5b9 NR voicing": Next Spelling

dynamicList(5) = "101010101000" 'R 2 3 #4 #5 D9b5, D9(+11), D Prometheus Scale, E9#5
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 075" & Space(3) & _
                                            II(Spelling) & "9b5   or" & Space(3) & _
                                            II(Spelling) & "9(+11)   or" & Space(3) & _
                                            II(Spelling) & " Prometheus" & Space(3) & _
                                           III(Spelling) & "9#5": Next Spelling
 
dynamicList(6) = "001011001001" '2 3 4 #5 7  E7b9, (G13b9 NR voicing), D Hybrid
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 076" & Space(3) & _
                                           III(Spelling) & "7b9" & Space(3) & _
                                             V(Spelling) & "13b9 NR voicing" & Space(3) & _
                                            II(Spelling) & " Hybrid": Next Spelling
                                           
dynamicList(7) = "001011010001" '2 3 4 5 7   E Pelog, E Hybrid, G13 voicing,
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 077" & Space(3) & _
                                           III(Spelling) & " Pelog" & Space(3) & _
                                             V(Spelling) & "13 voicing": Next Spelling
                                           
dynamicList(8) = "100011010100" 'R 3 4 5 6   FM9, G13 voicing
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 078" & Space(3) & _
                                            IV(Spelling) & "M9" & Space(3) & _
                                             V(Spelling) & "13 voicing": Next Spelling
                                           
dynamicList(9) = "101011000100" 'R 2 3 4 6   Dm9
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 079" & Space(3) & _
                                            II(Spelling) & "m9": Next Spelling
                                           
dynamicList(10) = "001001010101" '2 4 5 6 7  G9, G Pentatonic Dominant Mode, F6b5add9, FM13(+11)
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 080" & Space(3) & _
                                             V(Spelling) & "9   or" & Space(3) & _
                                             V(Spelling) & " Pentatonic Dominant Mode" & Space(3) & _
                                            IV(Spelling) & "6b5add9   or" & Space(3) & _
                                            IV(Spelling) & "M13(+11)": Next Spelling
 
dynamicList(11) = "100001010101" 'R 4 5 6 7  Am9#5
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 081" & Space(3) & _
                                            VI(Spelling) & "m9#5": Next Spelling

dynamicList(12) = "000011010101" ' 3 4 5 6 7 FM9(+11), G13 voicing
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 082" & Space(3) & _
                                            IV(Spelling) & "M9(+11)" & Space(3) & _
                                             V(Spelling) & "13 voicing": Next Spelling
 
dynamicList(13) = "101001010001" 'R 2 4 5 7  G7add4, G Hybrid, CM9sus4, CM11, Dm13 voicing, D Hybrid,
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 083" & Space(3) & _
                                             V(Spelling) & "7add4   or" & Space(3) & _
                                             V(Spelling) & " Hybrid" & Space(3) & _
                                             I(Spelling) & "M9sus4   or" & Space(3) & _
                                             I(Spelling) & "M11" & Space(3) & _
                                            II(Spelling) & "m13 voicing   or" & Space(3) & _
                                            II(Spelling) & " Hybrid": Next Spelling
 
dynamicList(14) = "101011000001" 'R 2 3 4 7   Dm13voicing
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 084" & Space(3) & _
                                            II(Spelling) & "m13 voicing": Next Spelling

dynamicList(15) = "101001000101" 'R 2 4 6 7  Am11#5 voicing
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 085" & Space(3) & _
                                            VI(Spelling) & "m11#5 voicing": Next Spelling

dynamicList(16) = "101011010000" 'R 2 3 4 5  Dm11voicing
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 086" & Space(3) & _
                                            II(Spelling) & "m11 voicing": Next Spelling

dynamicList(17) = "001011010100" '2 3 4 5 6  G13voicing
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 087" & Space(3) & _
                                             V(Spelling) & "13 voicing": Next Spelling

dynamicList(18) = "100001001101" 'R 4 #5 6 7 AmM9#5 PT
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 088" & Space(3) & _
                                            VI(Spelling) & "mM9#5": Next Spelling

dynamicList(19) = "100011001001" 'R 3 4 #5 7 FmM7(+11) PT
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 089" & Space(3) & _
                                            IV(Spelling) & "mM7(+11)": Next Spelling

dynamicList(20) = "100011001100" 'R 3 4 #5 6 FM7#9 PT
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 090" & Space(3) & _
                                            IV(Spelling) & "M7#9": Next Spelling

dynamicList(21) = "000011001101" '3 4 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 091": Next Spelling

dynamicList(22) = "001011001100" '2 3 4 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 092": Next Spelling

dynamicList(23) = "101001001100" 'R 2 4 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 093": Next Spelling

dynamicList(24) = "101001001001" 'R 2 4 #5 7 C Hybrid, B Hybrid
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 094" & Space(3) & _
                                             I(Spelling) & " Hybrid" & Space(3) & _
                                           VII(Spelling) & " Hybrid": Next Spelling

dynamicList(25) = "101011001000" 'R 2 3 4 #5 E7#5b9, D hybrid
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 095" & Space(3) & _
                                           III(Spelling) & "7#5b9" & Space(3) & _
                                            II(Spelling) & " Hybrid": Next Spelling

dynamicList(26) = "101000001101" 'R 2 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 096": Next Spelling
                                            
dynamicList(27) = "101010001001" 'R 2 3 #5 7 E7add#5
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 097" & Space(3) & _
                                           III(Spelling) & "7add#5": Next Spelling

dynamicList(28) = "100010001101" 'R 3 #5 6 7 AmM9
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 098" & Space(3) & _
                                            VI(Spelling) & "mM9": Next Spelling

dynamicList(29) = "101010001100" 'R 2 3 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 099": Next Spelling

dynamicList(30) = "101000101001" 'R 2 #4 #5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 100": Next Spelling

dynamicList(31) = "101000101100" 'R 2 #4 #5 6 G#7b5b9, D7#11
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 101" & Space(3) & _
                                        Vsharp(Spelling) & "7b5b9   " & _
                                            II(Spelling) & "7#11 ": Next Spelling

dynamicList(32) = "100000101101" 'R #4 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 102": Next Spelling

dynamicList(33) = "100010101001" 'R 3 #4 #5 7 G#7#5#9"
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 103" & Space(3) & _
                                             V(Spelling) & "7#5#9": Next Spelling

dynamicList(34) = "000011011001" '3 4 5 #5 7 G13b9, F Hybrid"
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 104" & Space(3) & _
                                             V(Spelling) & "13b9" & Space(3) & _
                                            IV(Spelling) & " Hybrid": Next Spelling

dynamicList(35) = "100010011001" 'R 3 5 #5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 105": Next Spelling

dynamicList(36) = "000111001100" 'b3 3 4 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 106": Next Spelling

dynamicList(37) = "000111001001" 'b3 3 4 #5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 107": Next Spelling

dynamicList(38) = "100111001000" 'R b3 3 4 #5
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 108": Next Spelling

dynamicList(39) = "100111000001" 'R b3 3 4 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 109": Next Spelling

dynamicList(40) = "000111000101" 'b3 3 4 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 110": Next Spelling

dynamicList(41) = "100111000100" 'R b3 3 4 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 111": Next Spelling

dynamicList(42) = "100010001110" 'R 3 #5 6 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 112": Next Spelling

dynamicList(43) = "101000001110" 'R 2 #5 6 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 113": Next Spelling

dynamicList(44) = "100001001110" 'R 4 #5 6 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 114": Next Spelling

dynamicList(45) = "101001100001" 'R 2 4 #4 7 D13#9 voicing
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 115" & Space(3) & _
                                            II(Spelling) & "13#9 voicing": Next Spelling

dynamicList(46) = "100011100001" 'R 3 4 #4 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 116": Next Spelling

dynamicList(47) = "100001100101" 'R 4 #4 6 7 B7b9(+11), (D13#9 NR voicing)
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 117" & Space(3) & _
                                           VII(Spelling) & "7b9(+11)" & Space(3) & _
                                            II(Spelling) & "13#9 NR voicing": Next Spelling

dynamicList(48) = "001011100001" '2 3 4 #4 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 118": Next Spelling

dynamicList(49) = "101011100000" 'R 2 3 4 #4
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 119": Next Spelling

dynamicList(50) = "001011100100" '2 3 4 #4 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 120": Next Spelling

dynamicList(51) = "101000011100" 'R 2 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 121": Next Spelling

dynamicList(52) = "100000111100" 'R #4 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 122": Next Spelling

dynamicList(53) = "100010111000" 'R 3 #4 5 #5
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 123": Next Spelling

dynamicList(54) = "001000111100" '2 #4 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 124": Next Spelling

dynamicList(55) = "000010111100" '3 #4 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 125": Next Spelling

dynamicList(56) = "001010011100" '2 3 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 126": Next Spelling

dynamicList(57) = "100001110001" 'R 4 #4 5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 127": Next Spelling

dynamicList(58) = "100001110100" 'R 4 #4 5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 128": Next Spelling

dynamicList(59) = "000001011101" '4 5 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 129": Next Spelling

dynamicList(60) = "100000011101" 'R 5 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 130": Next Spelling

dynamicList(61) = "000011011100" '3 4 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 131": Next Spelling

dynamicList(62) = "001111000001" '2 b3 3 4 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 132": Next Spelling

dynamicList(63) = "001111000100" '2 b3 3 4 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 133": Next Spelling

dynamicList(64) = "000011110100" '3 4 #4 5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 134": Next Spelling

dynamicList(65) = "100000001111" 'R #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 135": Next Spelling

End Sub

Private Sub Search_Hexatonics()

'========================================
'HEXATONICS (6 NOTE SCALES) (Total = 924)
'========================================

ReDim dynamicList(79)

dynamicList(0) = "110011001100" 'R #1 3 4 #5 6  C, E, G# Six Tone Symmetrical, C#, F, A Augmented Scale
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 136" & Space(3) & _
                                            VI(Spelling) & " Augmented Scale" & Space(3) & _
                                            IV(Spelling) & " Augmented Scale" & Space(3) & _
                                             Isharp(Spelling) & " Augmented Scale" & Space(81) & _
                                             I(Spelling) & " Six Tone Symmetrical" & Space(3) & _
                                           III(Spelling) & " Six Tone Symmetrical" & Space(3) & _
                                        Vsharp(Spelling) & " Six Tone Symmetrical": Next Spelling

dynamicList(1) = "001011100101" '2 3 4 #4 6 7 B Blues Scale, B Blues Minor Scale
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 137" & Space(3) & _
                                           VII(Spelling) & " Blues Scale": Next Spelling

dynamicList(2) = "001001011101" '2 4 5 #5 6 7 D Blues Major Scale, D Blues Mode
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 138": Next Spelling
                                            
dynamicList(3) = "101010010101" 'R 2 3 5 6 7 D Pyongio, CM9add6, Em9sus, Em11
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 139" & Space(3) & _
                                            II(Spelling) & " Pyongio" & Space(3) & _
                                             I(Spelling) & "M9add6" & Space(3) & _
                                           III(Spelling) & "m9sus   or" & Space(3) & _
                                           III(Spelling) & "m11": Next Spelling

dynamicList(4) = "101010101010" 'R 2 3 #4 #5 b7 C D E F# G# Bb Whole Tone Scale"
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 140" & Space(3) & _
                                             I(Spelling) & " Whole Tone Scale" & Space(3) & _
                                            II(Spelling) & " Whole Tone Scale" & Space(3) & _
                                           III(Spelling) & " Whole Tone Scale" & Space(3) & _
                                       IVsharp(Spelling) & " Whole Tone Scale" & Space(3) & _
                                        Vsharp(Spelling) & " Whole Tone Scale" & Space(3) & _
                                       VIIflat(Spelling) & " Whole Tone Scale": Next Spelling

dynamicList(5) = "100101100101" 'R b3 4 #4 6 7 F7b9(+11), B7b9(+11)
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 141" & Space(3) & _
                                            IV(Spelling) & "7b9(+11)" & Space(3) & _
                                           VII(Spelling) & "7b9(+11)": Next Spelling

dynamicList(6) = "101001101001" 'R 2 4 #4 #5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 142": Next Spelling

dynamicList(7) = "101011010001" 'R 2 3 4 5 7 (questionable Maj 11th)"
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 143": Next Spelling

dynamicList(8) = "101011000101" 'R 2 3 4 6 7 F6M7#11
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 144" & Space(3) & _
                                              IV(Spelling) & "6M7#11": Next Spelling

dynamicList(9) = "100011010101" 'R 3 4 5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 145": Next Spelling

dynamicList(10) = "001011010101" '2 3 4 5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 146": Next Spelling

dynamicList(11) = "101001010101" 'R 2 4 5 6 7 CM13 voicing, Am11#5"
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 147" & Space(3) & _
                                              I(Spelling) & "M13" & Space(3) & _
                                             VI(Spelling) & "m11#5": Next Spelling

dynamicList(12) = "101001001101" 'R 2 4 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 148": Next Spelling

dynamicList(13) = "100011001101" 'R 3 4 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 149": Next Spelling

dynamicList(14) = "101011001001" 'R 2 3 4 #5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 150": Next Spelling

dynamicList(15) = "001011001101" '2 3 4 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 151": Next Spelling

dynamicList(16) = "101011001100" 'R 2 3 4 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 152": Next Spelling

dynamicList(17) = "101010001101" 'R 2 3 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 153": Next Spelling

dynamicList(18) = "101000101101" 'R 2 #4 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 154": Next Spelling

dynamicList(19) = "101010101001" 'R 2 3 #4 #5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 155": Next Spelling

dynamicList(20) = "101010101100" 'R 2 3 #4 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 156": Next Spelling

dynamicList(21) = "100010101101" 'R 3 #4 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 157": Next Spelling

dynamicList(22) = "001011011001" '2 3 4 5 #5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 158": Next Spelling

dynamicList(23) = "101010011001" 'R 2 3 5 #5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 159": Next Spelling

dynamicList(24) = "100011011001" 'R 3 4 5 #5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 160": Next Spelling

dynamicList(25) = "000111001101" 'b3 3 4 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 161": Next Spelling

dynamicList(26) = "100111001001" 'R b3 3 4 #5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 162": Next Spelling

dynamicList(27) = "100101001101" 'R b3 4 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 163": Next Spelling

dynamicList(28) = "100111001100" 'R b3 3 4 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 164": Next Spelling

dynamicList(29) = "100111000101" 'R b3 3 4 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 165": Next Spelling

dynamicList(30) = "101001001110" 'R 2 4 #5 6 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 166": Next Spelling

dynamicList(31) = "100011001110" 'R 3 4 #5 6 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 167": Next Spelling

dynamicList(32) = "101010001110" 'R 2 3 #5 6 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 168": Next Spelling

dynamicList(33) = "100011100101" 'R 3 4 #4 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 169": Next Spelling

dynamicList(34) = "101011100001" 'R 2 3 4 #4 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 170": Next Spelling

dynamicList(35) = "101001100101" 'R 2 4 #4 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 171": Next Spelling

dynamicList(36) = "101011100100" 'R 2 3 4 #4 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 172": Next Spelling

dynamicList(37) = "101000111100" 'R 2 #4 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 173": Next Spelling

dynamicList(38) = "100010111100" 'R 3 #4 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 174": Next Spelling

dynamicList(39) = "001010111100" '2 3 #4 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 175": Next Spelling

dynamicList(40) = "101010011100" 'R 2 3 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 176": Next Spelling

dynamicList(41) = "101010111000" 'R 2 3 #4 5 #5
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 177": Next Spelling

dynamicList(42) = "101000101110" 'R 2 #4 #5 6 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 178": Next Spelling

dynamicList(43) = "100010101110" 'R 3 #4 #5 6 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 179": Next Spelling

dynamicList(44) = "100001101101" 'R 4 #4 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 180": Next Spelling

dynamicList(45) = "100001110101" 'R 4 #4 5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 181": Next Spelling

dynamicList(46) = "100101001110" 'R b3 4 #5 6 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 182": Next Spelling

dynamicList(47) = "100111000110" 'R b3 3 4 6 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 183": Next Spelling

dynamicList(48) = "100110001110" 'R b3 3 #5 6 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 184": Next Spelling

dynamicList(49) = "000111001110" 'b3 3 4 #5 6 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 185": Next Spelling

dynamicList(50) = "000011011101" '3 4 5 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 186": Next Spelling

dynamicList(51) = "100001011101" 'R 4 5 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 187": Next Spelling

dynamicList(52) = "101000011101" 'R 2 5 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 188": Next Spelling

dynamicList(53) = "100010011101" 'R 3 5 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 189": Next Spelling

dynamicList(54) = "100011011100" 'R 3 4 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 190": Next Spelling

dynamicList(55) = "001011011100" '2 3 4 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 191": Next Spelling

dynamicList(56) = "101111000001" 'R 2 b3 3 4 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 192": Next Spelling

dynamicList(57) = "001111010001" '2 b3 3 4 5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 193": Next Spelling

dynamicList(58) = "001111000101" '2 b3 3 4 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 194": Next Spelling

dynamicList(59) = "001111010100" '2 b3 3 4 5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 195": Next Spelling

dynamicList(60) = "101111000100" 'R 2 b3 3 4 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 196": Next Spelling

dynamicList(61) = "101111010000" 'R 2 b3 3 4 5
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 197": Next Spelling

dynamicList(62) = "100011110001" 'R 3 4 #4 5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 198": Next Spelling

dynamicList(63) = "100011110100" 'R 3 4 #4 5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 199": Next Spelling

dynamicList(64) = "111011000001" 'R #1 2 3 4 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 200": Next Spelling

dynamicList(65) = "110011000101" 'R #1 3 4 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 201": Next Spelling

dynamicList(66) = "101110001001" 'R 2 b3 3 #5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 202": Next Spelling

dynamicList(67) = "101110001100" 'R 2 b3 3 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 203": Next Spelling

dynamicList(68) = "101000001111" 'R 2 #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 204": Next Spelling

dynamicList(69) = "100011101100" 'R 3 4 #4 #5 6 (13#5b9)
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 205": Next Spelling

dynamicList(70) = "100010001111" 'R 3 #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 206": Next Spelling

dynamicList(71) = "001111001100" '2 b3 3 4 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 207": Next Spelling

dynamicList(72) = "000011001111" '3 4 #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 208": Next Spelling

dynamicList(73) = "100001001111" 'R 4 #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 209": Next Spelling

dynamicList(74) = "000011111100" '3 4 #4 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 210": Next Spelling

dynamicList(75) = "001011111000" '2 3 4 #4 5 #5
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 211": Next Spelling

dynamicList(76) = "001001111001" '2 4 #4 5 #5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 212": Next Spelling

dynamicList(77) = "100001111100" 'R 4 #4 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 213": Next Spelling

dynamicList(78) = "000111011100" 'b3 3 4 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 214": Next Spelling

dynamicList(79) = "001110001110" '2 b3 3 #5 6 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 215": Next Spelling

End Sub

Private Sub Search_Heptatonics()

'==========================
' HEPTATONICS (Total = 792)
'==========================

ReDim dynamicList(65)

dynamicList(0) = "101011010101" 'R 2 3 4 5 6 7  C Major Scale, C Ionian Mode, D Dorian Mode, Dm13, E Phrygian Mode, F Lydian Mode, G Mixolydian, G13, A Aeolian Mode, A Natural Minor Scale, B Locrian"
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 216" & Space(3) & _
                                             I(Spelling) & " Major Scale   or" & Space(3) & _
                                             I(Spelling) & " Ionian Mode" & Space(3) & _
                                            II(Spelling) & " Dorian Mode" & Space(3) & _
                                           III(Spelling) & " Phrygian Mode" & Space(3) & _
                                            IV(Spelling) & " Lydian Mode" & Space(47) & _
                                             V(Spelling) & " Mixolydian Mode" & Space(3) & _
                                            VI(Spelling) & " Aeolian Mode   or" & Space(3) & _
                                            VI(Spelling) & " Natural Minor Scale" & Space(3) & _
                                           VII(Spelling) & " Locrian Mode": Next Spelling
 
dynamicList(1) = "101011001101" 'R 2 3 4 #5 6 7 C Harmonic Major, D Romanian, D Ukrainian, D Dorian Sharp 4 Mode, E Phrygian Major, E Jewish Scale, E Espana, E Spanish Scale, F Lydian Sharp 2 mode,  G# Ultra Locrian, A Harmonic Minor Scale, B Locrian Nat 6 Mode"
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 217" & Space(3) & _
                                            VI(Spelling) & " Harmonic Minor Scale" & Space(3) & _
                                           VII(Spelling) & " Locrian Nat 6 Mode" & Space(3) & _
                                             I(Spelling) & " Harmonic Major Scale" & Space(3) & _
                                            II(Spelling) & " Romanian   or" & Space(3) & _
                                            II(Spelling) & " Ukrainian" & Space(3) & _
                                            II(Spelling) & " Dorian Sharp 4" & Space(3) & _
                                           III(Spelling) & " Phrygian Major   or" & Space(3) & _
                                           III(Spelling) & " Jewish Scale   or" & Space(3) & _
                                           III(Spelling) & " Espana   or" & Space(3) & _
                                           III(Spelling) & " Spanish Scale" & Space(3) & _
                                            IV(Spelling) & " Lydian Sharp 2 Mode" & Space(3) & _
                                        Vsharp(Spelling) & " Ultra Locrian": Next Spelling

dynamicList(2) = "101010101101" 'R 2 3 #4 #5 6 7 C Lydian Augmented, D Dominant Lydian, D Lydian Dominant,D Bartok Scale, D Lydian Flat 7 Minor Mode, D Overtone Scale, E Hindu(India), E Mixolydian Flat 6 Minor Mode, F# Locrian Natural 2nd, G# Super Locrian,  G# Diminished Whole Tone, A Melodic Minor Scale (Asc), A Hawaiian Scale, A Jazz Minor Mode,  B Javanese Scale
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 218" & Space(3) & _
                                            VI(Spelling) & " Ascending Melodic Minor   or" & Space(3) & _
                                            VI(Spelling) & " Hawaiian Scale   or" & Space(3) & _
                                            VI(Spelling) & " Jazz Minor Mode" & Space(3) & _
                                           VII(Spelling) & " Dorian b2" & Space(3) & _
                                           VII(Spelling) & " Javanese Scale" & Space(3) & _
                                             I(Spelling) & " Lydian Augmented" & Space(3) & _
                                            II(Spelling) & " Dominant Lydian   or" & Space(3) & _
                                            II(Spelling) & " Lydian Dominant or" & Space(3) & _
                                            II(Spelling) & " Bartok Scale or" & Space(3) & _
                                            II(Spelling) & " Lydian Flat 7 Minor Mode   or" & Space(3) & _
                                            II(Spelling) & " Overtone" & Space(3) & _
                                           III(Spelling) & " Hindu(India)   or" & Space(3) & _
                                           III(Spelling) & " Mixolydian Flat 6 Minor Mode" & Space(3) & _
                                       IVsharp(Spelling) & " Locrian Natural 2nd" & Space(3) & _
                                        Vsharp(Spelling) & " Super Locrian   or" & Space(3) & _
                                        Vsharp(Spelling) & " Diminished Whole Tone ": Next Spelling

dynamicList(3) = "101011011001" 'R 2 3 4 5 #5 7 C Ethiopian, E Indian, G Gypsy Scale, G13b9
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 219" & Space(3) & _
                                             I(Spelling) & " Ethiopian" & Space(3) & _
                                           III(Spelling) & " Indian" & Space(3) & _
                                             V(Spelling) & " Gypsy   or" & Space(3) & _
                                             V(Spelling) & "13b9": Next Spelling
 
dynamicList(4) = "100111001101" 'R b3 3 4 #5 6 7 C Byzantine, E Double Harmonic, A Algerian, A Hungarian  Minor, B Oriental, B13b5b9
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 220" & Space(3) & _
                                            VI(Spelling) & " Hungarian Minor" & Space(3) & _
                                            VI(Spelling) & " Algerian" & Space(3) & _
                                           VII(Spelling) & " Oriental" & Space(3) & _
                                           VII(Spelling) & "13b5b9" & Space(3) & _
                                             I(Spelling) & " Byzantine" & Space(3) & _
                                           III(Spelling) & " Double Harmonic" & Space(3): Next Spelling

dynamicList(5) = "101011001110" 'R 2 3 4 #5 6 b7 C Mixolydian Augmented, D Hungarian Gypsy,  A Neopolitan Minor
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 221" & Space(3) & _
                                            VI(Spelling) & " Neopolitan Minor" & Space(3) & _
                                             I(Spelling) & " Mixolydian Augmented" & Space(3) & _
                                            II(Spelling) & " Hungarian Gypsy": Next Spelling

dynamicList(6) = "101011100101" 'R 2 3 4 #4 6 7 Marva(Indian)
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 222" & Space(3) & _
                                             IV(Spelling) & " Marva": Next Spelling

dynamicList(7) = "101010111100" 'R 2 3 #4 5 #5 6 Enigmatic
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 223" & Space(3) & _
                                         Vsharp(Spelling) & " Enigmatic": Next Spelling
 
dynamicList(8) = "101001101101" 'R 2 4 #4 #5 6 7 D Hungarian Major
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 224" & Space(3) & _
                                            II(Spelling) & " Hungarian Major": Next Spelling
                                                           
dynamicList(9) = "101010101110" 'R 2 3 #4 #5 6 b7 C Leading Whole Tone Scale, D Lydian Minor, E Arabian, E Locrian Major, A Neopolitan Major,A Dorian b2 Minor Mode
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 225" & Space(3) & _
                                            VI(Spelling) & " Neopolitan Major" & Space(3) & _
                                             I(Spelling) & " Leading Whole Tone" & Space(3) & _
                                            II(Spelling) & " Lydian Minor" & Space(3) & _
                                           III(Spelling) & " Arabian   or" & Space(3) & _
                                           III(Spelling) & " Locrian Major": Next Spelling

dynamicList(10) = "100111001110" 'R b3 3 4 #5 6 b7 E Persian, A Todi(Indian)
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 226" & Space(3) & _
                                            VI(Spelling) & " Todi" & Space(3) & _
                                           III(Spelling) & " Persian": Next Spelling
                                            
dynamicList(11) = "101001110101" 'R 2 4 #4 5 6 7 Rock 'n Roll
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 227" & Space(3) & _
                                            II(Spelling) & " Rock n Roll": Next Spelling

dynamicList(12) = "101001011101" 'R 2 4 5 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 228": Next Spelling

dynamicList(13) = "100011011101" 'R 3 4 5 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 229": Next Spelling

dynamicList(14) = "001011011101" '2 3 4 5 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 230": Next Spelling

dynamicList(15) = "101010011101" 'R 2 3 5 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 231": Next Spelling

dynamicList(16) = "101011011100" 'R 2 3 4 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 232": Next Spelling

dynamicList(17) = "001111010101" '2 b3 3 4 5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 233": Next Spelling

dynamicList(18) = "101111010001" 'R 2 b3 3 4 5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 234": Next Spelling

dynamicList(19) = "101111000101" 'R 2 b3 3 4 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 235": Next Spelling

dynamicList(20) = "101111010100" 'R 2 b3 3 4 5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 236": Next Spelling

dynamicList(21) = "100011110101" 'R 3 4 #4 5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 237": Next Spelling

dynamicList(22) = "101011110001" 'R 2 3 4 #4 5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 238": Next Spelling

dynamicList(23) = "101011110100" 'R 2 3 4 #4 5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 239": Next Spelling

dynamicList(24) = "110110110010" 'R #1 b3 3 #4 5 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 240": Next Spelling

dynamicList(25) = "111011000101" 'R #1 2 3 4 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 241": Next Spelling

dynamicList(26) = "101111001001" 'R 2 b3 3 4 #5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 242": Next Spelling

dynamicList(27) = "111011001001" 'R #1 2 3 4 #5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 243": Next Spelling

dynamicList(28) = "101011001011" 'R 2 3 4 #5 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 244": Next Spelling

dynamicList(29) = "100011001111" 'R 3 4 #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 245": Next Spelling

dynamicList(30) = "110011001101" 'R #1 3 4 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 246": Next Spelling

dynamicList(31) = "101111001100" 'R 2 b3 3 4 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 247": Next Spelling

dynamicList(32) = "101001001111" 'R 2 4 #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 248": Next Spelling

dynamicList(33) = "101010001111" 'R 2 3 #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 249": Next Spelling

dynamicList(34) = "101110001101" 'R 2 b3 3 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 250": Next Spelling

dynamicList(35) = "100011101101" 'R 3 4 #4 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 251": Next Spelling

dynamicList(36) = "101011101001" 'R 2 3 4 #4 #5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 252": Next Spelling

dynamicList(37) = "101011101100" 'R 2 3 4 #4 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 253": Next Spelling

dynamicList(38) = "101000101111" 'R 2 #4 #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 254": Next Spelling

dynamicList(39) = "101000111101" 'R 2 #4 5 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 255": Next Spelling

dynamicList(40) = "101010111001" 'R 2 3 #4 5 #5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 256": Next Spelling

dynamicList(41) = "101011111000" 'R 2 3 4 #4 5 #5
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 257": Next Spelling

dynamicList(42) = "101001111001" 'R 2 4 #4 5 #5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 258": Next Spelling

dynamicList(43) = "100011111001" 'R 3 4 #4 5 #5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 259": Next Spelling

dynamicList(44) = "001011111001" '2 3 4 #4 5 #5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 260": Next Spelling

dynamicList(45) = "100111011001" 'R b3 3 4 5 #5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 261": Next Spelling

dynamicList(46) = "000111011101" 'b3 3 4 5 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 262": Next Spelling

dynamicList(47) = "100111011100" 'R b3 3 4 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 263": Next Spelling

dynamicList(48) = "100111100101" 'R b3 3 4 #4 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 264": Next Spelling

dynamicList(49) = "110111000101" 'R #1 b3 3 4 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 265": Next Spelling

dynamicList(50) = "100111000111" 'R b3 3 4 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 266": Next Spelling

dynamicList(51) = "000111001111" 'b3 3 4 #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 267": Next Spelling

dynamicList(52) = "000011111101" '3 4 #4 5 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 268": Next Spelling

dynamicList(53) = "100001111101" 'R 4 #4 5 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 269": Next Spelling

dynamicList(54) = "101001100111" 'R 2 4 #4 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 270": Next Spelling

dynamicList(55) = "100001110111" 'R 4 #4 5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 271": Next Spelling

dynamicList(56) = "001111110001" '2 b3 3 4 #4 5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 272": Next Spelling

dynamicList(57) = "101111100001" 'R 2 b3 3 4 #4 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 273": Next Spelling

dynamicList(58) = "100011111100" 'R 3 4 #4 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 274": Next Spelling

dynamicList(59) = "001011111100" '2 3 4 #4 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 275": Next Spelling

dynamicList(60) = "101001111100" 'R 2 4 #4 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 276": Next Spelling

dynamicList(61) = "101110011100" 'R 2 b3 3 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 277": Next Spelling

dynamicList(62) = "001111011100" '2 b3 3 4 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 278": Next Spelling

dynamicList(63) = "111111000001" 'R #1 2 b3 3 4 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 279": Next Spelling

dynamicList(64) = "101110001110" 'R 2 b3 3 #5 6 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 280": Next Spelling

dynamicList(65) = "101001101110" 'R 2 4 #4 #5 6 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 281": Next Spelling

End Sub

Private Sub Search_Eight()

'================================
' EIGHT NOTE SCALES (Total = 495)
'================================

ReDim dynamicList(42)

dynamicList(0) = "101101101101" 'R 2 b3 4 #4 #5 6 7    C, Eb, F#, A  8 Note Diminished/Whole-Half Scale       D, F, G#, B  8 Note Dominant/Symmetrical/Half-Whole Scale
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 282" & Space(3) & _
                                             I(Spelling) & " Whole-half Diminished" & Space(3) & _
                                       IIIflat(Spelling) & " Whole-half Diminished" & Space(3) & _
                                       IVsharp(Spelling) & " Whole-half Diminished" & Space(3) & _
                                            VI(Spelling) & " Whole-half Diminished" & Space(3) & _
                                            II(Spelling) & " Half-whole Diminished" & Space(3) & _
                                            IV(Spelling) & " Half-whole Diminished" & Space(3) & _
                                        Vsharp(Spelling) & " Half-whole Diminished" & Space(3) & _
                                           VII(Spelling) & " Half-whole Diminished": Next Spelling

dynamicList(1) = "101011011101" 'R 2 3 4 5 #5 6 7 C Augmented fifth Scale, E Flamenco"
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 283" & Space(3) & _
                                             I(Spelling) & " Augmented fifth Scale  " & Space(3) & _
                                           III(Spelling) & " Flamenco": Next Spelling

dynamicList(2) = "101111010101" 'R 2 b3 3 4 5 6 7 B Spanish 8 Tone Scale"
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 284" & Space(3) & _
                                           VII(Spelling) & " Spanish 8 Tone": Next Spelling

dynamicList(3) = "101011110101" 'R 2 3 4 #4 5 6 7 G Bebop"
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 285" & Space(3) & _
                                             V(Spelling) & " Bebop": Next Spelling
 
dynamicList(4) = "101010111101" 'R 2 3 #4 5 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 286": Next Spelling
 
dynamicList(5) = "101111001101" 'R 2 b3 3 4 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 287": Next Spelling

dynamicList(6) = "111011001101" 'R #1 2 3 4 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 288": Next Spelling

dynamicList(7) = "101011001111" 'R 2 3 4 #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 289": Next Spelling

dynamicList(8) = "101011101101" 'R 2 3 4 #4 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 290": Next Spelling

dynamicList(9) = "101010111110" 'R 2 3 #4 5 #5 6 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 291": Next Spelling

dynamicList(10) = "101110101101" 'R 2 b3 3 #4 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 292": Next Spelling

dynamicList(11) = "101011111001" 'R 2 3 4 #4 5 #5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 293": Next Spelling

dynamicList(12) = "111011011001" 'R #1 2 3 4 5 #5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 294": Next Spelling

dynamicList(13) = "110111001101" 'R #1 b3 3 4 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 295": Next Spelling

dynamicList(14) = "100111001111" 'R b3 3 4 #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 296": Next Spelling

dynamicList(15) = "100111011101" 'R b3 3 4 5 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 297": Next Spelling

dynamicList(16) = "101111001110" 'R 2 b3 3 4 #5 6 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 298": Next Spelling

dynamicList(17) = "101011101110" 'R 2 3 4 #4 #5 6 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 299": Next Spelling

dynamicList(18) = "101001101111" 'R 2 4 #4 #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 300": Next Spelling

dynamicList(19) = "101001111101" 'R 2 4 #4 5 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 301": Next Spelling

dynamicList(20) = "101110111100" 'R 2 b3 3 #4 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 302": Next Spelling

dynamicList(21) = "101011111100" 'R 2 3 4 #4 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 303": Next Spelling

dynamicList(22) = "101111100101" 'R 2 b3 3 4 #4 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 304": Next Spelling

dynamicList(23) = "101011100111" 'R 2 3 4 #4 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 305": Next Spelling

dynamicList(24) = "111011100101" 'R #1 2 3 4 #4 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 306": Next Spelling

dynamicList(25) = "100111011110" 'R b3 3 4 5 #5 6 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 307": Next Spelling

dynamicList(26) = "100111101110" 'R b3 3 4 #4 #5 6 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 308": Next Spelling

dynamicList(27) = "110111001110" 'R #1 b3 3 4 #5 6 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 309": Next Spelling

dynamicList(28) = "101110101110" 'R 2 b3 3 #4 #5 6 b7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 310": Next Spelling

dynamicList(29) = "111111000101" 'R #1 2 b3 3 4 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 311": Next Spelling

dynamicList(30) = "111111010001" 'R #1 2 b3 3 4 5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 312": Next Spelling

dynamicList(31) = "101111000111" 'R 2 b3 3 4 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 313": Next Spelling

dynamicList(32) = "101111110001" 'R 2 b3 3 4 #4 5 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 314": Next Spelling

dynamicList(33) = "001111110101" '2 b3 3 4 #4 5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 315": Next Spelling

dynamicList(34) = "100011111101" 'R 3 4 #4 5 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 316": Next Spelling

dynamicList(35) = "001011111101" '2 3 4 #4 5 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 317": Next Spelling

dynamicList(36) = "100001111111" 'R 4 #4 5 #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 318": Next Spelling

dynamicList(37) = "100111111100" 'R b3 3 4 #4 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 319": Next Spelling

dynamicList(38) = "100011101111" 'R 3 4 #4 #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 320": Next Spelling

dynamicList(39) = "101111101100" 'R 2 b3 3 4 #4 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 321": Next Spelling

dynamicList(40) = "101110001111" 'R 2 b3 3 #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 322": Next Spelling

dynamicList(41) = "111111001100" 'R #1 2 b3 3 4 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 323": Next Spelling

dynamicList(42) = "001111001111" '2 b3 3 4 #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 324": Next Spelling

End Sub

Private Sub Search_Nine()

'===============================
' NINE NOTE SCALES (Total = 220)
'===============================

ReDim dynamicList(18)


dynamicList(0) = "111011011101" 'R #1 2 3 4 5 #5 6 7   F 9 TONE SCALE
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 325   " & Space(3) & _
                                             IV(Spelling) & " 9-Tone Scale": Next Spelling

dynamicList(1) = "111111010101" 'R #1 2 b3 3 4 5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 326": Next Spelling
 
dynamicList(2) = "101111011101" 'R 2 b3 3 4 5 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 327": Next Spelling
 
dynamicList(3) = "101011111101" 'R 2 3 4 #4 5 #5 6 7   "BLUES SCALE STEW"
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 328" & Space(3) & _
                                             II(Spelling) & " blues scale stew": Next Spelling
 
dynamicList(4) = "101111110101" 'R 2 b3 3 4 #4 5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 329": Next Spelling
 
dynamicList(5) = "111011110101" 'R #1 2 3 4 #4 5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 330": Next Spelling
 
dynamicList(6) = "111111001101" 'R #1 2 b3 3 4 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 331": Next Spelling
 
dynamicList(7) = "100111111101" 'R b3 3 4 #4 5 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 332": Next Spelling
 
dynamicList(8) = "101111001111" 'R 2 b3 3 4 #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 333": Next Spelling

dynamicList(9) = "101011101111" 'R 2 3 4 #4 #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 334": Next Spelling

dynamicList(10) = "101111101101" 'R 2 b3 3 4 #4 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 335": Next Spelling

dynamicList(11) = "101110101111" 'R 2 b3 3 #4 #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 336": Next Spelling

dynamicList(12) = "101111100111" 'R 2 b3 3 4 #4 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 337": Next Spelling

dynamicList(13) = "100111011111" 'R b3 3 4 5 #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 338": Next Spelling

dynamicList(14) = "110111011101" 'R #1 b3 3 4 5 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 339": Next Spelling

dynamicList(15) = "111011100111" 'R #1 2 3 4 #4 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 340": Next Spelling

dynamicList(16) = "101111111100" 'R 2 b3 3 4 #4 5 #5 6
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 341": Next Spelling

dynamicList(17) = "101001111111" 'R 2 4 #4 5 #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 342": Next Spelling

dynamicList(18) = "100011111111" 'R 3 4 #4 5 #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 343": Next Spelling
                
End Sub

Private Sub Search_Ten()

'=============================
' TEN NOTE SCALES (Total = 66)
'=============================

ReDim dynamicList(5)

dynamicList(0) = "111111011101" 'R #1 2 b3 3 4 5 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 344": Next Spelling

dynamicList(1) = "101011111111" 'R 2 3 4 #4 5 #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 345": Next Spelling

dynamicList(2) = "101111111101" 'R 2 b3 3 4 #4 5 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 346": Next Spelling

dynamicList(3) = "111011111101" 'R #1 2 3 4 #4 5 #5 6 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 347": Next Spelling

dynamicList(4) = "111111001111" 'R #1 2 b3 3 4 #5 6 b7 7 Bb Chromatic Minor Pentatonic"
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 348" & Space(3) & _
                                           VIIflat(Spelling) & " Chromatic Minor Pentatonic": Next Spelling

dynamicList(5) = "101111101111" 'R 2 b3 3 4 #4 #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 349": Next Spelling

End Sub

Private Sub Search_Eleven()

'==============================
' ELEVEN NOTE dots (Total = 12)
'==============================
ReDim dynamicList(0)

dynamicList(0) = "101111111111" 'R 2 b3 3 4 #4 5 #5 6 b7 7
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 350": Next Spelling

End Sub

Private Sub Search_Twelve()

'===============
'CHROMATIC SCALE
'===============
ReDim dynamicList(0)

dynamicList(0) = "111111111111" 'Chromatic"
' oops! (12 items) to avoid index mismatch ...
For Spelling = 0 To 11: lstAppendix.AddItem "FORMULA 351" & Space(3) & _
                                             I(Spelling) & " Chromatic Scale" & Space(158) & "(names based on: I, II, III, IV, V, VI, VII and #I, bIII, #IV, #V, bVII": Next Spelling

End Sub

Private Sub lstPitch_Click()
   
   Call DrawObjects
   Call RemoveUnison
   
End Sub

Private Sub lstPitch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
      
   Dim n As Integer
   
   For n = 0 To lstPitch.ListCount - 1
      If lstPitch.Selected(n) = True Then
      cmdDigit_Click (n)
      End If
   Next n
   '====================
   lstRoot.ListIndex = -1 ' RESET INDEX or remove listbox highlight
   Call CompareIndex_Click
   
   '====================
End Sub


Private Sub cmdDigit_Click(Index As Integer)
   
   If lstRoot.ListIndex <> -1 Then
      cmdEnter.Enabled = True
      cmdEnter.Caption = "="
   ElseIf lstRoot.ListIndex = -1 Then
      cmdEnter.Enabled = False
      cmdEnter.Caption = ""
   End If
   
   Digits(Index) = 1
   txtInput.Text = ""
      
   For ctr2 = LBound(Digits) To UBound(Digits)
      txtInput.Text = txtInput.Text & Digits(ctr2)
   Next ctr2
   
   Call ConvertBinaryToFormula
   
End Sub

Private Sub CountNoteInput()
   
   Dim textCharacter As Integer
   
   CountNotes = 0
   
   For textCharacter = 1 To Len(txtInput.Text)
      If Mid(txtInput.Text, textCharacter, 1) = 1 Then
         CountNotes = CountNotes + 1
      End If
   Next textCharacter
   
End Sub

Private Sub cmdClear_Click()
   
   For ctr2 = LBound(Digits) To UBound(Digits) ' CLEAR ARRAY contents
         Digits(ctr2) = 0
   Next ctr2
   
   For ctr2 = 0 To lstPitch.ListCount - 1
      lstPitch.Selected(ctr2) = False
   Next ctr2
   
   txtFormula.Text = ""
   txtInput.Text = ""
   txtPosition.Text = ""
   List1.Clear
   lstRoot.ListIndex = -1
   cmdEnter.Enabled = False

End Sub

Private Sub cmdClear2_Click()
   Call cmdClear_Click
End Sub

Private Sub CompareIndex_Click()

Dim ctr1 As Integer
   Dim n As Integer
   
   Call CountNoteInput
   optNoteCount.Item(CountNotes - 1).Value = True
   
   ' =====EXTRACT ALL "EXATONIC" INVERSIONS AND FIND INDEX LOCATION==================
   List1.Clear
   For n = LBound(dynamicList) To UBound(dynamicList)
      For ctr1 = Len(dynamicList(n)) To 1 Step -1
        List1.AddItem Mid(dynamicList(n), ctr1 + 1, Len(dynamicList(n))) & Mid(dynamicList(n), 1, ctr1)  ' >>>>>>>>>>>>>>>>>>> lists inversions in an ASCENDING HALF STEP sequence ...
      Next ctr1

   Next n
   
   For n = 0 To List1.ListCount - 1
         
      If txtInput.Text = List1.List(n) Then
         lstAppendix_Click
         lstAppendix.ListIndex = n
         lstRoot.ListIndex = n Mod 12
         'Exit Sub
      End If
   Next n
   
   'cmdEnter.Enabled = False
End Sub


' (slight variation of the above procedure)
Private Sub cmdEnter_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   Dim ctr1 As Integer
   Dim n As Integer
   Dim KeySignature As Integer
   
   Call CountNoteInput
   optNoteCount.Item(CountNotes - 1).Value = True
   
   ' =====REEXTRACT ALL EXATONIC INVERSIONS AND FIND INDEX LOCATION==================
   List1.Clear
   For n = LBound(dynamicList) To UBound(dynamicList)
      For ctr1 = Len(dynamicList(n)) To 1 Step -1
        List1.AddItem Mid(dynamicList(n), ctr1 + 1, Len(dynamicList(n))) & Mid(dynamicList(n), 1, ctr1)  ' >>>>>>>>>>>>>>>>>>> lists inversions in an ASCENDING HALF STEP sequence ...
      Next ctr1

   Next n
   
   If lstRoot.ListIndex = -1 Then
      lstRoot.ListIndex = 0
   End If
   
   For n = 0 To List1.ListCount - 1
      
         
      If txtInput.Text = List1.List(n) Then
         lstAppendix_Click
         
         KeySignature = n Mod 12
         
         Select Case KeySignature
            Case Is = lstRoot.ListIndex
               lstAppendix.ListIndex = n
            Case Is > lstRoot.ListIndex
               lstAppendix.ListIndex = n - (KeySignature - lstRoot.ListIndex)
            Case Is < lstRoot.ListIndex
               lstAppendix.ListIndex = n + (lstRoot.ListIndex - KeySignature)
         End Select
            
      End If
   
   Next n
   
   cmdEnter.Enabled = False
   cmdEnter.Caption = ""
   
End Sub

Private Sub lstRoot_Click()
      
      If txtFormula.Text <> "" Then
         cmdEnter.Enabled = True
         cmdEnter.Caption = "="
      ElseIf txtFormula.Text = "" Then
         cmdEnter.Enabled = False
         cmdEnter.Caption = ""
      End If
      
      Call ChangeNames
      Call Rename

End Sub

Private Sub ConvertBinaryToFormula()
   
   Dim characterposition  As Integer
   
   txtFormula.Text = ""
   
      For characterposition = 1 To Len(txtInput.Text)
      
      If Mid(txtInput.Text, characterposition, 1) = 1 Then
         Select Case characterposition
         Case 1
         txtFormula.Text = txtFormula.Text & cmdDigit.Item(0).Caption & Space(1)
         Case 2
         txtFormula.Text = txtFormula.Text & cmdDigit.Item(1).Caption & Space(1)
         Case 3
         txtFormula.Text = txtFormula.Text & cmdDigit.Item(2).Caption & Space(1)
         Case 4
         txtFormula.Text = txtFormula.Text & cmdDigit.Item(3).Caption & Space(1)
         Case 5
         txtFormula.Text = txtFormula.Text & cmdDigit.Item(4).Caption & Space(1)
         Case 6
         txtFormula.Text = txtFormula.Text & cmdDigit.Item(5).Caption & Space(1)
         Case 7
         txtFormula.Text = txtFormula.Text & cmdDigit.Item(6).Caption & Space(1)
         Case 8
         txtFormula.Text = txtFormula.Text & cmdDigit.Item(7).Caption & Space(1)
         Case 9
         txtFormula.Text = txtFormula.Text & cmdDigit.Item(8).Caption & Space(1)
         Case 10
         txtFormula.Text = txtFormula.Text & cmdDigit.Item(9).Caption & Space(1)
         Case 11
         txtFormula.Text = txtFormula.Text & cmdDigit.Item(10).Caption & Space(1)
         Case 12
         txtFormula.Text = txtFormula.Text & cmdDigit.Item(11).Caption & Space(1)
         End Select
         
       End If
   
       Next characterposition
       
End Sub

Private Sub txtFormula_Click()
   HScroll1.SetFocus
End Sub


' The Midi's and the Drop tunings are still being analyzed
' and the code might still undergo improvements and revisions. This will be
' reflected in future uploads.

' Many thanks.................................................................................................................................................................. .............. ................................................................................................................................     :>








