VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Print Graph Paper"
   ClientHeight    =   4545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7185
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   303
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   479
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PIC 
      BackColor       =   &H80000013&
      Height          =   3900
      Left            =   180
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   196
      TabIndex        =   0
      Top             =   75
      Width           =   3000
      Begin VB.CheckBox chkPrint 
         BackColor       =   &H00FF8080&
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   345
         Left            =   975
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3390
         Width           =   1050
      End
      Begin VB.CheckBox chkExit 
         BackColor       =   &H008080FF&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2610
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   " Close "
         Top             =   30
         Width           =   285
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H80000013&
         Caption         =   "Heavy line spacing"
         Height          =   735
         Left            =   1140
         TabIndex        =   13
         Top             =   1875
         Width           =   1650
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "Form1.frx":0442
            Left            =   405
            List            =   "Form1.frx":045B
            TabIndex        =   14
            Text            =   "0"
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H80000013&
         Caption         =   "Color"
         Height          =   720
         Left            =   105
         TabIndex        =   11
         Top             =   1875
         Width           =   960
         Begin VB.Label LabColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   180
            TabIndex        =   12
            Top             =   225
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H80000013&
         Caption         =   "Lines"
         Height          =   930
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   2655
         Begin VB.Frame Frame6 
            BackColor       =   &H80000013&
            Height          =   420
            Left            =   270
            TabIndex        =   15
            Top             =   420
            Width           =   2130
            Begin VB.OptionButton optSolidDotted 
               BackColor       =   &H80000013&
               Caption         =   "Dotted"
               Height          =   195
               Index           =   1
               Left            =   900
               TabIndex        =   17
               Top             =   165
               Width           =   1005
            End
            Begin VB.OptionButton optSolidDotted 
               BackColor       =   &H80000013&
               Caption         =   "Solid"
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   16
               Top             =   165
               Width           =   675
            End
         End
         Begin VB.OptionButton optLines 
            BackColor       =   &H80000013&
            Caption         =   "Both"
            Height          =   195
            Index           =   2
            Left            =   1560
            TabIndex        =   10
            Top             =   210
            Width           =   840
         End
         Begin VB.OptionButton optLines 
            BackColor       =   &H80000013&
            Caption         =   "Vert"
            Height          =   195
            Index           =   1
            Left            =   840
            TabIndex        =   9
            Top             =   210
            Width           =   840
         End
         Begin VB.OptionButton optLines 
            BackColor       =   &H80000013&
            Caption         =   "Horz"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   210
            Width           =   840
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000013&
         Caption         =   "Scale"
         Height          =   465
         Left            =   105
         TabIndex        =   4
         Top             =   450
         Width           =   2655
         Begin VB.OptionButton optScale 
            BackColor       =   &H80000013&
            Caption         =   "cm"
            Height          =   210
            Index           =   1
            Left            =   1185
            TabIndex        =   6
            Top             =   195
            Width           =   870
         End
         Begin VB.OptionButton optScale 
            BackColor       =   &H80000013&
            Caption         =   "in"
            Height          =   210
            Index           =   0
            Left            =   195
            TabIndex        =   5
            Top             =   195
            Width           =   870
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H80000013&
         Caption         =   "Interval"
         Height          =   660
         Left            =   105
         TabIndex        =   1
         Top             =   2625
         Width           =   2685
         Begin VB.HScrollBar HSCR 
            Height          =   255
            Left            =   75
            Max             =   16
            Min             =   1
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   255
            Value           =   1
            Width           =   1755
         End
         Begin VB.Label LabIncr 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LanIncr"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1905
            TabIndex        =   3
            Top             =   255
            Width           =   690
         End
      End
      Begin VB.Label LabMove 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "   Graph Settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   285
         Left            =   15
         TabIndex        =   18
         Top             =   15
         Width           =   2910
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Print Graph Paper by Robert Rayment

' Modified & Extended from code by:  ------------------------
' Author:  Clint M. LaFever - [lafeverc@saic.com]
' Date: September,30 2002 @ 15:42:51
'------------------------------------------------------------

Private ICScale As Long    ' 0 in, 1 cm
Private HVBLines As Long   ' 0 Horz, 1 Vert, 2 Both
Private HeavyLineSpacing As Long
Private zLineInterval As Single
Private zMul As Single     ' 0.0625=1/16 in, 0.1= 1/10 cm

Private APErr As Boolean
Private xPIC As Single, yPIC As Single
Private xLab As Single, yLab As Single
Private aLabDown As Boolean
Private MoverSC As Long

Private STX As Long, STY As Long


Private Sub Form_Initialize()
   STX = Screen.TwipsPerPixelX
   STY = Screen.TwipsPerPixelY
   HVBLines = 2            ' Both
   HeavyLineSpacing = 0    ' No thick lines
   zMul = 0.0625           ' 1/16
   MoverSC = 1440          ' Twips/inch
End Sub
Private Sub Form_Load()
   Me.BackColor = vbWhite
   Me.DrawStyle = 0
   optScale(0).Value = True
   optSolidDotted(0).Value = True
   optLines(2).Value = True
   LabColor.BackColor = 0
   Me.ForeColor = LabColor.BackColor

   aLabDown = False
   xPIC = 0: yPIC = 0
   PIC.Move xPIC, yPIC
End Sub


' #### Graph Settings ############################################
Private Sub optScale_Click(Index As Integer)
   ICScale = Index   ' 0 in, 1 cm
   If Index = 0 Then
      zMul = 0.0625
      HSCR.Min = 1
      HSCR.Max = 16
      MoverSC = 1440 ' Twips/inch
   Else
      zMul = 0.1
      HSCR.Min = 1
      HSCR.Max = 10
      MoverSC = 567  ' Twips/cm
   End If
   DrawGrid
End Sub
Private Sub optLines_Click(Index As Integer)
  HVBLines = Index  ' 0 Horz, 1 Vert, 2 Both
  DrawGrid
End Sub
Private Sub optSolidDotted_Click(Index As Integer)
  If Index = 0 Then
     Me.DrawStyle = 0
  Else
     Me.DrawStyle = 2
  End If
  DrawGrid
End Sub
Private Sub LabColor_Click()
Dim CD As CDialog
Dim TheColor As Long
   Set CD = New CDialog
   If CD.VBChooseColor(TheColor, , , , Me.hWnd) Then
      Me.LabColor.BackColor = TheColor
   End If
   Set CD = Nothing
   DrawGrid
End Sub
Private Sub Combo1_Click()
   HeavyLineSpacing = Val(Combo1.Text)
   PIC.SetFocus
   DrawGrid
End Sub
Private Sub HSCR_Change()
   DrawGrid
End Sub
Private Sub HSCR_Scroll()
   DrawGrid
End Sub
' ################################################################


' #### DRAW GRID #################################################
Private Sub DrawGrid()
Dim x As Double, y As Double, n As Long
On Error GoTo ErrorDrawGrid
   y = 0
   x = 0
   'ICScale = Index   ' 0 in, 1 cm
   If ICScale = 0 Then
      Me.ScaleMode = vbInches
   Else
      Me.ScaleMode = vbCentimeters
   End If
   
   Me.Cls
   
   If HVBLines = 1 Or HVBLines = 2 Then
      n = 0    ' Vertical lines
      While x < Me.ScaleWidth
         If HeavyLineSpacing > 0 Then
            If n Mod HeavyLineSpacing = 0 Then
               Me.DrawWidth = 2
            Else
               Me.DrawWidth = 1
            End If
         Else
            Me.DrawWidth = 1
         End If
         ' zMul:  0.0625 = 1/16 in, 0.0 1/10 cm
         Me.Line (x, 0)-(x, Me.ScaleHeight), Me.LabColor.BackColor
         x = x + (zMul * Me.HSCR.Value)
         n = n + 1
      Wend
   End If
   
   If HVBLines = 0 Or HVBLines = 2 Then
      n = 0    ' Horizontal lines
      While y < Me.ScaleHeight
         If HeavyLineSpacing > 0 Then
            If n Mod HeavyLineSpacing = 0 Then
               Me.DrawWidth = 2
            Else
               Me.DrawWidth = 1
            End If
         Else
            Me.DrawWidth = 1
         End If
         ' zMul:  0.0625 = 1/16 in, 0.0 1/10 cm
         Me.Line (0, y)-(Me.ScaleWidth, y), Me.LabColor.BackColor
         y = y + (zMul * Me.HSCR.Value)
         n = n + 1
      Wend
   End If
   If ICScale = 0 Then
      'LabIncr = Str$(zMul * Me.HSCR.Value) & " in"
      Select Case Me.HSCR.Value
      Case 1, 3, 5, 7, 9, 11, 13, 15
         LabIncr = Str$(Me.HSCR.Value) & "/16 in"
      Case 2
         LabIncr = "1/8 in"
      Case 4
         LabIncr = "1/4 in"
      Case 6
         LabIncr = "3/8 in"
      Case 8
         LabIncr = "1/2 in"
      Case 10
         LabIncr = "5/8 in"
      Case 12
         LabIncr = "3/4 in"
      Case 14
         LabIncr = "7/8 in"
      Case 16
         LabIncr = "1 in"
      Case Else
         LabIncr = Str$(Me.HSCR.Value) & "/16 in"
      End Select
      'If Me.HSCR.Value < 16 Then
      '   LabIncr = Str$(Me.HSCR.Value) & "/16 in"
      'Else
      '   LabIncr = "1 in"
      'End If
   Else
      'LabIncr = Str$(zMul * Me.HSCR.Value) & " cm"
      If Me.HSCR.Value < 10 Then
         LabIncr = Str$(Me.HSCR.Value) & " mm"
      Else
         LabIncr = "1 cm"
      End If
      
   End If
   Exit Sub
'=======
ErrorDrawGrid:
   MsgBox Err & ":Error in call to DrawGrid()." _
   & vbCrLf & vbCrLf & "Error Description: " & Err.Description, vbCritical, "Warning"
   Exit Sub
End Sub



' #### PRINT #####################################################
Private Sub chkPrint_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   chkPrint.Value = Unchecked
   PRINTGRAPH
End Sub
Private Sub PRINTGRAPH()
Dim n As Long
Dim x As Double, y As Double
Dim res As Long
    
   res = MsgBox("IS PRINTER LIVE ?", vbQuestion + vbYesNo, "PRINTER")
   If res = vbNo Then
      Exit Sub
   End If
   
   ShowPrinter Me, APErr
   If Not APErr Then
   
      DrawGrid
      Printer.DrawStyle = Me.DrawStyle
      Printer.ScaleMode = Me.ScaleMode
      
      If HVBLines = 1 Or HVBLines = 2 Then
         n = 0    ' Vertical lines
         While x < Printer.ScaleWidth
            If HeavyLineSpacing > 0 Then
               If n Mod HeavyLineSpacing = 0 Then
                Printer.DrawWidth = 8
               Else
                Printer.DrawWidth = 1
               End If
            Else
               Printer.DrawWidth = 1
            End If
               
            Printer.Line (x, 0)-(x, Printer.ScaleHeight), Me.LabColor.BackColor
            x = x + (zMul * Me.HSCR.Value)
            n = n + 1
         Wend
      End If
      
      If HVBLines = 0 Or HVBLines = 2 Then
         n = 0    ' Horizontal lines
         While y < Printer.ScaleHeight
         
            If HeavyLineSpacing > 0 Then
               If n Mod HeavyLineSpacing = 0 Then
                  Printer.DrawWidth = 8
               Else
                  Printer.DrawWidth = 1
               End If
            Else
               Printer.DrawWidth = 1
            End If
            
            Printer.Line (0, y)-(Printer.ScaleWidth, y), Me.LabColor.BackColor
            y = y + (zMul * Me.HSCR.Value)
            n = n + 1
         Wend
      End If
      Printer.DrawStyle = 0
      Printer.EndDoc
   Else
      MsgBox "NO PRINT", vbCritical, ""
   End If
End Sub




Private Sub LabMove_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   aLabDown = True
   xLab = x
   yLab = y
End Sub
Private Sub LabMove_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim XSC As Single, YSC As Single
   XSC = MoverSC / STX
   YSC = MoverSC / STY
   If aLabDown Then
      xPIC = xPIC - (xLab - x) / XSC
      If xPIC < 0 Then xPIC = 0
      'Label1 = Str$(xPIC + PIC.Width * XSC) & Str$(Me.Width / STX)
      If xPIC + PIC.Width * XSC > Me.Width / STX Then
         xPIC = (Me.Width / STX - PIC.Width * XSC)
      End If
      yPIC = yPIC - (yLab - y) / YSC
      If yPIC < 0 Then yPIC = 0
      If yPIC + PIC.Height * YSC > Me.Height / STY Then
         yPIC = (Me.Height / STY - PIC.Height * YSC)
      End If
      PIC.Left = xPIC / XSC
      PIC.Top = yPIC / YSC
   End If
End Sub
Private Sub LabMove_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   aLabDown = False
End Sub


Private Sub Form_Resize()
   xPIC = 0: yPIC = 0
   PIC.Move xPIC, yPIC
   DrawGrid
End Sub

Private Sub chkExit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   chkExit.Value = Unchecked
   Form_Unload 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
   Unload Me
   End
End Sub

