VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Calculator"
   ClientHeight    =   2085
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Height          =   720
      Left            =   3075
      TabIndex        =   1
      Top             =   30
      Width           =   1740
   End
   Begin VB.Frame Frame3 
      Caption         =   "Basic"
      Height          =   1365
      Left            =   3075
      TabIndex        =   17
      Top             =   720
      Width           =   1740
      Begin VB.CommandButton Command4 
         Caption         =   "8"
         Height          =   270
         Index           =   16
         Left            =   660
         TabIndex        =   34
         Top             =   465
         Width           =   285
      End
      Begin VB.CommandButton Command4 
         Caption         =   "9"
         Height          =   270
         Index           =   15
         Left            =   975
         TabIndex        =   33
         Top             =   465
         Width           =   285
      End
      Begin VB.CommandButton Command4 
         Caption         =   "0"
         Height          =   270
         Index           =   14
         Left            =   1290
         TabIndex        =   32
         Top             =   465
         Width           =   285
      End
      Begin VB.CommandButton Command4 
         Caption         =   "1"
         Height          =   270
         Index           =   13
         Left            =   30
         TabIndex        =   31
         Top             =   165
         Width           =   285
      End
      Begin VB.CommandButton Command4 
         Caption         =   "2"
         Height          =   270
         Index           =   12
         Left            =   345
         TabIndex        =   30
         Top             =   165
         Width           =   285
      End
      Begin VB.CommandButton Command4 
         Caption         =   "3"
         Height          =   270
         Index           =   11
         Left            =   660
         TabIndex        =   29
         Top             =   165
         Width           =   285
      End
      Begin VB.CommandButton Command4 
         Caption         =   "4"
         Height          =   270
         Index           =   10
         Left            =   975
         TabIndex        =   28
         Top             =   165
         Width           =   285
      End
      Begin VB.CommandButton Command4 
         Caption         =   "5"
         Height          =   270
         Index           =   9
         Left            =   1290
         TabIndex        =   27
         Top             =   165
         Width           =   285
      End
      Begin VB.CommandButton Command4 
         Caption         =   "6"
         Height          =   270
         Index           =   8
         Left            =   30
         TabIndex        =   26
         Top             =   465
         Width           =   285
      End
      Begin VB.CommandButton Command4 
         Caption         =   "7"
         Height          =   270
         Index           =   7
         Left            =   345
         TabIndex        =   25
         Top             =   465
         Width           =   285
      End
      Begin VB.CommandButton Command4 
         Caption         =   "^"
         Height          =   270
         Index           =   6
         Left            =   660
         TabIndex        =   24
         Top             =   1050
         Width           =   285
      End
      Begin VB.CommandButton Command4 
         Caption         =   ")"
         Height          =   270
         Index           =   5
         Left            =   345
         TabIndex        =   23
         Top             =   1050
         Width           =   285
      End
      Begin VB.CommandButton Command4 
         Caption         =   "("
         Height          =   270
         Index           =   4
         Left            =   30
         TabIndex        =   22
         Top             =   1050
         Width           =   285
      End
      Begin VB.CommandButton Command4 
         Caption         =   "/"
         Height          =   270
         Index           =   3
         Left            =   975
         TabIndex        =   21
         Top             =   750
         Width           =   285
      End
      Begin VB.CommandButton Command4 
         Caption         =   "*"
         Height          =   270
         Index           =   2
         Left            =   660
         TabIndex        =   20
         Top             =   750
         Width           =   285
      End
      Begin VB.CommandButton Command4 
         Caption         =   "-"
         Height          =   270
         Index           =   1
         Left            =   345
         TabIndex        =   19
         Top             =   750
         Width           =   285
      End
      Begin VB.CommandButton Command4 
         Caption         =   "+"
         Height          =   270
         Index           =   0
         Left            =   30
         TabIndex        =   18
         Top             =   750
         Width           =   285
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Other"
      Height          =   555
      Left            =   45
      TabIndex        =   12
      Top             =   1530
      Width           =   2970
      Begin VB.CommandButton Command3 
         Caption         =   "Ln"
         Height          =   255
         Index           =   3
         Left            =   2175
         TabIndex        =   16
         Top             =   225
         Width           =   720
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Log"
         Height          =   255
         Index           =   2
         Left            =   1470
         TabIndex        =   15
         Top             =   225
         Width           =   720
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Exp"
         Height          =   255
         Index           =   1
         Left            =   765
         TabIndex        =   14
         Top             =   225
         Width           =   720
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Sqr"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   13
         Top             =   225
         Width           =   720
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Trigonometry"
      Height          =   1110
      Left            =   60
      TabIndex        =   3
      Top             =   420
      Width           =   2940
      Begin VB.CheckBox Check2 
         Caption         =   "Hyperbolic"
         Height          =   210
         Left            =   1065
         TabIndex        =   11
         Top             =   225
         Width           =   1140
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Inverse"
         Height          =   210
         Left            =   105
         TabIndex        =   10
         Top             =   210
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cot"
         CausesValidation=   0   'False
         Height          =   285
         Index           =   7
         Left            =   1920
         TabIndex        =   9
         Top             =   750
         Width           =   930
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Sec"
         CausesValidation=   0   'False
         Height          =   285
         Index           =   6
         Left            =   1005
         TabIndex        =   8
         Top             =   750
         Width           =   930
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cosec"
         CausesValidation=   0   'False
         Height          =   285
         Index           =   5
         Left            =   90
         TabIndex        =   7
         Top             =   750
         Width           =   930
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Tan"
         CausesValidation=   0   'False
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   6
         Top             =   480
         Width           =   930
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cos"
         CausesValidation=   0   'False
         Height          =   285
         Index           =   1
         Left            =   1005
         TabIndex        =   5
         Top             =   480
         Width           =   930
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Sin"
         CausesValidation=   0   'False
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   480
         Width           =   930
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   2730
   End
   Begin VB.Label Label1 
      Caption         =   "="
      Height          =   195
      Left            =   2910
      TabIndex        =   2
      Top             =   75
      Width           =   165
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The Evaluator is a real reccursive function
'That breakes the equation in to smaller pieces
Private Sub Command1_Click()
  Command1.Caption = Fractionize(Evaluate(Text1.Text))
End Sub

Function Fractionize(Value As Double) As String
  'Returns rationals as fractions
  Dim a As Currency, n%, d%
  a = Value - Fix(Value)
  If a <> 0 Then
    For n = 1 To 1000
      For d = n + 1 To 1000
        If a = n / d Then
           If Fix(Value) <> 0 Then
              Fractionize = Fix(Value) & " + " & n & "/" & d & vbCrLf & (d * Fix(Value) + n) & "/" & d & vbCrLf & Value
           Else
              Fractionize = n & "/" & d & vbCrLf & Value
           End If
           Exit Function
        End If
      Next
    Next
  End If
  Fractionize = Value
End Function

Function Evaluate(ByVal Equation As String) As Double
   Dim i&, j%, P%, Si&, d$, BB$, AB$, BE$, SpecialFunction$
   'i& variable to loop through each letter in equation
   'j% variable to loop and check if there are letters before the bracket
   'Si is the possition of the outermost open bracket
   'd$,BB$, AB$, BE$ are an intermediate variables to keep line length smaller see a couple of lines above for their contents
   'SpecialFunction$ holds the text found before the outermost bracket
   Equation = Replace(Replace(Equation, ",", "."), "--", "")
   'Parenthesis Begins with outermost
   For i = 1 To Len(Equation)
     If Mid(Equation, i, 1) = "(" Then
        P = P + 1
        If P = 1 Then Si = i
     End If
     If Mid(Equation, i, 1) = ")" Then
        If P = 1 Then
           'We found the outermost bracket. Check if it a function (it is a function when there are letters before the open bracket "("
           For j = Si - 1 To 1 Step -1
                d = Mid(Equation, j, 1)
                If d = "+" Or d = "-" Or d = "*" Or d = "/" Or d = "(" Or d = ")" Then
                   SpecialFunction = Trim(Mid(Equation, j + 1, Si - j - 1))
                   Exit For
                ElseIf j = 1 Then
                   SpecialFunction = Trim(Mid(Equation, j, Si - j))
                End If
           Next
           BB = Mid(Equation, 1, Si - 1 - Len(SpecialFunction)) 'Before Bracket
           AB = Mid(Equation, i + 1) 'After Bracket
           BE = Mid(Equation, Si + 1, i - 1 - Si) 'Bracket Expression
           Select Case UCase(SpecialFunction)
             Case "SIN": Evaluate = Evaluate(BB & Sin(Evaluate(BE)) & AB)
             Case "COS": Evaluate = Evaluate(BB & Cos(Evaluate(BE)) & AB)
             Case "TAN": Evaluate = Evaluate(BB & Tan(Evaluate(BE)) & AB)
             Case "SEC": Evaluate = Evaluate(BB & Sec(Evaluate(BE)) & AB)
             Case "COSEC": Evaluate = Evaluate(BB & Cosec(Evaluate(BE)) & AB)
             Case "COT": Evaluate = Evaluate(BB & Cot(Evaluate(BE)) & AB)
             
             Case "ASIN": Evaluate = Evaluate(BB & Arcsin(Evaluate(BE)) & AB)
             Case "ACOS": Evaluate = Evaluate(BB & Arccos(Evaluate(BE)) & AB)
             Case "ATAN": Evaluate = Evaluate(BB & Atn(Evaluate(BE)) & AB)
             Case "ASEC": Evaluate = Evaluate(BB & Arcsec(Evaluate(BE)) & AB)
             Case "ACOSEC": Evaluate = Evaluate(BB & Arccosec(Evaluate(BE)) & AB)
             Case "ACOT": Evaluate = Evaluate(BB & Arccot(Evaluate(BE)) & AB)
             
             Case "HSIN": Evaluate = Evaluate(BB & HSin(Evaluate(BE)) & AB)
             Case "HCOS": Evaluate = Evaluate(BB & HCos(Evaluate(BE)) & AB)
             Case "HTAN": Evaluate = Evaluate(BB & HTan(Evaluate(BE)) & AB)
             Case "HSEC": Evaluate = Evaluate(BB & HSec(Evaluate(BE)) & AB)
             Case "HCOSEC": Evaluate = Evaluate(BB & HCosec(Evaluate(BE)) & AB)
             Case "HCOT": Evaluate = Evaluate(BB & HCotan(Evaluate(BE)) & AB)
             
             Case "IHSIN": Evaluate = Evaluate(BB & HArcsin(Evaluate(BE)) & AB)
             Case "IHCOS": Evaluate = Evaluate(BB & HArccos(Evaluate(BE)) & AB)
             Case "IHTAN": Evaluate = Evaluate(BB & HArctan(Evaluate(BE)) & AB)
             Case "IHSEC": Evaluate = Evaluate(BB & HArcsec(Evaluate(BE)) & AB)
             Case "IHCOSEC": Evaluate = Evaluate(BB & HArccosec(Evaluate(BE)) & AB)
             Case "IHCOT": Evaluate = Evaluate(BB & HArccotan(Evaluate(BE)) & AB)
             
             Case "SQR": Evaluate = Evaluate(BB & Sqr(Evaluate(BE)) & AB)
             Case "EXP": Evaluate = Evaluate(BB & Exp(Evaluate(BE)) & AB)
             Case "LOG": Evaluate = Evaluate(BB & Log(Evaluate(BE)) & AB)
             Case "LN": Evaluate = Evaluate(BB & Log(Evaluate(BE)) / Log(Exp(1)) & AB)
             Case "": Evaluate = Evaluate(BB & Evaluate(BE) & AB)
             Case Else: MsgBox "Invalid function name / number before bracket", vbExclamation
           End Select
           
           Exit Function
        End If
        P = P - 1
     End If
   Next
   
   'Addition / Substruction
   For i = Len(Equation) To 1 Step -1
     If Mid(Equation, i, 1) = "+" Then
        If i > 1 Then
           If Mid(Equation, i - 1, 1) <> "*" And Mid(Equation, i - 1, 1) <> "/" Then
              Evaluate = Evaluate(Mid(Equation, 1, i - 1)) + Evaluate(Mid(Equation, i + 1))
              Exit Function
           End If
        Else
           Evaluate = Evaluate(Mid(Equation, 1, i - 1)) + Evaluate(Mid(Equation, i + 1))
           Exit Function
        End If
     End If
     If Mid(Equation, i, 1) = "-" Then
        If i > 1 Then
           If Mid(Equation, i - 1, 1) <> "*" And Mid(Equation, i - 1, 1) <> "/" Then
              Evaluate = Evaluate(Mid(Equation, 1, i - 1)) - Evaluate(Mid(Equation, i + 1))
              Exit Function
           End If
        Else
           Evaluate = Evaluate(Mid(Equation, 1, i - 1)) - Evaluate(Mid(Equation, i + 1))
           Exit Function
        End If
     End If
   Next
   
   'Multiplication / Division
   For i = Len(Equation) To 1 Step -1
     If Mid(Equation, i, 1) = "*" Then
        Evaluate = Evaluate(Mid(Equation, 1, i - 1)) * Evaluate(Mid(Equation, i + 1))
        Exit Function
     End If
     If Mid(Equation, i, 1) = "/" Then
        Evaluate = Evaluate(Mid(Equation, 1, i - 1)) / Evaluate(Mid(Equation, i + 1))
        Exit Function
     End If
   Next
   
   'Powers
   For i = Len(Equation) To 1 Step -1
     If Mid(Equation, i, 1) = "^" Then
        Evaluate = Evaluate(Mid(Equation, 1, i - 1)) ^ Evaluate(Mid(Equation, i + 1))
        Exit Function
     End If
   Next
   
   Evaluate = Val(Equation)
End Function

Private Sub Command2_Click(Index As Integer)
  If Check1.Value = 1 And Check2.Value = 0 Then Text1.SelText = "A" & Command2(Index).Caption & "("
  If Check1.Value = 0 And Check2.Value = 1 Then Text1.SelText = "H" & Command2(Index).Caption & "("
  If Check1.Value = 1 And Check2.Value = 1 Then Text1.SelText = "IH" & Command2(Index).Caption & "("
  If Check1.Value = 0 And Check2.Value = 0 Then Text1.SelText = Command2(Index).Caption & "("
  Check1.Value = 0
  Check2.Value = 0
  Text1.SetFocus
End Sub

Private Sub Command3_Click(Index As Integer)
  Text1.SelText = Command3(Index).Caption & "("
  Text1.SetFocus
End Sub

Private Sub Command4_Click(Index As Integer)
  Text1.SelText = Command4(Index).Caption
  Text1.SetFocus
End Sub
