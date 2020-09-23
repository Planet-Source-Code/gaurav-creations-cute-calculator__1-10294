VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Calculator"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1560
   Icon            =   "smallcal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   1560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command22 
      BackColor       =   &H0080FFFF&
      Caption         =   "ABOUT"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton Command21 
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   21
      ToolTipText     =   "Tangent of an angle "
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command20 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   20
      ToolTipText     =   "Cosine of an angle "
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command19 
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   19
      ToolTipText     =   "Sine of an angle "
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Sqrt"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   18
      ToolTipText     =   "Squareroot of a number"
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Command17 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   17
      ToolTipText     =   "Display result"
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton Command16 
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   840
      TabIndex        =   16
      ToolTipText     =   "Clear Display"
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   15
      ToolTipText     =   "Division"
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton Command14 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   14
      ToolTipText     =   "Multiplication"
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton Command13 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "Subtraction"
      Top             =   1320
      Width           =   270
   End
   Begin VB.CommandButton Command12 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   12
      ToolTipText     =   "Addition"
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton Command11 
      Caption         =   "'"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   11
      ToolTipText     =   "Decimal"
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton Command10 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   10
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton Command9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   8
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   7
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   6
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1200
      TabIndex        =   4
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   3
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   2
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   600
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   0
      Text            =   "0."
      ToolTipText     =   "Display Output & Accept Input"
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This Calculator is made by Gaurav Dhup
'www.question.chatbook.com
Dim i As Double
Dim a(100) As String
Dim add As Double
Dim d As Integer
Dim e As Double
Dim su As Double
Dim s As Integer
Dim g As Double
Dim mul As Double
Dim f As Integer
Dim m As Integer
Dim n As Double
Dim div As Double
Dim o As Integer
Dim p As Double
Dim sq As Double

Dim si As Double
Dim co As Double
Dim ta As Double

Private Sub Command1_Click()
Text1.Text = "1"
a(i) = (Text1.Text)
Beep
i = i + 1
For b = 0 To i - 1
    c = c + a(b)
    Text1.Text = c
Next b

End Sub

Private Sub Command10_Click()
Text1.Text = "0"
a(i) = (Text1.Text)
Beep
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 Text1.Text = c
 Next b
End Sub

Private Sub Command11_Click()
If f = 0 Then
Text1.Text = "."
a(i) = (Text1.Text)
Beep
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 Text1.Text = c
 Next b
f = 1
Else
MsgBox "DON't TRY TO USE TWO DECIMAL POINTS", vbCritical, "ERROR"
End If
End Sub

Private Sub Command12_Click()
f = 0
add = Text1.Text
Text1.Text = "0."
Beep
d = 1
For b = 0 To i - 1
a(b) = 0
 Next b
 i = 0
End Sub

Private Sub Command13_Click()
f = 0
su = Text1.Text
Text1.Text = "0."
Beep
s = 1
For b = 0 To i - 1
a(b) = 0
 Next b
 i = 0
End Sub

Private Sub Command14_Click()
f = 0
mul = Text1.Text
Text1.Text = "0."
Beep
m = 1
For b = 0 To i - 1
a(b) = 0
 Next b
 i = 0
End Sub

Private Sub Command15_Click()
f = 0
div = Text1.Text
Text1.Text = "0."
Beep
o = 1
For b = 0 To i - 1
a(b) = 0
 Next b
 i = 0
End Sub

Private Sub Command16_Click()
For b = 0 To i - 1
a(b) = 0
Next b
i = 0
Text1.Text = "0."
Beep
End Sub

Private Sub Command17_Click()
Beep
If d = 1 Then
e = Text1.Text + add
Text1.Text = e
d = 0
ElseIf s = 1 Then
g = su - Text1.Text
Text1.Text = g
s = 0
ElseIf m = 1 Then
n = mul * Text1.Text
Text1.Text = n
m = 0
ElseIf o = 1 Then
If (Text1.Text > 0) Then
p = div / Text1.Text
Text1.Text = p
o = 0
ElseIf (Text1.Text) = 0 Then
MsgBox ("INVALID CAN'T DIVIDE BY 0"), vbCritical, "ERROR"
For b = 0 To i - 1
a(b) = 0
Next b
Text1.Text = "0."
End If
End If

For b = 0 To i - 1
a(b) = 0
Next b
End Sub

Private Sub Command18_Click()
If (Text1.Text) > 0 Then
sq = Sqr(Text1.Text)
Text1.Text = sq
Beep
ElseIf Text1.Text = 0 Then
MsgBox ("INVALID: NOT DEFINED"), vbCritical, "ERROR"
For b = 0 To i - 1
a(b) = 0
Next b
Text1.Text = "0."
End If
End Sub

Private Sub Command19_Click()
si = Sin(Text1.Text)
Text1.Text = si
Beep
End Sub

Private Sub Command2_Click()
Text1.Text = "2"
a(i) = (Text1.Text)
Beep
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 Text1.Text = c
 Next b
End Sub

Private Sub Command20_Click()
co = Cos(Text1.Text)
Text1.Text = co
Beep
End Sub

Private Sub Command21_Click()
ta = Tan(Text1.Text)
Text1.Text = ta
Beep
End Sub

Private Sub Command22_Click()
Form2.Show
End Sub

Private Sub Command3_Click()
Text1.Text = "3"

a(i) = (Text1.Text)
Beep
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 Text1.Text = c
 Next b
End Sub

Private Sub Command4_Click()
Text1.Text = "4"

a(i) = (Text1.Text)
Beep
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 Text1.Text = c
 Next b
End Sub

Private Sub Command5_Click()
Text1.Text = "5"

a(i) = (Text1.Text)
Beep
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 Text1.Text = c
 Next b
End Sub

Private Sub Command6_Click()
Text1.Text = "6"

a(i) = (Text1.Text)
Beep
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 Text1.Text = c
 Next b
End Sub

Private Sub Command7_Click()
Text1.Text = "7"
a(i) = (Text1.Text)
Beep
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 Text1.Text = c
 Next b
End Sub

Private Sub Command8_Click()
Text1.Text = "8"

a(i) = (Text1.Text)
Beep
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 Text1.Text = c
 Next b
End Sub

Private Sub Command9_Click()
Text1.Text = "9"

a(i) = (Text1.Text)
Beep
i = i + 1
For b = 0 To i - 1
 c = c + a(b)
 Text1.Text = c
 Next b
End Sub


Private Sub Text1_Change()

End Sub
