VERSION 5.00
Begin VB.Form frmBestFit 
   Caption         =   "Form1"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4860
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   4860
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   3720
      TabIndex        =   0
      Top             =   2700
      Width           =   45
   End
End
Attribute VB_Name = "frmBestFit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Line of Best Fit Demo
'By Bill Soo
'October 16, 2002
'Calculates the line of best fit for points that you place on the form
'by clicking
'The line is calculated using the method of sum of least squares
'double click to erase the line and start again.

'NOTE: This demo may crash if the line is very steep since this can cause
'a division by zero or overflow error.

Option Explicit
Dim Sx As Double
Dim Sxy As Double
Dim Sy As Double
Dim Sx2 As Double
Dim n As Long
Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double
Dim bDblClick As Boolean
Private Const XMAX = 100
Private Const YMAX = 100
Private Const XMIN = -100
Private Const YMIN = -100

Private Sub Form_DblClick()
Me.Cls
Sx = 0: Sxy = 0: Sy = 0: Sx2 = 0: n = 0
Me.Caption = ""
bDblClick = True
End Sub

Private Sub Form_Load()
Me.Scale (XMIN, YMAX)-(XMAX, YMIN)
Me.AutoRedraw = True
Me.BackColor = vbWhite
Me.DrawMode = vbNotXorPen   'this allows us to erase lines by drawing over them
Me.MousePointer = 2
Label1.AutoSize = True
Label1.BackStyle = vbTransparent
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'this shows a label on the form with the current mouse coordinates
'the label moves to avoid getting in the way of the mouse
Label1 = x & ":" & y
If y > 0 Then
    Label1.Move Me.ScaleLeft, Me.ScaleTop + Me.ScaleHeight + Label1.Height
Else
    Label1.Move Me.ScaleLeft, Me.ScaleTop
End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'unless we did a double click (clearing the data), this event places new points on
'the graph. It then calculates the new line of best fit and plots it.
Dim m As Double
Dim b As Double
If bDblClick Then
    bDblClick = False
Else
    Me.PSet (x, y)
    n = n + 1                    'n is the number of points
    Sx = Sx + x                  'Sx is the SUM of all X coordinates
    Sy = Sy + y                  'Sy is the SUM of all Y coordinates
    Sxy = Sxy + x * y            'Sxy is the SUM of all x*y
    Sx2 = Sx2 + x * x            'Sx2 is the SUM of all x*x
    If n > 1 Then
        If n > 2 Then Me.Line (x1, y1)-(x2, y2) 'erase old line
        'calculate new line using least squares
        'm and b define the line using y=m*x+b
        m = (n * Sxy - Sx * Sy) / (n * Sx2 - Sx * Sx)
        b = (Sy - m * Sx) / n
        Me.Caption = "Y = " & m & " * X + " & b
        CalcLine m, b, x1, y1, x2, y2
        Me.Line (x1, y1)-(x2, y2)
    End If
End If
End Sub

Private Sub CalcLine(ByVal m#, ByVal b#, x1#, y1#, x2#, y2#)
'given a line defined by m and b, calculates two points on that line
'that intersect the bounding box of (100,100)-(-100,-100)
Dim x#, y#

y = m * XMIN + b
If (y >= YMIN) And (y <= YMAX) Then
    x1 = XMIN
    y1 = y
Else
    If m > 0 Then
        y1 = YMIN
    Else
        y1 = YMAX
    End If
    x1 = (y1 - b) / m
End If
y = m * XMAX + b
If (y >= YMIN) And (y <= YMAX) Then
    x2 = XMAX
    y2 = y
Else
    If m > 0 Then
        y2 = YMAX
    Else
        y2 = YMIN
    End If
    x2 = (y2 - b) / m
End If
End Sub
