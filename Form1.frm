VERSION 5.00
Begin VB.Form Form1 
   Caption         =   " "
   ClientHeight    =   3315
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3315
   ScaleWidth      =   9735
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "SET"
      Height          =   285
      Left            =   6180
      TabIndex        =   13
      Top             =   2160
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SET"
      Height          =   285
      Left            =   6180
      TabIndex        =   12
      Top             =   1830
      Width           =   735
   End
   Begin VB.TextBox Mark2 
      Height          =   285
      Left            =   4800
      TabIndex        =   9
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox Mark1 
      Height          =   285
      Left            =   4800
      TabIndex        =   8
      Top             =   1830
      Width           =   1335
   End
   Begin VB.TextBox txtInterval 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   2580
      TabIndex        =   7
      Top             =   1770
      Width           =   765
   End
   Begin VB.ListBox List2 
      Height          =   645
      Left            =   1500
      TabIndex        =   4
      Top             =   2400
      Width           =   1020
   End
   Begin VB.TextBox txtFIN 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   1320
      TabIndex        =   3
      Top             =   1770
      Width           =   990
   End
   Begin VB.TextBox txtINI 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   285
      TabIndex        =   2
      Top             =   1770
      Width           =   990
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   240
      TabIndex        =   1
      Top             =   2400
      Width           =   1155
   End
   Begin Proyecto1.AxRangeSlider AxRangeSlider1 
      Height          =   1260
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   2223
      Enabled         =   -1  'True
      Style           =   1
      BorderColor     =   16744576
      BackColor       =   -2147483633
      GradientColor1  =   255
      GradientColor2  =   49152
      BorderWidth     =   1
      CornerCurve     =   10
      ValuesForeColor =   9257492
      MarksForeColor  =   16777215
      BeginProperty ValuesFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty MarksFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorRightMark  =   255
      ColorLeftMark   =   16761024
      ValuesLineColor =   9257492
      Min             =   "0"
      Max             =   "100"
      Interval        =   3
      ValueType       =   1
      DateValueIntervalBy=   0
      ColorSelector   =   12582912
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Right Mark"
      Height          =   195
      Left            =   3975
      TabIndex        =   11
      Top             =   2205
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Left Mark"
      Height          =   195
      Left            =   4065
      TabIndex        =   10
      Top             =   1875
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ValueType           DateValueInterval by"
      Height          =   195
      Left            =   285
      TabIndex        =   6
      Top             =   2190
      Width           =   2775
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Min                  Max                     Interval"
      Height          =   195
      Left            =   315
      TabIndex        =   5
      Top             =   1575
      Width           =   2865
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AxRangeSlider1_ChangeMarks(vLeftMark As String, vRightMark As String)
Mark1.Text = vLeftMark
Mark2.Text = vRightMark
End Sub


Private Sub Command1_Click()
AxRangeSlider1.SetMarkLValue Mark1.Text
End Sub

Private Sub Command2_Click()
AxRangeSlider1.SetMarkRValue Mark2.Text
End Sub

Private Sub Form_Load()

List1.AddItem "eDateValue"
List1.AddItem "eNumValue"
List1.AddItem "eLetterValue"

List2.AddItem "byDay"
List2.AddItem "byMonth"
List2.AddItem "byYear"

With AxRangeSlider1
  txtINI.Text = .Min
  txtFIN.Text = .Max
  txtInterval.Text = .Interval

End With
End Sub

Private Sub Form_Resize()
AxRangeSlider1.Move 150, 150, Me.ScaleWidth - 300, 1250
End Sub

Private Sub List1_Click()
AxRangeSlider1.ValueType = List1.ListIndex
txtINI.Text = AxRangeSlider1.Min
txtFIN.Text = AxRangeSlider1.Max
End Sub

Private Sub List2_Click()
AxRangeSlider1.DateValueIntervalBy = List2.ListIndex
End Sub

Private Sub txtFIN_Change()
On Error Resume Next
AxRangeSlider1.Max = txtFIN.Text
End Sub

Private Sub txtINI_Change()
On Error Resume Next
AxRangeSlider1.Min = txtINI.Text
End Sub

Private Sub txtInterval_Change()
On Error Resume Next
AxRangeSlider1.Interval = txtInterval.Text
End Sub
