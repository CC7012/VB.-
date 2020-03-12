VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   7860
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9075
   LinkTopic       =   "Form2"
   ScaleHeight     =   7860
   ScaleWidth      =   9075
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "计算"
      Height          =   615
      Left            =   120
      TabIndex        =   27
      Top             =   6600
      Width           =   3615
   End
   Begin VB.Frame Frame2 
      Caption         =   "星/角接"
      Height          =   615
      Left            =   1680
      TabIndex        =   24
      Top             =   5760
      Width           =   2055
      Begin VB.OptionButton Option4 
         Caption         =   "角"
         Height          =   180
         Left            =   1200
         TabIndex        =   26
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option3 
         Caption         =   "星"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "单/双层"
      Height          =   615
      Left            =   1680
      TabIndex        =   21
      Top             =   5040
      Width           =   2055
      Begin VB.OptionButton Option2 
         Caption         =   "双"
         Height          =   255
         Left            =   1200
         TabIndex        =   23
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton Option1 
         Caption         =   "单"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   1680
      TabIndex        =   19
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   1680
      TabIndex        =   17
      Top             =   4080
      Width           =   2055
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   1680
      TabIndex        =   15
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   1680
      TabIndex        =   13
      Top             =   3120
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   1680
      TabIndex        =   7
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label12 
      Height          =   6015
      Left            =   4080
      TabIndex        =   28
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "绕组型式："
      Height          =   180
      Left            =   120
      TabIndex        =   20
      Top             =   5280
      Width           =   900
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "起动电流倍数："
      Height          =   180
      Left            =   120
      TabIndex        =   18
      Top             =   4560
      Width           =   1260
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "起动转矩倍数："
      Height          =   180
      Left            =   120
      TabIndex        =   16
      Top             =   4080
      Width           =   1260
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "失步转矩倍数："
      Height          =   180
      Left            =   120
      TabIndex        =   14
      Top             =   3600
      Width           =   1260
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "额定功率因数："
      Height          =   180
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   1260
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "额定效率："
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "极对数："
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "额定频率(HZ)："
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "额定线电压(V)："
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "相数："
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "额定功率（KW）："
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1440
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1.Value Then
DSC = 1
Else
DSC = 2
End If

If Option3.Value Then           '星接或者角接时相电压计算'
UN = UN1 / Sqr(3)
Else
UN = UN1
End If
EJ = "额定相电压(V)：" & UN & vbCrLf

NI = (PN * 1000) / (m * UN * eta * GLYS)    '相电流计算'
EJ = EJ & "额定相电流(A)：" & NI & vbCrLf

nN = 60 * f / p              '额定转速计算'
EJ = EJ & "额定转速(r/min)：" & nN & vbCrLf

TN = 9.549 * PN * 1000 / nN '额定转矩计算'
EJ = EJ & "额定转矩(N*m)" & TN & vbCrLf

EJ = EJ & "绝缘等级：B级" & vbCrLf


Label12.Caption = EJ

End Sub

Private Sub Text1_LostFocus()
PN = Text1.Text
End Sub
Private Sub Text2_LostFocus()
m = Text2.Text
End Sub
Private Sub Text3_LostFocus()
UN1 = Text3.Text
End Sub
Private Sub Text4_LostFocus()
f = Text4.Text
End Sub
Private Sub Text5_LostFocus()
p = Text5.Text
End Sub
Private Sub Text6_LostFocus()
eta = Text6.Text
End Sub
Private Sub Text7_LostFocus()
GLYS = Text7.Text
End Sub
Private Sub Text8_LostFocus()
SBZJBS = Text8.Text
End Sub
Private Sub Text9_LostFocus()
QDZJBS = Text9.Text
End Sub
Private Sub Text10_LostFocus()
QDDLBS = Text10.Text
End Sub





