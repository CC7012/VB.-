VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "额定数据和技术要求"
   ClientHeight    =   5955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7050
   Icon            =   "Form2.frx":0000
   ScaleHeight     =   5955
   ScaleWidth      =   7050
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
EJ = "额定相电压:" & UN & vbCrLf

NI = (PN * 1000) / (m * UN * eta * GLYS)
EJ = EJ & "额定相电流:" & NI & vbCrLf



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




