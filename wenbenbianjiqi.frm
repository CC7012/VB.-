VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command6 
      Caption         =   "½áÊø"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "20ºÅ"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   2040
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ºÚÌå"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   2640
      Picture         =   "wenbenbianjiqi.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Õ³Ìù"
      Top             =   1080
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   2640
      Picture         =   "wenbenbianjiqi.frx":0532
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "¸´ÖÆ"
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   2640
      Picture         =   "wenbenbianjiqi.frx":0A64
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "¼ôÇÐ"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   2535
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "wenbenbianjiqi.frx":0F96
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim st$
Private Sub Command1_Click()
st = Text1.SelText
Text1.SelText = ""
End Sub
Private Sub Command2_Click()
st = Text1.SelText
End Sub
Private Sub Command3_Click()
Text1.SelText = st
End Sub
Private Sub Command4_Click()
Text1.FontName = "ºÚÌå"
End Sub
Private Sub Command5_Click()
Text1.FontSize = 20
End Sub
Private Sub Command6_Click()
End
End Sub

