VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "异步起动永磁同步电动机"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14640
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   14640
   StartUpPosition =   3  '窗口缺省
   Begin TabDlg.SSTab 转子参数 
      Height          =   8535
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   15055
      _Version        =   393216
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "额定数据和技术要求"
      TabPicture(0)   =   "Y.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(9)=   "Label10"
      Tab(0).Control(10)=   "Label11"
      Tab(0).Control(11)=   "Label63"
      Tab(0).Control(12)=   "Label64"
      Tab(0).Control(13)=   "Frame4"
      Tab(0).Control(14)=   "txtPN"
      Tab(0).Control(15)=   "txtm"
      Tab(0).Control(16)=   "txtUN1"
      Tab(0).Control(17)=   "txtf"
      Tab(0).Control(18)=   "txtGLYS"
      Tab(0).Control(19)=   "txteta"
      Tab(0).Control(20)=   "txtSBZJBS"
      Tab(0).Control(21)=   "txtQDZJBS"
      Tab(0).Control(22)=   "txtQDDLBS"
      Tab(0).Control(23)=   "Frame1"
      Tab(0).Control(24)=   "Frame2"
      Tab(0).Control(25)=   "txtp"
      Tab(0).Control(26)=   "Frame6"
      Tab(0).Control(27)=   "txtPSNX"
      Tab(0).ControlCount=   28
      TabCaption(1)   =   "定子参数"
      TabPicture(1)   =   "Y.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label52"
      Tab(1).Control(1)=   "Label51"
      Tab(1).Control(2)=   "Label50"
      Tab(1).Control(3)=   "Label49"
      Tab(1).Control(4)=   "Label48"
      Tab(1).Control(5)=   "Label47"
      Tab(1).Control(6)=   "Label46"
      Tab(1).Control(7)=   "Label24"
      Tab(1).Control(8)=   "Label23"
      Tab(1).Control(9)=   "Label22"
      Tab(1).Control(10)=   "Label21"
      Tab(1).Control(11)=   "Label20"
      Tab(1).Control(12)=   "Label19"
      Tab(1).Control(13)=   "Label18"
      Tab(1).Control(14)=   "Label17"
      Tab(1).Control(15)=   "txtQ17"
      Tab(1).Control(16)=   "Label16"
      Tab(1).Control(17)=   "Label15"
      Tab(1).Control(18)=   "Label14"
      Tab(1).Control(19)=   "Label13"
      Tab(1).Control(20)=   "Label58"
      Tab(1).Control(21)=   "txtd"
      Tab(1).Control(22)=   "txty"
      Tab(1).Control(23)=   "txth"
      Tab(1).Control(24)=   "txtCi"
      Tab(1).Control(25)=   "txthd2"
      Tab(1).Control(26)=   "txtd12"
      Tab(1).Control(27)=   "txtNt2"
      Tab(1).Control(28)=   "txthd1"
      Tab(1).Control(29)=   "txtd11"
      Tab(1).Control(30)=   "txtNt1"
      Tab(1).Control(31)=   "txta1"
      Tab(1).Control(32)=   "txtNs"
      Tab(1).Control(33)=   "txtKFe"
      Tab(1).Control(34)=   "txtdelta"
      Tab(1).Control(35)=   "txtalpha1"
      Tab(1).Control(36)=   "txth12"
      Tab(1).Control(37)=   "txtr1"
      Tab(1).Control(38)=   "txtb1"
      Tab(1).Control(39)=   "txtb01"
      Tab(1).Control(40)=   "txth01"
      Tab(1).Control(41)=   "txtQ1"
      Tab(1).Control(42)=   "txtDi1"
      Tab(1).Control(43)=   "txtL1"
      Tab(1).Control(44)=   "txtD1"
      Tab(1).Control(45)=   "Picture1"
      Tab(1).Control(46)=   "txtmur"
      Tab(1).ControlCount=   47
      TabCaption(2)   =   "转子参数"
      TabPicture(2)   =   "Y.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "SSTab2"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "性能指标"
      TabPicture(3)   =   "Y.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "计算过程"
      TabPicture(4)   =   "Y.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Text1"
      Tab(4).Control(1)=   "CommandTC"
      Tab(4).Control(2)=   "CommandJS"
      Tab(4).ControlCount=   3
      Begin VB.TextBox txtPSNX 
         Height          =   375
         Left            =   -69720
         TabIndex        =   166
         Text            =   "0.008"
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Frame Frame6 
         Caption         =   "防护式/封闭型自扇冷式"
         Height          =   975
         Left            =   -70920
         TabIndex        =   162
         Top             =   480
         Width           =   3135
         Begin VB.OptionButton FBXZSLS 
            Caption         =   "封闭型自扇冷式"
            Height          =   180
            Left            =   1560
            TabIndex        =   164
            Top             =   480
            Width           =   1695
         End
         Begin VB.OptionButton FHS 
            Caption         =   "防护式"
            Height          =   375
            Left            =   120
            TabIndex        =   163
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.TextBox txtmur 
         Height          =   375
         Left            =   -69720
         TabIndex        =   148
         Text            =   "1.05"
         Top             =   5400
         Width           =   1815
      End
      Begin VB.TextBox txtp 
         Height          =   390
         Left            =   -73080
         TabIndex        =   48
         Text            =   "3"
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Frame Frame2 
         Caption         =   "星/角接"
         Height          =   1455
         Left            =   -72000
         TabIndex        =   45
         Top             =   5460
         Width           =   975
         Begin VB.OptionButton JIAO 
            Caption         =   "角"
            Height          =   375
            Left            =   240
            TabIndex        =   47
            Top             =   840
            Width           =   615
         End
         Begin VB.OptionButton XING 
            Caption         =   "星"
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "单/双层"
         Height          =   1455
         Left            =   -73080
         TabIndex        =   42
         Top             =   5460
         Width           =   975
         Begin VB.OptionButton SHUANG 
            Caption         =   "双"
            Height          =   375
            Left            =   240
            TabIndex        =   44
            Top             =   840
            Width           =   615
         End
         Begin VB.OptionButton DAN 
            Caption         =   "单"
            Height          =   255
            Left            =   240
            TabIndex        =   43
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.TextBox txtQDDLBS 
         Height          =   375
         Left            =   -73080
         TabIndex        =   41
         Text            =   "9.7"
         Top             =   4980
         Width           =   2055
      End
      Begin VB.TextBox txtQDZJBS 
         Height          =   375
         Left            =   -73080
         TabIndex        =   40
         Text            =   "3.0"
         Top             =   4500
         Width           =   2055
      End
      Begin VB.TextBox txtSBZJBS 
         Height          =   375
         Left            =   -73080
         TabIndex        =   39
         Text            =   "1.8"
         Top             =   4020
         Width           =   2055
      End
      Begin VB.TextBox txteta 
         Height          =   375
         Left            =   -73080
         TabIndex        =   38
         Text            =   "0.94"
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox txtGLYS 
         Height          =   375
         Left            =   -73080
         TabIndex        =   37
         Text            =   "0.95"
         Top             =   3540
         Width           =   2055
      End
      Begin VB.TextBox txtf 
         Height          =   375
         Left            =   -73080
         TabIndex        =   36
         Text            =   "50"
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox txtUN1 
         Height          =   375
         Left            =   -73080
         TabIndex        =   35
         Text            =   "380"
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox txtm 
         Height          =   375
         Left            =   -73080
         TabIndex        =   34
         Text            =   "3"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox txtPN 
         Height          =   375
         Left            =   -73080
         TabIndex        =   33
         Text            =   "30"
         Top             =   600
         Width           =   2055
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   4665
         Left            =   -64920
         Picture         =   "Y.frx":008C
         ScaleHeight     =   4605
         ScaleWidth      =   3855
         TabIndex        =   32
         Top             =   1320
         Width           =   3915
      End
      Begin VB.TextBox txtD1 
         Height          =   375
         Left            =   -72960
         TabIndex        =   31
         Text            =   "40"
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtL1 
         Height          =   375
         Left            =   -72960
         TabIndex        =   30
         Text            =   "21"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtDi1 
         Height          =   375
         Left            =   -72960
         TabIndex        =   29
         Text            =   "28.5"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtQ1 
         Height          =   375
         Left            =   -72960
         TabIndex        =   28
         Text            =   "72"
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txth01 
         Height          =   375
         Left            =   -69720
         TabIndex        =   27
         Text            =   "0.1"
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtb01 
         Height          =   375
         Left            =   -69720
         TabIndex        =   26
         Text            =   "0.38"
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtb1 
         Height          =   375
         Left            =   -69720
         TabIndex        =   25
         Text            =   "0.68"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtr1 
         Height          =   375
         Left            =   -69720
         TabIndex        =   24
         Text            =   "0.45"
         Top             =   2160
         Width           =   1815
      End
      Begin VB.TextBox txth12 
         Height          =   375
         Left            =   -69720
         TabIndex        =   23
         Text            =   "2.1"
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtalpha1 
         Height          =   375
         Left            =   -69720
         TabIndex        =   22
         Text            =   "35"
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox txtdelta 
         Height          =   375
         Left            =   -72960
         TabIndex        =   21
         Text            =   "0.07"
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox txtKFe 
         Height          =   375
         Left            =   -72960
         TabIndex        =   20
         Text            =   "0.93"
         Top             =   3120
         Width           =   1695
      End
      Begin VB.TextBox txtNs 
         Height          =   375
         Left            =   -72960
         TabIndex        =   19
         Text            =   "32"
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox txta1 
         Height          =   375
         Left            =   -72960
         TabIndex        =   18
         Text            =   "6"
         Top             =   4560
         Width           =   1695
      End
      Begin VB.TextBox txtNt1 
         Height          =   375
         Left            =   -72960
         TabIndex        =   17
         Text            =   "2"
         Top             =   5040
         Width           =   735
      End
      Begin VB.TextBox txtd11 
         Height          =   375
         Left            =   -72120
         TabIndex        =   16
         Text            =   "1.3"
         Top             =   5040
         Width           =   855
      End
      Begin VB.TextBox txthd1 
         Height          =   375
         Left            =   -72960
         TabIndex        =   15
         Text            =   "0.08"
         Top             =   5520
         Width           =   735
      End
      Begin VB.TextBox txtNt2 
         Height          =   375
         Left            =   -72960
         TabIndex        =   14
         Text            =   "0"
         Top             =   6000
         Width           =   735
      End
      Begin VB.TextBox txtd12 
         Height          =   375
         Left            =   -72120
         TabIndex        =   13
         Text            =   "0"
         Top             =   6000
         Width           =   855
      End
      Begin VB.TextBox txthd2 
         Height          =   375
         Left            =   -72960
         TabIndex        =   12
         Text            =   "0"
         Top             =   6480
         Width           =   735
      End
      Begin VB.TextBox txtCi 
         Height          =   375
         Left            =   -72960
         TabIndex        =   11
         Text            =   "0.035"
         Top             =   6960
         Width           =   1695
      End
      Begin VB.TextBox txth 
         Height          =   375
         Left            =   -72960
         TabIndex        =   10
         Text            =   "0.2"
         Top             =   7440
         Width           =   1695
      End
      Begin VB.TextBox txty 
         Height          =   375
         Left            =   -69720
         TabIndex        =   9
         Text            =   "11"
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Frame Frame4 
         Height          =   1215
         Left            =   -73080
         TabIndex        =   5
         Top             =   6960
         Width           =   2775
         Begin VB.OptionButton FZD 
            Caption         =   "分组的"
            Height          =   180
            Left            =   1680
            TabIndex        =   160
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton TXS 
            Caption         =   "同心式"
            Height          =   255
            Left            =   480
            TabIndex        =   8
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton JCS 
            Caption         =   "交叉式"
            Height          =   255
            Left            =   480
            TabIndex        =   7
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton QT 
            Caption         =   "其他"
            Height          =   255
            Left            =   480
            TabIndex        =   6
            Top             =   840
            Width           =   855
         End
      End
      Begin VB.TextBox txtd 
         Height          =   375
         Left            =   -69720
         TabIndex        =   4
         Text            =   "1.5"
         Top             =   4560
         Width           =   1815
      End
      Begin VB.CommandButton CommandJS 
         Caption         =   "计算"
         Height          =   735
         Left            =   -73800
         TabIndex        =   3
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton CommandTC 
         Caption         =   "退出"
         Height          =   735
         Left            =   -73800
         TabIndex        =   2
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   5535
         Left            =   -69840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "Y.frx":3AC3D
         Top             =   480
         Width           =   5895
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   7695
         Left            =   -240
         TabIndex        =   49
         Top             =   360
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   13573
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "槽形参数"
         TabPicture(0)   =   "Y.frx":3AC48
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label34"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label33"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label32"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label31"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label30"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Label29"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "Label28"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Label27"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Label26"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "Label25"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "Label59"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "Label60"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "Label61"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Label62"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).Control(14)=   "txtalpha2"
         Tab(0).Control(14).Enabled=   0   'False
         Tab(0).Control(15)=   "txthr12"
         Tab(0).Control(15).Enabled=   0   'False
         Tab(0).Control(16)=   "txtbr2"
         Tab(0).Control(16).Enabled=   0   'False
         Tab(0).Control(17)=   "txtbr1"
         Tab(0).Control(17).Enabled=   0   'False
         Tab(0).Control(18)=   "txtb02"
         Tab(0).Control(18).Enabled=   0   'False
         Tab(0).Control(19)=   "txth02"
         Tab(0).Control(19).Enabled=   0   'False
         Tab(0).Control(20)=   "txtQ2"
         Tab(0).Control(20).Enabled=   0   'False
         Tab(0).Control(21)=   "txtL2"
         Tab(0).Control(21).Enabled=   0   'False
         Tab(0).Control(22)=   "txtDi2"
         Tab(0).Control(22).Enabled=   0   'False
         Tab(0).Control(23)=   "Picture2"
         Tab(0).Control(23).Enabled=   0   'False
         Tab(0).Control(24)=   "txtLB"
         Tab(0).Control(24).Enabled=   0   'False
         Tab(0).Control(25)=   "txtAB"
         Tab(0).Control(25).Enabled=   0   'False
         Tab(0).Control(26)=   "TXTDR"
         Tab(0).Control(26).Enabled=   0   'False
         Tab(0).Control(27)=   "txtAR"
         Tab(0).Control(27).Enabled=   0   'False
         Tab(0).Control(28)=   "Frame5"
         Tab(0).Control(28).Enabled=   0   'False
         Tab(0).ControlCount=   29
         TabCaption(1)   =   "永磁体参数"
         TabPicture(1)   =   "Y.frx":3AC64
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label57"
         Tab(1).Control(1)=   "Label56"
         Tab(1).Control(2)=   "Label12"
         Tab(1).Control(3)=   "Label45"
         Tab(1).Control(4)=   "Label44"
         Tab(1).Control(5)=   "Label43"
         Tab(1).Control(6)=   "Label42"
         Tab(1).Control(7)=   "Label41"
         Tab(1).Control(8)=   "Label40"
         Tab(1).Control(9)=   "Label39"
         Tab(1).Control(10)=   "Label38"
         Tab(1).Control(11)=   "Label37"
         Tab(1).Control(12)=   "Label36"
         Tab(1).Control(13)=   "Label35"
         Tab(1).Control(14)=   "txtdelta2"
         Tab(1).Control(15)=   "txtw2"
         Tab(1).Control(16)=   "txtw1"
         Tab(1).Control(17)=   "txtYCTMD"
         Tab(1).Control(18)=   "Frame3"
         Tab(1).Control(19)=   "txtLM"
         Tab(1).Control(20)=   "txtbM"
         Tab(1).Control(21)=   "txthM"
         Tab(1).Control(22)=   "txtt"
         Tab(1).Control(23)=   "txtIL"
         Tab(1).Control(24)=   "txtalphaBr"
         Tab(1).Control(25)=   "txtHc20"
         Tab(1).Control(26)=   "txtBr20"
         Tab(1).Control(27)=   "YCCLMC"
         Tab(1).ControlCount=   28
         Begin VB.Frame Frame5 
            Caption         =   "转子类型"
            Height          =   1095
            Left            =   360
            TabIndex        =   157
            Top             =   2280
            Width           =   3255
            Begin VB.OptionButton TTZZ 
               Caption         =   "铜条转子"
               Height          =   255
               Left            =   1680
               TabIndex        =   159
               Top             =   480
               Width           =   1095
            End
            Begin VB.OptionButton ZLZZ 
               Caption         =   "铸铝转子"
               Height          =   255
               Left            =   240
               TabIndex        =   158
               Top             =   480
               Width           =   1095
            End
         End
         Begin VB.TextBox txtAR 
            Height          =   375
            Left            =   2040
            TabIndex        =   156
            Text            =   "500"
            Top             =   5280
            Width           =   1455
         End
         Begin VB.TextBox TXTDR 
            Height          =   375
            Left            =   2040
            TabIndex        =   154
            Text            =   "25.96"
            Top             =   4800
            Width           =   1455
         End
         Begin VB.TextBox txtAB 
            Height          =   375
            Left            =   2040
            TabIndex        =   152
            Text            =   "55"
            Top             =   4320
            Width           =   1455
         End
         Begin VB.TextBox txtLB 
            Height          =   375
            Left            =   2040
            TabIndex        =   150
            Text            =   "21"
            Top             =   3840
            Width           =   1455
         End
         Begin VB.PictureBox Picture2 
            AutoSize        =   -1  'True
            Height          =   5715
            Left            =   9240
            Picture         =   "Y.frx":3AC80
            ScaleHeight     =   5655
            ScaleWidth      =   4755
            TabIndex        =   91
            Top             =   720
            Width           =   4815
         End
         Begin VB.TextBox txtDi2 
            Height          =   375
            Left            =   2040
            TabIndex        =   90
            Text            =   "10"
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox txtL2 
            Height          =   375
            Left            =   2040
            TabIndex        =   89
            Text            =   "21"
            Top             =   1200
            Width           =   1575
         End
         Begin VB.TextBox txtQ2 
            Height          =   375
            Left            =   2040
            TabIndex        =   88
            Text            =   "54"
            Top             =   1680
            Width           =   1575
         End
         Begin VB.TextBox txth02 
            Height          =   375
            Left            =   4920
            TabIndex        =   87
            Text            =   "0.08"
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtb02 
            Height          =   375
            Left            =   4920
            TabIndex        =   86
            Text            =   "0.15"
            Top             =   1200
            Width           =   1695
         End
         Begin VB.TextBox txtbr1 
            Height          =   375
            Left            =   4920
            TabIndex        =   85
            Text            =   "0.32"
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox txtbr2 
            Height          =   375
            Left            =   4920
            TabIndex        =   84
            Text            =   "0.3"
            Top             =   2160
            Width           =   1695
         End
         Begin VB.TextBox txthr12 
            Height          =   375
            Left            =   4920
            TabIndex        =   83
            Text            =   "1.8"
            Top             =   2640
            Width           =   1695
         End
         Begin VB.TextBox txtalpha2 
            Height          =   375
            Left            =   4920
            TabIndex        =   82
            Text            =   "30"
            Top             =   3120
            Width           =   1695
         End
         Begin VB.TextBox YCCLMC 
            Height          =   375
            Left            =   -72360
            TabIndex        =   81
            Text            =   "烧结钕铁硼永磁"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtBr20 
            Height          =   375
            Left            =   -72360
            TabIndex        =   80
            Text            =   "1.18"
            Top             =   1560
            Width           =   1695
         End
         Begin VB.TextBox txtHc20 
            Height          =   375
            Left            =   -72360
            TabIndex        =   79
            Text            =   "898"
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox txtalphaBr 
            Height          =   375
            Left            =   -72360
            TabIndex        =   78
            Text            =   "-0.0012"
            Top             =   2520
            Width           =   1695
         End
         Begin VB.TextBox txtIL 
            Height          =   375
            Left            =   -72360
            TabIndex        =   77
            Text            =   "0"
            Top             =   3000
            Width           =   1695
         End
         Begin VB.TextBox txtt 
            Height          =   375
            Left            =   -72360
            TabIndex        =   76
            Text            =   "60"
            Top             =   3480
            Width           =   1695
         End
         Begin VB.TextBox txthM 
            Height          =   390
            Left            =   -68400
            TabIndex        =   75
            Text            =   "0.42"
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtbM 
            Height          =   375
            Left            =   -68400
            TabIndex        =   74
            Text            =   "12.4"
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtLM 
            Height          =   375
            Left            =   -68400
            TabIndex        =   73
            Text            =   "21"
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Frame Frame3 
            Caption         =   "转子磁路结构"
            Height          =   2415
            Left            =   -74640
            TabIndex        =   54
            Top             =   4920
            Width           =   12855
            Begin VB.PictureBox Picture3 
               Height          =   1035
               Left            =   1680
               Picture         =   "Y.frx":74708
               ScaleHeight     =   975
               ScaleWidth      =   1125
               TabIndex        =   69
               Top             =   120
               Width           =   1185
            End
            Begin VB.OptionButton CLJGa 
               Caption         =   "a"
               Height          =   975
               Left            =   2880
               TabIndex        =   68
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox Picture4 
               AutoSize        =   -1  'True
               Height          =   1080
               Left            =   3600
               Picture         =   "Y.frx":77C79
               ScaleHeight     =   1020
               ScaleWidth      =   1065
               TabIndex        =   67
               Top             =   120
               Width           =   1125
            End
            Begin VB.OptionButton CLJGb 
               Caption         =   "b"
               Height          =   1095
               Left            =   4800
               TabIndex        =   66
               Top             =   120
               Width           =   375
            End
            Begin VB.PictureBox Picture5 
               Height          =   1095
               Left            =   5280
               Picture         =   "Y.frx":AA255
               ScaleHeight     =   1035
               ScaleWidth      =   1035
               TabIndex        =   65
               Top             =   120
               Width           =   1095
            End
            Begin VB.OptionButton CLJGc 
               Caption         =   "c"
               Height          =   375
               Left            =   6480
               TabIndex        =   64
               Top             =   480
               Width           =   375
            End
            Begin VB.PictureBox Picture6 
               Height          =   1095
               Left            =   6960
               Picture         =   "Y.frx":DCE38
               ScaleHeight     =   1035
               ScaleWidth      =   1155
               TabIndex        =   63
               Top             =   120
               Width           =   1215
            End
            Begin VB.OptionButton CLJGd 
               Caption         =   "d"
               Height          =   255
               Left            =   8400
               TabIndex        =   62
               Top             =   600
               Width           =   375
            End
            Begin VB.PictureBox Picture7 
               Height          =   1095
               Left            =   1680
               Picture         =   "Y.frx":1102DA
               ScaleHeight     =   1035
               ScaleWidth      =   1035
               TabIndex        =   61
               Top             =   1200
               Width           =   1095
            End
            Begin VB.PictureBox Picture8 
               Height          =   1095
               Left            =   3600
               Picture         =   "Y.frx":143245
               ScaleHeight     =   1035
               ScaleWidth      =   1035
               TabIndex        =   60
               Top             =   1200
               Width           =   1095
            End
            Begin VB.PictureBox Picture9 
               Height          =   1095
               Left            =   5280
               Picture         =   "Y.frx":14760F
               ScaleHeight     =   1035
               ScaleWidth      =   1035
               TabIndex        =   59
               Top             =   1200
               Width           =   1095
            End
            Begin VB.OptionButton BLJGa 
               Caption         =   "a"
               Height          =   375
               Left            =   2880
               TabIndex        =   58
               Top             =   1560
               Width           =   375
            End
            Begin VB.OptionButton BLJGb 
               Caption         =   "b"
               Height          =   255
               Left            =   4800
               TabIndex        =   57
               Top             =   1560
               Width           =   495
            End
            Begin VB.OptionButton BLJGc 
               Caption         =   "c"
               Height          =   255
               Left            =   6480
               TabIndex        =   56
               Top             =   1560
               Width           =   495
            End
            Begin VB.TextBox txtqm 
               Height          =   375
               Left            =   11160
               TabIndex        =   55
               Text            =   "8"
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label53 
               AutoSize        =   -1  'True
               Caption         =   "每极永磁体所跨转子槽数："
               Height          =   180
               Left            =   8880
               TabIndex        =   72
               Top             =   240
               Width           =   2160
            End
            Begin VB.Label Label54 
               AutoSize        =   -1  'True
               Caption         =   "串联结构："
               Height          =   180
               Left            =   240
               TabIndex        =   71
               Top             =   240
               Width           =   900
            End
            Begin VB.Label Label55 
               Caption         =   "并联结构："
               Height          =   375
               Left            =   240
               TabIndex        =   70
               Top             =   1320
               Width           =   975
            End
         End
         Begin VB.TextBox txtYCTMD 
            Height          =   375
            Left            =   -72360
            TabIndex        =   53
            Text            =   "7.45"
            Top             =   3960
            Width           =   1695
         End
         Begin VB.TextBox txtw1 
            Height          =   375
            Left            =   -64200
            TabIndex        =   52
            Text            =   "0.15"
            Top             =   600
            Width           =   1575
         End
         Begin VB.TextBox txtw2 
            Height          =   375
            Left            =   -64200
            TabIndex        =   51
            Text            =   "0.15"
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox txtdelta2 
            Height          =   375
            Left            =   -64200
            TabIndex        =   50
            Text            =   "0.015"
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label Label62 
            AutoSize        =   -1  'True
            Caption         =   "端环截面积(mm^2)："
            Height          =   180
            Left            =   240
            TabIndex        =   155
            Top             =   5280
            Width           =   1620
         End
         Begin VB.Label Label61 
            AutoSize        =   -1  'True
            Caption         =   "端环平均直径(cm)："
            Height          =   180
            Left            =   240
            TabIndex        =   153
            Top             =   4800
            Width           =   1620
         End
         Begin VB.Label Label60 
            AutoSize        =   -1  'True
            Caption         =   "导条截面积(mm^2)："
            Height          =   180
            Left            =   240
            TabIndex        =   151
            Top             =   4320
            Width           =   1620
         End
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            Caption         =   "转子导条长度(cm)："
            Height          =   180
            Left            =   240
            TabIndex        =   149
            Top             =   3840
            Width           =   1620
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "转子槽形"
            Height          =   180
            Left            =   11520
            TabIndex        =   115
            Top             =   6600
            Width           =   720
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "转子内径(cm)："
            Height          =   180
            Left            =   240
            TabIndex        =   114
            Top             =   720
            Width           =   1260
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "转子铁心长度(cm)："
            Height          =   180
            Left            =   240
            TabIndex        =   113
            Top             =   1200
            Width           =   1620
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "转子槽数："
            Height          =   180
            Left            =   360
            TabIndex        =   112
            Top             =   1680
            Width           =   900
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "h02(cm)："
            Height          =   180
            Left            =   3960
            TabIndex        =   111
            Top             =   720
            Width           =   810
         End
         Begin VB.Label Label30 
            Caption         =   "b02(cm)："
            Height          =   375
            Left            =   3960
            TabIndex        =   110
            Top             =   1200
            Width           =   855
         End
         Begin VB.Label Label31 
            Caption         =   "br1(cm)："
            Height          =   375
            Left            =   3960
            TabIndex        =   109
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "br2(cm)："
            Height          =   180
            Left            =   3960
            TabIndex        =   108
            Top             =   2160
            Width           =   810
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "hr12(cm)："
            Height          =   180
            Left            =   3960
            TabIndex        =   107
            Top             =   2640
            Width           =   900
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "α2："
            Height          =   180
            Left            =   4080
            TabIndex        =   106
            Top             =   3120
            Width           =   450
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "永磁材料名称："
            Height          =   180
            Left            =   -74760
            TabIndex        =   105
            Top             =   600
            Width           =   1260
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "20摄氏度时"
            Height          =   180
            Left            =   -74760
            TabIndex        =   104
            Top             =   1080
            Width           =   900
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "剩磁(T)："
            Height          =   180
            Left            =   -74760
            TabIndex        =   103
            Top             =   1560
            Width           =   810
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "矫顽力(kA/m)："
            Height          =   180
            Left            =   -74760
            TabIndex        =   102
            Top             =   2040
            Width           =   1260
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "剩磁可逆温度系数(K^(-1))："
            Height          =   180
            Left            =   -74760
            TabIndex        =   101
            Top             =   2520
            Width           =   2340
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "剩磁不可逆损失率(K^(-1))："
            Height          =   180
            Left            =   -74760
            TabIndex        =   100
            Top             =   3000
            Width           =   2340
         End
         Begin VB.Label Label41 
            Caption         =   "预计永磁体工作温度："
            Height          =   255
            Left            =   -74760
            TabIndex        =   99
            Top             =   3480
            Width           =   2175
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            Caption         =   "磁化方向长度(cm)："
            Height          =   180
            Left            =   -70320
            TabIndex        =   98
            Top             =   600
            Width           =   1620
         End
         Begin VB.Label Label43 
            AutoSize        =   -1  'True
            Caption         =   "每极永磁体宽度(cm)："
            Height          =   180
            Left            =   -70320
            TabIndex        =   97
            Top             =   1080
            Width           =   1800
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "永磁体轴向长度(cm)："
            Height          =   180
            Left            =   -70320
            TabIndex        =   96
            Top             =   1560
            Width           =   1800
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "永磁体密度（g/cm^3）:"
            Height          =   180
            Left            =   -74760
            TabIndex        =   95
            Top             =   3960
            Width           =   1890
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "隔磁磁桥1的宽度(cm)："
            Height          =   180
            Left            =   -66240
            TabIndex        =   94
            Top             =   600
            Width           =   1890
         End
         Begin VB.Label Label56 
            AutoSize        =   -1  'True
            Caption         =   "隔磁磁桥2的宽度(cm)："
            Height          =   180
            Left            =   -66240
            TabIndex        =   93
            Top             =   1200
            Width           =   1890
         End
         Begin VB.Label Label57 
            Caption         =   "永磁体沿磁化方向与永磁体槽间的间隙(cm)："
            Height          =   495
            Left            =   -66240
            TabIndex        =   92
            Top             =   1560
            Width           =   1815
         End
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         Caption         =   "杂散损耗："
         Height          =   180
         Left            =   -70920
         TabIndex        =   165
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label Label63 
         Height          =   495
         Left            =   -70920
         TabIndex        =   161
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         Caption         =   "相对导磁系数："
         Height          =   180
         Left            =   -71280
         TabIndex        =   147
         Top             =   5400
         Width           =   1260
      End
      Begin VB.Label Label11 
         Caption         =   "极对数"
         Height          =   255
         Left            =   -74760
         TabIndex        =   146
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "绕组型式："
         Height          =   180
         Left            =   -74760
         TabIndex        =   145
         Top             =   5460
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "起动电流倍数："
         Height          =   180
         Left            =   -74760
         TabIndex        =   144
         Top             =   4980
         Width           =   1260
      End
      Begin VB.Label Label8 
         Caption         =   "起动转矩倍数："
         Height          =   255
         Left            =   -74760
         TabIndex        =   143
         Top             =   4500
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "失步转矩倍数："
         Height          =   255
         Left            =   -74760
         TabIndex        =   142
         Top             =   4020
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "额定效率："
         Height          =   255
         Left            =   -74760
         TabIndex        =   141
         Top             =   3000
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "额定功率因数："
         Height          =   375
         Left            =   -74760
         TabIndex        =   140
         Top             =   3600
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "额定频率(HZ)："
         Height          =   255
         Left            =   -74760
         TabIndex        =   139
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "额定线电压（V）："
         Height          =   180
         Left            =   -74760
         TabIndex        =   138
         Top             =   1680
         Width           =   1530
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "相数："
         Height          =   180
         Left            =   -74760
         TabIndex        =   137
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "额定功率(kW)："
         Height          =   180
         Left            =   -74760
         TabIndex        =   136
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "定子槽形"
         Height          =   180
         Left            =   -63360
         TabIndex        =   135
         Top             =   5760
         Width           =   720
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "定子外径(cm)："
         Height          =   180
         Left            =   -74760
         TabIndex        =   134
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "定子铁心长度（cm）："
         Height          =   180
         Left            =   -74760
         TabIndex        =   133
         Top             =   1680
         Width           =   1800
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "定子内径(cm)："
         Height          =   180
         Left            =   -74760
         TabIndex        =   132
         Top             =   1200
         Width           =   1260
      End
      Begin VB.Label txtQ17 
         AutoSize        =   -1  'True
         Caption         =   "定子槽数："
         Height          =   180
         Left            =   -74760
         TabIndex        =   131
         Top             =   2160
         Width           =   900
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "h01(cm)："
         Height          =   180
         Left            =   -70680
         TabIndex        =   130
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "b01(cm)："
         Height          =   180
         Left            =   -70680
         TabIndex        =   129
         Top             =   1200
         Width           =   810
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "b1(cm)："
         Height          =   180
         Left            =   -70680
         TabIndex        =   128
         Top             =   1680
         Width           =   720
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "r1(cm)："
         Height          =   180
         Left            =   -70680
         TabIndex        =   127
         Top             =   2160
         Width           =   720
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "h12(cm)："
         Height          =   180
         Left            =   -70680
         TabIndex        =   126
         Top             =   2640
         Width           =   810
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "α1："
         Height          =   180
         Left            =   -70560
         TabIndex        =   125
         Top             =   3120
         Width           =   450
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "气隙长度(cm)："
         Height          =   180
         Left            =   -74760
         TabIndex        =   124
         Top             =   2640
         Width           =   1260
      End
      Begin VB.Label Label24 
         Caption         =   "铁心叠压系数："
         Height          =   255
         Left            =   -74760
         TabIndex        =   123
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "每槽导体数："
         Height          =   180
         Left            =   -74760
         TabIndex        =   122
         Top             =   4080
         Width           =   1080
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "并联支路数："
         Height          =   180
         Left            =   -74760
         TabIndex        =   121
         Top             =   4560
         Width           =   1080
      End
      Begin VB.Label Label48 
         Caption         =   "并绕根数-线径(mm)-双边绝缘厚度(mm)："
         Height          =   540
         Left            =   -74760
         TabIndex        =   120
         Top             =   5040
         Width           =   1680
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "槽绝缘厚度(cm)："
         Height          =   180
         Left            =   -74640
         TabIndex        =   119
         Top             =   6960
         Width           =   1440
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "槽楔厚度(cm)："
         Height          =   180
         Left            =   -74640
         TabIndex        =   118
         Top             =   7440
         Width           =   1260
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "节距："
         Height          =   180
         Left            =   -70680
         TabIndex        =   117
         Top             =   4080
         Width           =   540
      End
      Begin VB.Label Label52 
         Caption         =   "绕组直线部分伸出长(cm)："
         Height          =   615
         Left            =   -70680
         TabIndex        =   116
         Top             =   4560
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandJS_Click()
'1.额定功率
PN = Val(txtPN.Text)
'2.相数
m = Val(txtm.Text)
'3.额定线电压
UN1 = Val(txtUN1.Text)
'4.额定频率
f = Val(txtf.Text)
'5.极对数
p = Val(txtp.Text)
'6.额定效率
eta = Val(txteta.Text)
'7.额定功率因数
GLYS = Val(txtGLYS.Text)
'8.失步转矩倍数
SBZJBS = Val(txtSBZJBS.Text)
'9.起动转矩倍数
QDZJBS = Val(txtQDZJBS.Text)
'10.起动电流倍数
QDDLBS = Val(txtQDDLBS.Text)

''''''''''''''''''''''''''''''''''''''''''''''''''
'11.绕组形式
If DAN.Value Then
DSC = 1
ElseIf SHUANG.Value Then
DSC = 2
Else
MsgBox ("请选择绕组形式")
Exit Sub
End If

'12.额定相电压
If XING.Value Then           '星接或者角接时相电压计算'
UN = UN1 / Sqr(3)
ElseIf JIAO.Value Then
UN = UN1
Else
MsgBox ("请选择绕组形式")
Exit Sub
End If
EJ = "额定相电压(V)：" & UN & vbCrLf

'13.额定相电流
NI = (PN * 1000) / (m * UN * eta * GLYS)    '相电流计算'
EJ = EJ & "额定相电流(A)：" & NI & vbCrLf

'14.额度转速
nN = 60 * f / p              '额定转速计算'
EJ = EJ & "额定转速(r/min)：" & nN & vbCrLf

'15.额定转矩
TN = 9.549 * PN * 1000 / nN '额定转矩计算'
EJ = EJ & "额定转矩(N*m)" & TN & vbCrLf

'16
EJ = EJ & "绝缘等级：B级" & vbCrLf



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'17
If CLJGb.Value Or CLJGd.Value Or CLJGc.Value Or CLJGa.Value Then
EJ = EJ & "转子磁路结构: 串联式结构" & vbCrLf
ElseIf BLJGa.Value Or BLJGb.Value Or BLJGc.Value Then
EJ = EJ & "转子磁路结构: 并联式结构" & vbCrLf
Else
MsgBox ("选择永磁体结构")
Exit Sub
End If



'定转子参数
'18 气隙长度
delta = Val(txtdelta.Text)
'19定子外径
D1 = Val(txtD1.Text)
'20定子内径
Di1 = Val(txtDi1.Text)
'21转子外径
D2 = Di1 - 2 * delta
EJ = EJ & "转子外径：" & D2 & vbCrLf

'22转子内径
Di2 = Val(txtDi2.Text)
'23定转子铁心长度
L1 = Val(txtL1.Text)
L2 = Val(txtL2.Text)
'24电枢计算长度
If L1 < L2 Then
La = L1
Else
La = L2
End If
Lef = La + 2 * delta
EJ = EJ & "电枢计算长度：" & Lef & vbCrLf

'25定转子槽数
Q1 = Val(txtQ1.Text)
Q2 = Val(txtQ2.Text)
'26定子每极每相槽数（60度相带）
q = Q1 / (2 * m * p)
EJ = EJ & "定子每极每相槽数：" & q & vbCrLf

'27极距
tao1 = 3.1415926 * Di1 / (2 * p)
EJ = EJ & "极距：" & tao1 & vbCrLf

'28硅钢片质量
KFe = Val(txtKFe.Text)
If L1 > L2 Then
LB = L1
Else
LB = L2
End If
mFe = 7.8 * LB * KFe * (D1 + 0.5) * (D1 + 0.5) * 0.001
EJ = EJ & "硅钢片质量：" & mFe & vbCrLf

''''''''''''''''''''''''''''
'永磁体计算
Br20 = Val(txtBr20.Text)
Hc20 = Val(txtHc20.Text)
alphaBr = Val(txtalphaBr.Text)
IL = Val(txtIL.Text)
t = Val(txtt.Text)
'29永磁材料
EJ = EJ & "永磁材料："
EJ = EJ & "20摄氏度时，" & YCCLMC.Text & ",剩磁" & txtBr20.Text & ",矫顽力" & txtHc20.Text & vbCrLf
'30计算剩磁密度
Br = (1 + (t - 20) * alphaBr) * (1 - IL / 100) * Br20
EJ = EJ & "计算剩磁密度：" & Br & vbCrLf
'31
Hc = (1 + (t - 20) * alphaBr) * (1 - IL / 100) * Hc20
EJ = EJ & "计算矫顽力：" & Hc & vbCrLf
'32
'真空磁导率=4*3.1415926*0.0000001
mu = Br20 / (4 * 3.1415926 * 0.0000001 * Hc20 * 1000)
EJ = EJ & "相对回复磁导率：" & mu & vbCrLf
'33磁化方向长度
hM = Val(txthM.Text)
'34每极永磁体宽度
bM = Val(txtbM.Text)
'35永磁体轴向长度
LM = Val(txtLM.Text)
'36提供每极磁通的截面积
If CLJGb.Value Or CLJGd.Value Or CLJGc.Value Or CLJGa.Value Then
Am = bM * LM
ElseIf BLJGa.Value Or BLJGb.Value Or BLJGc.Value Then
Am = 2 * bM * LM
Else
MsgBox ("选择永磁体结构")
Exit Sub
End If
EJ = EJ & "提供每极磁通的截面积：" & Am & vbCrLf
'37
YCTMD = Val(txtYCTMD.Text)
mm = 2 * p * bM * hM * LM * YCTMD * 0.001
EJ = EJ & "永磁体质量：" & mm & vbCrLf

'38定子槽形
h01 = Val(txth01.Text)
b01 = Val(txtb01.Text)
B1 = Val(txtb1.Text)
r1 = Val(txtr1.Text)
h12 = Val(txth12.Text)
alpha1 = Val(txtalpha1.Text)
'39转子槽形
h02 = Val(txth02.Text)
b02 = Val(txtb02.Text)
br1 = Val(txtbr1.Text)
br2 = Val(txtbr2.Text)
hr12 = Val(txthr12.Text)
alpha2 = Val(txtalpha2.Text)
'40
T1 = 3.1415926 * Di1 / Q1
EJ = EJ & "定子齿距：" & T1 & vbCrLf
'41
tsk = T1
EJ = EJ & "定子斜槽距离：" & tsk & vbCrLf
'42
h1 = (B1 - b01) * Tan(alpha1 * 3.1415926 / 180) / 2
bt11 = 3.1415926 * (Di1 + 2 * (h01 + h12)) / Q1 - 2 * r1
bt12 = 3.1415926 * (Di1 + 2 * (h01 + h1)) / Q1 - B1
If bt12 <= bt11 Then
bt1 = bt12 + (bt11 - bt12) / 3
Else
bt1 = bt11 + (bt12 - bt11) / 3
End If
EJ = EJ & "定子计算齿宽：" & bt1 & vbCrLf
'43
hj1 = (D1 - Di1) / 2 - (h01 + h12 + 2 / 3 * r1)
EJ = EJ & "定子轭计算高度：" & hj1 & vbCrLf

'44
ht1 = h12 + r1 / 3
EJ = EJ & "定子齿磁路计算长度：" & ht1 & vbCrLf

'45
Lj1 = 3.1415926 * (D1 - hj1) / (4 * p)
EJ = EJ & "定子轭磁路计算长度：" & Lj1 & vbCrLf

'46
Vt1 = Q1 * L1 * KFe * ht1 * bt1
EJ = EJ & "定子齿体积：" & Vt1 & vbCrLf

'47
VJ1 = 3.1415926 * L1 * KFe * hj1 * (D1 - hj1)
EJ = EJ & "定子轭体积：" & VJ1 & vbCrLf

'48
t2 = 3.1415926 * D2 / Q2
EJ = EJ & "转子齿距：" & t2 & vbCrLf

'49
ht2 = hr12
EJ = EJ & "转子齿磁路计算长度：" & ht2 & vbCrLf

'50
If CLJGb.Value Or CLJGd.Value Or CLJGc.Value Or CLJGa.Value Then
'仅平底槽
hj2 = (D2 - Di2) / 2 - (h02 + hr12) - hM
If p = 1 Then
hj2 = hj2 + Di2 / 3
End If
ElseIf BLJGa.Value Or BLJGb.Value Or BLJGc.Value Then
hj2 = bM
Else
MsgBox ("选择永磁体结构")
Exit Sub
End If
EJ = EJ & "转子轭计算高度：" & hj2 & vbCrLf

'51
Lj2 = 3.1415926 * (Di2 + hj2) / 4 / p
EJ = EJ & "转子轭磁路计算长度：" & Lj2 & vbCrLf

'52每槽导体数
Ns = Val(txtNs.Text)
'53并联支路数
a1 = Val(txta1.Text)
'54并绕根数-线径-双边绝缘厚度
Nt1 = Val(txtNt1.Text)
d11 = Val(txtd11.Text)
hd1 = Val(txthd1.Text)
Nt2 = Val(txtNt2.Text)
d12 = Val(txtd12.Text)
hd2 = Val(txthd2.Text)
'55
N = Ns * Q1 / (2 * m * a1)
EJ = EJ & "每相绕组串联匝数：" & N & vbCrLf
'56
H = Val(txth.Text)
As1 = (2 * r1 + B1) / 2 * (h12 - H) + 3.1415926 * r1 * r1 / 2
'EJ = EJ & As1 & vbCrLf
Ci = Val(txtCi.Text)
If DAN.Value Then
Ai = Ci * (2 * h12 + 3.14159236 * r1)
ElseIf SHUANG.Value Then
Ai = Ci * (2 * h12 + 3.14159236 * r1 + 2 * r1 + B1)
Else
MsgBox ("请选择绕组形式")
Exit Sub
End If
'EJ = EJ & Ai & vbCrLf
Aef = As1 - Ai
'EJ = EJ & Aef & vbCrLf
Sf = Ns * (Nt1 * (d11 + hd1) * (d11 + hd1) + Nt2 * (d12 + hd2) * (d12 + hd2)) / Aef
EJ = EJ & "槽满率计算：" & Sf & vbCrLf
'57节距
y = Val(txty.Text)
'58
beta = y / (m * q)
Kp1 = Sin(3.1415926 * beta / 2)
EJ = EJ & "绕组节距因数：" & Kp1 & vbCrLf
'59
alpha3 = 2 * p * 3.1415926 / Q1
Kd1 = Sin(q * alpha3 / 2) / q / Sin(alpha3 / 2)
EJ = EJ & "绕组分布因数：" & Kd1 & vbCrLf
'60
alphas = tsk / tao1 * 3.1415926
Ksk1 = 2 * Sin(alphas / 2) / alphas
EJ = EJ & "斜槽因数：" & Ksk1 & vbCrLf
'61
Kdp = Kd1 * Kp1 * Ksk1
EJ = EJ & "绕组因数：" & Kdp & vbCrLf
'62
If p = 1 Then
k = 0.58
ElseIf p = 2 Or p = 3 Then
k = 0.6
ElseIf p = 4 Then
k = 0.625
Else
MsgBox ("只适合2，4，6，8极电机")
End If
ssinalpha0 = (B1 + 2 * r1) / (B1 + 2 * r1 + 2 * bt1)
ccosalpha0 = Sqr(1 - ssinalpha0 * ssinalpha0)
If DAN.Value Then
If TXS.Value Or JCS.Value Then
MsgBox ("抱歉，目前无法计算单层同心式和单层交叉式线圈")
End If
Else
beta0 = beta
End If
taoy = 3.1415926 * (Di1 + 2 * h01 + h1 + h12 + r1) * beta0 / 2 / p
If DAN.Value Then
LEp = k * taoy
ElseIf SHUANG.Value Then
LEp = taoy / (2 * ccosalpha0)
Else
MsgBox ("请选择绕组形式")
Exit Sub
End If
d = Val(txtd.Text)
Lav = L1 + 2 * (d + LEp)
'EJ = EJ & taoy & vbCrLf
'EJ = EJ & LEp & vbCrLf
EJ = EJ & "线圈平均半匝长：" & Lav & vbCrLf
'63
fd = LEp * ssinalpha0
EJ = EJ & "线圈端部轴向投影长：" & fd & vbCrLf
'64
LE = 2 * (d + LEp)
EJ = EJ & "线圈端部平均长：" & LE & vbCrLf
'65
mcu = 1.05 * 3.1415926 * 8.9 * Q1 * Ns * Lav * (Nt1 * d11 * d11 + Nt2 * d12 * d12) / 4 * 0.00001
EJ = EJ & "定子导线质量(Kg)：" & mcu & vbCrLf
'66
If CLJGb.Value Or CLJGd.Value Or CLJGc.Value Then
qm = Val(txtqm.Text)
If qm = 0 Then
MsgBox ("每极永磁体所跨转子槽数")
End If
alphap = qm / (Q2 / 2 / p)
ElseIf CLJGa.Value Or BLJGa.Value Or BLJGb.Value Or BLJGc.Value Then
tao2 = 3.1415926 * D2 / 2 / p
alphap = (tao2 - b02) / tao2
txtqm = ""
Else
MsgBox ("选择永磁体结构")
Exit Sub
End If
EJ = EJ & "极弧系数：" & alphap & vbCrLf
'67
alphai = alphap
EJ = EJ & "计算极弧系数：" & alphai & vbCrLf
'68
KF = 4 * Sin(alphai * 3.1415926 / 2) / 3.1415926
EJ = EJ & "气隙磁密波形系数：" & KF & vbCrLf
'69
Kphi = 8 * Sin(alphai * 3.1415926 / 2) / 3.1415926 / 3.1415926 / alphai
EJ = EJ & "气隙磁通波形系数：" & Kphi & vbCrLf
'70
Kdelta1 = T1 * (4.4 * delta + 0.75 * b01) / (T1 * (4.4 * delta + 0.75 * b01) - b01 * b01)
Kdelta2 = t2 * (4.4 * delta + 0.75 * b02) / (t2 * (4.4 * delta + 0.75 * b02) - b02 * b02)
Kdelta = Kdelta1 * Kdelta2
EJ = EJ & "气隙系数：" & Kdelta & vbCrLf
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
cycle2:
'71永磁体空载工作点假定值
bm01 = 0.791
cycle1:
'72空载漏磁系数假定值
sigma01 = 1.25
'73
phidelta0 = bm01 * Br * Am * 0.0001 / sigma01
EJ = EJ & "空载主磁通：" & phidelta0 & vbCrLf
'74
Bdelta = phidelta0 * 10000 / alphai / tao1 / Lef
EJ = EJ & "气隙磁密：" & Bdelta & vbCrLf
'75
delta2 = Val(txtdelta2.Text)
'直轴磁路
Fdelta = 2 * Bdelta * (delta2 + Kdelta * delta) * 0.01 / (4 * 3.1415926 * 0.0000001)
EJ = EJ & "直轴磁路气隙磁位差：" & Fdelta & vbCrLf
'交轴磁路
Fdeltaq = 2 * Bdelta * Kdelta * delta * 0.01 / (4 * 3.1415926 * 0.0000001)
EJ = EJ & "交轴磁路气隙磁位差：" & Fdeltaq & vbCrLf
'76
Bbt1 = Bdelta * T1 * Lef / (bt1 * KFe * L1)
EJ = EJ & "定子齿磁密：" & Bbt1 & vbCrLf
'77
'HHt1 = CHQXf(Bbt1)
HHt1 = 41.9
EJ = EJ & "Ht1：" & HHt1 & vbCrLf
Ft1 = 2 * HHt1 * ht1
EJ = EJ & "定子齿磁位差：" & Ft1 & vbCrLf
'78
bbj1 = phidelta0 * 10000 / (2 * L1 * KFe * hj1)
EJ = EJ & "定子轭部磁密：" & bbj1 & vbCrLf
'79
'HHj1 = CHQXf(bbj1)
HHj1 = 16.3
EJ = EJ & "Hj1：" & HHj1 & vbCrLf
HJtao = hj1 / tao1             '校正系数曲线的X轴计算
'If p >= 3 Then
'C1 = EBJZXS6Sf(bbj1, HJtao)
'ElseIf p = 1 Then
'C1 = EBJZXS2Sf(bbj1, HJtao)
'ElseIf p = 2 Then
'C1 = EBJZXS4Sf(bbj1, HJtao)
'End If
C1 = 0.38
EJ = EJ & "定子轭部校正系数：" & C1 & vbCrLf
Fj1 = 2 * C1 * HHj1 * Lj1
EJ = EJ & "定子轭磁位差：" & Fj1 & vbCrLf
'80
bt21 = 3.1415926 * (D2 - 2 * (h02 + hr1)) / Q2 - br1
bt22 = 3.1415926 * (D2 - 2 * (hr12 + h02)) / Q2 - br2
If bt22 < bt21 Then
Bt2 = bt22 + (bt21 - bt22) / 3
Else
Bt2 = bt21 + (bt22 - bt21) / 3
End If
BBt2 = Bdelta * t2 * Lef / (Bt2 * KFe * L2)
EJ = EJ & "转子齿磁密：" & BBt2 & vbCrLf
'81
'HHt2 = CHQXf(BBt2)
HHt2 = 3.74
EJ = EJ & "Ht2：" & HHt2 & vbCrLf
Ft2 = 2 * HHt2 * ht2
EJ = EJ & "转子齿磁位差：" & Ft2 & vbCrLf
'82
bbj2 = phidelta0 * 10000 / (2 * L2 * KFe * hj2)
EJ = EJ & "转子轭磁密：" & bbj2 & vbCrLf
'83
'HHj2 = CHQXf(bbj2)
HHj2 = 2.06
EJ = EJ & "Hj2：" & HHj2 & vbCrLf
HJtao = hj2 / tao1           '校正系数曲线的"x"轴计算


'If p >= 3 Then
'C2 = EBJZXS6Rf(bbj2, HJtao)
'ElseIf p = 1 Then
'C2 = EBJZXS2Rf(bbj2, HJtao)
'ElseIf p = 2 Then
'C2 = EBJZXS4Rf(bbj2, HJtao)
'End If

C2 = 0.6
Fj2 = 2 * C2 * HHj2 * Lj2
EJ = EJ & "转子轭磁位差：" & Fj2 & vbCrLf
'84
sigmaF = Fj2 + Ft2 + Fj1 + Ft1 + Fdelta
EJ = EJ & "每对极总磁位差：" & sigmaF & vbCrLf
F0 = (Fj2 + Ft2 + Fj1 + Ft1 + Fdeltaq) / 2
EJ = EJ & "每极总磁位差：" & F0 & vbCrLf
'85空载漏磁系数
'(1)通过转子槽的漏磁通
hr1 = (br1 - b02) / 2 * Tan(alpha2 * 3.1415926 / 180)
hr2 = hr12 - hr1
phir = 2 * 4 * 3.1415926 * 0.0000001 * F0 * (h02 / b02 + 2 * hr1 / (b02 + br1) + 2 * hr2 / (br2 + br1)) * Lef * 0.01
EJ = EJ & "通过转子槽的漏磁通：" & phir & vbCrLf
'(2)
w1 = Val(txtw1.Text)
w2 = Val(txtw2.Text)
If br2 > (hM + delta2) Then
Min = hM + delta2
Else
Min = br2
End If
HHb1 = F0 / (Min)
EJ = EJ & "Hb1：" & HHb1 & vbCrLf
Bb1 = 2.24 + 0.15 * (HHb1 - 1500) / 1500
EJ = EJ & "Bb1：" & Bb1 & vbCrLf
phix1 = 2 * Bb1 * w1 * Lef * 0.0001
EJ = EJ & "phix1：" & phix1 & vbCrLf
HHb2 = F0 / (hM + delta2)
EJ = EJ & "Hb2：" & HHb2 & vbCrLf
Bb2 = 2.24 + 0.15 * (HHb2 - 1500) / 1500
EJ = EJ & "Bb2：" & Bb2 & vbCrLf
phix2 = Bb2 * w2 * Lef * 0.0001
EJ = EJ & "phix2：" & phix2 & vbCrLf
phix = phix1 + phix2
EJ = EJ & "通过隔磁磁桥的磁通：" & phix & vbCrLf
'(3)
sigma1 = (phix + phir + phidelta0) / phidelta0    '总磁通/主磁通
EJ = EJ & "转子内部漏磁系数：" & sigma1 & vbCrLf
'(4)
If CLJGb.Value Or CLJGd.Value Or CLJGc.Value Or CLJGa.Value Then
bMP = bM
hMp = hM
ElseIf BLJGa.Value Or BLJGb.Value Or BLJGc.Value Then
bMP = 2 * bM
hMp = hM / 2
Else
MsgBox ("选择永磁体结构")
Exit Sub
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''图2-31
sigma2P = 0.282
EJ = EJ & "sigma2P：" & sigma2P & vbCrLf
tao2 = 3.1415926 * D2 / 2 / p
sigma2 = 1 + (sigma2P * bMP) / L2 / tao2
EJ = EJ & "转子端部漏磁系数：" & sigma2 & vbCrLf
'(5)
sigma0 = sigma2 + sigma1 - 1
If (sigma0 - sigma0p) / sigma0 > 0.01 Then
sigma0p = sigma0 - (sigma0 - sigma0p) / 3
GoTo cycle1
Else
GoTo line1
End If
line1:
EJ = EJ & "空载漏磁系数：" & sigma0 & vbCrLf
'86
Kst = (Fdeltaq + Ft1 + Ft2) / (Fdeltaq)
EJ = EJ & "齿磁路饱和系数：" & Kst & vbCrLf
'87
llambdadelta = phidelta0 / sigmaF
EJ = EJ & "主磁导：" & llambdadelta & vbCrLf
'88
mur = Val(txtmur.Text)
If CLJGb.Value Or CLJGd.Value Or CLJGc.Value Or CLJGa.Value Then
lambdadelta = 2 * llambdadelta * hM * 100 / (mur * 4 * 3.1415926 * 0.0000001 * Am)
ElseIf BLJGa.Value Or BLJGb.Value Or BLJGc.Value Then
lambdadelta = llambdadelta * hM * 100 / (mur * mu0 * Am)
Else
MsgBox ("选择永磁体结构")
Exit Sub
End If
EJ = EJ & "主磁导标幺值：" & lambdadelta & vbCrLf
'89
 lambdan = sigma0 * lambdadelta
EJ = EJ & "外磁路总磁导标幺值：" & lambdan & vbCrLf
'90
lambdasigma = (sigma0 - 1) * lambdadelta
EJ = EJ & "漏磁导标幺值：" & lambdasigma & vbCrLf
'91
bm0 = lambdan / (lambdan + 1)
If (bm0 - bm0p) / bm0 > 0.01 Then
bm0p = bm0 - (bm0 - bm0p) / 3
GoTo cycle2
Else
GoTo line2
End If
line2:
'92
EJ = EJ & " 永磁体空载工作点：" & bm0 & vbCrLf
Bdelta1 = KF * phidelta0 * 10000 / (tao1 * Lef * alphai)
EJ = EJ & " 气隙磁密基波幅值：" & Bdelta1 & vbCrLf
'93
E0 = 4.44 * f * Kdp * N * phidelta0 * Kphi
EJ = EJ & " 空载反电动势：" & E0 & vbCrLf

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''参数计算
'94
rr1 = 2.17 * 0.0001 * 2 * Lav * N / (3.1415926 * a1 * (Nt1 * (d11 / 2) * (d11 / 2) + Nt2 * (d12 / 2) * (d12 / 2)))
EJ = EJ & " 定子直流电阻：" & rr1 & vbCrLf
'95
LB = Val(txtLB.Text)
AB = Val(txtAB.Text)
DR = Val(TXTDR.Text)
AR = Val(txtAR.Text)
If ZLZZ.Value Then
KB = 1.04
ElseIf TTZZ.Value Then
KB = 1
Else
MsgBox ("请选择转子类型")
End If
kc = 4 * m * (N * Kdp) * (N * Kdp) / Q2
RB = KB * kc * 4.34 * 0.0001 * LB / AB
RR = kc * Q2 * 4.34 * 0.0001 * DR / 2 / 3.1415926 / p / p / AR
rr2 = RB + RR
EJ = EJ & " 转子折算电阻：" & rr2 & vbCrLf
'96
If ZLZZ.Value Then
mcu2 = 2.7 * (Q2 * AB * LB + 2 * AR * 3.1415926 * DR) * 0.00001
ElseIf TTZZ.Value Then
mcu2 = 8.9 * (Q2 * AB * LB + 2 * AR * 3.1415926 * DR) * 0.00001
Else
MsgBox ("请选择转子类型")
End If
EJ = EJ & " 转子绕组质量：" & mcu2 & vbCrLf
'97
CX = 4 * 3.1415926 * f * 4 * 3.1415926 * 0.0000001 * Lef * Kdp * Kdp * N * N * 0.01 / p
EJ = EJ & " 漏抗系数：" & CX & vbCrLf
'98
If beta >= 0 And beta <= 1 / 3 Then
KU1 = 3 * beta / 4
KL1 = (9 * beta + 1) / 16
ElseIf beta >= 1 / 3 And beta <= 2 / 3 Then
KU1 = (6 * beta - 1) / 4
KL1 = (18 * beta + 1) / 16
ElseIf beta >= 2 / 3 And beta <= 1 Then
KU1 = (3 * beta + 1) / 4
KL1 = (9 * beta + 7) / 16
End If
EJ = EJ & " KU1：" & KU1 & vbCrLf
EJ = EJ & " Kl1：" & KL1 & vbCrLf
lambdau1 = h01 / b01 + 2 * h1 / (b01 + B1)
EJ = EJ & "lambdaU1：" & lambdau1 & vbCrLf
alphaalpha = B1 / 2 / r1
betabetas = (h12 - h1) / 2 / r1
Kr1 = 1 / 3 - (1 - alphaalpha) * (1 / 4 + 1 / 3 / (1 - alphaalpha) + 1 / 2 / (1 - alphaalpha) / (1 - alphaalpha) + 1 / (1 - alphaalpha) / (1 - alphaalpha) / (1 - alphaalpha) + Log(alphaalpha) / (1 - alphaalpha) / (1 - alphaalpha) / (1 - alphaalpha) / (1 - alphaalpha)) / 4
EJ = EJ & "Kr1：" & Kr1 & vbCrLf
Kr2 = (2 * 3.1415926 * 3.1415926 * 3.1415926 - 9 * 3.1415926) / 1536 / betabetas / betabetas / betabetas + 3.1415926 / 16 / betabetas - 3.1415926 / 8 / betabetas / (1 - alphaalpha) - (3.1415926 * 3.1415926 / 64 / betabetas / betabetas / (1 - alphaalpha) + 3.1415926 / 8 / (1 - alphaalpha) / (1 - alphaalpha) / betabetas) * Log(alphaalpha)
EJ = EJ & "Kr2：" & Kr2 & vbCrLf
lambdaL1 = betabetas * (Kr1 + Kr2) / (3.1415926 / 8 / betabetas + (1 + alphaalpha) / 2) / (3.1415926 / 8 / betabetas + (1 + alphaalpha) / 2)
EJ = EJ & "lambdaL1：" & lambdaL1 & vbCrLf
lambdas1 = KU1 * lambdau1 + KL1 * lambdaL1
EJ = EJ & "定子槽比漏磁导：" & lambdas1 & vbCrLf
'99
Xs1 = 2 * p * m * L1 * lambdas1 * CX / Lef / Q1 / Kdp / Kdp
EJ = EJ & "定子槽漏抗：" & Xs1 & vbCrLf
'100
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''图4-6
sigmasigmas = 0.0066
Xd1 = m * tao1 * sigmasigmas * CX / 3.1415926 / 3.1415926 / Kdelta / delta / Kst / Kdp / Kdp
EJ = EJ & "定子谐波漏抗：" & Xd1 & vbCrLf
'101
If DAN.Value And QT.Value Then
XE1 = 0.2 * LE * CX / Lef / Kdp / Kdp
ElseIf SHUANG.Value And QT.Value Then
XE1 = 1.2 * (d + 0.5 * fd) * CX / Lef
ElseIf DAN.Value And JCS.Value Then
XE1 = 0.47 * LE - 0.64 * taoy * CX / Lef / Kdp / Kdp
ElseIf DAN.Value And TXS.Value Then                       ''''''''''''''''''适用于分组的同心式
XE1 = 0.47 * LE - 0.64 * taoy * CX / Lef / Kdp / Kdp
ElseIf DAN.Value And TXS.Value And FZD.Value Then                       ''''''''''''''''''适用于分组的同心式
XE1 = 0.67 * LE - 0.64 * taoy * CX / Lef / Kdp / Kdp
End If
EJ = EJ & "定子端部漏抗：" & XE1 & vbCrLf
'102
Xsk = 0.5 * (tsk / T1) * (tsk / T1) * Xd1
EJ = EJ & "定子斜槽漏抗：" & Xsk & vbCrLf
'103
X1 = Xs1 + Xd1 + XE1 + Xsk
EJ = EJ & "定子漏抗：" & X1 & vbCrLf
'104
lambdaU2 = h02 / b02
EJ = EJ & "lambdau2：" & lambdaU2 & vbCrLf
betabetar = hr2 / br2
EJ = EJ & "betabetar：" & betabetar & vbCrLf
alphaalphar = br1 / br2
EJ = EJ & "alphaalphar：" & alphaalphar & vbCrLf
Kr1 = 1 / 3 - (1 - alphaalphar) / 4 * (1 / 4 + 1 / 3 / (1 - alphaalphar) + 1 / 2 / (1 - alphaalphar) / (1 - alphaalphar) + 1 / (1 - alphaalphar) / (1 - alphaalphar) / (1 - alphaalphar) + Log(alphaalphar) / (1 - alphaalphar) / (1 - alphaalphar) / (1 - alphaalphar) / (1 - alphaalphar))
EJ = EJ & "Kr1：" & Kr1 & vbCrLf
lambdaL2 = 2 * hr1 / (b02 + br1) + 4 * betabetar * Kr1 / (1 + alphaalphar) / (1 + alphaalphar)
EJ = EJ & "lambdaL2：" & lambdaL2 & vbCrLf
lambdas2 = lambdaL2 + lambdaU2
EJ = EJ & "转子槽比漏磁导：" & lambdas2 & vbCrLf
'105
Xs2 = 2 * m * p * L2 * lambdas2 * CX / Lef / Q2
EJ = EJ & "转子槽漏抗：" & Xs2 & vbCrLf
'106
sigmasigmaR = 3.1415926 * 3.1415926 * (2 * p / Q2) * (2 * p / Q2) / 12
Xd2 = m * tao1 * sigmasigmaR * CX / 3.1415926 / 3.1415926 / Kdelta / delta / Kst
EJ = EJ & "转子谐波漏抗：" & Xd2 & vbCrLf
'107
XE2 = 0.757 * ((LB - L2) / 1.13 + DR / 2 / p) * CX / Lef
EJ = EJ & "转子端部漏抗：" & XE2 & vbCrLf
'108
X2 = Xs2 + Xd2 + XE2
EJ = EJ & "转子漏抗：" & X2 & vbCrLf
'109
Kad = 1 / KF
EJ = EJ & "直轴电枢磁动势折算系数：" & Kad & vbCrLf
'110
Idd = NI * 0.5
EJ = EJ & "Id：" & Idd & vbCrLf
Fad1 = 0.45 * m * Kad * Kdp * N * Idd / p
EJ = EJ & "Fad1：" & Fad1 & vbCrLf
If CLJGb.Value Or CLJGd.Value Or CLJGc.Value Or CLJGa.Value Then
fap = Fad1 / (sigma0 * hM * Hc * 10)
ElseIf BLJGa.Value Or BLJGb.Value Or BLJGc.Value Then
fap = 2 * Fad1 / (sigma0 * hM * Hc * 10)
End If
EJ = EJ & "fap：" & fap & vbCrLf
bmN = lambdan * (1 - fap) / (lambdan + 1)
EJ = EJ & "bmN：" & bmN & vbCrLf
phideltaN = (bmN - (1 - bmN) * lambdasigma) * Am * Br * 0.0001
EJ = EJ & "phideltaN：" & phideltaN & vbCrLf
Ed = 4.44 * f * Kdp * N * phideltaN * Kphi
EJ = EJ & "Ed：" & Ed & vbCrLf
xad = Abs(E0 - Ed) / Idd
EJ = EJ & "直轴电枢反应电抗：" & xad & vbCrLf
'111
xd = xad + X1
EJ = EJ & "直轴同步电抗：" & xd & vbCrLf
'112
xaq = xad * (1 + lambdadelta / (1 + lambdasigma))
EJ = EJ & "交轴电枢反应电抗：" & xaq & vbCrLf
'113
xq = xaq + X1
EJ = EJ & "交轴同步电抗：" & xq & vbCrLf
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''交轴磁化曲线
'phiaq = 0.35 * phidelat0
'Do While (phiaq < 0.85 * phidelat0)
'phiaq = phiaq + 1
'Loop

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''目前由于二元插值问题无法解决
'114
If FHS.Value Then
If p = 1 Then
PFW = 5.5 * (3 / p) * (3 / p) * D2 * D2 * D2 * 0.001
Else
PFW = 6.5 * (3 / p) * (3 / p) * D2 * D2 * D2 * 0.001
End If
ElseIf FNXZSLS.Value Then
If p = 1 Then
PFW = 13 * (1 - D1 * 0.01) * (3 / p) * (3 / p) * D1 * D1 * D1 * D1 * 0.00001
Else
PFW = (3 / p) * (3 / p) * D1 * D1 * D1 * D1 * 0.0001
End If
Else
MsgBox ("请选择防护式或者封闭型自扇冷式")
Exit Sub
End If
EJ = EJ & "机械损耗：" & PFW & vbCrLf
'115
theta = 56.5 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''注意
'116
P1 = m * (E0 * UN * (xq * Sin(theta * 3.1415926 / 180) - rr1 * Cos(theta * 3.1415926 / 180)) + rr1 * UN * UN + 0.5 * UN * UN * (xd - xq) * Sin(2 * theta * 3.1415926 / 180)) / (xd * xq + rr1 * rr1)
EJ = EJ & "输入功率：" & P1 & vbCrLf
'117
Id = (rr1 * UN * Sin(theta * 3.1415926 / 180) + xq * (E0 - UN * Cos(theta * 3.1415926 / 180))) / (xd * xq + rr1 * rr1)
EJ = EJ & "直轴电流：" & Id & vbCrLf
'118
Iq = (xd * UN * Sin(theta * 3.1415926 / 180) - rr1 * (E0 - UN * Cos(theta * 3.1415926 / 180))) / (xd * xq + rr1 * rr1)
EJ = EJ & "交轴电流：" & Iq & vbCrLf
'119
psi = Atn(Id / Iq)
phiPHI = theta * 3.1415926 / 180 - psi
GLYS = Cos(phiPHI)
EJ = EJ & "功率因数：" & GLYS & vbCrLf
'120
I1 = Sqr(Id * Id + Iq * Iq)
EJ = EJ & "定子电流：" & I1 & vbCrLf
'121
PCU = m * I1 * I1 * rr1
EJ = EJ & "定子电阻损耗：" & PCU & vbCrLf
'122
Edelta = Sqr((E0 - Id * xad) * (E0 - Id * xad) + Iq * Iq * xaq * xaq)
EJ = EJ & "Edelta：" & Edelta & vbCrLf
phidelta = Edelta / (4.44 * f * Kdp * N * Kphi)
EJ = EJ & "负载气隙磁通：" & phidelta & vbCrLf
'123
Bdeltad = phidelta * 10000 / tao1 / Lef / alphai
EJ = EJ & "负载气隙磁密：" & Bdeltad & vbCrLf
'124
Bt1d = Bdeltad * T1 * Lef / bt1 / KFe / L1
EJ = EJ & "负载定子齿磁密：" & Bt1d & vbCrLf
'125
Bj1d = phidelta * 10000 / 2 / L1 / KFe / hj1

EJ = EJ & "负载定子轭磁密：" & Bj1d & vbCrLf
'126
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
pt1d = SHQXf(Bt1d) * 0.001
'pt1d = 0.0429
EJ = EJ & "pt1d ：" & pt1d & vbCrLf
pj1d = SHQXf(Bj1d) * 0.001
'pj1d = 0.0351
EJ = EJ & "pj1d ：" & pj1d & vbCrLf
PFE = 2.5 * pt1d * Vt1 + 2 * pj1d * VJ1
EJ = EJ & "铁耗：" & PFE & vbCrLf

'127
PSNX = Val(txtPSNX.Text)
'PS = (I1 / I1) * (I1 / I1) * PN * 1000 * PSNX
PS = (I1 / NI) * (I1 / NI) * PN * 1000 * PSNX
EJ = EJ & "杂散损耗：" & PS & vbCrLf
'128
sigmaP = PCU + PFE + PFW + PS
EJ = EJ & "总损耗：" & sigmaP & vbCrLf
'129
P2 = P1 - sigmaP

EJ = EJ & "输出功率：" & P2 & vbCrLf
'130
eat = P2 / P1
EJ = EJ & "效率：" & eat & vbCrLf

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'132
ZJBL1 = 2 * m * UN * UN * (1 / xq - 1 / xd)
ZJBL2 = m * UN * E0 / xd
ZJBL3 = m * UN * UN * (1 / xd - 1 / xq)
ZJBL4 = ZJBL2 ^ 2 - 4 * ZJBL1 * ZJBL3

If ZJBL4 < 0 Then
MsgBox ("该方程无实数根")
Exit Sub
Else
ZJBL5 = (-ZJBL2 + Sqr(ZJBL2 ^ 2 - 4 * ZJBL1 * ZJBL3)) / (2 * ZJBL1)
ZJBL6 = (-ZJBL2 - Sqr(ZJBL2 ^ 2 - 4 * ZJBL1 * ZJBL3)) / (2 * ZJBL1)
End If
If ZJBL5 >= -1 And ZJBL5 <= 1 Then
ZJBL55 = Sqr(1 - ZJBL5 * ZJBL5)
End If
If ZJBL6 >= -1 And ZJBL6 <= 1 Then
ZJBL66 = Sqr(1 - ZJBL6 * ZJBL6)
End If

ZJBL555 = ZJBL55 / ZJBL5
ZJBL666 = ZJBL66 / ZJBL6
thetatheta1 = Atn(ZJBL555) * 180 / 3.1415926 + 180
thetatheta2 = Atn(ZJBL666) * 180 / 3.1415926
EJ = EJ & "thetatheta1：" & thetatheta1 & vbCrLf
EJ = EJ & "thetatheta2：" & thetatheta2 & vbCrLf
If thetatheta1 > 0 And thetatheta1 < 180 Then
thetatheta = thetatheta1
Else
thetatheta = thetatheta2
End If
EJ = EJ & "最大功率出现在功角为：" & thetatheta & vbCrLf
Pemmax = m * UN * E0 * Sin(thetatheta * 3.1415926 / 180) / xd + m / 2 * UN * UN * (1 / xq - 1 / xd) * Sin(2 * thetatheta * 3.1415926 / 180)
EJ = EJ & "Pemmax ：" & Pemmax & vbCrLf
TP0X = Pemmax / PN / 1000
EJ = EJ & "失步转矩倍数：" & TP0X & vbCrLf
'133
IdN = Id
If CLJGb.Value Or CLJGd.Value Or CLJGc.Value Or CLJGa.Value Then
faNp = 0.45 * m * Kad * Kdp * N * IdN / (p * sigma0 * Hc * hM * 10)
ElseIf BLJGa.Value Or BLJGb.Value Or BLJGc.Value Then
faNp = 0.9 * m * Kad * Kdp * N * IdN / (p * sigma0 * Hc * hM * 10)
End If
EJ = EJ & "faNp：" & faNp & vbCrLf
bmN = lambdan * (1 - faNp) / (lambdan + 1)
EJ = EJ & "永磁体额定负载工作点：" & bmN & vbCrLf
'134
a = 2 * m * N * I1 / 3.1415926 / Di1
EJ = EJ & "电负荷：" & a & vbCrLf
'135
J1 = I1 / (a1 * 3.1415926 * (Nt1 * (d11 / 2) * (d11 / 2) + Nt2 * (d12 / 2) * (d12 / 2)))
EJ = EJ & "电流密度：" & J1 & vbCrLf
'136
AJ1 = a * J1
EJ = EJ & "热负荷：" & AJ1 & vbCrLf
'137
Iadh = (E0 * xd + Sqr(E0 * E0 * xd * xd - (rr1 * rr1 + xd * xd) * (E0 * E0 - UN * UN))) / (rr1 * rr1 + xd * xd)
EJ = EJ & "Iadh：" & Iadh & vbCrLf
If CLJGb.Value Or CLJGd.Value Or CLJGc.Value Or CLJGa.Value Then
fadhp = 0.45 * m * Kad * Kdp * N * Iadh / p / sigma0 / Hc / hM / 10
ElseIf BLJGa.Value Or BLJGb.Value Or BLJGc.Value Then
fadhp = 0.9 * m * Kad * Kdp * N * Iadh / p / sigma0 / Hc / hM / 10
End If
EJ = EJ & "fadhp：" & fadhp & vbCrLf
bmh = lambdan * (1 - fadhp) / (lambdan + 1)
EJ = EJ & "永磁体最大去磁工作点：" & bmh & vbCrLf
cycle3:
'138
istp = 460
'''''''''''''''''''''''''''''''''''''''''图4-15
Ks = 0.6
betac = 0.64 + 2.5 * Sqr(delta / (T1 + t2))
EJ = EJ & "betac：" & betac & vbCrLf
Fst = 0.707 * istp * Ns * (KU1 + Kd1 * Kd1 * Kp1 * Q1 / Q2) * E0 / UN / a1
EJ = EJ & "Fst：" & Fst & vbCrLf
BL = 4 * 3.1415926 * 0.0000001 * Fst / (2 * betac * delta * 0.01)
EJ = EJ & "漏抗饱和系数：" & BL & vbCrLf
'140
Cs1 = (T1 - b01) * (1 - Ks)
EJ = EJ & "齿顶漏磁饱和引起定子齿顶宽度的减小：" & Cs1 & vbCrLf
'141
CS2 = (t2 - b02) * (1 - Ks)
EJ = EJ & "齿顶漏磁饱和引起转子齿顶宽度的减小：" & CS2 & vbCrLf
'142
deltalambdaU1 = (h01 + 0.58 * h1) * (Cs1 / (Cs1 + 1.5 * b01)) / b01
EJ = EJ & "deltalambdaU1：" & deltalambdaU1 & vbCrLf
lambdas1st = KU1 * (lambdau1 - deltalambdaU1) + KL1 * lambdaL1
EJ = EJ & "起动时定子槽比磁导：" & lambdas1st & vbCrLf
'143
Xs1st = lambdas1st * Xs1 / lambdas1
EJ = EJ & "起动时定子槽漏抗：" & Xs1st & vbCrLf
'144
Xd1st = Ks * Xd1
EJ = EJ & "起动时定子谐波漏抗：" & Xd1st & vbCrLf
'145
Xskst = Ks * Xsk
EJ = EJ & "起动时定子斜槽波漏抗：" & Xskst & vbCrLf
'146
x1st = Xskst + Xd1st + Xs1st + XE1
EJ = EJ & "起动时定子漏抗：" & x1st & vbCrLf
'147
If ZLZZ.Value Then
hhbb = hr12
BBBS = 1
ElseIf TTZZ.Value Then
hhbb = hr12 + h02
BBBS = 0.9
End If
xi = 2 * 3.1415926 * hhbb * Sqr(BBBS * f / (4.34 * 0.0001 * 10000000))
EJ = EJ & "考虑挤流效应的转子导条相对高度：" & xi & vbCrLf
'148
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''图4-17
Ka = 1
sh2xi = (Exp(2 * xi) - Exp(-2 * xi)) / 2
ch2xi = (Exp(2 * xi) + Exp(-2 * xi)) / 2
PHIXI = xi * (sh2xi + Sin(2 * xi)) / (ch2xi - Cos(2 * xi))
EJ = EJ & " phixi：" & PHIXI & vbCrLf
hpr = hhbb * Ka / PHIXI
EJ = EJ & " 导条电阻等效高度：" & hpr & vbCrLf
'149
psixi = 3 * (sh2xi - Sin(2 * xi)) / (ch2xi - Cos(2 * xi)) / 2 / xi
EJ = EJ & " psixi：" & psixi & vbCrLf
hpX = hhbb * psixi * Ka
EJ = EJ & "槽漏抗等效高度：" & hpX & vbCrLf
'150
KRr = (1 + alphaalphar) * PHIXI * PHIXI / (1 + alphaalphar * ((2 * PHIXI) - 1))
EJ = EJ & "起动转子电阻增大系数：" & KRr & vbCrLf
'151

Kr1 = 1 / 3 - (1 - alphaalphar) / 4 * (1 / 4 + 1 / 3 / (1 - alphaalphar) + 1 / 2 / (1 - alphaalphar) / (1 - alphaalphar) + 1 / (1 - alphaalphar) / (1 - alphaalphar) / (1 - alphaalphar) + Log(alphaalphar) / (1 - alphaalphar) / (1 - alphaalphar) / (1 - alphaalphar) / (1 - alphaalphar))
EJ = EJ & "Kr1：" & Kr1 & vbCrLf
bpx = br1 + (br2 - br1) * psixi
EJ = EJ & "bpx：" & bpx & vbCrLf
alphaalpharp = br1 / bpx
EJ = EJ & "alphaalpharp：" & alphaalpharp & vbCrLf
Kr1p = 1 / 3 - (1 - alphaalpharp) / 4 * (1 / 4 + 1 / 3 / (1 - alphaalpharp) + 1 / 2 / (1 - alphaalpharp) / (1 - alphaalpharp) + 1 / (1 - alphaalpharp) / (1 - alphaalpharp) / (1 - alphaalpharp) + Log(alphaalpharp) / (1 - alphaalpharp) / (1 - alphaalpharp) / (1 - alphaalpharp) / (1 - alphaalpharp))
EJ = EJ & "Kr1p：" & Kr1p & vbCrLf
Kx = br2 * (1 + alphaalphar) * (1 + alphaalphar) * psixi * Kr1p / bpx / (1 + alphaalpharp) / (1 + alphaalphar) / Kr1
EJ = EJ & "起动时的转子漏抗减小系数：" & Kx & vbCrLf
'152
lambdal2st = Kx * lambdaL2
EJ = EJ & "起动时转子槽下部漏磁导：" & lambdal2st & vbCrLf
'153
lambdaU2 = h02 / b02
DELTAlambdau2 = h02 * (CS2 / (CS2 + b02)) / b02
lambdau2st = lambdaU2 - DELTAlambdau2
lambdas2st = lambdau2st + lambdal2st
EJ = EJ & "起动时的转子槽比漏磁导：" & lambdas2st & vbCrLf
'154
XS2ST = lambdas2st * Xs2 / lambdas2
EJ = EJ & "起动时的转子槽漏抗：" & XS2ST & vbCrLf
'155
Xd2st = Ks * Xd2
EJ = EJ & "起动时转子谐波漏抗：" & Xd2st & vbCrLf
'156
x2st = XS2ST + Xd2st + XE2
EJ = EJ & "转子起动漏抗：" & x2st & vbCrLf
'157
xst = x1st + x2st
EJ = EJ & "起动总漏抗：" & xst & vbCrLf
'158
r2st = (KRr * L2 / LB + (LB - L2) / LB) * RB + RR
EJ = EJ & "转子起动电阻：" & r2st & vbCrLf
'159
RST = rr1 + r2st
EJ = EJ & "起动时总电阻：" & RST & vbCrLf
'160
ZST = Sqr(RST * RST + xst * xst)
EJ = EJ & "起动时总阻抗：" & ZST & vbCrLf
'161
ist = UN / ZST
EJ = EJ & "起动电流：" & ist & vbCrLf

If (ist - istp) / ist > 0.003 Then
istp = ist - (ist - istp) / 3
GoTo cycle3
Else
GoTo line3
End If
line3:

'162
istx = ist / NI
EJ = EJ & "起动电流倍数：" & istx & vbCrLf





'163
s = 1
x2sts = (x2st - X2) * Sqr(s) + X2
x1sts = (x1st - X1) * Sqr(s) + X1
r2sts = (r2st - rr2) * Sqr(s) + rr2
xm = (2 * xad * xaq) / (xad + xaq)
c1s = 1 + xsts / xm
tC = m * p * UN * UN * r2sts / s / 2 / 3.1415926 / f / ((rr1 + c1s * r2sts / s) * (rr1 + c1s * r2sts / s) + (x1sts + c1s * x2sts) * (x1sts + c1s * x2sts))
tCX = tC / TN
EJ = EJ & "异步起动转矩倍数：" & tCX & vbCrLf
'164
Tg = -m * p * E0 * E0 * rr1 * (1 - s) * (rr1 * rr1 + (1 - s) * (1 - s) * xq * xq) / 2 / 3.1415926 / f / (rr1 * rr1 + (1 - s) * (1 - s) * xd * xq) * (rr1 * rr1 + (1 - s) * (1 - s) * xd * xq)
Tgx = Tg / TN
EJ = EJ & "永磁体发电制动转矩倍数：" & Tgx & vbCrLf

'165
Tav = tC + Tg
EJ = EJ & "合成起动转矩曲线：" & Tav & vbCrLf
'166
TSTX = Tav / TN
EJ = EJ & "起动转矩倍数：" & TSTX & vbCrLf

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''T-S特性曲线



'Dim yxyx%
's = 1
'For yxyx = 1 To 10
'EJ = EJ & "电机转差率：" & s & vbCrLf
'x2sts = (x2st - X2) * Sqr(s) + X2
'x1sts = (x1st - X1) * Sqr(s) + X1
'r2sts = (r2st - rr2) * Sqr(s) + rr2
'xm = (2 * xad * xaq) / (xad + xaq)
'c1s = 1 + xsts / xm
'tC = m * p * UN * UN * r2sts / s / 2 / 3.1415926 / f / ((rr1 + c1s * r2sts / s) * (rr1 + c1s * r2sts / s) + (x1sts + c1s * x2sts) * (x1sts + c1s * x2sts))
'tCX = tC / TN
'EJ = EJ & "异步起动转矩倍数：" & tCX & vbCrLf
''164
'Tg = m * p * E0 * E0 * rr1 * (1 - s) * (rr1 * rr1 + (1 - s) * (1 - s) * xq * xq) / 2 / 3.1415926 / f / (rr1 * rr1 + (1 - s) * (1 - s) * xd * xq) / (rr1 * rr1 + (1 - s) * (1 - s) * xd * xq)
'Tgx = Tg / TN
'EJ = EJ & "永磁体发电制动转矩倍数：" & Tgx & vbCrLf
'
''165
'Tav = tC - Tg
''EJ = EJ & "合成起动转矩曲线：" & Tav & vbCrLf
''166
'TSTX = Tav / TN
'EJ = EJ & "起动转矩倍数：" & TSTX & vbCrLf
's = s - 0.1
'Next yxyx




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''电动机的工作特性
'Dim xyxy%
'theta = 20
'For xyxy = 1 To 15
'
'EJ = EJ & "theta：" & theta & vbCrLf
''116
'P1 = m * (E0 * UN * (Xq * Sin(theta * 3.1415926 / 180) - RR1 * Cos(theta * 3.1415926 / 180)) + RR1 * UN * UN + 0.5 * UN * UN * (Xd - Xq) * Sin(2 * theta * 3.1415926 / 180)) / (Xd * Xq + RR1 * RR1)
'EJ = EJ & "输入功率：" & P1 & vbCrLf
''117
'Id = (RR1 * UN * Sin(theta * 3.1415926 / 180) + Xq * (E0 - UN * Cos(theta * 3.1415926 / 180))) / (Xd * Xq + RR1 * RR1)
'EJ = EJ & "直轴电流：" & Id & vbCrLf
''118
'Iq = (Xd * UN * Sin(theta * 3.1415926 / 180) - RR1 * (E0 - UN * Cos(theta * 3.1415926 / 180))) / (Xd * Xq + RR1 * RR1)
'EJ = EJ & "交轴电流：" & Iq & vbCrLf
''119
'psi = Atn(Id / Iq)
'phiPHI = theta * 3.1415926 / 180 - psi
'GLYS = Cos(phiPHI)
'EJ = EJ & "功率因数：" & GLYS & vbCrLf
''120
'I1 = Sqr(Id * Id + Iq * Iq)
'EJ = EJ & "定子电流：" & I1 & vbCrLf
''121
'PCU = m * I1 * I1 * RR1
''EJ = EJ & "定子电阻损耗：" & PCU & vbCrLf
''122
'Edelta = Sqr((E0 - Id * Xad) * (E0 - Id * Xad) + Iq * Iq * Xaq * Xaq)
''EJ = EJ & "Edelta：" & Edelta & vbCrLf
'phidelta = Edelta / (4.44 * f * Kdp * N * Kphi)
''EJ = EJ & "负载气隙磁通：" & phidelta & vbCrLf
''123
'Bdeltad = phidelta * 10000 / tao1 / Lef / alphai
''EJ = EJ & "负载气隙磁密：" & Bdeltad & vbCrLf
''124
'Bt1d = Bdeltad * t1 * Lef / bt1 / KFe / L1
''EJ = EJ & "负载定子齿磁密：" & Bt1d & vbCrLf
''125
'Bj1d = phidelta * 10000 / 2 / L1 / KFe / hj1
'
''EJ = EJ & "负载定子轭磁密：" & Bj1d & vbCrLf
''126
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'pt1d = SHQXf(Bt1d) * 0.001
''pt1d = 0.0429
''EJ = EJ & "pt1d ：" & pt1d & vbCrLf
'pj1d = SHQXf(Bj1d) * 0.001
''pj1d = 0.0351
''EJ = EJ & "pj1d ：" & pj1d & vbCrLf
'PFE = 2.5 * pt1d * Vt1 + 2 * pj1d * VJ1
''EJ = EJ & "铁耗：" & PFE & vbCrLf
'
''127
'PSNX = Val(txtPSNX.Text)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''PS = (I1 / I1) * (I1 / I1) * PN * 1000 * PSNX
'PS = (I1 / NI) * (I1 / NI) * PN * 1000 * PSNX
''EJ = EJ & "杂散损耗：" & PS & vbCrLf
''128
'sigmaP = PCU + PFE + PFW + PS
''EJ = EJ & "总损耗：" & sigmaP & vbCrLf
''129
'P2 = P1 - sigmaP
'
'EJ = EJ & "输出功率：" & P2 & vbCrLf
''130
'eat = P2 / P1
'EJ = EJ & "效率：" & eat & vbCrLf
'theta = theta + 5
'Next xyxy






Text1.Text = EJ
End Sub






Private Sub CommandTC_Click()
End
End Sub





