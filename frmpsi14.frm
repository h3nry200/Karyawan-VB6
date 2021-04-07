VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmpsi14 
   Caption         =   "Form1"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox tanggal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   3
      Top             =   810
      Width           =   3195
   End
   Begin VB.TextBox nama 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   2
      Top             =   390
      Width           =   3195
   End
   Begin VB.TextBox id 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1170
      TabIndex        =   1
      Top             =   0
      Width           =   3195
   End
   Begin VB.CommandButton nextbtn 
      Caption         =   "NEXT >>>>>"
      DisabledPicture =   "frmpsi14.frx":0000
      DownPicture     =   "frmpsi14.frx":13EA
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   18240
      Picture         =   "frmpsi14.frx":27D4
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8415
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   20055
      _ExtentX        =   35375
      _ExtentY        =   14843
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "SOAL NOMOR 1 - 10"
      TabPicture(0)   =   "frmpsi14.frx":3BBE
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame9"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame6"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame1"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "SOAL NOMOR 11 - 20"
      TabPicture(1)   =   "frmpsi14.frx":3BDA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame20"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame19"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame18"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame17"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Frame16"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame15"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Frame14"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Frame13"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Frame12"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Frame11"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "SOAL NOMOR 21 - 30"
      TabPicture(2)   =   "frmpsi14.frx":3BF6
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Frame21"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame22"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame23"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Frame24"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Frame25"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Frame26"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Frame27"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "Frame28"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Frame29"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "Frame30"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      Begin VB.Frame Frame30 
         Height          =   735
         Left            =   120
         TabIndex        =   191
         Top             =   6960
         Width           =   19815
         Begin VB.TextBox Text30 
            Height          =   495
            Left            =   18360
            TabIndex        =   194
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b30 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5640
            TabIndex        =   193
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton a30 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   192
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label73 
            Caption         =   "30. Douglas Mark Arthur"
            Height          =   255
            Left            =   120
            TabIndex        =   196
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label72 
            Caption         =   "Douglas Mark Arthurr"
            Height          =   255
            Left            =   6960
            TabIndex        =   195
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame29 
         Height          =   735
         Left            =   120
         TabIndex        =   185
         Top             =   6240
         Width           =   19815
         Begin VB.TextBox Text29 
            Height          =   495
            Left            =   18360
            TabIndex        =   188
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b29 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5640
            TabIndex        =   187
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a29 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   186
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label71 
            Caption         =   "29. General Insurance"
            Height          =   255
            Left            =   120
            TabIndex        =   190
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label70 
            Caption         =   "General Assurance"
            Height          =   255
            Left            =   6960
            TabIndex        =   189
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame28 
         Height          =   735
         Left            =   120
         TabIndex        =   179
         Top             =   5520
         Width           =   19815
         Begin VB.TextBox Text28 
            Height          =   495
            Left            =   18360
            TabIndex        =   182
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b28 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5640
            TabIndex        =   181
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a28 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   180
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label69 
            Caption         =   "28. ttfo: 45 jurgenpark kopen hagen 66"
            Height          =   255
            Left            =   120
            TabIndex        =   184
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label68 
            Caption         =   "ttfo: 45 jurgenpark kopen hogen 66"
            Height          =   255
            Left            =   6960
            TabIndex        =   183
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame27 
         Height          =   735
         Left            =   120
         TabIndex        =   173
         Top             =   4800
         Width           =   19815
         Begin VB.TextBox Text27 
            Height          =   495
            Left            =   18360
            TabIndex        =   176
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b27 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5640
            TabIndex        =   175
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton a27 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   174
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label66 
            Caption         =   "27. bunga_padi@wasantara.net.id"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   178
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label65 
            Caption         =   "bungapadi@wasantara.net.id"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6960
            TabIndex        =   177
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame26 
         Height          =   735
         Left            =   120
         TabIndex        =   167
         Top             =   4080
         Width           =   19815
         Begin VB.TextBox Text26 
            Height          =   495
            Left            =   18360
            TabIndex        =   170
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b26 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5640
            TabIndex        =   169
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton a26 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   168
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label64 
            Caption         =   "26. R47867543279"
            Height          =   255
            Left            =   120
            TabIndex        =   172
            Top             =   240
            Width           =   2655
         End
         Begin VB.Label Label63 
            Caption         =   "R47867543279"
            Height          =   255
            Left            =   6960
            TabIndex        =   171
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame25 
         Height          =   735
         Left            =   120
         TabIndex        =   161
         Top             =   3360
         Width           =   19815
         Begin VB.TextBox Text25 
            Height          =   495
            Left            =   18360
            TabIndex        =   164
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b25 
            Caption         =   "(TS)"
            Height          =   555
            Left            =   5640
            TabIndex        =   163
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton a25 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   162
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label62 
            Caption         =   "25. Plaza Mall, Senin, 5/9 2003, 21:35"
            Height          =   255
            Left            =   120
            TabIndex        =   166
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label61 
            Caption         =   "Plaza Mall, Senin, 5/6 2003, 21.25"
            Height          =   255
            Left            =   6960
            TabIndex        =   165
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame24 
         Height          =   735
         Left            =   120
         TabIndex        =   155
         Top             =   2640
         Width           =   19815
         Begin VB.TextBox Text24 
            Height          =   495
            Left            =   18360
            TabIndex        =   158
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b24 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5640
            TabIndex        =   157
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton a24 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   156
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label60 
            Caption         =   "24. 12,6 + 87,24 + 36% X 21"
            Height          =   255
            Left            =   120
            TabIndex        =   160
            Top             =   240
            Width           =   2775
         End
         Begin VB.Label Label59 
            Caption         =   "12,6 + 87,24 + 36% X 21"
            Height          =   255
            Left            =   6960
            TabIndex        =   159
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame23 
         Height          =   735
         Left            =   120
         TabIndex        =   149
         Top             =   1920
         Width           =   19815
         Begin VB.TextBox Text23 
            Height          =   495
            Left            =   18360
            TabIndex        =   152
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b23 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5640
            TabIndex        =   151
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton a23 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   150
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label58 
            Caption         =   "23. Jl. Jend. Anumerta A. Yani 87 C"
            Height          =   255
            Left            =   120
            TabIndex        =   154
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label Label57 
            Caption         =   "Jl. Jend Anumerta A Yani 87 C"
            Height          =   255
            Left            =   6960
            TabIndex        =   153
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame22 
         Height          =   735
         Left            =   120
         TabIndex        =   143
         Top             =   1200
         Width           =   19815
         Begin VB.TextBox Text22 
            Height          =   495
            Left            =   18360
            TabIndex        =   146
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b22 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5640
            TabIndex        =   145
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton a22 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   144
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label56 
            Caption         =   "22. Defisit 50 US $"
            Height          =   255
            Left            =   120
            TabIndex        =   148
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label55 
            Caption         =   "Profit 50 US $"
            Height          =   255
            Left            =   6960
            TabIndex        =   147
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame21 
         Height          =   735
         Left            =   120
         TabIndex        =   137
         Top             =   480
         Width           =   19815
         Begin VB.TextBox Text21 
            Height          =   495
            Left            =   18360
            TabIndex        =   140
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b21 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5640
            TabIndex        =   139
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton a21 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   138
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label54 
            Caption         =   "21. 12,6% per tahun : 2 = bunga"
            Height          =   255
            Left            =   120
            TabIndex        =   142
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label53 
            Caption         =   "12,6% per tahun : 2 = bunga"
            Height          =   255
            Left            =   6960
            TabIndex        =   141
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -74880
         TabIndex        =   100
         Top             =   660
         Width           =   19815
         Begin VB.OptionButton a1 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   103
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton b1 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5400
            TabIndex        =   102
            Top             =   120
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   495
            Left            =   18600
            TabIndex        =   101
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label29 
            Caption         =   "Drs.Jemblung Sudjatmiko,MM"
            Height          =   375
            Left            =   6480
            TabIndex        =   116
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label Label3 
            Caption         =   "1. Drs.Jemblung Sudjatmiko,MM"
            Height          =   375
            Left            =   120
            TabIndex        =   104
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   -74880
         TabIndex        =   95
         Top             =   1380
         Width           =   19815
         Begin VB.OptionButton a2 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   98
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton b2 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5400
            TabIndex        =   97
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox Text2 
            Height          =   495
            Left            =   18600
            TabIndex        =   96
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label30 
            Caption         =   "Prof.Edi Dwi Efendi, SH, LLM"
            Height          =   375
            Left            =   6480
            TabIndex        =   117
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label Label10 
            Caption         =   "2.  Prof.Edi Dwi Effendi, SH, LLM"
            Height          =   255
            Left            =   120
            TabIndex        =   99
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   -74880
         TabIndex        =   90
         Top             =   2100
         Width           =   19815
         Begin VB.OptionButton a3 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   93
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton b3 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5400
            TabIndex        =   92
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox Text3 
            Height          =   495
            Left            =   18600
            TabIndex        =   91
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label33 
            Caption         =   "I Ketut Sugiarso, M.Hum."
            Height          =   375
            Left            =   6480
            TabIndex        =   118
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label Label11 
            Caption         =   "3.  I Ketut Sugiarso, M.Hum"
            Height          =   375
            Left            =   120
            TabIndex        =   94
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   -74880
         TabIndex        =   85
         Top             =   2820
         Width           =   19815
         Begin VB.OptionButton a4 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   88
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton b4 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5400
            TabIndex        =   87
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text4 
            Height          =   525
            Left            =   18600
            TabIndex        =   86
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label34 
            Caption         =   "PT Bank Panindo Sejahtera"
            Height          =   375
            Left            =   6480
            TabIndex        =   119
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label Label12 
            Caption         =   "4.   PT Bank Panin Sejahtera"
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Frame Frame5 
         Height          =   735
         Left            =   -74880
         TabIndex        =   80
         Top             =   3540
         Width           =   19815
         Begin VB.OptionButton a5 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   83
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton b5 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5400
            TabIndex        =   82
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text5 
            Height          =   495
            Left            =   18600
            TabIndex        =   81
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label35 
            Caption         =   "Capital Mal 55"
            Height          =   375
            Left            =   6480
            TabIndex        =   120
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label Label13 
            Caption         =   "5.  Capital Mall 555"
            Height          =   255
            Left            =   120
            TabIndex        =   84
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame Frame6 
         Height          =   735
         Left            =   -74880
         TabIndex        =   75
         Top             =   4260
         Width           =   19815
         Begin VB.OptionButton a6 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   78
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton b6 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5400
            TabIndex        =   77
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox Text6 
            Height          =   495
            Left            =   18600
            TabIndex        =   76
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label36 
            Caption         =   "1185 Ampera Ridge"
            Height          =   375
            Left            =   6480
            TabIndex        =   121
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label Label14 
            Caption         =   "6.  1185 Ampera Bridge"
            Height          =   255
            Left            =   120
            TabIndex        =   79
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame7 
         Height          =   735
         Left            =   -74880
         TabIndex        =   70
         Top             =   4980
         Width           =   19815
         Begin VB.OptionButton a7 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   73
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton b7 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5400
            TabIndex        =   72
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox Text7 
            Height          =   495
            Left            =   18600
            TabIndex        =   71
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label37 
            Caption         =   "08156726874"
            Height          =   375
            Left            =   6480
            TabIndex        =   122
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label Label15 
            Caption         =   "7.  08156728874"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame8 
         Height          =   735
         Left            =   -74880
         TabIndex        =   65
         Top             =   5700
         Width           =   19815
         Begin VB.OptionButton a8 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   68
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton b8 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5400
            TabIndex        =   67
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text8 
            Height          =   495
            Left            =   18600
            TabIndex        =   66
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label39 
            Caption         =   "Madisson Square ' 288 USA"
            Height          =   375
            Left            =   6480
            TabIndex        =   123
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label Label16 
            Caption         =   "8. Madisson Square ' 228 USA"
            Height          =   255
            Left            =   120
            TabIndex        =   69
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame9 
         Height          =   735
         Left            =   -74880
         TabIndex        =   60
         Top             =   6420
         Width           =   19815
         Begin VB.OptionButton a9 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   63
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton b9 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5400
            TabIndex        =   62
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox Text9 
            Height          =   495
            Left            =   18600
            TabIndex        =   61
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label40 
            Caption         =   "Dr. Whani Darmawan, SE, Akt"
            Height          =   375
            Left            =   6480
            TabIndex        =   124
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label Label17 
            Caption         =   "9.  Dr. Whani Sudarmawan, SE, Akt"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame10 
         Height          =   735
         Left            =   -74880
         TabIndex        =   55
         Top             =   7140
         Width           =   19815
         Begin VB.OptionButton a10 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   58
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton b10 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5400
            TabIndex        =   57
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text10 
            Height          =   495
            Left            =   18600
            TabIndex        =   56
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label41 
            Caption         =   "Jhon M. Echols and Hassan Sadhily"
            Height          =   375
            Left            =   6480
            TabIndex        =   125
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label Label18 
            Caption         =   "10. John M.Echols and Hassan Sadhily"
            Height          =   375
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame Frame11 
         Height          =   735
         Left            =   -74880
         TabIndex        =   50
         Top             =   660
         Width           =   19815
         Begin VB.OptionButton a11 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   53
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton b11 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5640
            TabIndex        =   52
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox Text11 
            Height          =   495
            Left            =   18360
            TabIndex        =   51
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label42 
            Caption         =   "Cornell University Press"
            Height          =   255
            Left            =   6960
            TabIndex        =   126
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label19 
            Caption         =   "11.  Cornell University Press"
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame12 
         Height          =   735
         Left            =   -74880
         TabIndex        =   45
         Top             =   1380
         Width           =   19815
         Begin VB.OptionButton a12 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   48
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton b12 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5640
            TabIndex        =   47
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox Text12 
            Height          =   495
            Left            =   18360
            TabIndex        =   46
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label43 
            Caption         =   "PT Barumun Indojaya, ltd"
            Height          =   255
            Left            =   6960
            TabIndex        =   127
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label20 
            Caption         =   "12. PT Barumun Indojaya, Ltd"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame13 
         Height          =   735
         Left            =   -74880
         TabIndex        =   40
         Top             =   2100
         Width           =   19815
         Begin VB.OptionButton a13 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   43
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton b13 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5640
            TabIndex        =   42
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox Text13 
            Height          =   495
            Left            =   18360
            TabIndex        =   41
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label44 
            Caption         =   "Shouteast Asia Program"
            Height          =   255
            Left            =   6960
            TabIndex        =   128
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label21 
            Caption         =   "13.   Shoutes Asia Program"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame14 
         Height          =   735
         Left            =   -74880
         TabIndex        =   35
         Top             =   2820
         Width           =   19815
         Begin VB.OptionButton a14 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   38
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton b14 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5640
            TabIndex        =   37
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox Text14 
            Height          =   495
            Left            =   18360
            TabIndex        =   36
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label45 
            Caption         =   "Beksan Langen Mandra Wanara"
            Height          =   255
            Left            =   6960
            TabIndex        =   129
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label22 
            Caption         =   "14.   Beksan langen Mandra Wanara"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame Frame15 
         Height          =   735
         Left            =   -74880
         TabIndex        =   30
         Top             =   3540
         Width           =   19815
         Begin VB.OptionButton a15 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   33
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton b15 
            Caption         =   "(TS)"
            Height          =   555
            Left            =   5640
            TabIndex        =   32
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox Text15 
            Height          =   495
            Left            =   18360
            TabIndex        =   31
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label46 
            Caption         =   "Jl. Gathot Subroto I Kav 5 No. 97"
            Height          =   255
            Left            =   6960
            TabIndex        =   130
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label23 
            Caption         =   "15.   Jl. Gathot Subroto I Kav 5 No. 67"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame Frame16 
         Height          =   735
         Left            =   -74880
         TabIndex        =   25
         Top             =   4260
         Width           =   19815
         Begin VB.OptionButton a16 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   28
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton b16 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5640
            TabIndex        =   27
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox Text16 
            Height          =   495
            Left            =   18360
            TabIndex        =   26
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label47 
            Caption         =   "CV Inti Mas Makmur Abadi"
            Height          =   255
            Left            =   6960
            TabIndex        =   131
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label24 
            Caption         =   "16.  CV Inti Mas Makmur Abadi"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame17 
         Height          =   735
         Left            =   -74880
         TabIndex        =   20
         Top             =   4980
         Width           =   19815
         Begin VB.OptionButton a17 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   23
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton b17 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5640
            TabIndex        =   22
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox Text17 
            Height          =   495
            Left            =   18360
            TabIndex        =   21
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label48 
            Caption         =   "Curriculum Vitae"
            Height          =   255
            Left            =   6960
            TabIndex        =   132
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label25 
            Caption         =   "17.  Curriculum Vitae"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame18 
         Height          =   735
         Left            =   -74880
         TabIndex        =   15
         Top             =   5700
         Width           =   19815
         Begin VB.OptionButton a18 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   18
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton b18 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5640
            TabIndex        =   17
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text18 
            Height          =   495
            Left            =   18360
            TabIndex        =   16
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label50 
            Caption         =   "Mount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   8880
            TabIndex        =   134
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label49 
            Caption         =   "The inspiration from Merapi"
            Height          =   255
            Left            =   6960
            TabIndex        =   133
            Top             =   240
            Width           =   1935
         End
         Begin VB.Label Label26 
            Caption         =   "18.  The inspiration from Merapi Mount"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame19 
         Height          =   735
         Left            =   -74880
         TabIndex        =   10
         Top             =   6420
         Width           =   19815
         Begin VB.OptionButton a19 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   13
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton b19 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5640
            TabIndex        =   12
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text19 
            Height          =   495
            Left            =   18360
            TabIndex        =   11
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label51 
            Caption         =   "Kloter 10 BPS No. Kursi 0021"
            Height          =   255
            Left            =   6960
            TabIndex        =   135
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label27 
            Caption         =   "19.  Kloter 10 BSP No. Kursi 0021"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame Frame20 
         Height          =   735
         Left            =   -74880
         TabIndex        =   5
         Top             =   7140
         Width           =   19815
         Begin VB.OptionButton a20 
            Caption         =   "(S)"
            Height          =   495
            Left            =   3480
            TabIndex        =   8
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton b20 
            Caption         =   "(TS)"
            Height          =   495
            Left            =   5640
            TabIndex        =   7
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox Text20 
            Height          =   495
            Left            =   18360
            TabIndex        =   6
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label52 
            Caption         =   "Sardana W.Kusuma"
            Height          =   255
            Left            =   6960
            TabIndex        =   136
            Top             =   240
            Width           =   3015
         End
         Begin VB.Label Label28 
            Caption         =   "20. Sardono W.Kusumo"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   2775
         End
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Waktu 10 menit Soal 20 Nomor."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   115
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "Petunjuk Khusus :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   114
      Top             =   1200
      Width           =   2085
   End
   Begin VB.Label Label6 
      Caption         =   "ID                 :"
      Height          =   285
      Left            =   30
      TabIndex        =   113
      Top             =   0
      Width           =   1125
   End
   Begin VB.Label Label7 
      Caption         =   "NAMA           :"
      Height          =   165
      Left            =   0
      TabIndex        =   112
      Top             =   420
      Width           =   1245
   End
   Begin VB.Label Label8 
      Caption         =   "Tanggal Test :"
      Height          =   195
      Left            =   0
      TabIndex        =   111
      Top             =   870
      Width           =   1305
   End
   Begin VB.Label Label5 
      Caption         =   $"frmpsi14.frx":3C12
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   90
      TabIndex        =   110
      Top             =   1650
      Width           =   19755
   End
   Begin VB.Label Label2 
      Caption         =   "Test Ketelitian"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9600
      TabIndex        =   109
      Top             =   720
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "TEST XIV"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   9150
      TabIndex        =   108
      Top             =   120
      Width           =   2355
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   20250
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Label Label31 
      Caption         =   "12.   3, 7, 15, 31, 63, ..., ..."
      Height          =   255
      Left            =   600
      TabIndex        =   107
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Label Label32 
      Caption         =   "12.   3, 7, 15, 31, 63, ..., ..."
      Height          =   255
      Left            =   600
      TabIndex        =   106
      Top             =   5400
      Width           =   3015
   End
   Begin VB.Label Label38 
      Caption         =   "12.   3, 7, 15, 31, 63, ..., ..."
      Height          =   255
      Left            =   600
      TabIndex        =   105
      Top             =   6120
      Width           =   3015
   End
End
Attribute VB_Name = "frmpsi14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim koneksi As New rdoConnection
Dim rQuery As New rdoQuery
Dim rs As rdoResultset

Private Sub Form_Load()
koneksi.Connect = "Driver={MySQL ODBC 5.1 Driver};SERVER=192.168.2.11;PWD=123456;UID=BRIWIRA;PORT=3306;DATABASE=karyawan;"
koneksi.EstablishConnection
End Sub

Private Sub a1_Click()
Text1.Text = "S"
End Sub

Private Sub b1_Click()
Text1.Text = "TS"
End Sub

Private Sub a2_Click()
Text2.Text = "S"
End Sub

Private Sub b2_Click()
Text2.Text = "TS"
End Sub

Private Sub a3_Click()
Text3.Text = "S"
End Sub

Private Sub b3_Click()
Text3.Text = "TS"
End Sub

Private Sub a4_Click()
Text4.Text = "S"
End Sub

Private Sub b4_Click()
Text4.Text = "TS"
End Sub

Private Sub a5_Click()
Text5.Text = "S"
End Sub

Private Sub b5_Click()
Text5.Text = "TS"
End Sub

Private Sub a6_Click()
Text6.Text = "S"
End Sub

Private Sub b6_Click()
Text6.Text = "TS"
End Sub

Private Sub a7_Click()
Text7.Text = "S"
End Sub

Private Sub b7_Click()
Text7.Text = "TS"
End Sub

Private Sub a8_Click()
Text8.Text = "S"
End Sub
Private Sub b8_Click()
Text8.Text = "TS"
End Sub

Private Sub a9_Click()
Text9.Text = "S"
End Sub

Private Sub b9_Click()
Text9.Text = "TS"
End Sub

Private Sub a10_Click()
Text10.Text = "S"
End Sub

Private Sub b10_Click()
Text10.Text = "TS"
End Sub

Private Sub a11_Click()
Text11.Text = "S"
End Sub

Private Sub b11_Click()
Text11.Text = "TS"
End Sub

Private Sub a12_Click()
Text12.Text = "S"
End Sub

Private Sub b12_Click()
Text12.Text = "TS"
End Sub

Private Sub a13_Click()
Text13.Text = "S"
End Sub

Private Sub b13_Click()
Text13.Text = "TS"
End Sub

Private Sub a14_Click()
Text14.Text = "S"
End Sub

Private Sub b14_Click()
Text14.Text = "TS"
End Sub

Private Sub a15_Click()
Text15.Text = "S"
End Sub

Private Sub b15_Click()
Text15.Text = "TS"
End Sub

Private Sub a16_Click()
Text16.Text = "S"
End Sub

Private Sub b16_Click()
Text16.Text = "TS"
End Sub

Private Sub a17_Click()
Text17.Text = "S"
End Sub

Private Sub b17_Click()
Text17.Text = "TS"
End Sub

Private Sub a18_Click()
Text18.Text = "S"
End Sub

Private Sub b18_Click()
Text18.Text = "TS"
End Sub

Private Sub a19_Click()
Text19.Text = "S"
End Sub

Private Sub b19_Click()
Text19.Text = "TS"
End Sub

Private Sub a20_Click()
Text20.Text = "S"
nextbtn.Enabled = True
End Sub

Private Sub b20_Click()
Text20.Text = "TS"
nextbtn.Enabled = True
End Sub

Private Sub a21_Click()
Text21.Text = "S"
nextbtn.Enabled = True
End Sub

Private Sub b21_Click()
Text21.Text = "TS"
nextbtn.Enabled = True
End Sub

Private Sub a22_Click()
Text22.Text = "S"
nextbtn.Enabled = True
End Sub

Private Sub b22_Click()
Text22.Text = "TS"
nextbtn.Enabled = True
End Sub

Private Sub a23_Click()
Text23.Text = "S"
nextbtn.Enabled = True
End Sub

Private Sub b23_Click()
Text23.Text = "TS"
nextbtn.Enabled = True
End Sub

Private Sub a24_Click()
Text24.Text = "S"
nextbtn.Enabled = True
End Sub

Private Sub b24_Click()
Text24.Text = "TS"
nextbtn.Enabled = True
End Sub

Private Sub a25_Click()
Text25.Text = "S"
nextbtn.Enabled = True
End Sub

Private Sub b25_Click()
Text25.Text = "TS"
nextbtn.Enabled = True
End Sub

Private Sub a26_Click()
Text26.Text = "S"
nextbtn.Enabled = True
End Sub

Private Sub b26_Click()
Text26.Text = "TS"
nextbtn.Enabled = True
End Sub

Private Sub a27_Click()
Text27.Text = "S"
nextbtn.Enabled = True
End Sub

Private Sub b27_Click()
Text27.Text = "TS"
nextbtn.Enabled = True
End Sub

Private Sub a28_Click()
Text28.Text = "S"
nextbtn.Enabled = True
End Sub

Private Sub b28_Click()
Text28.Text = "TS"
nextbtn.Enabled = True
End Sub

Private Sub a29_Click()
Text29.Text = "S"
nextbtn.Enabled = True
End Sub

Private Sub b29_Click()
Text29.Text = "TS"
nextbtn.Enabled = True
End Sub

Private Sub a30_Click()
Text30.Text = "S"
nextbtn.Enabled = True
End Sub

Private Sub b30_Click()
Text30.Text = "TS"
nextbtn.Enabled = True
End Sub


Private Sub nextbtn_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Or Text11.Text = "" Or Text12.Text = "" Or Text13.Text = "" Or Text14.Text = "" Or Text15.Text = "" Or Text16.Text = "" Or Text17.Text = "" Or Text18.Text = "" Or Text19.Text = "" Or Text20.Text = "" Or Text21.Text = "" Or Text22.Text = "" Or Text23.Text = "" Or Text24.Text = "" Or Text25.Text = "" Or Text26.Text = "" Or Text27.Text = "" Or Text28.Text = "" Or Text29.Text = "" Or Text30.Text = "" Then
    MsgBox "Data Belum Lengkap"
    Else
    Dim tambahdata As String
        tambahdata = "Insert Into hasiltest14 values ('" & id.Text & "','" & nama.Text & "','" & tanggal.Text & "','" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "','" & Text7.Text & "','" & Text8.Text & "','" & Text9.Text & "','" & Text10.Text & "','" & Text11.Text & "','" & Text12.Text & "','" & Text13.Text & "','" & Text14.Text & "','" & Text15.Text & "','" & Text16.Text & "','" & Text17.Text & "','" & Text18.Text & "','" & Text19.Text & "','" & Text20.Text & "','" & Text21.Text & "','" & Text22.Text & "','" & Text23.Text & "','" & Text24.Text & "','" & Text25.Text & "','" & Text26.Text & "','" & Text27.Text & "','" & Text28.Text & "','" & Text29.Text & "','" & Text30.Text & "')"
        koneksi.Execute tambahdata
        MsgBox "TEST 14 Berhasil Di simpan, Silakan lanjutkan ke TEST 15", vbInformation, "Pemberitahuan"
    frmpsi15.Show
    koneksi.Close
    Unload Me
    End If

End Sub

Private Sub Option12_Click()

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

If SSTab1.Caption = "SOAL NOMOR 21 - 30" Then
nextbtn.Enabled = True
Else
nextbtn.Enabled = False
End If
End Sub






