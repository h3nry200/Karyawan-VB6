VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmpsi10 
   Caption         =   "Form1"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton nextbtn 
      Caption         =   "NEXT >>>>>"
      DisabledPicture =   "frmpsi10.frx":0000
      DownPicture     =   "frmpsi10.frx":13EA
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
      Left            =   18210
      Picture         =   "frmpsi10.frx":27D4
      TabIndex        =   3
      Top             =   210
      Width           =   1695
   End
   Begin VB.TextBox tanggal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1170
      TabIndex        =   2
      Top             =   810
      Width           =   3195
   End
   Begin VB.TextBox nama 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1170
      TabIndex        =   1
      Top             =   390
      Width           =   3195
   End
   Begin VB.TextBox id 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1170
      TabIndex        =   0
      Top             =   0
      Width           =   3195
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7815
      Left            =   0
      TabIndex        =   4
      Top             =   2640
      Width           =   20055
      _ExtentX        =   35375
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "SOAL NOMOR 1 - 20"
      TabPicture(0)   =   "frmpsi10.frx":3BBE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame13"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame14"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame15"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame16"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame17"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame18"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame19"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame20"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame6"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame7"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame8"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame9"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame10"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Frame21"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Frame22"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Frame23"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Frame24"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Frame25"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "SOAL NOMOR 21 - 40"
      TabPicture(1)   =   "frmpsi10.frx":3BDA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(3)=   "Frame4"
      Tab(1).Control(4)=   "Frame5"
      Tab(1).Control(5)=   "Frame26"
      Tab(1).Control(6)=   "Frame27"
      Tab(1).Control(7)=   "Frame28"
      Tab(1).Control(8)=   "Frame29"
      Tab(1).Control(9)=   "Frame30"
      Tab(1).Control(10)=   "Frame31"
      Tab(1).Control(11)=   "Frame32"
      Tab(1).Control(12)=   "Frame33"
      Tab(1).Control(13)=   "Frame34"
      Tab(1).Control(14)=   "Frame35"
      Tab(1).Control(15)=   "Frame36"
      Tab(1).Control(16)=   "Frame37"
      Tab(1).Control(17)=   "Frame38"
      Tab(1).Control(18)=   "Frame39"
      Tab(1).Control(19)=   "Frame40"
      Tab(1).ControlCount=   20
      Begin VB.Frame Frame40 
         Height          =   735
         Left            =   -75000
         TabIndex        =   211
         Top             =   360
         Width           =   6255
         Begin VB.TextBox Text21 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   214
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b21 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   213
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a21 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   212
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label51 
            Caption         =   "21. Harmoni - sumbang"
            Height          =   495
            Left            =   120
            TabIndex        =   215
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.Frame Frame39 
         Height          =   735
         Left            =   -75000
         TabIndex        =   206
         Top             =   1080
         Width           =   6255
         Begin VB.TextBox Text22 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   209
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b22 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   208
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton a22 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   207
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label50 
            Caption         =   "22. Fertil - subur"
            Height          =   495
            Left            =   120
            TabIndex        =   210
            Top             =   120
            Width           =   2295
         End
      End
      Begin VB.Frame Frame38 
         Height          =   735
         Left            =   -75000
         TabIndex        =   201
         Top             =   1800
         Width           =   6255
         Begin VB.TextBox Text23 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   204
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b23 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   203
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a23 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   202
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label49 
            Caption         =   "23. Apex - zenit"
            Height          =   495
            Left            =   120
            TabIndex        =   205
            Top             =   120
            Width           =   2415
         End
      End
      Begin VB.Frame Frame37 
         Height          =   735
         Left            =   -75000
         TabIndex        =   196
         Top             =   2520
         Width           =   6255
         Begin VB.TextBox Text24 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   199
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b24 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   198
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a24 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   197
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label48 
            Caption         =   "24. Donasi - kontribusi"
            Height          =   495
            Left            =   120
            TabIndex        =   200
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.Frame Frame36 
         Height          =   735
         Left            =   -75000
         TabIndex        =   191
         Top             =   3240
         Width           =   6255
         Begin VB.TextBox Text25 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   194
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b25 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   193
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a25 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   192
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label47 
            Caption         =   "25. Inklusif - terbuka"
            Height          =   375
            Left            =   120
            TabIndex        =   195
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame Frame35 
         Height          =   735
         Left            =   -75000
         TabIndex        =   186
         Top             =   3960
         Width           =   6255
         Begin VB.TextBox Text26 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   189
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b26 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   188
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a26 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   187
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label46 
            Caption         =   "26. Kulit - isi"
            Height          =   375
            Left            =   120
            TabIndex        =   190
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame34 
         Height          =   735
         Left            =   -75000
         TabIndex        =   181
         Top             =   4680
         Width           =   6255
         Begin VB.OptionButton b27 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   184
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a27 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   183
            Top             =   120
            Width           =   1455
         End
         Begin VB.TextBox Text27 
            Height          =   495
            Left            =   5400
            TabIndex        =   182
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label45 
            Caption         =   "27. Animo - kekuatan"
            Height          =   375
            Left            =   120
            TabIndex        =   185
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame33 
         Height          =   735
         Left            =   -75000
         TabIndex        =   176
         Top             =   5400
         Width           =   6255
         Begin VB.TextBox Text28 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   179
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b28 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   178
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a28 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   177
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label44 
            Caption         =   "28. Mukzizat - karamah"
            Height          =   375
            Left            =   120
            TabIndex        =   180
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame32 
         Height          =   735
         Left            =   -75000
         TabIndex        =   171
         Top             =   6120
         Width           =   6255
         Begin VB.TextBox Text29 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   174
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b29 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   173
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton a29 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   172
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label43 
            Caption         =   "29. Langit - bumi"
            Height          =   375
            Left            =   120
            TabIndex        =   175
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame31 
         Height          =   735
         Left            =   -75000
         TabIndex        =   166
         Top             =   6840
         Width           =   6255
         Begin VB.TextBox Text30 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   169
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b30 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   168
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a30 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   167
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label42 
            Caption         =   "30. Vertikal - horizontal"
            Height          =   375
            Left            =   120
            TabIndex        =   170
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame Frame30 
         Height          =   735
         Left            =   -62640
         TabIndex        =   161
         Top             =   360
         Width           =   6255
         Begin VB.OptionButton a31 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   164
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton b31 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   163
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text31 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   162
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label41 
            Caption         =   "31. Gaji - honor"
            Height          =   495
            Left            =   120
            TabIndex        =   165
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.Frame Frame29 
         Height          =   735
         Left            =   -62640
         TabIndex        =   156
         Top             =   1080
         Width           =   6255
         Begin VB.OptionButton a32 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   159
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton b32 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   158
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox Text32 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   157
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label40 
            Caption         =   "32. Abdi - majikan"
            Height          =   495
            Left            =   120
            TabIndex        =   160
            Top             =   120
            Width           =   2295
         End
      End
      Begin VB.Frame Frame28 
         Height          =   735
         Left            =   -62640
         TabIndex        =   151
         Top             =   1800
         Width           =   6255
         Begin VB.OptionButton a33 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   154
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton b33 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   153
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text33 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   152
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label39 
            Caption         =   "33. Abadi - fana"
            Height          =   495
            Left            =   120
            TabIndex        =   155
            Top             =   120
            Width           =   2415
         End
      End
      Begin VB.Frame Frame27 
         Height          =   735
         Left            =   -62640
         TabIndex        =   146
         Top             =   2520
         Width           =   6255
         Begin VB.OptionButton a34 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   149
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton b34 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   148
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text34 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   147
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label38 
            Caption         =   "34. Tangkal - cegah"
            Height          =   495
            Left            =   120
            TabIndex        =   150
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.Frame Frame26 
         Height          =   735
         Left            =   -62640
         TabIndex        =   141
         Top             =   3240
         Width           =   6255
         Begin VB.OptionButton a35 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   144
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton b35 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   143
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text35 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   142
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label37 
            Caption         =   "35. Fakta - opini"
            Height          =   375
            Left            =   120
            TabIndex        =   145
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame Frame5 
         Height          =   735
         Left            =   -62640
         TabIndex        =   136
         Top             =   3960
         Width           =   6255
         Begin VB.OptionButton a36 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   139
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton b36 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   138
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text36 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   137
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label24 
            Caption         =   "36. Apatis - aktif"
            Height          =   375
            Left            =   120
            TabIndex        =   140
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   -62640
         TabIndex        =   131
         Top             =   4680
         Width           =   6255
         Begin VB.TextBox Text37 
            Height          =   495
            Left            =   5400
            TabIndex        =   134
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton a37 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   133
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton b37 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   132
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label23 
            Caption         =   "37. Kolusi - korupsi"
            Height          =   375
            Left            =   120
            TabIndex        =   135
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   -62640
         TabIndex        =   126
         Top             =   5400
         Width           =   6255
         Begin VB.OptionButton a38 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   129
            Top             =   120
            Width           =   1335
         End
         Begin VB.OptionButton b38 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   128
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text38 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   127
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label22 
            Caption         =   "38. Ide - gagasan"
            Height          =   375
            Left            =   120
            TabIndex        =   130
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   -62640
         TabIndex        =   121
         Top             =   6120
         Width           =   6255
         Begin VB.OptionButton a39 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   124
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton b39 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   123
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox Text39 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   122
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label21 
            Caption         =   "39. Kampiun - juara"
            Height          =   375
            Left            =   120
            TabIndex        =   125
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -62640
         TabIndex        =   116
         Top             =   6840
         Width           =   6255
         Begin VB.OptionButton a40 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   119
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton b40 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   118
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text40 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   117
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label20 
            Caption         =   "40. Khusus - spesifik"
            Height          =   375
            Left            =   120
            TabIndex        =   120
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame Frame25 
         Height          =   735
         Left            =   12480
         TabIndex        =   111
         Top             =   6840
         Width           =   6255
         Begin VB.TextBox Text20 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   114
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b20 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   113
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a20 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   112
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label36 
            Caption         =   "20. Pra - pasca"
            Height          =   375
            Left            =   120
            TabIndex        =   115
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame Frame24 
         Height          =   735
         Left            =   12480
         TabIndex        =   106
         Top             =   6120
         Width           =   6255
         Begin VB.TextBox Text19 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   109
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b19 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   108
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton a19 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   107
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label35 
            Caption         =   "19. Depresi - resesi"
            Height          =   375
            Left            =   120
            TabIndex        =   110
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame23 
         Height          =   735
         Left            =   12480
         TabIndex        =   101
         Top             =   5400
         Width           =   6255
         Begin VB.TextBox Text18 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   104
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b18 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   103
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a18 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   102
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label34 
            Caption         =   "18. Hukuman - denda"
            Height          =   375
            Left            =   120
            TabIndex        =   105
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame22 
         Height          =   735
         Left            =   12480
         TabIndex        =   96
         Top             =   4680
         Width           =   6255
         Begin VB.OptionButton b17 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   99
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a17 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   98
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox Text17 
            Height          =   495
            Left            =   5400
            TabIndex        =   97
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label33 
            Caption         =   "17. Meluas - menciut"
            Height          =   375
            Left            =   120
            TabIndex        =   100
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame21 
         Height          =   735
         Left            =   12480
         TabIndex        =   91
         Top             =   3960
         Width           =   6255
         Begin VB.TextBox Text16 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   94
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b16 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   93
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a16 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   92
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label32 
            Caption         =   "16. Izin - biarkan"
            Height          =   375
            Left            =   120
            TabIndex        =   95
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame10 
         Height          =   735
         Left            =   12480
         TabIndex        =   86
         Top             =   3240
         Width           =   6255
         Begin VB.TextBox Text15 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   89
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b15 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   88
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a15 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   87
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label29 
            Caption         =   "15. Bukit - lembah"
            Height          =   375
            Left            =   120
            TabIndex        =   90
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame Frame9 
         Height          =   735
         Left            =   12480
         TabIndex        =   81
         Top             =   2520
         Width           =   6255
         Begin VB.TextBox Text14 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   84
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b14 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   83
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a14 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   82
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label28 
            Caption         =   "14. Gemuk - kurus"
            Height          =   495
            Left            =   120
            TabIndex        =   85
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.Frame Frame8 
         Height          =   735
         Left            =   12480
         TabIndex        =   76
         Top             =   1800
         Width           =   6255
         Begin VB.TextBox Text13 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   79
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b13 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   78
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a13 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   77
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label27 
            Caption         =   "13. Eksklusif - tertentu"
            Height          =   495
            Left            =   120
            TabIndex        =   80
            Top             =   120
            Width           =   2415
         End
      End
      Begin VB.Frame Frame7 
         Height          =   735
         Left            =   12480
         TabIndex        =   71
         Top             =   1080
         Width           =   6255
         Begin VB.TextBox Text12 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   74
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b12 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   73
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton a12 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   72
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label26 
            Caption         =   "12. Tinggi - rendah"
            Height          =   495
            Left            =   120
            TabIndex        =   75
            Top             =   120
            Width           =   2295
         End
      End
      Begin VB.Frame Frame6 
         Height          =   735
         Left            =   12480
         TabIndex        =   66
         Top             =   360
         Width           =   6255
         Begin VB.TextBox Text11 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   69
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b11 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   68
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a11 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   67
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label25 
            Caption         =   "11. Licik - cerdik"
            Height          =   495
            Left            =   120
            TabIndex        =   70
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.Frame Frame20 
         Height          =   735
         Left            =   120
         TabIndex        =   50
         Top             =   6840
         Width           =   6255
         Begin VB.OptionButton a10 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   53
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton b10 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   52
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text10 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   51
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label19 
            Caption         =   "10. Jonggar - tegas"
            Height          =   375
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame Frame19 
         Height          =   735
         Left            =   120
         TabIndex        =   45
         Top             =   6120
         Width           =   6255
         Begin VB.OptionButton a9 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   48
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton b9 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   47
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox Text9 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   46
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label18 
            Caption         =   "9. Bonus - diskon"
            Height          =   375
            Left            =   120
            TabIndex        =   49
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame18 
         Height          =   735
         Left            =   120
         TabIndex        =   40
         Top             =   5400
         Width           =   6255
         Begin VB.OptionButton a8 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   43
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton b8 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   42
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text8 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   41
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label17 
            Caption         =   "8.Gasal - ganjil"
            Height          =   375
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame17 
         Height          =   735
         Left            =   120
         TabIndex        =   35
         Top             =   4680
         Width           =   6255
         Begin VB.TextBox Text7 
            Height          =   495
            Left            =   5400
            TabIndex        =   38
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton a7 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   37
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton b7 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   36
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label31 
            Caption         =   "7. Feminin - maskulin"
            Height          =   375
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame16 
         Height          =   735
         Left            =   120
         TabIndex        =   30
         Top             =   3960
         Width           =   6255
         Begin VB.OptionButton a6 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   33
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton b6 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   32
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text6 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   31
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "6. Ilegal - sah"
            Height          =   375
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame15 
         Height          =   735
         Left            =   120
         TabIndex        =   25
         Top             =   3240
         Width           =   6255
         Begin VB.OptionButton a5 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   28
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton b5 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   27
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text5 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   26
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label14 
            Caption         =   "5. Elastis - kaku"
            Height          =   375
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Width           =   2415
         End
      End
      Begin VB.Frame Frame14 
         Height          =   735
         Left            =   120
         TabIndex        =   20
         Top             =   2520
         Width           =   6255
         Begin VB.OptionButton a4 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   23
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton b4 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   22
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text4 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   21
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label13 
            Caption         =   "4.Primordia - sectarian"
            Height          =   495
            Left            =   120
            TabIndex        =   24
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.Frame Frame13 
         Height          =   735
         Left            =   120
         TabIndex        =   15
         Top             =   1800
         Width           =   6255
         Begin VB.OptionButton a3 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   18
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton b3 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   17
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text3 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   16
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label12 
            Caption         =   "3. Urban - kota"
            Height          =   495
            Left            =   120
            TabIndex        =   19
            Top             =   120
            Width           =   2415
         End
      End
      Begin VB.Frame Frame12 
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Top             =   1080
         Width           =   6255
         Begin VB.OptionButton a2 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   13
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton b2 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   12
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   11
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label11 
            Caption         =   "2. Pelopor - pewaris"
            Height          =   495
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Width           =   2295
         End
      End
      Begin VB.Frame Frame11 
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   6255
         Begin VB.OptionButton a1 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   8
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton b1 
            Caption         =   "(A)"
            Height          =   495
            Left            =   4080
            TabIndex        =   7
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   6
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label10 
            Caption         =   "1.Kelompok - group"
            Height          =   495
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.Label Label16 
         Caption         =   "7. Anjing adalah binatang yang berguna karena?"
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
         Left            =   3360
         TabIndex        =   55
         Top             =   -2160
         Width           =   4575
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Contoh : Besar - Kecil (S) - (A)                        Jawaban : (A)"
      Height          =   255
      Left            =   120
      TabIndex        =   65
      Top             =   2040
      Width           =   4455
   End
   Begin VB.Label Label9 
      Caption         =   "Waktu 8 menit Soal 20 Nomor."
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
      TabIndex        =   64
      Top             =   2400
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
      TabIndex        =   63
      Top             =   1200
      Width           =   2085
   End
   Begin VB.Label Label6 
      Caption         =   "ID                 :"
      Height          =   285
      Left            =   30
      TabIndex        =   62
      Top             =   0
      Width           =   1125
   End
   Begin VB.Label Label7 
      Caption         =   "NAMA           :"
      Height          =   165
      Left            =   0
      TabIndex        =   61
      Top             =   420
      Width           =   1245
   End
   Begin VB.Label Label8 
      Caption         =   "Tanggal Test :"
      Height          =   195
      Left            =   0
      TabIndex        =   60
      Top             =   870
      Width           =   1305
   End
   Begin VB.Label Label30 
      Caption         =   "8. Mengapa remaja sebaiknya diajarkan pelajaran seksual?"
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
      Left            =   450
      TabIndex        =   59
      Top             =   7410
      Width           =   4575
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   20250
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Label Label5 
      Caption         =   $"frmpsi10.frx":3BF6
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   58
      Top             =   1650
      Width           =   19995
   End
   Begin VB.Label Label2 
      Caption         =   "Test Posisi kata Sinonim - Antonim"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   57
      Top             =   720
      Width           =   3405
   End
   Begin VB.Label Label1 
      Caption         =   "TEST X"
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
      TabIndex        =   56
      Top             =   120
      Width           =   1875
   End
End
Attribute VB_Name = "frmpsi10"
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
Text1.Text = "A"
End Sub

Private Sub a2_Click()
Text2.Text = "S"
End Sub

Private Sub b2_Click()
Text2.Text = "A"
End Sub

Private Sub a3_Click()
Text3.Text = "S"
End Sub

Private Sub b3_Click()
Text3.Text = "A"
End Sub

Private Sub a4_Click()
Text4.Text = "S"
End Sub

Private Sub b4_Click()
Text4.Text = "A"
End Sub

Private Sub a5_Click()
Text5.Text = "S"
End Sub

Private Sub b5_Click()
Text5.Text = "A"
End Sub

Private Sub a6_Click()
Text6.Text = "S"
End Sub

Private Sub b6_Click()
Text6.Text = "A"
End Sub

Private Sub a7_Click()
Text7.Text = "S"
End Sub

Private Sub b7_Click()
Text7.Text = "A"
End Sub

Private Sub a8_Click()
Text8.Text = "S"
End Sub
Private Sub b8_Click()
Text8.Text = "A"
End Sub

Private Sub a9_Click()
Text9.Text = "S"
End Sub

Private Sub b9_Click()
Text9.Text = "A"
End Sub

Private Sub a10_Click()
Text10.Text = "S"
End Sub

Private Sub b10_Click()
Text10.Text = "A"
End Sub

Private Sub a11_Click()
Text11.Text = "S"
End Sub

Private Sub b11_Click()
Text11.Text = "A"
End Sub

Private Sub a12_Click()
Text12.Text = "S"
End Sub

Private Sub b12_Click()
Text12.Text = "A"
End Sub

Private Sub a13_Click()
Text13.Text = "S"
End Sub

Private Sub b13_Click()
Text13.Text = "A"
End Sub

Private Sub a14_Click()
Text14.Text = "S"
End Sub

Private Sub b14_Click()
Text14.Text = "A"
End Sub

Private Sub a15_Click()
Text15.Text = "S"
End Sub

Private Sub b15_Click()
Text15.Text = "A"
End Sub

Private Sub a16_Click()
Text16.Text = "S"
End Sub

Private Sub b16_Click()
Text16.Text = "A"
End Sub

Private Sub a17_Click()
Text17.Text = "S"
End Sub

Private Sub b17_Click()
Text17.Text = "A"
End Sub

Private Sub a18_Click()
Text18.Text = "S"
End Sub

Private Sub b18_Click()
Text18.Text = "A"
End Sub

Private Sub a19_Click()
Text19.Text = "S"
End Sub

Private Sub b19_Click()
Text19.Text = "A"
End Sub

Private Sub a20_Click()
Text20.Text = "S"
End Sub

Private Sub b20_Click()
Text20.Text = "A"
End Sub

Private Sub a21_Click()
Text21.Text = "S"
End Sub

Private Sub b21_Click()
Text21.Text = "A"
End Sub

Private Sub a22_Click()
Text22.Text = "S"
End Sub

Private Sub b22_Click()
Text22.Text = "A"
End Sub

Private Sub a23_Click()
Text23.Text = "S"
End Sub

Private Sub b23_Click()
Text23.Text = "A"
End Sub

Private Sub a24_Click()
Text24.Text = "S"
End Sub

Private Sub b24_Click()
Text24.Text = "A"
End Sub

Private Sub a25_Click()
Text25.Text = "S"
End Sub

Private Sub b25_Click()
Text25.Text = "A"
End Sub

Private Sub a26_Click()
Text26.Text = "S"
End Sub

Private Sub b26_Click()
Text26.Text = "A"
End Sub

Private Sub a27_Click()
Text27.Text = "S"
End Sub

Private Sub b27_Click()
Text27.Text = "A"
End Sub

Private Sub a28_Click()
Text28.Text = "S"
End Sub

Private Sub b28_Click()
Text28.Text = "A"
End Sub

Private Sub a29_Click()
Text29.Text = "S"
End Sub

Private Sub b29_Click()
Text29.Text = "A"
End Sub

Private Sub a30_Click()
Text30.Text = "S"
End Sub

Private Sub b30_Click()
Text30.Text = "A"
End Sub

Private Sub a31_Click()
Text31.Text = "S"
End Sub

Private Sub b31_Click()
Text31.Text = "A"
End Sub

Private Sub a32_Click()
Text32.Text = "S"
End Sub

Private Sub b32_Click()
Text32.Text = "A"
End Sub

Private Sub a33_Click()
Text33.Text = "S"
End Sub

Private Sub b33_Click()
Text33.Text = "A"
End Sub

Private Sub a34_Click()
Text34.Text = "S"
End Sub

Private Sub b34_Click()
Text34.Text = "A"
End Sub

Private Sub a35_Click()
Text35.Text = "S"
End Sub

Private Sub b35_Click()
Text35.Text = "A"
End Sub

Private Sub a36_Click()
Text36.Text = "S"
End Sub

Private Sub b36_Click()
Text36.Text = "A"
End Sub

Private Sub a37_Click()
Text37.Text = "S"
End Sub

Private Sub b37_Click()
Text37.Text = "A"
End Sub

Private Sub a38_Click()
Text38.Text = "S"
End Sub

Private Sub b38_Click()
Text38.Text = "A"
End Sub

Private Sub a39_Click()
Text39.Text = "S"
End Sub

Private Sub b39_Click()
Text39.Text = "A"
End Sub

Private Sub a40_Click()
Text40.Text = "S"
End Sub

Private Sub b40_Click()
Text40.Text = "A"
End Sub

Private Sub nextbtn_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Or Text11.Text = "" Or Text12.Text = "" Or Text13.Text = "" Or Text14.Text = "" Or Text15.Text = "" Or Text16.Text = "" Or Text17.Text = "" Or Text18.Text = "" Or Text19.Text = "" Or Text20.Text = "" Or Text21.Text = "" Or Text22.Text = "" Or Text23.Text = "" Or Text24.Text = "" Or Text25.Text = "" Or Text26.Text = "" Or Text27.Text = "" Or Text28.Text = "" Or Text29.Text = "" Or Text30.Text = "" Or Text31.Text = "" Or Text32.Text = "" Or Text33.Text = "" Or Text34.Text = "" Or Text35.Text = "" Or Text36.Text = "" Or Text37.Text = "" Or Text38.Text = "" Or Text39.Text = "" Or Text40.Text = "" Then
    MsgBox "Data Belum Lengkap"
    Else
    Dim tambahdata As String
        tambahdata = "Insert Into hasiltest10 values ('" & id.Text & "','" & nama.Text & "','" & tanggal.Text & "','" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "','" & Text7.Text & "','" & Text8.Text & "','" & Text9.Text & "','" & Text10.Text & "','" & Text11.Text & "','" & Text12.Text & "','" & Text13.Text & "','" & Text14.Text & "','" & Text15.Text & "','" & Text16.Text & "','" & Text17.Text & "','" & Text18.Text & "','" & Text19.Text & "','" & Text20.Text & "','" & Text21.Text & "','" & Text22.Text & "','" & Text23.Text & "','" & Text24.Text & "','" & Text25.Text & "','" & Text26.Text & "','" & Text27.Text & "','" & Text28.Text & "','" & Text29.Text & "','" & Text30.Text & "','" & Text31.Text & "','" & Text32.Text & "','" & Text33.Text & "','" & Text34.Text & "','" & Text35.Text & "','" & Text36.Text & "','" & Text37.Text & "','" & Text38.Text & "','" & Text39.Text & "','" & Text40.Text & "')"
        koneksi.Execute tambahdata
        MsgBox "TEST 10 Berhasil Di simpan, Silakan lanjutkan ke TEST 11", vbInformation, "Pemberitahuan"
    frmpsi11.Show
    koneksi.Close
    Unload Me
    End If

End Sub

Private Sub Option12_Click()

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

If SSTab1.Caption = "SOAL NOMOR 21 - 40" Then
nextbtn.Enabled = True
Else
nextbtn.Enabled = False
End If
End Sub



