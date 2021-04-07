VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmpsi7 
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
      DisabledPicture =   "frmpsi7.frx":0000
      DownPicture     =   "frmpsi7.frx":13EA
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
      Picture         =   "frmpsi7.frx":27D4
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox id 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1170
      TabIndex        =   2
      Top             =   0
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
   Begin VB.TextBox tanggal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1170
      TabIndex        =   0
      Top             =   810
      Width           =   3195
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   120
      TabIndex        =   4
      Top             =   2520
      Width           =   20055
      _ExtentX        =   35375
      _ExtentY        =   14631
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "SOAL NOMOR 1 - 10"
      TabPicture(0)   =   "frmpsi7.frx":3BBE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame10"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "SOAL NOMOR 11 - 20"
      TabPicture(1)   =   "frmpsi7.frx":3BDA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame11"
      Tab(1).Control(1)=   "Frame12"
      Tab(1).Control(2)=   "Frame13"
      Tab(1).Control(3)=   "Frame14"
      Tab(1).Control(4)=   "Frame15"
      Tab(1).Control(5)=   "Frame16"
      Tab(1).Control(6)=   "Frame17"
      Tab(1).Control(7)=   "Frame18"
      Tab(1).Control(8)=   "Frame19"
      Tab(1).Control(9)=   "Frame20"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "SOAL NOMOR 21 - 25"
      TabPicture(2)   =   "frmpsi7.frx":3BF6
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame21"
      Tab(2).Control(1)=   "Frame22"
      Tab(2).Control(2)=   "Frame23"
      Tab(2).Control(3)=   "Frame24"
      Tab(2).Control(4)=   "Frame25"
      Tab(2).ControlCount=   5
      Begin VB.Frame Frame25 
         Height          =   735
         Left            =   -74880
         TabIndex        =   197
         Top             =   1200
         Width           =   19815
         Begin VB.TextBox Text22 
            Height          =   495
            Left            =   18240
            TabIndex        =   203
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton e22 
            Caption         =   "Tujuan"
            Height          =   495
            Left            =   12480
            TabIndex        =   202
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton d22 
            Caption         =   "Rencana"
            Height          =   495
            Left            =   10200
            TabIndex        =   201
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton c22 
            Caption         =   "Pelaksanaan"
            Height          =   495
            Left            =   8160
            TabIndex        =   200
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b22 
            Caption         =   "Terlaksana"
            Height          =   495
            Left            =   5880
            TabIndex        =   199
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton a22 
            Caption         =   "Tertata"
            Height          =   495
            Left            =   3720
            TabIndex        =   198
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label Label33 
            Caption         =   "22.   APLIKASI :"
            Height          =   255
            Left            =   240
            TabIndex        =   204
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Frame Frame24 
         Height          =   735
         Left            =   -74880
         TabIndex        =   189
         Top             =   480
         Width           =   19815
         Begin VB.TextBox Text21 
            Height          =   495
            Left            =   18240
            TabIndex        =   195
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton e21 
            Caption         =   "Pembinaan"
            Height          =   495
            Left            =   12480
            TabIndex        =   194
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton d21 
            Caption         =   "Pimpinan"
            Height          =   495
            Left            =   10200
            TabIndex        =   193
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton c21 
            Caption         =   "Pengelolaan"
            Height          =   495
            Left            =   8160
            TabIndex        =   192
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b21 
            Caption         =   "Pemasaran"
            Height          =   495
            Left            =   5880
            TabIndex        =   191
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton a21 
            Caption         =   "Akuntansi"
            Height          =   495
            Left            =   3720
            TabIndex        =   190
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label Label34 
            Caption         =   "21.   MANAJEMEN :"
            Height          =   255
            Left            =   240
            TabIndex        =   196
            Top             =   240
            Width           =   3375
         End
      End
      Begin VB.Frame Frame23 
         Height          =   735
         Left            =   -74880
         TabIndex        =   181
         Top             =   1920
         Width           =   19815
         Begin VB.TextBox Text23 
            Height          =   495
            Left            =   18240
            TabIndex        =   187
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton e23 
            Caption         =   "Percobaan"
            Height          =   495
            Left            =   12480
            TabIndex        =   186
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton d23 
            Caption         =   "Makmur"
            Height          =   495
            Left            =   10200
            TabIndex        =   185
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton c23 
            Caption         =   "Akal"
            Height          =   495
            Left            =   8160
            TabIndex        =   184
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b23 
            Caption         =   "Cerdas"
            Height          =   495
            Left            =   5880
            TabIndex        =   183
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton a23 
            Caption         =   "Kreatif"
            Height          =   495
            Left            =   3720
            TabIndex        =   182
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label Label35 
            Caption         =   "23.    INOVATIF :"
            Height          =   255
            Left            =   240
            TabIndex        =   188
            Top             =   240
            Width           =   3375
         End
      End
      Begin VB.Frame Frame22 
         Height          =   735
         Left            =   -74880
         TabIndex        =   173
         Top             =   2640
         Width           =   19815
         Begin VB.TextBox Text24 
            Height          =   495
            Left            =   18240
            TabIndex        =   179
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton e24 
            Caption         =   "Ghaib"
            Height          =   495
            Left            =   12480
            TabIndex        =   178
            Top             =   120
            Width           =   2295
         End
         Begin VB.OptionButton d24 
            Caption         =   "Dugaan"
            Height          =   495
            Left            =   10200
            TabIndex        =   177
            Top             =   120
            Width           =   2295
         End
         Begin VB.OptionButton c24 
            Caption         =   "Keajaiban"
            Height          =   495
            Left            =   8160
            TabIndex        =   176
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b24 
            Caption         =   "Rahasia"
            Height          =   495
            Left            =   5880
            TabIndex        =   175
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton a24 
            Caption         =   "Histeri"
            Height          =   495
            Left            =   3720
            TabIndex        =   174
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label Label36 
            Caption         =   "24.    MISTERI :"
            Height          =   255
            Left            =   240
            TabIndex        =   180
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Frame Frame21 
         Height          =   735
         Left            =   -74880
         TabIndex        =   165
         Top             =   3360
         Width           =   19815
         Begin VB.TextBox Text25 
            Height          =   495
            Left            =   18240
            TabIndex        =   171
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton e25 
            Caption         =   "Jasmani"
            Height          =   495
            Left            =   12480
            TabIndex        =   170
            Top             =   120
            Width           =   2415
         End
         Begin VB.OptionButton d25 
            Caption         =   "Ciri - Ciri"
            Height          =   495
            Left            =   10200
            TabIndex        =   169
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton c25 
            Caption         =   "Teladan"
            Height          =   495
            Left            =   8160
            TabIndex        =   168
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton b25 
            Caption         =   "Sidik Jari"
            Height          =   495
            Left            =   5880
            TabIndex        =   167
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton a25 
            Caption         =   "Gambar"
            Height          =   495
            Left            =   3720
            TabIndex        =   166
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label37 
            Caption         =   "25.    IDENTITAS :"
            Height          =   255
            Left            =   240
            TabIndex        =   172
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Frame Frame20 
         Height          =   735
         Left            =   -74880
         TabIndex        =   157
         Top             =   6960
         Width           =   19815
         Begin VB.TextBox Text20 
            Height          =   495
            Left            =   18360
            TabIndex        =   163
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton e20 
            Caption         =   "Terkendali"
            Height          =   495
            Left            =   12240
            TabIndex        =   162
            Top             =   120
            Width           =   1815
         End
         Begin VB.OptionButton d20 
            Caption         =   "Adil"
            Height          =   495
            Left            =   9960
            TabIndex        =   161
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton c20 
            Caption         =   "Sejahtera"
            Height          =   495
            Left            =   7800
            TabIndex        =   160
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton b20 
            Caption         =   "Nyaman"
            Height          =   495
            Left            =   5640
            TabIndex        =   159
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton a20 
            Caption         =   "Aman"
            Height          =   495
            Left            =   3480
            TabIndex        =   158
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label28 
            Caption         =   "20.  MAKMUR :"
            Height          =   255
            Left            =   120
            TabIndex        =   164
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame Frame19 
         Height          =   735
         Left            =   -74880
         TabIndex        =   149
         Top             =   6240
         Width           =   19815
         Begin VB.TextBox Text19 
            Height          =   495
            Left            =   18360
            TabIndex        =   155
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton e19 
            Caption         =   "Akuntabilitas"
            Height          =   495
            Left            =   12240
            TabIndex        =   154
            Top             =   120
            Width           =   1815
         End
         Begin VB.OptionButton d19 
            Caption         =   "Efektifitas"
            Height          =   495
            Left            =   9960
            TabIndex        =   153
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton c19 
            Caption         =   "Efektifitas"
            Height          =   495
            Left            =   7800
            TabIndex        =   152
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton b19 
            Caption         =   "Keseriusan"
            Height          =   495
            Left            =   5640
            TabIndex        =   151
            Top             =   120
            Width           =   1815
         End
         Begin VB.OptionButton a19 
            Caption         =   "Tantangan"
            Height          =   495
            Left            =   3480
            TabIndex        =   150
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label27 
            Caption         =   "19.  PROBLEMA :"
            Height          =   255
            Left            =   120
            TabIndex        =   156
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame Frame18 
         Height          =   735
         Left            =   -74880
         TabIndex        =   141
         Top             =   5520
         Width           =   19815
         Begin VB.TextBox Text18 
            Height          =   495
            Left            =   18360
            TabIndex        =   147
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton e18 
            Caption         =   "Efektifitas"
            Height          =   495
            Left            =   12240
            TabIndex        =   146
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton d18 
            Caption         =   "Pengelolaan"
            Height          =   495
            Left            =   9960
            TabIndex        =   145
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton c18 
            Caption         =   "Evaluasi"
            Height          =   495
            Left            =   7800
            TabIndex        =   144
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton b18 
            Caption         =   "Control"
            Height          =   495
            Left            =   5640
            TabIndex        =   143
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton a18 
            Caption         =   "Planning"
            Height          =   495
            Left            =   3480
            TabIndex        =   142
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label Label26 
            Caption         =   "18.  RENCANA :"
            Height          =   255
            Left            =   120
            TabIndex        =   148
            Top             =   240
            Width           =   2535
         End
      End
      Begin VB.Frame Frame17 
         Height          =   735
         Left            =   -74880
         TabIndex        =   133
         Top             =   4800
         Width           =   19815
         Begin VB.TextBox Text17 
            Height          =   495
            Left            =   18360
            TabIndex        =   139
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton e17 
            Caption         =   "Rutinitas"
            Height          =   555
            Left            =   12240
            TabIndex        =   138
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton d17 
            Caption         =   "Produk"
            Height          =   555
            Left            =   9960
            TabIndex        =   137
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton c17 
            Caption         =   "Bisnis"
            Height          =   495
            Left            =   7800
            TabIndex        =   136
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton b17 
            Caption         =   "Perdagangan"
            Height          =   495
            Left            =   5640
            TabIndex        =   135
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton a17 
            Caption         =   "Masyarakat"
            Height          =   495
            Left            =   3480
            TabIndex        =   134
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label Label25 
            Caption         =   "17.  PASAR :"
            Height          =   255
            Left            =   120
            TabIndex        =   140
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame16 
         Height          =   735
         Left            =   -74880
         TabIndex        =   125
         Top             =   4080
         Width           =   19815
         Begin VB.TextBox Text16 
            Height          =   495
            Left            =   18360
            TabIndex        =   131
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton e16 
            Caption         =   "Metode"
            Height          =   495
            Left            =   12240
            TabIndex        =   130
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton d16 
            Caption         =   "Taktik"
            Height          =   555
            Left            =   9960
            TabIndex        =   129
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton c16 
            Caption         =   "Pemikiran"
            Height          =   495
            Left            =   7800
            TabIndex        =   128
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton b16 
            Caption         =   "Kepemimpinan"
            Height          =   495
            Left            =   5640
            TabIndex        =   127
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton a16 
            Caption         =   "Manajemen"
            Height          =   495
            Left            =   3480
            TabIndex        =   126
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label Label24 
            Caption         =   "16.  STRATEGI :"
            Height          =   255
            Left            =   120
            TabIndex        =   132
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame15 
         Height          =   735
         Left            =   -74880
         TabIndex        =   117
         Top             =   3360
         Width           =   19815
         Begin VB.TextBox Text15 
            Height          =   495
            Left            =   18360
            TabIndex        =   123
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton e15 
            Caption         =   "Perkiraan"
            Height          =   495
            Left            =   12240
            TabIndex        =   122
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton d15 
            Caption         =   "Gagasan"
            Height          =   495
            Left            =   9960
            TabIndex        =   121
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton c15 
            Caption         =   "Rincian"
            Height          =   495
            Left            =   7800
            TabIndex        =   120
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b15 
            Caption         =   "Kemajuan"
            Height          =   555
            Left            =   5640
            TabIndex        =   119
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton a15 
            Caption         =   "Perkembangan"
            Height          =   495
            Left            =   3480
            TabIndex        =   118
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label Label23 
            Caption         =   "15.   IDE :"
            Height          =   255
            Left            =   120
            TabIndex        =   124
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame Frame14 
         Height          =   735
         Left            =   -74880
         TabIndex        =   109
         Top             =   2640
         Width           =   19815
         Begin VB.TextBox Text14 
            Height          =   495
            Left            =   18360
            TabIndex        =   115
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton e14 
            Caption         =   "Tercapai"
            Height          =   495
            Left            =   12240
            TabIndex        =   114
            Top             =   120
            Width           =   1575
         End
         Begin VB.OptionButton d14 
            Caption         =   "Karir"
            Height          =   555
            Left            =   9960
            TabIndex        =   113
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton c14 
            Caption         =   "Aman"
            Height          =   495
            Left            =   7800
            TabIndex        =   112
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b14 
            Caption         =   "Sentausa"
            Height          =   495
            Left            =   5640
            TabIndex        =   111
            Top             =   120
            Width           =   1815
         End
         Begin VB.OptionButton a14 
            Caption         =   "Sehat"
            Height          =   495
            Left            =   3480
            TabIndex        =   110
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label Label22 
            Caption         =   "14.   SUKSES :"
            Height          =   255
            Left            =   120
            TabIndex        =   116
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame Frame13 
         Height          =   735
         Left            =   -74880
         TabIndex        =   101
         Top             =   1920
         Width           =   19815
         Begin VB.TextBox Text13 
            Height          =   495
            Left            =   18360
            TabIndex        =   107
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton e13 
            Caption         =   "Cerai"
            Height          =   495
            Left            =   12240
            TabIndex        =   106
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton d13 
            Caption         =   "Laris"
            Height          =   495
            Left            =   9960
            TabIndex        =   105
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton c13 
            Caption         =   "Gundah"
            Height          =   555
            Left            =   7800
            TabIndex        =   104
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton b13 
            Caption         =   "Mewah"
            Height          =   495
            Left            =   5640
            TabIndex        =   103
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton a13 
            Caption         =   "Gaduh"
            Height          =   495
            Left            =   3480
            TabIndex        =   102
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label21 
            Caption         =   "13.   RAMAI :"
            Height          =   255
            Left            =   120
            TabIndex        =   108
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame12 
         Height          =   735
         Left            =   -74880
         TabIndex        =   93
         Top             =   1200
         Width           =   19815
         Begin VB.TextBox Text12 
            Height          =   495
            Left            =   18360
            TabIndex        =   99
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton e12 
            Caption         =   "Inspirasi"
            Height          =   495
            Left            =   12240
            TabIndex        =   98
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton d12 
            Caption         =   "Inovatif"
            Height          =   555
            Left            =   9960
            TabIndex        =   97
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton c12 
            Caption         =   "Ide"
            Height          =   495
            Left            =   7800
            TabIndex        =   96
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b12 
            Caption         =   "Menulis"
            Height          =   495
            Left            =   5640
            TabIndex        =   95
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton a12 
            Caption         =   "Membuat"
            Height          =   495
            Left            =   3480
            TabIndex        =   94
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label20 
            Caption         =   "12.   KREATIF :"
            Height          =   255
            Left            =   120
            TabIndex        =   100
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame11 
         Height          =   735
         Left            =   -74880
         TabIndex        =   85
         Top             =   480
         Width           =   19815
         Begin VB.TextBox Text11 
            Height          =   495
            Left            =   18360
            TabIndex        =   91
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton e11 
            Caption         =   "Biaya"
            Height          =   495
            Left            =   12240
            TabIndex        =   90
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton d11 
            Caption         =   "Kontainer"
            Height          =   495
            Left            =   9960
            TabIndex        =   89
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton c11 
            Caption         =   "Mebelair"
            Height          =   555
            Left            =   7800
            TabIndex        =   88
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b11 
            Caption         =   "Tekstil"
            Height          =   495
            Left            =   5640
            TabIndex        =   87
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton a11 
            Caption         =   "Jual"
            Height          =   495
            Left            =   3480
            TabIndex        =   86
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label19 
            Caption         =   "11.   EKSPOR :"
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame10 
         Height          =   735
         Left            =   120
         TabIndex        =   77
         Top             =   6960
         Width           =   19815
         Begin VB.TextBox Text10 
            Height          =   495
            Left            =   18600
            TabIndex        =   83
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton e10 
            Caption         =   "Enggan"
            Height          =   495
            Left            =   11880
            TabIndex        =   82
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton d10 
            Caption         =   "Malas"
            Height          =   495
            Left            =   9480
            TabIndex        =   81
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton c10 
            Caption         =   "Bimbang"
            Height          =   495
            Left            =   7320
            TabIndex        =   80
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b10 
            Caption         =   "Sedih"
            Height          =   495
            Left            =   5400
            TabIndex        =   79
            Top             =   120
            Width           =   1815
         End
         Begin VB.OptionButton a10 
            Caption         =   "Rajin"
            Height          =   495
            Left            =   3480
            TabIndex        =   78
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label18 
            Caption         =   "10.  RAGU - RAGU :"
            Height          =   255
            Left            =   120
            TabIndex        =   84
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame9 
         Height          =   735
         Left            =   120
         TabIndex        =   69
         Top             =   6240
         Width           =   19815
         Begin VB.TextBox Text9 
            Height          =   495
            Left            =   18600
            TabIndex        =   75
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton e9 
            Caption         =   "Sederhana"
            Height          =   495
            Left            =   11880
            TabIndex        =   74
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton d9 
            Caption         =   "Mirip"
            Height          =   555
            Left            =   9480
            TabIndex        =   73
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton c9 
            Caption         =   "Kecewa"
            Height          =   495
            Left            =   7320
            TabIndex        =   72
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b9 
            Caption         =   "Khawatir"
            Height          =   495
            Left            =   5400
            TabIndex        =   71
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton a9 
            Caption         =   "Ragu - Ragu"
            Height          =   495
            Left            =   3480
            TabIndex        =   70
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label Label17 
            Caption         =   "9.  CEMAS :"
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame8 
         Height          =   735
         Left            =   120
         TabIndex        =   61
         Top             =   5520
         Width           =   19815
         Begin VB.TextBox Text8 
            Height          =   495
            Left            =   18600
            TabIndex        =   67
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton e8 
            Caption         =   "Persatuan"
            Height          =   555
            Left            =   11880
            TabIndex        =   66
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton d8 
            Caption         =   "Pandangan"
            Height          =   495
            Left            =   9480
            TabIndex        =   65
            Top             =   120
            Width           =   2295
         End
         Begin VB.OptionButton c8 
            Caption         =   "Kemajuan"
            Height          =   495
            Left            =   7320
            TabIndex        =   64
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b8 
            Caption         =   "Global"
            Height          =   495
            Left            =   5400
            TabIndex        =   63
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton a8 
            Caption         =   "Modern"
            Height          =   495
            Left            =   3480
            TabIndex        =   62
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label16 
            Caption         =   "8.  UNIVERSAL :"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame7 
         Height          =   735
         Left            =   120
         TabIndex        =   53
         Top             =   4800
         Width           =   19815
         Begin VB.TextBox Text7 
            Height          =   495
            Left            =   18600
            TabIndex        =   59
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton e7 
            Caption         =   "Laten"
            Height          =   495
            Left            =   11880
            TabIndex        =   58
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton d7 
            Caption         =   "Mewah"
            Height          =   495
            Left            =   9480
            TabIndex        =   57
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton c7 
            Caption         =   "Seperti"
            Height          =   495
            Left            =   7320
            TabIndex        =   56
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b7 
            Caption         =   "Kerap Kali"
            Height          =   495
            Left            =   5400
            TabIndex        =   55
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton a7 
            Caption         =   "Jarang"
            Height          =   495
            Left            =   3480
            TabIndex        =   54
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label15 
            Caption         =   "7.   SPORADIS :"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame6 
         Height          =   735
         Left            =   120
         TabIndex        =   45
         Top             =   4080
         Width           =   19815
         Begin VB.TextBox Text6 
            Height          =   495
            Left            =   18600
            TabIndex        =   51
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton e6 
            Caption         =   "Cinta"
            Height          =   495
            Left            =   11880
            TabIndex        =   50
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton d6 
            Caption         =   "Anak - Istri"
            Height          =   435
            Left            =   9480
            TabIndex        =   49
            Top             =   120
            Width           =   2295
         End
         Begin VB.OptionButton c6 
            Caption         =   "Suami - Istri"
            Height          =   495
            Left            =   7320
            TabIndex        =   48
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b6 
            Caption         =   "Pernikahan"
            Height          =   495
            Left            =   5400
            TabIndex        =   47
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton a6 
            Caption         =   "Rumah Tangga"
            Height          =   495
            Left            =   3480
            TabIndex        =   46
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label Label14 
            Caption         =   "6.   PERKAWINAN :"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame5 
         Height          =   735
         Left            =   120
         TabIndex        =   37
         Top             =   3360
         Width           =   19815
         Begin VB.TextBox Text5 
            Height          =   495
            Left            =   18600
            TabIndex        =   43
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton e5 
            Caption         =   "Matematika"
            Height          =   495
            Left            =   11880
            TabIndex        =   42
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton d5 
            Caption         =   "Perdagangan"
            Height          =   495
            Left            =   9480
            TabIndex        =   41
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton c5 
            Caption         =   "Akuntansi"
            Height          =   495
            Left            =   7320
            TabIndex        =   40
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b5 
            Caption         =   "Organisasi"
            Height          =   495
            Left            =   5400
            TabIndex        =   39
            Top             =   120
            Width           =   1815
         End
         Begin VB.OptionButton a5 
            Caption         =   "Manajemen"
            Height          =   495
            Left            =   3480
            TabIndex        =   38
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label Label13 
            Caption         =   "5.   PERUSAHAAN :"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   120
         TabIndex        =   29
         Top             =   2640
         Width           =   19815
         Begin VB.TextBox Text4 
            Height          =   525
            Left            =   18600
            TabIndex        =   35
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton e4 
            Caption         =   "Suksesi"
            Height          =   495
            Left            =   11880
            TabIndex        =   34
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton d4 
            Caption         =   "Aturan"
            Height          =   495
            Left            =   9480
            TabIndex        =   33
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton c4 
            Caption         =   "Presiden"
            Height          =   495
            Left            =   7320
            TabIndex        =   32
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b4 
            Caption         =   "Undang - Undang"
            Height          =   495
            Left            =   5400
            TabIndex        =   31
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton a4 
            Caption         =   "Negara"
            Height          =   495
            Left            =   3480
            TabIndex        =   30
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label12 
            Caption         =   "4.   PEMERINTAH :"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   120
         TabIndex        =   21
         Top             =   1920
         Width           =   19815
         Begin VB.TextBox Text3 
            Height          =   495
            Left            =   18600
            TabIndex        =   27
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton e3 
            Caption         =   "Stereo"
            Height          =   495
            Left            =   11880
            TabIndex        =   26
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton d3 
            Caption         =   "Merasa"
            Height          =   495
            Left            =   9480
            TabIndex        =   25
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton c3 
            Caption         =   "Beda"
            Height          =   435
            Left            =   7320
            TabIndex        =   24
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b3 
            Caption         =   "Ajeng"
            Height          =   495
            Left            =   5400
            TabIndex        =   23
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton a3 
            Caption         =   "Satu"
            Height          =   495
            Left            =   3480
            TabIndex        =   22
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label11 
            Caption         =   "3.  MONOTON :"
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   120
         TabIndex        =   13
         Top             =   1200
         Width           =   19815
         Begin VB.TextBox Text2 
            Height          =   495
            Left            =   18600
            TabIndex        =   19
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton e2 
            Caption         =   "Terlambat"
            Height          =   495
            Left            =   11880
            TabIndex        =   18
            Top             =   120
            Width           =   1815
         End
         Begin VB.OptionButton d2 
            Caption         =   "Praktis"
            Height          =   495
            Left            =   9480
            TabIndex        =   17
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton c2 
            Caption         =   "Bersahaja"
            Height          =   495
            Left            =   7320
            TabIndex        =   16
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton b2 
            Caption         =   "Canggih"
            Height          =   495
            Left            =   5400
            TabIndex        =   15
            Top             =   120
            Width           =   1815
         End
         Begin VB.OptionButton a2 
            Caption         =   "Kuno"
            Height          =   495
            Left            =   3480
            TabIndex        =   14
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label10 
            Caption         =   "2.  SEDERHANA : "
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   19815
         Begin VB.TextBox Text1 
            Height          =   495
            Left            =   18600
            TabIndex        =   11
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton e1 
            Caption         =   "Swakarsa"
            Height          =   495
            Left            =   11880
            TabIndex        =   10
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton d1 
            Caption         =   "Berdiri"
            Height          =   495
            Left            =   9480
            TabIndex        =   9
            Top             =   120
            Width           =   2295
         End
         Begin VB.OptionButton c1 
            Caption         =   "Mengikuti"
            Height          =   495
            Left            =   7320
            TabIndex        =   8
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton b1 
            Caption         =   "Mandiri"
            Height          =   495
            Left            =   5400
            TabIndex        =   7
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton a1 
            Caption         =   "Roboh"
            Height          =   495
            Left            =   3480
            TabIndex        =   6
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "1.  SWADAYA :"
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   3135
         End
      End
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
      TabIndex        =   215
      Top             =   1200
      Width           =   2085
   End
   Begin VB.Label Label6 
      Caption         =   "ID                 :"
      Height          =   285
      Left            =   30
      TabIndex        =   214
      Top             =   0
      Width           =   1125
   End
   Begin VB.Label Label7 
      Caption         =   "NAMA           :"
      Height          =   165
      Left            =   0
      TabIndex        =   213
      Top             =   420
      Width           =   1245
   End
   Begin VB.Label Label8 
      Caption         =   "Tanggal Test :"
      Height          =   195
      Left            =   0
      TabIndex        =   212
      Top             =   870
      Width           =   1305
   End
   Begin VB.Label Label38 
      Caption         =   "12.   3, 7, 15, 31, 63, ..., ..."
      Height          =   255
      Left            =   600
      TabIndex        =   211
      Top             =   6120
      Width           =   3015
   End
   Begin VB.Label Label32 
      Caption         =   "12.   3, 7, 15, 31, 63, ..., ..."
      Height          =   255
      Left            =   600
      TabIndex        =   210
      Top             =   5400
      Width           =   3015
   End
   Begin VB.Label Label31 
      Caption         =   "12.   3, 7, 15, 31, 63, ..., ..."
      Height          =   255
      Left            =   600
      TabIndex        =   209
      Top             =   4680
      Width           =   3015
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   20250
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "TEST VII"
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
      TabIndex        =   208
      Top             =   120
      Width           =   2115
   End
   Begin VB.Label Label2 
      Caption         =   "Test Diksi"
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
      TabIndex        =   207
      Top             =   720
      Width           =   1005
   End
   Begin VB.Label Label5 
      Caption         =   $"frmpsi7.frx":3C12
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   90
      TabIndex        =   206
      Top             =   1650
      Width           =   12795
   End
   Begin VB.Label Label9 
      Caption         =   "Waktu 8 menit Soal 25 Nomor."
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
      Left            =   3600
      TabIndex        =   205
      Top             =   1200
      Width           =   3135
   End
End
Attribute VB_Name = "frmpsi7"
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
Text1.Text = "A"
End Sub

Private Sub b1_Click()
Text1.Text = "B"
End Sub

Private Sub c1_Click()
Text1.Text = "C"
End Sub

Private Sub d1_Click()
Text1.Text = "D"
End Sub

Private Sub e1_Click()
Text1.Text = "E"
End Sub

Private Sub a2_Click()
Text2.Text = "A"
End Sub

Private Sub b2_Click()
Text2.Text = "B"
End Sub

Private Sub c2_Click()
Text2.Text = "C"
End Sub

Private Sub d2_Click()
Text2.Text = "D"
End Sub

Private Sub e2_Click()
Text2.Text = "E"
End Sub

Private Sub a3_Click()
Text3.Text = "A"
End Sub

Private Sub b3_Click()
Text3.Text = "B"
End Sub

Private Sub c3_Click()
Text3.Text = "C"
End Sub

Private Sub d3_Click()
Text3.Text = "D"
End Sub

Private Sub e3_Click()
Text3.Text = "E"
End Sub
Private Sub a4_Click()
Text4.Text = "A"
End Sub

Private Sub b4_Click()
Text4.Text = "B"
End Sub

Private Sub c4_Click()
Text4.Text = "C"
End Sub

Private Sub d4_Click()
Text4.Text = "D"
End Sub

Private Sub e4_Click()
Text4.Text = "E"
End Sub

Private Sub a5_Click()
Text5.Text = "A"
End Sub

Private Sub b5_Click()
Text5.Text = "B"
End Sub

Private Sub c5_Click()
Text5.Text = "C"
End Sub

Private Sub d5_Click()
Text5.Text = "D"
End Sub

Private Sub e5_Click()
Text5.Text = "E"
End Sub

Private Sub a6_Click()
Text6.Text = "A"
End Sub

Private Sub b6_Click()
Text6.Text = "B"
End Sub

Private Sub c6_Click()
Text6.Text = "C"
End Sub

Private Sub d6_Click()
Text6.Text = "D"
End Sub

Private Sub e6_Click()
Text6.Text = "E"
End Sub

Private Sub a7_Click()
Text7.Text = "A"
End Sub

Private Sub b7_Click()
Text7.Text = "B"
End Sub

Private Sub c7_Click()
Text7.Text = "C"
End Sub

Private Sub d7_Click()
Text7.Text = "D"
End Sub

Private Sub e7_Click()
Text7.Text = "E"
End Sub

Private Sub a8_Click()
Text8.Text = "A"
End Sub

Private Sub b8_Click()
Text8.Text = "B"
End Sub

Private Sub c8_Click()
Text8.Text = "C"
End Sub

Private Sub d8_Click()
Text8.Text = "D"
End Sub

Private Sub e8_Click()
Text8.Text = "E"
End Sub

Private Sub a9_Click()
Text9.Text = "A"
End Sub

Private Sub b9_Click()
Text9.Text = "B"
End Sub

Private Sub c9_Click()
Text9.Text = "C"
End Sub

Private Sub d9_Click()
Text9.Text = "D"
End Sub

Private Sub e9_Click()
Text9.Text = "E"
End Sub

Private Sub a10_Click()
Text10.Text = "A"
End Sub

Private Sub b10_Click()
Text10.Text = "B"
End Sub

Private Sub c10_Click()
Text10.Text = "C"
End Sub

Private Sub d10_Click()
Text10.Text = "D"
End Sub

Private Sub e10_Click()
Text10.Text = "E"
End Sub

Private Sub a11_Click()
Text11.Text = "A"
End Sub

Private Sub b11_Click()
Text11.Text = "B"
End Sub

Private Sub c11_Click()
Text11.Text = "C"
End Sub

Private Sub d11_Click()
Text11.Text = "D"
End Sub

Private Sub e11_Click()
Text11.Text = "E"
End Sub

Private Sub a12_Click()
Text12.Text = "A"
End Sub

Private Sub b12_Click()
Text12.Text = "B"
End Sub

Private Sub c12_Click()
Text12.Text = "C"
End Sub

Private Sub d12_Click()
Text12.Text = "D"
End Sub

Private Sub e12_Click()
Text12.Text = "E"
End Sub

Private Sub a13_Click()
Text13.Text = "A"
End Sub

Private Sub b13_Click()
Text13.Text = "B"
End Sub

Private Sub c13_Click()
Text13.Text = "C"
End Sub

Private Sub d13_Click()
Text13.Text = "D"
End Sub

Private Sub e13_Click()
Text13.Text = "E"
End Sub

Private Sub a14_Click()
Text14.Text = "A"
End Sub

Private Sub b14_Click()
Text14.Text = "B"
End Sub

Private Sub c14_Click()
Text14.Text = "C"
End Sub

Private Sub d14_Click()
Text14.Text = "D"
End Sub

Private Sub e14_Click()
Text14.Text = "E"
End Sub

Private Sub a15_Click()
Text15.Text = "A"
End Sub

Private Sub b15_Click()
Text15.Text = "B"
End Sub

Private Sub c15_Click()
Text15.Text = "C"
End Sub

Private Sub d15_Click()
Text15.Text = "D"
End Sub

Private Sub e15_Click()
Text15.Text = "E"
End Sub

Private Sub a16_Click()
Text16.Text = "A"
End Sub

Private Sub b16_Click()
Text16.Text = "B"
End Sub

Private Sub c16_Click()
Text16.Text = "C"
End Sub

Private Sub d16_Click()
Text16.Text = "D"
End Sub

Private Sub e16_Click()
Text16.Text = "E"
End Sub

Private Sub a17_Click()
Text17.Text = "A"
End Sub

Private Sub b17_Click()
Text17.Text = "B"
End Sub

Private Sub c17_Click()
Text17.Text = "C"
End Sub

Private Sub d17_Click()
Text17.Text = "D"
End Sub

Private Sub e17_Click()
Text17.Text = "E"
End Sub

Private Sub a18_Click()
Text18.Text = "A"
End Sub

Private Sub b18_Click()
Text18.Text = "B"
End Sub

Private Sub c18_Click()
Text18.Text = "C"
End Sub

Private Sub d18_Click()
Text18.Text = "D"
End Sub

Private Sub e18_Click()
Text18.Text = "E"
End Sub

Private Sub a19_Click()
Text19.Text = "A"
End Sub

Private Sub b19_Click()
Text19.Text = "B"
End Sub

Private Sub c19_Click()
Text19.Text = "C"
End Sub

Private Sub d19_Click()
Text19.Text = "D"
End Sub

Private Sub e19_Click()
Text19.Text = "E"
End Sub

Private Sub a20_Click()
Text20.Text = "A"
End Sub

Private Sub b20_Click()
Text20.Text = "B"
End Sub

Private Sub c20_Click()
Text20.Text = "C"
End Sub

Private Sub d20_Click()
Text20.Text = "D"
End Sub

Private Sub e20_Click()
Text20.Text = "E"
End Sub

Private Sub a21_Click()
Text21.Text = "A"
End Sub

Private Sub b21_Click()
Text21.Text = "B"
End Sub

Private Sub c21_Click()
Text21.Text = "C"
End Sub

Private Sub d21_Click()
Text21.Text = "D"
End Sub

Private Sub e21_Click()
Text21.Text = "E"
End Sub

Private Sub a22_Click()
Text22.Text = "A"
End Sub

Private Sub b22_Click()
Text22.Text = "B"
End Sub

Private Sub c22_Click()
Text22.Text = "C"
End Sub

Private Sub d22_Click()
Text22.Text = "D"
End Sub

Private Sub e22_Click()
Text22.Text = "E"
End Sub

Private Sub a23_Click()
Text23.Text = "A"
End Sub

Private Sub b23_Click()
Text23.Text = "B"
End Sub

Private Sub c23_Click()
Text23.Text = "C"
End Sub

Private Sub d23_Click()
Text23.Text = "D"
End Sub

Private Sub e23_Click()
Text23.Text = "E"
End Sub

Private Sub a24_Click()
Text24.Text = "A"
End Sub

Private Sub b24_Click()
Text24.Text = "B"
End Sub

Private Sub c24_Click()
Text24.Text = "C"
End Sub

Private Sub d24_Click()
Text24.Text = "D"
End Sub

Private Sub e24_Click()
Text24.Text = "E"
End Sub

Private Sub a25_Click()
Text25.Text = "A"
End Sub

Private Sub b25_Click()
Text25.Text = "B"
End Sub

Private Sub c25_Click()
Text25.Text = "C"
End Sub

Private Sub d25_Click()
Text25.Text = "D"
End Sub

Private Sub e25_Click()
Text25.Text = "E"
End Sub

Private Sub nextbtn_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Or Text11.Text = "" Or Text12.Text = "" Or Text13.Text = "" Or Text14.Text = "" Or Text15.Text = "" Or Text16.Text = "" Or Text17.Text = "" Or Text18.Text = "" Or Text19.Text = "" Or Text20.Text = "" Or Text21.Text = "" Or Text22.Text = "" Or Text23.Text = "" Or Text24.Text = "" Or Text25.Text = "" Then
    MsgBox "Data Belum Lengkap"
    Else
    Dim tambahdata As String
        tambahdata = "Insert Into hasiltest7 values ('" & id.Text & "','" & nama.Text & "','" & tanggal.Text & "','" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "','" & Text7.Text & "','" & Text8.Text & "','" & Text9.Text & "','" & Text10.Text & "','" & Text11.Text & "','" & Text12.Text & "','" & Text13.Text & "','" & Text14.Text & "','" & Text15.Text & "','" & Text16.Text & "','" & Text17.Text & "','" & Text18.Text & "','" & Text19.Text & "','" & Text20.Text & "','" & Text21.Text & "','" & Text22.Text & "','" & Text23.Text & "','" & Text24.Text & "','" & Text25.Text & "')"
        koneksi.Execute tambahdata
        MsgBox "TEST 7 Berhasil Di simpan, Silakan lanjutkan ke TEST 8", vbInformation, "Pemberitahuan"
    frmpsi8.Show
    koneksi.Close
    Unload Me
    End If

End Sub

Private Sub Option12_Click()

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

If SSTab1.Caption = "SOAL NOMOR 21 - 25" Then
nextbtn.Enabled = True
Else
nextbtn.Enabled = False
End If
End Sub



