VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmpsi12 
   Caption         =   "Form1"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox id 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1170
      TabIndex        =   3
      Top             =   0
      Width           =   3195
   End
   Begin VB.TextBox nama 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1170
      TabIndex        =   2
      Top             =   390
      Width           =   3195
   End
   Begin VB.TextBox tanggal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1170
      TabIndex        =   1
      Top             =   810
      Width           =   3195
   End
   Begin VB.CommandButton nextbtn 
      Caption         =   "NEXT >>>>>"
      DisabledPicture =   "frmpsi12.frx":0000
      DownPicture     =   "frmpsi12.frx":13EA
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
      Picture         =   "frmpsi12.frx":27D4
      TabIndex        =   0
      Top             =   210
      Width           =   1695
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
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "SOAL NOMOR 1 - 20"
      TabPicture(0)   =   "frmpsi12.frx":3BBE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame25"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame24"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame23"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame22"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame21"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame10"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame7"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame6"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame20"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Frame19"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame18"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame17"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Frame16"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Frame15"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Frame14"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Frame13"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Frame12"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Frame11"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      Begin VB.Frame Frame11 
         Height          =   735
         Left            =   120
         TabIndex        =   100
         Top             =   360
         Width           =   6255
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   103
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b1 
            Caption         =   "(B)"
            Height          =   495
            Left            =   4080
            TabIndex        =   102
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a1 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   101
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "1.Pelimbahan air jatuhnya cucuran ke juga"
            Height          =   495
            Left            =   120
            TabIndex        =   104
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.Frame Frame12 
         Height          =   735
         Left            =   120
         TabIndex        =   95
         Top             =   1080
         Width           =   6255
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   98
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b2 
            Caption         =   "(B)"
            Height          =   495
            Left            =   4080
            TabIndex        =   97
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton a2 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   96
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "2. Makhluk tidak hidup mati setiap akan"
            Height          =   495
            Left            =   120
            TabIndex        =   99
            Top             =   120
            Width           =   2295
         End
      End
      Begin VB.Frame Frame13 
         Height          =   735
         Left            =   120
         TabIndex        =   90
         Top             =   1800
         Width           =   6255
         Begin VB.TextBox Text3 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   93
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b3 
            Caption         =   "(B)"
            Height          =   495
            Left            =   4080
            TabIndex        =   92
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a3 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   91
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "3. Di mindanau ada pulau Filipina"
            Height          =   495
            Left            =   120
            TabIndex        =   94
            Top             =   120
            Width           =   2295
         End
      End
      Begin VB.Frame Frame14 
         Height          =   735
         Left            =   120
         TabIndex        =   85
         Top             =   2520
         Width           =   6255
         Begin VB.TextBox Text4 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   88
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b4 
            Caption         =   "(B)"
            Height          =   495
            Left            =   4080
            TabIndex        =   87
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a4 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   86
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "4.Pasak besar dari pada tiang"
            Height          =   495
            Left            =   120
            TabIndex        =   89
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.Frame Frame15 
         Height          =   735
         Left            =   120
         TabIndex        =   80
         Top             =   3240
         Width           =   6255
         Begin VB.TextBox Text5 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   83
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b5 
            Caption         =   "(B)"
            Height          =   495
            Left            =   4080
            TabIndex        =   82
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a5 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   81
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label14 
            Caption         =   "5. Dan kucing anjing bersahabat sangat"
            Height          =   495
            Left            =   120
            TabIndex        =   84
            Top             =   120
            Width           =   2415
         End
      End
      Begin VB.Frame Frame16 
         Height          =   735
         Left            =   120
         TabIndex        =   75
         Top             =   3960
         Width           =   6255
         Begin VB.TextBox Text6 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   78
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b6 
            Caption         =   "(B)"
            Height          =   495
            Left            =   4080
            TabIndex        =   77
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a6 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   76
            Top             =   120
            Width           =   855
         End
         Begin VB.Label Label15 
            Caption         =   "6. Penguin burung terbang bisa"
            Height          =   375
            Left            =   120
            TabIndex        =   79
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame17 
         Height          =   735
         Left            =   120
         TabIndex        =   70
         Top             =   4680
         Width           =   6255
         Begin VB.OptionButton b7 
            Caption         =   "(B)"
            Height          =   495
            Left            =   4080
            TabIndex        =   73
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a7 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   72
            Top             =   120
            Width           =   975
         End
         Begin VB.TextBox Text7 
            Height          =   495
            Left            =   5400
            TabIndex        =   71
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label31 
            Caption         =   "7. Kubus lima mempunyai sisi"
            Height          =   375
            Left            =   120
            TabIndex        =   74
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame18 
         Height          =   735
         Left            =   120
         TabIndex        =   65
         Top             =   5400
         Width           =   6255
         Begin VB.TextBox Text8 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   68
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b8 
            Caption         =   "(B)"
            Height          =   495
            Left            =   4080
            TabIndex        =   67
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a8 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   66
            Top             =   120
            Width           =   975
         End
         Begin VB.Label Label17 
            Caption         =   "8.Hitam asin sedunia paling laut"
            Height          =   375
            Left            =   120
            TabIndex        =   69
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame19 
         Height          =   735
         Left            =   120
         TabIndex        =   60
         Top             =   6120
         Width           =   6255
         Begin VB.TextBox Text9 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   63
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b9 
            Caption         =   "(B)"
            Height          =   495
            Left            =   4080
            TabIndex        =   62
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton a9 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   61
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label18 
            Caption         =   "9. Dua kalimantan Negara terdapat pulau"
            Height          =   495
            Left            =   120
            TabIndex        =   64
            Top             =   120
            Width           =   2055
         End
      End
      Begin VB.Frame Frame20 
         Height          =   735
         Left            =   120
         TabIndex        =   55
         Top             =   6840
         Width           =   6255
         Begin VB.TextBox Text10 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   58
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton b10 
            Caption         =   "(B)"
            Height          =   495
            Left            =   4080
            TabIndex        =   57
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton a10 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   56
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label19 
            Caption         =   "10. Adalah Indonesia Presiden ketiga Abdurrahman Wahid"
            Height          =   495
            Left            =   120
            TabIndex        =   59
            Top             =   120
            Width           =   2295
         End
      End
      Begin VB.Frame Frame6 
         Height          =   735
         Left            =   12480
         TabIndex        =   50
         Top             =   360
         Width           =   6255
         Begin VB.OptionButton a11 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   53
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton b11 
            Caption         =   "(B)"
            Height          =   495
            Left            =   4080
            TabIndex        =   52
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text11 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   51
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label25 
            Caption         =   "11. Ringan sama di pikul sama berat di jinjing"
            Height          =   495
            Left            =   120
            TabIndex        =   54
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.Frame Frame7 
         Height          =   735
         Left            =   12480
         TabIndex        =   45
         Top             =   1080
         Width           =   6255
         Begin VB.OptionButton a12 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   48
            Top             =   120
            Width           =   1215
         End
         Begin VB.OptionButton b12 
            Caption         =   "(B)"
            Height          =   495
            Left            =   4080
            TabIndex        =   47
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox Text12 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   46
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label26 
            Caption         =   "12. Tikus kucing takut pada"
            Height          =   495
            Left            =   120
            TabIndex        =   49
            Top             =   120
            Width           =   2295
         End
      End
      Begin VB.Frame Frame8 
         Height          =   735
         Left            =   12480
         TabIndex        =   40
         Top             =   1800
         Width           =   6255
         Begin VB.OptionButton a13 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   43
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton b13 
            Caption         =   "(B)"
            Height          =   495
            Left            =   4080
            TabIndex        =   42
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text13 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   41
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label27 
            Caption         =   "13.Di sity dimana dipijak bumi, langit di junjung"
            Height          =   495
            Left            =   120
            TabIndex        =   44
            Top             =   120
            Width           =   2415
         End
      End
      Begin VB.Frame Frame9 
         Height          =   735
         Left            =   12480
         TabIndex        =   35
         Top             =   2520
         Width           =   6255
         Begin VB.OptionButton a14 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   38
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton b14 
            Caption         =   "(B)"
            Height          =   495
            Left            =   4080
            TabIndex        =   37
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text14 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   36
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label28 
            Caption         =   "14. Afrika dari berasal Cendrawasih burung"
            Height          =   495
            Left            =   120
            TabIndex        =   39
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.Frame Frame10 
         Height          =   735
         Left            =   12480
         TabIndex        =   30
         Top             =   3240
         Width           =   6255
         Begin VB.OptionButton a15 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   33
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton b15 
            Caption         =   "(B)"
            Height          =   495
            Left            =   4080
            TabIndex        =   32
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text15 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   31
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label29 
            Caption         =   "15.Berakit-rakit berenang-renang ke hulu ke danau"
            Height          =   495
            Left            =   120
            TabIndex        =   34
            Top             =   120
            Width           =   2415
         End
      End
      Begin VB.Frame Frame21 
         Height          =   735
         Left            =   12480
         TabIndex        =   25
         Top             =   3960
         Width           =   6255
         Begin VB.OptionButton a16 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   28
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton b16 
            Caption         =   "(B)"
            Height          =   495
            Left            =   4080
            TabIndex        =   27
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text16 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   26
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label32 
            Caption         =   "16. Enam adalah serangga berkaki binatang"
            Height          =   495
            Left            =   120
            TabIndex        =   29
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.Frame Frame22 
         Height          =   735
         Left            =   12480
         TabIndex        =   20
         Top             =   4680
         Width           =   6255
         Begin VB.TextBox Text17 
            Height          =   495
            Left            =   5400
            TabIndex        =   23
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton a17 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   22
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton b17 
            Caption         =   "(B)"
            Height          =   495
            Left            =   4080
            TabIndex        =   21
            Top             =   120
            Width           =   1095
         End
         Begin VB.Label Label33 
            Caption         =   "17. Mata Malaysia dolar adalah uang"
            Height          =   495
            Left            =   120
            TabIndex        =   24
            Top             =   120
            Width           =   2055
         End
      End
      Begin VB.Frame Frame23 
         Height          =   735
         Left            =   12480
         TabIndex        =   15
         Top             =   5400
         Width           =   6255
         Begin VB.OptionButton a18 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   18
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton b18 
            Caption         =   "(B)"
            Height          =   495
            Left            =   4080
            TabIndex        =   17
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text18 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   16
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label34 
            Caption         =   "18. Menjadi nasi sudah bubur"
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame24 
         Height          =   735
         Left            =   12480
         TabIndex        =   10
         Top             =   6120
         Width           =   6255
         Begin VB.OptionButton a19 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   13
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton b19 
            Caption         =   "(B)"
            Height          =   495
            Left            =   4080
            TabIndex        =   12
            Top             =   120
            Width           =   1215
         End
         Begin VB.TextBox Text19 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   11
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label35 
            Caption         =   "19. Kemudian tak berguna sesal"
            Height          =   495
            Left            =   120
            TabIndex        =   14
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.Frame Frame25 
         Height          =   735
         Left            =   12480
         TabIndex        =   5
         Top             =   6840
         Width           =   6255
         Begin VB.OptionButton a20 
            Caption         =   "(S)"
            Height          =   495
            Left            =   2520
            TabIndex        =   8
            Top             =   120
            Width           =   1455
         End
         Begin VB.OptionButton b20 
            Caption         =   "(B)"
            Height          =   495
            Left            =   4080
            TabIndex        =   7
            Top             =   120
            Width           =   1095
         End
         Begin VB.TextBox Text20 
            Enabled         =   0   'False
            Height          =   495
            Left            =   5400
            TabIndex        =   6
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label36 
            Caption         =   "20. Ibukota Sudan Port Moresbu adalah"
            Height          =   495
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   2295
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
         TabIndex        =   105
         Top             =   -2160
         Width           =   4575
      End
   End
   Begin VB.Label Label54 
      Caption         =   "b. Dibenarkan : Tape dibuat dari gandum (S)"
      Height          =   375
      Left            =   6480
      TabIndex        =   118
      Top             =   2280
      Width           =   4095
   End
   Begin VB.Label Label53 
      Caption         =   "Contoh 2 :a. Gandum dari tape dibuat"
      Height          =   255
      Left            =   5760
      TabIndex        =   117
      Top             =   2040
      Width           =   4455
   End
   Begin VB.Label Label52 
      Caption         =   "b. Dibenarkan : Petani membajak sawah dengan sapi (B)"
      Height          =   375
      Left            =   840
      TabIndex        =   116
      Top             =   2280
      Width           =   4095
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
      Left            =   2400
      TabIndex        =   115
      Top             =   1200
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
      Height          =   315
      Left            =   120
      TabIndex        =   114
      Top             =   1200
      Width           =   2085
   End
   Begin VB.Label Label8 
      Caption         =   "Tanggal Test :"
      Height          =   195
      Left            =   0
      TabIndex        =   113
      Top             =   870
      Width           =   1305
   End
   Begin VB.Label Label7 
      Caption         =   "NAMA           :"
      Height          =   165
      Left            =   0
      TabIndex        =   112
      Top             =   420
      Width           =   1245
   End
   Begin VB.Label Label6 
      Caption         =   "ID                 :"
      Height          =   285
      Left            =   30
      TabIndex        =   111
      Top             =   0
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "TEST XII"
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
      TabIndex        =   110
      Top             =   120
      Width           =   2115
   End
   Begin VB.Label Label2 
      Caption         =   "Test Kalimat Tak Teratur"
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
      Left            =   9120
      TabIndex        =   109
      Top             =   720
      Width           =   2325
   End
   Begin VB.Label Label5 
      Caption         =   $"frmpsi12.frx":3BDA
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
      TabIndex        =   108
      Top             =   1530
      Width           =   19995
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   20250
      Y1              =   1140
      Y2              =   1140
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
      TabIndex        =   107
      Top             =   7410
      Width           =   4575
   End
   Begin VB.Label Label3 
      Caption         =   "Contoh 1 :a. Sawah membajak sapi dengan petani"
      Height          =   255
      Left            =   120
      TabIndex        =   106
      Top             =   2040
      Width           =   4455
   End
End
Attribute VB_Name = "frmpsi12"
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
Text1.Text = "B"
End Sub

Private Sub a2_Click()
Text2.Text = "S"
End Sub

Private Sub b2_Click()
Text2.Text = "B"
End Sub

Private Sub a3_Click()
Text3.Text = "S"
End Sub

Private Sub b3_Click()
Text3.Text = "B"
End Sub

Private Sub a4_Click()
Text4.Text = "S"
End Sub

Private Sub b4_Click()
Text4.Text = "B"
End Sub

Private Sub a5_Click()
Text5.Text = "S"
End Sub

Private Sub b5_Click()
Text5.Text = "B"
End Sub

Private Sub a6_Click()
Text6.Text = "S"
End Sub

Private Sub b6_Click()
Text6.Text = "B"
End Sub

Private Sub a7_Click()
Text7.Text = "S"
End Sub

Private Sub b7_Click()
Text7.Text = "B"
End Sub

Private Sub a8_Click()
Text8.Text = "S"
End Sub
Private Sub b8_Click()
Text8.Text = "B"
End Sub

Private Sub a9_Click()
Text9.Text = "S"
End Sub

Private Sub b9_Click()
Text9.Text = "B"
End Sub

Private Sub a10_Click()
Text10.Text = "S"
End Sub

Private Sub b10_Click()
Text10.Text = "B"
End Sub

Private Sub a11_Click()
Text11.Text = "S"
End Sub

Private Sub b11_Click()
Text11.Text = "B"
End Sub

Private Sub a12_Click()
Text12.Text = "S"
End Sub

Private Sub b12_Click()
Text12.Text = "B"
End Sub

Private Sub a13_Click()
Text13.Text = "S"
End Sub

Private Sub b13_Click()
Text13.Text = "B"
End Sub

Private Sub a14_Click()
Text14.Text = "S"
End Sub

Private Sub b14_Click()
Text14.Text = "B"
End Sub

Private Sub a15_Click()
Text15.Text = "S"
End Sub

Private Sub b15_Click()
Text15.Text = "B"
End Sub

Private Sub a16_Click()
Text16.Text = "S"
End Sub

Private Sub b16_Click()
Text16.Text = "B"
End Sub

Private Sub a17_Click()
Text17.Text = "S"
End Sub

Private Sub b17_Click()
Text17.Text = "B"
End Sub

Private Sub a18_Click()
Text18.Text = "S"
End Sub

Private Sub b18_Click()
Text18.Text = "B"
End Sub

Private Sub a19_Click()
Text19.Text = "S"
End Sub

Private Sub b19_Click()
Text19.Text = "B"
End Sub

Private Sub a20_Click()
Text20.Text = "S"
nextbtn.Enabled = True
End Sub

Private Sub b20_Click()
Text20.Text = "B"
nextbtn.Enabled = True
End Sub


Private Sub nextbtn_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Or Text11.Text = "" Or Text12.Text = "" Or Text13.Text = "" Or Text14.Text = "" Or Text15.Text = "" Or Text16.Text = "" Or Text17.Text = "" Or Text18.Text = "" Or Text19.Text = "" Or Text20.Text = "" Then
    MsgBox "Data Belum Lengkap"
    Else
    Dim tambahdata As String
        tambahdata = "Insert Into hasiltest12 values ('" & id.Text & "','" & nama.Text & "','" & tanggal.Text & "','" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "','" & Text7.Text & "','" & Text8.Text & "','" & Text9.Text & "','" & Text10.Text & "','" & Text11.Text & "','" & Text12.Text & "','" & Text13.Text & "','" & Text14.Text & "','" & Text15.Text & "','" & Text16.Text & "','" & Text17.Text & "','" & Text18.Text & "','" & Text19.Text & "','" & Text20.Text & "')"
        koneksi.Execute tambahdata
        MsgBox "TEST 12 Berhasil Di simpan, Silakan lanjutkan ke TEST 13", vbInformation, "Pemberitahuan"
    frmpsi13.Show
    koneksi.Close
    Unload Me
    End If

End Sub

Private Sub Option12_Click()

End Sub

'Private Sub SSTab1_Click(PreviousTab As Integer)
'If SSTab1.Caption = "SOAL NOMOR 21 - 40" Then
'nextbtn.Enabled = True
'Else
'nextbtn.Enabled = False
'End If
'End Sub





