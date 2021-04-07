VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmpsi2 
   Caption         =   "SOAL PSIKOTES TEST II"
   ClientHeight    =   10425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   ScaleHeight     =   10425
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton nextbtn 
      Caption         =   "NEXT >>>>>"
      DisabledPicture =   "frmpsi.frx":0000
      DownPicture     =   "frmpsi.frx":13EA
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
      Picture         =   "frmpsi.frx":27D4
      TabIndex        =   13
      Top             =   240
      Width           =   1695
   End
   Begin VB.TextBox tanggal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   10
      Top             =   840
      Width           =   3195
   End
   Begin VB.TextBox nama 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Top             =   420
      Width           =   3195
   End
   Begin VB.TextBox id 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   30
      Width           =   3195
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7815
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   20055
      _ExtentX        =   35375
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "SOAL NOMOR 1 - 10"
      TabPicture(0)   =   "frmpsi.frx":3BBE
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
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "SOAL NOMOR 11 - 20"
      TabPicture(1)   =   "frmpsi.frx":3BDA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(3)=   "Frame4"
      Tab(1).Control(4)=   "Frame5"
      Tab(1).Control(5)=   "Frame6"
      Tab(1).Control(6)=   "Frame7"
      Tab(1).Control(7)=   "Frame8"
      Tab(1).Control(8)=   "Frame9"
      Tab(1).Control(9)=   "Frame10"
      Tab(1).ControlCount=   10
      Begin VB.Frame Frame20 
         Height          =   735
         Left            =   120
         TabIndex        =   130
         Top             =   6840
         Width           =   19815
         Begin VB.OptionButton a10 
            Caption         =   "Karena matahari lebih dekat daripada rembulan."
            Height          =   375
            Left            =   4800
            TabIndex        =   134
            Top             =   240
            Width           =   3135
         End
         Begin VB.OptionButton b10 
            Caption         =   "Karena sinar matahari lebih besar daripada rembulan."
            Height          =   495
            Left            =   7920
            TabIndex        =   133
            Top             =   120
            Width           =   3375
         End
         Begin VB.OptionButton c10 
            Caption         =   "Karena sinar matahari langsung, sedangkan rembulan hanya pantulan."
            Height          =   495
            Left            =   11400
            TabIndex        =   132
            Top             =   120
            Width           =   3735
         End
         Begin VB.TextBox Text10 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   131
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label19 
            Caption         =   "10. Mengapa sinar matahari lebih terang daripada sinar rebulan?"
            Height          =   375
            Left            =   240
            TabIndex        =   135
            Top             =   240
            Width           =   4575
         End
      End
      Begin VB.Frame Frame19 
         Height          =   735
         Left            =   120
         TabIndex        =   124
         Top             =   6120
         Width           =   19815
         Begin VB.OptionButton a9 
            Caption         =   "Karena para pejabat korupsi."
            Height          =   495
            Left            =   4800
            TabIndex        =   128
            Top             =   120
            Width           =   3135
         End
         Begin VB.OptionButton b9 
            Caption         =   "Karena senang di percaya sebagai Negara donor."
            Height          =   495
            Left            =   7920
            TabIndex        =   127
            Top             =   120
            Width           =   3375
         End
         Begin VB.OptionButton c9 
            Caption         =   "Untuk mempercepat pembangunan bangsa."
            Height          =   495
            Left            =   11400
            TabIndex        =   126
            Top             =   120
            Width           =   3495
         End
         Begin VB.TextBox Text9 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   125
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label18 
            Caption         =   "9. Mengapa negara Indonesia ini yang kaya raya banyak hutang?"
            Height          =   495
            Left            =   240
            TabIndex        =   129
            Top             =   120
            Width           =   4575
         End
      End
      Begin VB.Frame Frame18 
         Height          =   735
         Left            =   120
         TabIndex        =   118
         Top             =   5400
         Width           =   19815
         Begin VB.OptionButton a8 
            Caption         =   "Supaya waktu senggangnya terisi."
            Height          =   495
            Left            =   4800
            TabIndex        =   122
            Top             =   120
            Width           =   3015
         End
         Begin VB.OptionButton b8 
            Caption         =   "Supaya bisa memahami etika hubungan dengan lawan jenis."
            Height          =   495
            Left            =   7920
            TabIndex        =   121
            Top             =   120
            Width           =   3135
         End
         Begin VB.OptionButton c8 
            Caption         =   "Supaya cepat tumbuh dewasa."
            Height          =   495
            Left            =   11400
            TabIndex        =   120
            Top             =   120
            Width           =   3375
         End
         Begin VB.TextBox Text8 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   119
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label17 
            Caption         =   "8. Mengapa remaja sebaiknya diajarkan pelajaran seksual?"
            Height          =   375
            Left            =   240
            TabIndex        =   123
            Top             =   240
            Width           =   4575
         End
      End
      Begin VB.Frame Frame17 
         Height          =   735
         Left            =   120
         TabIndex        =   110
         Top             =   4680
         Width           =   19815
         Begin VB.TextBox Text7 
            Height          =   495
            Left            =   18600
            TabIndex        =   116
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton a7 
            Caption         =   "Dapat dijadikan teman bermain."
            Height          =   255
            Left            =   4800
            TabIndex        =   115
            Top             =   240
            Width           =   2535
         End
         Begin VB.OptionButton b7 
            Caption         =   "Dapat menggonggong."
            Height          =   255
            Left            =   7920
            TabIndex        =   114
            Top             =   240
            Width           =   3255
         End
         Begin VB.OptionButton c7 
            Caption         =   "Dapat menjaga "
            Height          =   255
            Left            =   11400
            TabIndex        =   113
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label31 
            Caption         =   "7. Anjing adalah binatang yang berguna, karena ?"
            Height          =   375
            Left            =   240
            TabIndex        =   117
            Top             =   240
            Width           =   4455
         End
      End
      Begin VB.Frame Frame16 
         Height          =   735
         Left            =   120
         TabIndex        =   104
         Top             =   3960
         Width           =   19815
         Begin VB.OptionButton a6 
            Caption         =   "Karena orang desa penghasilannya rendah."
            Height          =   495
            Left            =   4800
            TabIndex        =   108
            Top             =   120
            Width           =   3135
         End
         Begin VB.OptionButton b6 
            Caption         =   "Karena ikatan persaudaraan dan persamaan di desa tinggi."
            Height          =   495
            Left            =   7920
            TabIndex        =   107
            Top             =   120
            Width           =   3135
         End
         Begin VB.OptionButton c6 
            Caption         =   "Karena orang desa tidak materialistis."
            Height          =   495
            Left            =   11400
            TabIndex        =   106
            Top             =   120
            Width           =   3375
         End
         Begin VB.TextBox Text6 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   105
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label15 
            Caption         =   "6. Mengapa di desa, gotong royong lebih kuat daripada di kota?"
            Height          =   375
            Left            =   240
            TabIndex        =   109
            Top             =   240
            Width           =   4575
         End
      End
      Begin VB.Frame Frame15 
         Height          =   735
         Left            =   120
         TabIndex        =   98
         Top             =   3240
         Width           =   19815
         Begin VB.OptionButton a5 
            Caption         =   "Rekreasi."
            Height          =   495
            Left            =   4800
            TabIndex        =   102
            Top             =   120
            Width           =   3015
         End
         Begin VB.OptionButton b5 
            Caption         =   "Cadangan kayu alam."
            Height          =   495
            Left            =   7920
            TabIndex        =   101
            Top             =   120
            Width           =   3015
         End
         Begin VB.OptionButton c5 
            Caption         =   "Sebagai cagar alam pelindung ekosistem."
            Height          =   495
            Left            =   11400
            TabIndex        =   100
            Top             =   120
            Width           =   3255
         End
         Begin VB.TextBox Text5 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   99
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label14 
            Caption         =   "5. Apa kegunaan hutan?"
            Height          =   375
            Left            =   240
            TabIndex        =   103
            Top             =   240
            Width           =   4575
         End
      End
      Begin VB.Frame Frame14 
         Height          =   735
         Left            =   120
         TabIndex        =   92
         Top             =   2520
         Width           =   19815
         Begin VB.OptionButton a4 
            Caption         =   "Supaya tidak dianggap sombong."
            Height          =   495
            Left            =   4800
            TabIndex        =   96
            Top             =   120
            Width           =   3015
         End
         Begin VB.OptionButton b4 
            Caption         =   "Pengiritan nasional."
            Height          =   495
            Left            =   7920
            TabIndex        =   95
            Top             =   120
            Width           =   3015
         End
         Begin VB.OptionButton c4 
            Caption         =   "Supaya jalan raya tidak terlalu sesak."
            Height          =   495
            Left            =   11400
            TabIndex        =   94
            Top             =   120
            Width           =   3375
         End
         Begin VB.TextBox Text4 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   93
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label13 
            Caption         =   "4. Mengapa sebaiknya orang naik bis daripada naik mobil pribadi?"
            Height          =   495
            Left            =   240
            TabIndex        =   97
            Top             =   120
            Width           =   4575
         End
      End
      Begin VB.Frame Frame13 
         Height          =   735
         Left            =   120
         TabIndex        =   86
         Top             =   1800
         Width           =   19815
         Begin VB.OptionButton a3 
            Caption         =   "Membuat suatu usaha atau perusahaan."
            Height          =   495
            Left            =   4800
            TabIndex        =   90
            Top             =   120
            Width           =   3015
         End
         Begin VB.OptionButton b3 
            Caption         =   "Membeli mobil,rumah dan menikah."
            Height          =   495
            Left            =   7920
            TabIndex        =   89
            Top             =   120
            Width           =   3015
         End
         Begin VB.OptionButton c3 
            Caption         =   "Mendermakan sebagian uang kepada fakir miskin."
            Height          =   495
            Left            =   11400
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   120
            Width           =   3255
         End
         Begin VB.TextBox Text3 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   87
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label12 
            Caption         =   "3. Bila anda mempunyai satu milyar rupiah di tangan, apa yang akan anda lakuan?"
            Height          =   495
            Left            =   240
            TabIndex        =   91
            Top             =   120
            Width           =   4575
         End
      End
      Begin VB.Frame Frame12 
         Height          =   735
         Left            =   120
         TabIndex        =   80
         Top             =   1080
         Width           =   19815
         Begin VB.OptionButton a2 
            Caption         =   "Karya palsu."
            Height          =   495
            Left            =   4800
            TabIndex        =   84
            Top             =   120
            Width           =   3015
         End
         Begin VB.OptionButton b2 
            Caption         =   "Penjiplak karya orang lain."
            Height          =   495
            Left            =   7920
            TabIndex        =   83
            Top             =   120
            Width           =   3015
         End
         Begin VB.OptionButton c2 
            Caption         =   "Dokumen yang diragukan keotentikannya."
            Height          =   495
            Left            =   11400
            TabIndex        =   82
            Top             =   120
            Width           =   3495
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   81
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label11 
            Caption         =   "2. Apa yang di maksud dengan PLAGIATOR ?"
            Height          =   375
            Left            =   240
            TabIndex        =   85
            Top             =   240
            Width           =   4575
         End
      End
      Begin VB.Frame Frame11 
         Height          =   735
         Left            =   120
         TabIndex        =   74
         Top             =   360
         Width           =   19815
         Begin VB.OptionButton a1 
            Caption         =   "Lari dan mencari penyebab kebakaran."
            Height          =   495
            Left            =   4800
            TabIndex        =   78
            Top             =   120
            Width           =   3015
         End
         Begin VB.OptionButton b1 
            Caption         =   "Keluar secara tenang melalui pintu darurat."
            Height          =   495
            Left            =   7920
            TabIndex        =   77
            Top             =   120
            Width           =   2895
         End
         Begin VB.OptionButton c1 
            Caption         =   "Cari pemadam kebakaran"
            Height          =   495
            Left            =   11400
            TabIndex        =   76
            Top             =   120
            Width           =   3375
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   75
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label10 
            Caption         =   "1. Apabila di rumah tetangga anda terjadi kebakaran, apa yang harus di lakukan?"
            Height          =   495
            Left            =   240
            TabIndex        =   79
            Top             =   120
            Width           =   4575
         End
      End
      Begin VB.Frame Frame10 
         Height          =   735
         Left            =   -74880
         TabIndex        =   37
         Top             =   6840
         Width           =   19815
         Begin VB.TextBox Text20 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   73
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton c20 
            Caption         =   "Sebagai pelindung kepala agar bila terjadi kecelakaan terlindung dari benturan benda keras"
            Height          =   495
            Left            =   11520
            TabIndex        =   63
            Top             =   120
            Width           =   3855
         End
         Begin VB.OptionButton b20 
            Caption         =   "Sebagai pelindung telinga dari bisingnya lalu lintas."
            Height          =   495
            Left            =   8280
            TabIndex        =   62
            Top             =   120
            Width           =   3015
         End
         Begin VB.OptionButton a20 
            Caption         =   "Agar tidak ditangkap polisi."
            Height          =   495
            Left            =   4800
            TabIndex        =   61
            Top             =   120
            Width           =   3255
         End
         Begin VB.Label Label29 
            Caption         =   "20. Apakah gunanya help bagi pengendara motor?"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   4455
         End
      End
      Begin VB.Frame Frame9 
         Height          =   735
         Left            =   -74880
         TabIndex        =   26
         Top             =   6120
         Width           =   19815
         Begin VB.TextBox Text19 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   72
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton c19 
            Caption         =   "Sebab sinar matahari jarang menyinari puncak gunung."
            Height          =   495
            Left            =   11520
            TabIndex        =   59
            Top             =   120
            Width           =   3975
         End
         Begin VB.OptionButton b19 
            Caption         =   "Makin tinggi suatu tempat, makin rendah suhu tempat itu."
            Height          =   495
            Left            =   8280
            TabIndex        =   58
            Top             =   120
            Width           =   3015
         End
         Begin VB.OptionButton a19 
            Caption         =   "Sebab puncak gunung jayawijaya dekat dengan awan."
            Height          =   495
            Left            =   4800
            TabIndex        =   57
            Top             =   120
            Width           =   3255
         End
         Begin VB.Label Label28 
            Caption         =   "19. Mengapa puncak gunung jayawijaya di selimuti salju?"
            Height          =   375
            Left            =   120
            TabIndex        =   56
            Top             =   240
            Width           =   4215
         End
      End
      Begin VB.Frame Frame8 
         Height          =   735
         Left            =   -74880
         TabIndex        =   25
         Top             =   5400
         Width           =   19815
         Begin VB.TextBox Text18 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   71
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton c18 
            Caption         =   "Untuk mengisi waktu luang."
            Height          =   495
            Left            =   11520
            TabIndex        =   55
            Top             =   120
            Width           =   3975
         End
         Begin VB.OptionButton b18 
            Caption         =   "Supaya mendapat ilmu yang dapat di amalkan di masa depan."
            Height          =   495
            Left            =   8160
            TabIndex        =   54
            Top             =   120
            Width           =   3135
         End
         Begin VB.OptionButton a18 
            Caption         =   "Supaya nantinya mudah mencari pekerjaan."
            Height          =   495
            Left            =   4800
            TabIndex        =   53
            Top             =   120
            Width           =   3375
         End
         Begin VB.Label Label27 
            Caption         =   "18. Mengapa belajar itu perlu?"
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   240
            Width           =   4335
         End
      End
      Begin VB.Frame Frame7 
         Height          =   735
         Left            =   -74880
         TabIndex        =   24
         Top             =   4680
         Width           =   19815
         Begin VB.TextBox Text17 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   70
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton c17 
            Caption         =   "Agar tidak di anggap kampungan."
            Height          =   495
            Left            =   11520
            TabIndex        =   51
            Top             =   120
            Width           =   3855
         End
         Begin VB.OptionButton b17 
            Caption         =   "Untuk menjaga harga diri dan perasaan orang lain."
            Height          =   495
            Left            =   8160
            TabIndex        =   50
            Top             =   120
            Width           =   3255
         End
         Begin VB.OptionButton a17 
            Caption         =   "Agar orang lain tidak membenci kita."
            Height          =   255
            Left            =   4800
            TabIndex        =   49
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label26 
            Caption         =   "17. Apa gunanya sopan santun?"
            Height          =   375
            Left            =   120
            TabIndex        =   48
            Top             =   240
            Width           =   4215
         End
      End
      Begin VB.Frame Frame6 
         Height          =   735
         Left            =   -74880
         TabIndex        =   23
         Top             =   3960
         Width           =   19815
         Begin VB.TextBox Text16 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   69
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton c16 
            Caption         =   "Karena emosi."
            Height          =   255
            Left            =   11520
            TabIndex        =   47
            Top             =   240
            Width           =   3855
         End
         Begin VB.OptionButton b16 
            Caption         =   "Karena malu kepada korban."
            Height          =   495
            Left            =   8160
            TabIndex        =   46
            Top             =   120
            Width           =   3255
         End
         Begin VB.OptionButton a16 
            Caption         =   "Karena untuk menyelamatkan diri."
            Height          =   495
            Left            =   4800
            TabIndex        =   45
            Top             =   120
            Width           =   3375
         End
         Begin VB.Label Label25 
            Caption         =   "16. Mengapa penjahat yang terjepit sering membunuh korban?"
            Height          =   375
            Left            =   120
            TabIndex        =   44
            Top             =   120
            Width           =   4095
         End
      End
      Begin VB.Frame Frame5 
         Height          =   735
         Left            =   -74880
         TabIndex        =   22
         Top             =   3240
         Width           =   19815
         Begin VB.TextBox Text15 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   68
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton c15 
            Caption         =   "Untuk memberi contoh kepada anak-anak."
            Height          =   495
            Left            =   11520
            TabIndex        =   43
            Top             =   120
            Width           =   3975
         End
         Begin VB.OptionButton b15 
            Caption         =   "Untuk menjaga gengsi dengan relasi bisnis."
            Height          =   495
            Left            =   8160
            TabIndex        =   42
            Top             =   120
            Width           =   3135
         End
         Begin VB.OptionButton a15 
            Caption         =   "Untuk menambah wawasan dan pengetahuan."
            Height          =   495
            Left            =   4800
            TabIndex        =   41
            Top             =   120
            Width           =   3255
         End
         Begin VB.Label Label24 
            Caption         =   "15. Mengapa membaca buku itu penting?"
            Height          =   375
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   3975
         End
      End
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   -74880
         TabIndex        =   21
         Top             =   2520
         Width           =   19815
         Begin VB.TextBox Text14 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   67
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton c14 
            Caption         =   "Banjir di musim hujan dan kering di musim panas."
            Height          =   375
            Left            =   11520
            TabIndex        =   39
            Top             =   120
            Width           =   3855
         End
         Begin VB.OptionButton b14 
            Caption         =   "Banyak tanah kosong untuk di buat perkebunan."
            Height          =   375
            Left            =   8160
            TabIndex        =   38
            Top             =   120
            Width           =   3135
         End
         Begin VB.OptionButton a14 
            Caption         =   "Fauna dan flora di muka bumi ini akan lenyap."
            Height          =   375
            Left            =   4800
            TabIndex        =   36
            Top             =   120
            Width           =   2775
         End
         Begin VB.Label Label23 
            Caption         =   "14. Apa yang terjadi bila hutan di gunduli?"
            Height          =   375
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   3855
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   -74880
         TabIndex        =   20
         Top             =   1800
         Width           =   19815
         Begin VB.TextBox Text13 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   66
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton c13 
            Caption         =   "Karena matahari lebih lama daripada di daerah kutub."
            Height          =   375
            Left            =   11520
            TabIndex        =   34
            Top             =   120
            Width           =   3615
         End
         Begin VB.OptionButton b13 
            Caption         =   "Karena matahari jatuhnya miring daripada di daerah kutub."
            Height          =   375
            Left            =   8160
            TabIndex        =   33
            Top             =   120
            Width           =   3015
         End
         Begin VB.OptionButton a13 
            Caption         =   "Karena di kutub banyak gunung es."
            Height          =   375
            Left            =   4800
            TabIndex        =   32
            Top             =   120
            Width           =   2655
         End
         Begin VB.Label Label22 
            Caption         =   "13. Mengapa di dekat kutub lebih dingin daripada di dekat khatulistiwa?"
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   120
            Width           =   3855
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   -74880
         TabIndex        =   19
         Top             =   1080
         Width           =   19815
         Begin VB.TextBox Text12 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   65
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton a12 
            Caption         =   "Karena tidak terlatih berbicara."
            Height          =   375
            Left            =   4800
            TabIndex        =   28
            Top             =   120
            Width           =   3015
         End
         Begin VB.OptionButton c12 
            Caption         =   "Karena orang tuli malas bicara."
            Height          =   255
            Left            =   11520
            TabIndex        =   30
            Top             =   240
            Width           =   3615
         End
         Begin VB.OptionButton b12 
            Caption         =   "Karena orang tuli tidak pernah mendengarkan ucapan orang lain."
            Height          =   375
            Left            =   8160
            TabIndex        =   29
            Top             =   120
            Width           =   3015
         End
         Begin VB.Label Label21 
            Caption         =   "12. Mengapa orang tuli biasanya juga bisu?"
            Height          =   375
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   4215
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -74880
         TabIndex        =   14
         Top             =   360
         Width           =   19815
         Begin VB.TextBox Text11 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   64
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton c11 
            Caption         =   "Karena apa yang seseorang kerjakan mencerminkan siapa dia sebenarnya."
            Height          =   495
            Left            =   11520
            TabIndex        =   18
            Top             =   120
            Width           =   3135
         End
         Begin VB.OptionButton b11 
            Caption         =   "Karena orang tuli tidak bisa mendengar."
            Height          =   495
            Left            =   8160
            TabIndex        =   17
            Top             =   120
            Width           =   2895
         End
         Begin VB.OptionButton a11 
            Caption         =   "Biasanya orang senang menceritakan kabar bohong."
            Height          =   495
            Left            =   4800
            TabIndex        =   16
            Top             =   120
            Width           =   2775
         End
         Begin VB.Label Label20 
            Caption         =   "11. Mengapa menilai orang dari apa yang di kerjakannya lebih baik daripada apa yang di katakanya?"
            Height          =   495
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   4215
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
         TabIndex        =   111
         Top             =   -2160
         Width           =   4575
      End
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
      Left            =   480
      TabIndex        =   112
      Top             =   7440
      Width           =   4575
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
      TabIndex        =   11
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label Label8 
      Caption         =   "Tanggal Test :"
      Height          =   195
      Left            =   90
      TabIndex        =   7
      Top             =   900
      Width           =   1305
   End
   Begin VB.Label Label7 
      Caption         =   "NAMA           :"
      Height          =   165
      Left            =   90
      TabIndex        =   6
      Top             =   450
      Width           =   1245
   End
   Begin VB.Label Label6 
      Caption         =   "ID                 :"
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   30
      Width           =   1125
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   20250
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label5 
      Caption         =   $"frmpsi.frx":3BF6
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
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   19995
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
      Left            =   90
      TabIndex        =   3
      Top             =   1200
      Width           =   2085
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   "Petunjuk Khusus :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   -1740
      TabIndex        =   2
      Top             =   -4920
      Width           =   1905
   End
   Begin VB.Label Label2 
      Caption         =   "Test Pemahaman Sederhana"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8430
      TabIndex        =   1
      Top             =   750
      Width           =   3405
   End
   Begin VB.Label Label1 
      Caption         =   "TEST II"
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
      Left            =   9180
      TabIndex        =   0
      Top             =   150
      Width           =   1635
   End
End
Attribute VB_Name = "frmpsi2"
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

Private Sub nextbtn_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Or Text11.Text = "" Or Text12.Text = "" Or Text13.Text = "" Or Text14.Text = "" Or Text15.Text = "" Or Text16.Text = "" Or Text17.Text = "" Or Text18.Text = "" Or Text19.Text = "" Or Text20.Text = "" Then
    MsgBox "Data Belum Lengkap"
    Else
    Dim tambahdata As String
        tambahdata = "Insert Into hasiltest2 values ('" & id.Text & "','" & nama.Text & "','" & tanggal.Text & "','" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "','" & Text7.Text & "','" & Text8.Text & "','" & Text9.Text & "','" & Text10.Text & "','" & Text11.Text & "','" & Text12.Text & "','" & Text13.Text & "','" & Text14.Text & "','" & Text15.Text & "','" & Text16.Text & "','" & Text17.Text & "','" & Text18.Text & "','" & Text19.Text & "','" & Text20.Text & "')"
        koneksi.Execute tambahdata
        MsgBox "TEST 2 Berhasil Di simpan, Silakan lanjutkan ke TEST 3", vbInformation, "Pemberitahuan"
    frmpsi3.Show
    koneksi.Close
    Unload Me
    End If

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If SSTab1.Caption = "SOAL NOMOR 1 - 10" Then
nextbtn.Enabled = False
Else
nextbtn.Enabled = True
End If
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

Private Sub a2_Click()
Text2.Text = "A"
a2.Value = True
End Sub

Private Sub b2_Click()
Text2.Text = "B"
End Sub

Private Sub c2_Click()
Text2.Text = "C"
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

Private Sub a4_Click()
Text4.Text = "A"
End Sub

Private Sub b4_Click()
Text4.Text = "B"
End Sub

Private Sub c4_Click()
Text4.Text = "C"
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

Private Sub a6_Click()
Text6.Text = "A"
End Sub

Private Sub b6_Click()
Text6.Text = "B"
End Sub

Private Sub c6_Click()
Text6.Text = "C"
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

Private Sub a8_Click()
Text8.Text = "A"
End Sub

Private Sub b8_Click()
Text8.Text = "B"
End Sub

Private Sub c8_Click()
Text8.Text = "C"
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

Private Sub a10_Click()
Text10.Text = "A"
End Sub

Private Sub b10_Click()
Text10.Text = "B"
End Sub

Private Sub c10_Click()
Text10.Text = "C"
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

Private Sub a12_Click()
Text12.Text = "A"
End Sub

Private Sub b12_Click()
Text12.Text = "B"
End Sub

Private Sub c12_Click()
Text12.Text = "C"
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

Private Sub a14_Click()
Text14.Text = "A"
End Sub

Private Sub b14_Click()
Text14.Text = "B"
End Sub

Private Sub c14_Click()
Text14.Text = "C"
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

Private Sub a16_Click()
Text16.Text = "A"
End Sub

Private Sub b16_Click()
Text16.Text = "B"
End Sub

Private Sub c16_Click()
Text16.Text = "C"
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

Private Sub a18_Click()
Text18.Text = "A"
End Sub

Private Sub b18_Click()
Text18.Text = "B"
End Sub

Private Sub c18_Click()
Text18.Text = "C"
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

Private Sub a20_Click()
Text20.Text = "A"
End Sub

Private Sub b20_Click()
Text20.Text = "B"
End Sub

Private Sub c20_Click()
Text20.Text = "C"
End Sub




