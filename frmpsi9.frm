VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmpsi9 
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
      Height          =   285
      Left            =   1170
      TabIndex        =   3
      Top             =   0
      Width           =   3195
   End
   Begin VB.TextBox nama 
      Height          =   285
      Left            =   1170
      TabIndex        =   2
      Top             =   390
      Width           =   3195
   End
   Begin VB.TextBox tanggal 
      Height          =   285
      Left            =   1170
      TabIndex        =   1
      Top             =   810
      Width           =   3195
   End
   Begin VB.CommandButton nextbtn 
      Caption         =   "NEXT >>>>>"
      DisabledPicture =   "frmpsi9.frx":0000
      DownPicture     =   "frmpsi9.frx":13EA
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
      Picture         =   "frmpsi9.frx":27D4
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
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "SOAL NOMOR 1 - 10"
      TabPicture(0)   =   "frmpsi9.frx":3BBE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame20"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame19"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame18"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame17"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame16"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame15"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame14"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame13"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame12"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Frame11"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "SOAL NOMOR 11 - 15"
      TabPicture(1)   =   "frmpsi9.frx":3BDA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(1)=   "Frame4"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(3)=   "Frame2"
      Tab(1).Control(4)=   "Frame1"
      Tab(1).ControlCount=   5
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   -74880
         TabIndex        =   89
         Top             =   360
         Width           =   19815
         Begin VB.OptionButton a11 
            Caption         =   "Mungkin Seto membenci Sudirman."
            Height          =   495
            Left            =   7920
            TabIndex        =   93
            Top             =   120
            Width           =   2775
         End
         Begin VB.OptionButton b11 
            Caption         =   "Seto tidak membenci Sudirman."
            Height          =   495
            Left            =   11160
            TabIndex        =   92
            Top             =   120
            Width           =   2895
         End
         Begin VB.OptionButton c11 
            Caption         =   "Seto tidak membenci Niken."
            Height          =   495
            Left            =   14520
            TabIndex        =   91
            Top             =   120
            Width           =   3135
         End
         Begin VB.TextBox Text11 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   90
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label20 
            Caption         =   "11. Seto membenci Shinta dan kawan-kawannya. Niken teman Shinta. Shinta telah bersuami. Sudirman adalah teman suaminya."
            Height          =   495
            Left            =   120
            TabIndex        =   94
            Top             =   120
            Width           =   7215
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   -74880
         TabIndex        =   83
         Top             =   1080
         Width           =   19815
         Begin VB.OptionButton b12 
            Caption         =   "Mungkin sate kelinci sama harganya dengan sate kambing."
            Height          =   375
            Left            =   11160
            TabIndex        =   87
            Top             =   120
            Width           =   3015
         End
         Begin VB.OptionButton c12 
            Caption         =   "Sate kambing harganya paling mahal."
            Height          =   375
            Left            =   14520
            TabIndex        =   86
            Top             =   120
            Width           =   3615
         End
         Begin VB.OptionButton a12 
            Caption         =   "Harga telur lebih mahal daripada sate kambing."
            Height          =   375
            Left            =   7920
            TabIndex        =   85
            Top             =   120
            Width           =   3015
         End
         Begin VB.TextBox Text12 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   84
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label21 
            Caption         =   $"frmpsi9.frx":3BF6
            Height          =   495
            Left            =   120
            TabIndex        =   88
            Top             =   120
            Width           =   7455
         End
      End
      Begin VB.Frame Frame3 
         Height          =   855
         Left            =   -74880
         TabIndex        =   77
         Top             =   1800
         Width           =   19815
         Begin VB.OptionButton a13 
            Caption         =   "Bisa jadi Wiwin bukan anak yang pandai dan rajin."
            Height          =   615
            Left            =   7920
            TabIndex        =   81
            Top             =   120
            Width           =   3135
         End
         Begin VB.OptionButton b13 
            Caption         =   "Wiwin anak yang pandai dan malas."
            Height          =   615
            Left            =   11160
            TabIndex        =   80
            Top             =   120
            Width           =   3255
         End
         Begin VB.OptionButton c13 
            Caption         =   "Wiwin mungkin anak yang malas."
            Height          =   615
            Left            =   14520
            TabIndex        =   79
            Top             =   120
            Width           =   3615
         End
         Begin VB.TextBox Text13 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   78
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label22 
            Caption         =   $"frmpsi9.frx":3C8B
            Height          =   615
            Left            =   120
            TabIndex        =   82
            Top             =   120
            Width           =   7575
         End
      End
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   -74880
         TabIndex        =   71
         Top             =   2640
         Width           =   19815
         Begin VB.OptionButton a14 
            Caption         =   "Ali tidak mungkin sedang main ketempat Anton."
            Height          =   495
            Left            =   7920
            TabIndex        =   75
            Top             =   120
            Width           =   2895
         End
         Begin VB.OptionButton b14 
            Caption         =   "Ali pasti sedang main ke rumah Endri."
            Height          =   495
            Left            =   11160
            TabIndex        =   74
            Top             =   120
            Width           =   3135
         End
         Begin VB.OptionButton c14 
            Caption         =   "Kemungkinan besar Ali main ke rumah Endri, tapi tidak menutup kemungkinan sedang di rumah Anton"
            Height          =   495
            Left            =   14520
            TabIndex        =   73
            Top             =   120
            Width           =   4095
         End
         Begin VB.TextBox Text14 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   72
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label23 
            Caption         =   $"frmpsi9.frx":3DAC
            Height          =   495
            Left            =   120
            TabIndex        =   76
            Top             =   120
            Width           =   7695
         End
      End
      Begin VB.Frame Frame5 
         Height          =   855
         Left            =   -74880
         TabIndex        =   65
         Top             =   3360
         Width           =   19815
         Begin VB.OptionButton a15 
            Caption         =   "Lampu neon buatan Orion lebih mahal daripada lampu neon buatan Maspion."
            Height          =   615
            Left            =   7920
            TabIndex        =   69
            Top             =   120
            Width           =   3255
         End
         Begin VB.OptionButton b15 
            Caption         =   "Lampu neon 10 watt lebih mahal daripada bola lampu 10 watt."
            Height          =   615
            Left            =   11160
            TabIndex        =   68
            Top             =   120
            Width           =   3135
         End
         Begin VB.OptionButton c15 
            Caption         =   "Bola lampu 10 watt buatan Orion lebih mahal daripada bola lampu 10 watt buatan Maspion."
            Height          =   615
            Left            =   14520
            TabIndex        =   67
            Top             =   120
            Width           =   3975
         End
         Begin VB.TextBox Text15 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   66
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label24 
            Caption         =   $"frmpsi9.frx":3E4B
            Height          =   615
            Left            =   120
            TabIndex        =   70
            Top             =   120
            Width           =   7815
         End
      End
      Begin VB.Frame Frame11 
         Height          =   735
         Left            =   120
         TabIndex        =   59
         Top             =   360
         Width           =   19815
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   63
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton c1 
            Caption         =   "Pada zaman krisis moneter, barang kebutuhan sehari-hari sulit didapat."
            Height          =   495
            Left            =   14400
            TabIndex        =   62
            Top             =   120
            Width           =   3375
         End
         Begin VB.OptionButton b1 
            Caption         =   "Harga beras dan telur tidak stabil."
            Height          =   495
            Left            =   11160
            TabIndex        =   61
            Top             =   120
            Width           =   3135
         End
         Begin VB.OptionButton a1 
            Caption         =   "Harga beras dan telur naik."
            Height          =   495
            Left            =   7920
            TabIndex        =   60
            Top             =   120
            Width           =   3135
         End
         Begin VB.Label Label10 
            Caption         =   "1. Pada zaman krisis moneter ini, harga barang kebutuhan sehari-hari tidak stabil, beras dan telur adalah kebutuhan sehari-hari."
            Height          =   495
            Left            =   120
            TabIndex        =   64
            Top             =   120
            Width           =   4815
         End
      End
      Begin VB.Frame Frame12 
         Height          =   735
         Left            =   120
         TabIndex        =   53
         Top             =   1080
         Width           =   19815
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   57
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton c2 
            Caption         =   "Menteri Luar Negeri Filipina sedang ikut konfrensi ASEAN di Jakarta."
            Height          =   495
            Left            =   14400
            TabIndex        =   56
            Top             =   120
            Width           =   3495
         End
         Begin VB.OptionButton b2 
            Caption         =   "Filipina belum tentu anggota ASEAN."
            Height          =   495
            Left            =   11160
            TabIndex        =   55
            Top             =   120
            Width           =   3015
         End
         Begin VB.OptionButton a2 
            Caption         =   "Filipina adalah anggota ASEAN."
            Height          =   495
            Left            =   7920
            TabIndex        =   54
            Top             =   120
            Width           =   3135
         End
         Begin VB.Label Label11 
            Caption         =   $"frmpsi9.frx":3F1E
            Height          =   495
            Left            =   120
            TabIndex        =   58
            Top             =   120
            Width           =   6135
         End
      End
      Begin VB.Frame Frame13 
         Height          =   735
         Left            =   120
         TabIndex        =   47
         Top             =   1800
         Width           =   19815
         Begin VB.TextBox Text3 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   51
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton c3 
            Caption         =   "Minyak tanah adalah salah satu sumber energi."
            Height          =   495
            Left            =   14400
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   120
            Width           =   3255
         End
         Begin VB.OptionButton b3 
            Caption         =   "Minyak tanah bisa di perbaharui."
            Height          =   495
            Left            =   11160
            TabIndex        =   49
            Top             =   120
            Width           =   3015
         End
         Begin VB.OptionButton a3 
            Caption         =   "Minyak tanah adalah sumber energi yang tidak bisa di perbaharui."
            Height          =   495
            Left            =   7920
            TabIndex        =   48
            Top             =   120
            Width           =   3135
         End
         Begin VB.Label Label12 
            Caption         =   "3. Semua sumber energi yang di gali dari dalam tanah tidak bisa diperbaharui. Minyak tanah adalah sumber energi dari dalam tanah."
            Height          =   495
            Left            =   120
            TabIndex        =   52
            Top             =   120
            Width           =   5895
         End
      End
      Begin VB.Frame Frame14 
         Height          =   735
         Left            =   120
         TabIndex        =   41
         Top             =   2520
         Width           =   19815
         Begin VB.TextBox Text4 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   45
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton c4 
            Caption         =   "Pak Cokro tidak beruban."
            Height          =   495
            Left            =   14400
            TabIndex        =   44
            Top             =   120
            Width           =   3375
         End
         Begin VB.OptionButton b4 
            Caption         =   "Pak Cokro hampir beruban."
            Height          =   495
            Left            =   11160
            TabIndex        =   43
            Top             =   120
            Width           =   3015
         End
         Begin VB.OptionButton a4 
            Caption         =   "Pak Cokro sudah beruban."
            Height          =   495
            Left            =   7920
            TabIndex        =   42
            Top             =   120
            Width           =   3015
         End
         Begin VB.Label Label13 
            Caption         =   "4.Pria yang sudah di atas 50 tahun selalu beruban. Pak Cokro adalah seoarang pria, ia berusia 49 tahun."
            Height          =   495
            Left            =   120
            TabIndex        =   46
            Top             =   120
            Width           =   4575
         End
      End
      Begin VB.Frame Frame15 
         Height          =   735
         Left            =   120
         TabIndex        =   35
         Top             =   3240
         Width           =   19815
         Begin VB.TextBox Text5 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   39
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton c5 
            Caption         =   "Intan adalah seorang siswa pelajar yang rajin mengerjakan PR."
            Height          =   495
            Left            =   14400
            TabIndex        =   38
            Top             =   120
            Width           =   3255
         End
         Begin VB.OptionButton b5 
            Caption         =   "Intan adalah seorang pelajar."
            Height          =   495
            Left            =   11160
            TabIndex        =   37
            Top             =   120
            Width           =   3015
         End
         Begin VB.OptionButton a5 
            Caption         =   "Intan adalah seorang pelajar."
            Height          =   495
            Left            =   7920
            TabIndex        =   36
            Top             =   120
            Width           =   3015
         End
         Begin VB.Label Label14 
            Caption         =   $"frmpsi9.frx":3FB9
            Height          =   375
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   5775
         End
      End
      Begin VB.Frame Frame16 
         Height          =   735
         Left            =   120
         TabIndex        =   29
         Top             =   3960
         Width           =   19815
         Begin VB.TextBox Text6 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   33
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton c6 
            Caption         =   "Piring tidak termasuk barang pecah belah."
            Height          =   495
            Left            =   14400
            TabIndex        =   32
            Top             =   120
            Width           =   3375
         End
         Begin VB.OptionButton b6 
            Caption         =   "Piring dari bahan melamin termasuk barang pecah belah."
            Height          =   495
            Left            =   11160
            TabIndex        =   31
            Top             =   120
            Width           =   3135
         End
         Begin VB.OptionButton a6 
            Caption         =   "Piring dari bahan melamin tidak termasuk barang pecah belah."
            Height          =   495
            Left            =   7920
            TabIndex        =   30
            Top             =   120
            Width           =   3135
         End
         Begin VB.Label Label15 
            Caption         =   $"frmpsi9.frx":405D
            Height          =   495
            Left            =   120
            TabIndex        =   34
            Top             =   120
            Width           =   7455
         End
      End
      Begin VB.Frame Frame17 
         Height          =   735
         Left            =   120
         TabIndex        =   23
         Top             =   4680
         Width           =   19815
         Begin VB.OptionButton c7 
            Caption         =   "Ibu hanya suka kain dari bahan katun."
            Height          =   255
            Left            =   14400
            TabIndex        =   27
            Top             =   240
            Width           =   3255
         End
         Begin VB.OptionButton b7 
            Caption         =   "Kemarin ibu hanya membeli kain dari bahan katun di Toko Condong."
            Height          =   495
            Left            =   11160
            TabIndex        =   26
            Top             =   120
            Width           =   3135
         End
         Begin VB.OptionButton a7 
            Caption         =   "Mungkin ibu mau membeli kain dari bahan katun di Toko Condong"
            Height          =   495
            Left            =   7920
            TabIndex        =   25
            Top             =   120
            Width           =   3135
         End
         Begin VB.TextBox Text7 
            Height          =   495
            Left            =   18600
            TabIndex        =   24
            Top             =   120
            Width           =   615
         End
         Begin VB.Label Label31 
            Caption         =   $"frmpsi9.frx":412D
            Height          =   495
            Left            =   120
            TabIndex        =   28
            Top             =   120
            Width           =   7575
         End
      End
      Begin VB.Frame Frame18 
         Height          =   735
         Left            =   120
         TabIndex        =   17
         Top             =   5400
         Width           =   19815
         Begin VB.TextBox Text8 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   21
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton c8 
            Caption         =   "Perbuatan Markum tak harus di hukum."
            Height          =   495
            Left            =   14400
            TabIndex        =   20
            Top             =   120
            Width           =   3375
         End
         Begin VB.OptionButton b8 
            Caption         =   "Perbuatan Markum belum tentu tak terpuji."
            Height          =   495
            Left            =   11160
            TabIndex        =   19
            Top             =   120
            Width           =   3135
         End
         Begin VB.OptionButton a8 
            Caption         =   "Sekarang markum wajib di hukum."
            Height          =   495
            Left            =   7920
            TabIndex        =   18
            Top             =   120
            Width           =   3015
         End
         Begin VB.Label Label17 
            Caption         =   "8. Menyontek itu perbuatan tak terpuji. Para penyontek wajib di hukum. Markum pernah menyontek waktu menjadi mahasiswa."
            Height          =   495
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   7455
         End
      End
      Begin VB.Frame Frame19 
         Height          =   735
         Left            =   120
         TabIndex        =   11
         Top             =   6120
         Width           =   19815
         Begin VB.TextBox Text9 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   15
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton c9 
            Caption         =   "Jamur adalah tanaman."
            Height          =   495
            Left            =   14400
            TabIndex        =   14
            Top             =   120
            Width           =   3495
         End
         Begin VB.OptionButton b9 
            Caption         =   "Jamur membutuhkan sinar matahari."
            Height          =   495
            Left            =   11160
            TabIndex        =   13
            Top             =   120
            Width           =   3375
         End
         Begin VB.OptionButton a9 
            Caption         =   "Jamur tidak selalu membutuhkan sinar matahari."
            Height          =   495
            Left            =   7920
            TabIndex        =   12
            Top             =   120
            Width           =   3135
         End
         Begin VB.Label Label18 
            Caption         =   "9. Semua tanaman membutuhkan sinar matahari. Jamur bukan tanaman."
            Height          =   495
            Left            =   120
            TabIndex        =   16
            Top             =   120
            Width           =   7695
         End
      End
      Begin VB.Frame Frame20 
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   6840
         Width           =   19815
         Begin VB.TextBox Text10 
            Enabled         =   0   'False
            Height          =   495
            Left            =   18600
            TabIndex        =   9
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton c10 
            Caption         =   "Mungkin ada petani tidak membutuhkan irigasi."
            Height          =   495
            Left            =   14400
            TabIndex        =   8
            Top             =   120
            Width           =   3735
         End
         Begin VB.OptionButton b10 
            Caption         =   "Mungkin ada petani yang membuat irigasi di luar OPA."
            Height          =   495
            Left            =   11160
            TabIndex        =   7
            Top             =   120
            Width           =   3135
         End
         Begin VB.OptionButton a10 
            Caption         =   "Semua petani taat di atur oleh OPA."
            Height          =   495
            Left            =   7920
            TabIndex        =   6
            Top             =   120
            Width           =   3135
         End
         Begin VB.Label Label19 
            Caption         =   $"frmpsi9.frx":41DD
            Height          =   495
            Left            =   120
            TabIndex        =   10
            Top             =   120
            Width           =   7695
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
         TabIndex        =   95
         Top             =   -2160
         Width           =   4575
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
      TabIndex        =   104
      Top             =   1200
      Width           =   2085
   End
   Begin VB.Label Label9 
      Caption         =   "Waktu 8 menit Soal 15 Nomor."
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
      TabIndex        =   103
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label Label6 
      Caption         =   "ID                 :"
      Height          =   285
      Left            =   30
      TabIndex        =   102
      Top             =   0
      Width           =   1125
   End
   Begin VB.Label Label7 
      Caption         =   "NAMA           :"
      Height          =   165
      Left            =   0
      TabIndex        =   101
      Top             =   420
      Width           =   1245
   End
   Begin VB.Label Label8 
      Caption         =   "Tanggal Test :"
      Height          =   195
      Left            =   0
      TabIndex        =   100
      Top             =   870
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "TEST IX"
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
      TabIndex        =   99
      Top             =   120
      Width           =   1875
   End
   Begin VB.Label Label2 
      Caption         =   "Test Silogisme Sederhana"
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
      Left            =   8880
      TabIndex        =   98
      Top             =   720
      Width           =   2565
   End
   Begin VB.Label Label5 
      Caption         =   $"frmpsi9.frx":4271
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
      Left            =   90
      TabIndex        =   97
      Top             =   1650
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
      TabIndex        =   96
      Top             =   7410
      Width           =   4575
   End
End
Attribute VB_Name = "frmpsi9"
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
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Or Text11.Text = "" Or Text12.Text = "" Or Text13.Text = "" Or Text14.Text = "" Or Text15.Text = "" Then
    MsgBox "Data Belum Lengkap"
    Else
    Dim tambahdata As String
        tambahdata = "Insert Into hasiltest9 values ('" & id.Text & "','" & nama.Text & "','" & tanggal.Text & "','" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "','" & Text7.Text & "','" & Text8.Text & "','" & Text9.Text & "','" & Text10.Text & "','" & Text11.Text & "','" & Text12.Text & "','" & Text13.Text & "','" & Text14.Text & "','" & Text15.Text & "')"
        koneksi.Execute tambahdata
        MsgBox "TEST 9 Berhasil Di simpan, Silakan lanjutkan ke TEST 10", vbInformation, "Pemberitahuan"
    frmpsi10.Show
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




