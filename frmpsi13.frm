VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmpsi13 
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
      DisabledPicture =   "frmpsi13.frx":0000
      DownPicture     =   "frmpsi13.frx":13EA
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
      Picture         =   "frmpsi13.frx":27D4
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
      Left            =   1200
      TabIndex        =   1
      Top             =   390
      Width           =   3195
   End
   Begin VB.TextBox tanggal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   810
      Width           =   3195
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
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "SOAL NOMOR 1 - 10"
      TabPicture(0)   =   "frmpsi13.frx":3BBE
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
      TabPicture(1)   =   "frmpsi13.frx":3BDA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame20"
      Tab(1).Control(1)=   "Frame19"
      Tab(1).Control(2)=   "Frame18"
      Tab(1).Control(3)=   "Frame17"
      Tab(1).Control(4)=   "Frame16"
      Tab(1).Control(5)=   "Frame15"
      Tab(1).Control(6)=   "Frame14"
      Tab(1).Control(7)=   "Frame13"
      Tab(1).Control(8)=   "Frame12"
      Tab(1).Control(9)=   "Frame11"
      Tab(1).ControlCount=   10
      Begin VB.Frame Frame20 
         Height          =   735
         Left            =   -74880
         TabIndex        =   138
         Top             =   6840
         Width           =   19815
         Begin VB.TextBox Text20 
            Height          =   495
            Left            =   18360
            TabIndex        =   143
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton d20 
            Caption         =   "159%"
            Height          =   495
            Left            =   9960
            TabIndex        =   142
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton c20 
            Caption         =   "847,5"
            Height          =   495
            Left            =   7800
            TabIndex        =   141
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton b20 
            Caption         =   "84,75"
            Height          =   495
            Left            =   5640
            TabIndex        =   140
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton a20 
            Caption         =   "84"
            Height          =   495
            Left            =   3480
            TabIndex        =   139
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label28 
            Caption         =   "20. 67 + 17 + 75%                           ="
            Height          =   255
            Left            =   120
            TabIndex        =   144
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame Frame19 
         Height          =   735
         Left            =   -74880
         TabIndex        =   131
         Top             =   6120
         Width           =   19815
         Begin VB.TextBox Text19 
            Height          =   495
            Left            =   18360
            TabIndex        =   136
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton d19 
            Caption         =   "0,49"
            Height          =   495
            Left            =   9960
            TabIndex        =   135
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton c19 
            Caption         =   "4,9"
            Height          =   495
            Left            =   7800
            TabIndex        =   134
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton b19 
            Caption         =   "0,5"
            Height          =   495
            Left            =   5640
            TabIndex        =   133
            Top             =   120
            Width           =   1815
         End
         Begin VB.OptionButton a19 
            Caption         =   "35,14"
            Height          =   495
            Left            =   3480
            TabIndex        =   132
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label27 
            Caption         =   "19.  0,14 + 35%                               ="
            Height          =   255
            Left            =   120
            TabIndex        =   137
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame Frame18 
         Height          =   735
         Left            =   -74880
         TabIndex        =   124
         Top             =   5400
         Width           =   19815
         Begin VB.TextBox Text18 
            Height          =   495
            Left            =   18360
            TabIndex        =   129
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton d18 
            Caption         =   "437,5"
            Height          =   495
            Left            =   9960
            TabIndex        =   128
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton c18 
            Caption         =   "43,76"
            Height          =   495
            Left            =   7800
            TabIndex        =   127
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton b18 
            Caption         =   "43%"
            Height          =   495
            Left            =   5640
            TabIndex        =   126
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton a18 
            Caption         =   "43,75"
            Height          =   495
            Left            =   3480
            TabIndex        =   125
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label Label26 
            Caption         =   "18.  87,5 X 5%                                 ="
            Height          =   255
            Left            =   120
            TabIndex        =   130
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame17 
         Height          =   735
         Left            =   -74880
         TabIndex        =   117
         Top             =   4680
         Width           =   19815
         Begin VB.TextBox Text17 
            Height          =   495
            Left            =   18360
            TabIndex        =   122
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton d17 
            Caption         =   "- 15,5"
            Height          =   555
            Left            =   9960
            TabIndex        =   121
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton c17 
            Caption         =   "15,05"
            Height          =   495
            Left            =   7800
            TabIndex        =   120
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton b17 
            Caption         =   "- 10"
            Height          =   495
            Left            =   5640
            TabIndex        =   119
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton a17 
            Caption         =   "54,8"
            Height          =   495
            Left            =   3480
            TabIndex        =   118
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label Label25 
            Caption         =   "17.  768 : 14 - 65                             ="
            Height          =   255
            Left            =   120
            TabIndex        =   123
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame16 
         Height          =   735
         Left            =   -74880
         TabIndex        =   110
         Top             =   3960
         Width           =   19815
         Begin VB.TextBox Text16 
            Height          =   495
            Left            =   18360
            TabIndex        =   115
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton d16 
            Caption         =   "37656"
            Height          =   555
            Left            =   9960
            TabIndex        =   114
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton c16 
            Caption         =   "696"
            Height          =   495
            Left            =   7800
            TabIndex        =   113
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton b16 
            Caption         =   "2016"
            Height          =   495
            Left            =   5640
            TabIndex        =   112
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton a16 
            Caption         =   "672"
            Height          =   495
            Left            =   3480
            TabIndex        =   111
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label Label24 
            Caption         =   "16.  56 X 12 + 24                            = "
            Height          =   255
            Left            =   120
            TabIndex        =   116
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame15 
         Height          =   735
         Left            =   -74880
         TabIndex        =   103
         Top             =   3240
         Width           =   19815
         Begin VB.TextBox Text15 
            Height          =   495
            Left            =   18360
            TabIndex        =   108
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton d15 
            Caption         =   "3,5"
            Height          =   495
            Left            =   9960
            TabIndex        =   107
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton c15 
            Caption         =   "10"
            Height          =   495
            Left            =   7800
            TabIndex        =   106
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b15 
            Caption         =   "10,6"
            Height          =   555
            Left            =   5640
            TabIndex        =   105
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton a15 
            Caption         =   "31,8"
            Height          =   495
            Left            =   3480
            TabIndex        =   104
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label Label23 
            Caption         =   "15.   699 : 66 X 3                            ="
            Height          =   255
            Left            =   120
            TabIndex        =   109
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame Frame14 
         Height          =   735
         Left            =   -74880
         TabIndex        =   96
         Top             =   2520
         Width           =   19815
         Begin VB.TextBox Text14 
            Height          =   495
            Left            =   18360
            TabIndex        =   101
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton d14 
            Caption         =   "22,5%"
            Height          =   555
            Left            =   9960
            TabIndex        =   100
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton c14 
            Caption         =   "23,5"
            Height          =   495
            Left            =   7800
            TabIndex        =   99
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b14 
            Caption         =   "2,9"
            Height          =   495
            Left            =   5640
            TabIndex        =   98
            Top             =   120
            Width           =   1815
         End
         Begin VB.OptionButton a14 
            Caption         =   "3,2"
            Height          =   495
            Left            =   3480
            TabIndex        =   97
            Top             =   120
            Width           =   2055
         End
         Begin VB.Label Label22 
            Caption         =   "14.   20% + 2,5 + 1,5                       ="
            Height          =   255
            Left            =   120
            TabIndex        =   102
            Top             =   240
            Width           =   2775
         End
      End
      Begin VB.Frame Frame13 
         Height          =   735
         Left            =   -74880
         TabIndex        =   89
         Top             =   1800
         Width           =   19815
         Begin VB.TextBox Text13 
            Height          =   495
            Left            =   18360
            TabIndex        =   94
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton d13 
            Caption         =   "2"
            Height          =   495
            Left            =   9960
            TabIndex        =   93
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton c13 
            Caption         =   "4/5"
            Height          =   555
            Left            =   7800
            TabIndex        =   92
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton b13 
            Caption         =   "10"
            Height          =   495
            Left            =   5640
            TabIndex        =   91
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton a13 
            Caption         =   "9/5"
            Height          =   495
            Left            =   3480
            TabIndex        =   90
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label21 
            Caption         =   "13.   3/5 + 7/5                                ="
            Height          =   255
            Left            =   120
            TabIndex        =   95
            Top             =   240
            Width           =   2655
         End
      End
      Begin VB.Frame Frame12 
         Height          =   735
         Left            =   -74880
         TabIndex        =   82
         Top             =   1080
         Width           =   19815
         Begin VB.TextBox Text12 
            Height          =   495
            Left            =   18360
            TabIndex        =   87
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton d12 
            Caption         =   "168"
            Height          =   555
            Left            =   9960
            TabIndex        =   86
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton c12 
            Caption         =   "281"
            Height          =   495
            Left            =   7800
            TabIndex        =   85
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b12 
            Caption         =   "181"
            Height          =   495
            Left            =   5640
            TabIndex        =   84
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton a12 
            Caption         =   "188"
            Height          =   495
            Left            =   3480
            TabIndex        =   83
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label20 
            Caption         =   "12. 87 + 43 + 51                             ="
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame11 
         Height          =   735
         Left            =   -74880
         TabIndex        =   75
         Top             =   360
         Width           =   19815
         Begin VB.TextBox Text11 
            Height          =   495
            Left            =   18360
            TabIndex        =   80
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton d11 
            Caption         =   "951,2"
            Height          =   495
            Left            =   9960
            TabIndex        =   79
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton c11 
            Caption         =   "951"
            Height          =   555
            Left            =   7800
            TabIndex        =   78
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b11 
            Caption         =   "8512"
            Height          =   495
            Left            =   5640
            TabIndex        =   77
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton a11 
            Caption         =   "9512"
            Height          =   495
            Left            =   3480
            TabIndex        =   76
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label Label19 
            Caption         =   "11.  656 X 14 1/2                           ="
            Height          =   255
            Left            =   120
            TabIndex        =   81
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame10 
         Height          =   735
         Left            =   120
         TabIndex        =   68
         Top             =   6840
         Width           =   19815
         Begin VB.TextBox Text10 
            Height          =   495
            Left            =   18600
            TabIndex        =   73
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton d10 
            Caption         =   "272,9"
            Height          =   495
            Left            =   9480
            TabIndex        =   72
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton c10 
            Caption         =   "272,8"
            Height          =   495
            Left            =   7320
            TabIndex        =   71
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b10 
            Caption         =   "272"
            Height          =   495
            Left            =   5400
            TabIndex        =   70
            Top             =   120
            Width           =   1815
         End
         Begin VB.OptionButton a10 
            Caption         =   "27,28"
            Height          =   495
            Left            =   3480
            TabIndex        =   69
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label18 
            Caption         =   "10. 3547 : 6 + 7                           ="
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame9 
         Height          =   735
         Left            =   120
         TabIndex        =   61
         Top             =   6120
         Width           =   19815
         Begin VB.TextBox Text9 
            Height          =   495
            Left            =   18600
            TabIndex        =   66
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton d9 
            Caption         =   "87"
            Height          =   555
            Left            =   9480
            TabIndex        =   65
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton c9 
            Caption         =   "7,76"
            Height          =   495
            Left            =   7320
            TabIndex        =   64
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b9 
            Caption         =   "8,76"
            Height          =   495
            Left            =   5400
            TabIndex        =   63
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton a9 
            Caption         =   "76"
            Height          =   495
            Left            =   3480
            TabIndex        =   62
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label Label17 
            Caption         =   "9.  7676 : 876                              ="
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame8 
         Height          =   735
         Left            =   120
         TabIndex        =   54
         Top             =   5400
         Width           =   19815
         Begin VB.TextBox Text8 
            Height          =   495
            Left            =   18600
            TabIndex        =   59
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton d8 
            Caption         =   "67764"
            Height          =   495
            Left            =   9480
            TabIndex        =   58
            Top             =   120
            Width           =   2295
         End
         Begin VB.OptionButton c8 
            Caption         =   "96573"
            Height          =   495
            Left            =   7320
            TabIndex        =   57
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b8 
            Caption         =   "65347"
            Height          =   495
            Left            =   5400
            TabIndex        =   56
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton a8 
            Caption         =   "96563"
            Height          =   495
            Left            =   3480
            TabIndex        =   55
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label16 
            Caption         =   "8. 87576  + 8997                         ="
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame7 
         Height          =   735
         Left            =   120
         TabIndex        =   47
         Top             =   4680
         Width           =   19815
         Begin VB.TextBox Text7 
            Height          =   495
            Left            =   18600
            TabIndex        =   52
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton d7 
            Caption         =   "7,68"
            Height          =   495
            Left            =   9480
            TabIndex        =   51
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton c7 
            Caption         =   "11,86"
            Height          =   495
            Left            =   7320
            TabIndex        =   50
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b7 
            Caption         =   "11,38"
            Height          =   495
            Left            =   5400
            TabIndex        =   49
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton a7 
            Caption         =   "11,68"
            Height          =   495
            Left            =   3480
            TabIndex        =   48
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label15 
            Caption         =   "7.  2,13 + 9,55                             ="
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame6 
         Height          =   735
         Left            =   120
         TabIndex        =   40
         Top             =   3960
         Width           =   19815
         Begin VB.TextBox Text6 
            Height          =   495
            Left            =   18600
            TabIndex        =   45
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton d6 
            Caption         =   "666"
            Height          =   435
            Left            =   9480
            TabIndex        =   44
            Top             =   120
            Width           =   2295
         End
         Begin VB.OptionButton c6 
            Caption         =   "777"
            Height          =   495
            Left            =   7320
            TabIndex        =   43
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b6 
            Caption         =   "678"
            Height          =   495
            Left            =   5400
            TabIndex        =   42
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton a6 
            Caption         =   "499"
            Height          =   495
            Left            =   3480
            TabIndex        =   41
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label Label14 
            Caption         =   "6.  0,75 X 888                              ="
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame5 
         Height          =   735
         Left            =   120
         TabIndex        =   33
         Top             =   3240
         Width           =   19815
         Begin VB.TextBox Text5 
            Height          =   495
            Left            =   18600
            TabIndex        =   38
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton d5 
            Caption         =   "121"
            Height          =   495
            Left            =   9480
            TabIndex        =   37
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton c5 
            Caption         =   "235"
            Height          =   495
            Left            =   7320
            TabIndex        =   36
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b5 
            Caption         =   "221"
            Height          =   495
            Left            =   5400
            TabIndex        =   35
            Top             =   120
            Width           =   1815
         End
         Begin VB.OptionButton a5 
            Caption         =   "122"
            Height          =   495
            Left            =   3480
            TabIndex        =   34
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label Label13 
            Caption         =   "5.  188 - 67                                  ="
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   120
         TabIndex        =   26
         Top             =   2520
         Width           =   19815
         Begin VB.TextBox Text4 
            Height          =   525
            Left            =   18600
            TabIndex        =   31
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton d4 
            Caption         =   "324"
            Height          =   495
            Left            =   9480
            TabIndex        =   30
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton c4 
            Caption         =   "248"
            Height          =   495
            Left            =   7320
            TabIndex        =   29
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b4 
            Caption         =   "348"
            Height          =   495
            Left            =   5400
            TabIndex        =   28
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton a4 
            Caption         =   "338"
            Height          =   495
            Left            =   3480
            TabIndex        =   27
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label12 
            Caption         =   "4.   182 + 76 + 80                        ="
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   3255
         End
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   19815
         Begin VB.TextBox Text3 
            Height          =   495
            Left            =   18600
            TabIndex        =   24
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton d3 
            Caption         =   "12"
            Height          =   495
            Left            =   9480
            TabIndex        =   23
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton c3 
            Caption         =   "9"
            Height          =   435
            Left            =   7320
            TabIndex        =   22
            Top             =   120
            Width           =   1935
         End
         Begin VB.OptionButton b3 
            Caption         =   "8"
            Height          =   495
            Left            =   5400
            TabIndex        =   21
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton a3 
            Caption         =   "7"
            Height          =   495
            Left            =   3480
            TabIndex        =   20
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label11 
            Caption         =   "3.  144 : 16                                 ="
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   1080
         Width           =   19815
         Begin VB.TextBox Text2 
            Height          =   495
            Left            =   18600
            TabIndex        =   17
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton d2 
            Caption         =   "86"
            Height          =   495
            Left            =   9480
            TabIndex        =   16
            Top             =   120
            Width           =   2175
         End
         Begin VB.OptionButton c2 
            Caption         =   "98"
            Height          =   495
            Left            =   7320
            TabIndex        =   15
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton b2 
            Caption         =   "96"
            Height          =   495
            Left            =   5400
            TabIndex        =   14
            Top             =   120
            Width           =   1815
         End
         Begin VB.OptionButton a2 
            Caption         =   "69"
            Height          =   495
            Left            =   3480
            TabIndex        =   13
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label10 
            Caption         =   "2.  12 X 8                                    ="
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   3015
         End
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   19815
         Begin VB.TextBox Text1 
            Height          =   495
            Left            =   18600
            TabIndex        =   10
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton d1 
            Caption         =   "97"
            Height          =   495
            Left            =   9480
            TabIndex        =   9
            Top             =   120
            Width           =   2295
         End
         Begin VB.OptionButton c1 
            Caption         =   "87"
            Height          =   495
            Left            =   7320
            TabIndex        =   8
            Top             =   120
            Width           =   2055
         End
         Begin VB.OptionButton b1 
            Caption         =   "89"
            Height          =   495
            Left            =   5400
            TabIndex        =   7
            Top             =   120
            Width           =   1695
         End
         Begin VB.OptionButton a1 
            Caption         =   "88"
            Height          =   495
            Left            =   3480
            TabIndex        =   6
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "1. 28 + 29 + 30                           ="
            Height          =   375
            Left            =   120
            TabIndex        =   11
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
      TabIndex        =   155
      Top             =   1200
      Width           =   2085
   End
   Begin VB.Label Label8 
      Caption         =   "Tanggal Test :"
      Height          =   195
      Left            =   0
      TabIndex        =   154
      Top             =   870
      Width           =   1305
   End
   Begin VB.Label Label7 
      Caption         =   "NAMA           :"
      Height          =   165
      Left            =   0
      TabIndex        =   153
      Top             =   420
      Width           =   1245
   End
   Begin VB.Label Label6 
      Caption         =   "ID                 :"
      Height          =   285
      Left            =   30
      TabIndex        =   152
      Top             =   0
      Width           =   1125
   End
   Begin VB.Label Label38 
      Caption         =   "12.   3, 7, 15, 31, 63, ..., ..."
      Height          =   255
      Left            =   600
      TabIndex        =   151
      Top             =   6120
      Width           =   3015
   End
   Begin VB.Label Label32 
      Caption         =   "12.   3, 7, 15, 31, 63, ..., ..."
      Height          =   255
      Left            =   600
      TabIndex        =   150
      Top             =   5400
      Width           =   3015
   End
   Begin VB.Label Label31 
      Caption         =   "12.   3, 7, 15, 31, 63, ..., ..."
      Height          =   255
      Left            =   600
      TabIndex        =   149
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
      Caption         =   "TEST XIII"
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
      TabIndex        =   148
      Top             =   120
      Width           =   2115
   End
   Begin VB.Label Label2 
      Caption         =   "Test Berhitung Cepat"
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
      Left            =   9240
      TabIndex        =   147
      Top             =   720
      Width           =   2325
   End
   Begin VB.Label Label5 
      Caption         =   "Hitunglah dengan cepat dan pilih jawaban yang benar."
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
      TabIndex        =   146
      Top             =   1650
      Width           =   5355
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
      TabIndex        =   145
      Top             =   2040
      Width           =   3135
   End
End
Attribute VB_Name = "frmpsi13"
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
        tambahdata = "Insert Into hasiltest13 values ('" & id.Text & "','" & nama.Text & "','" & tanggal.Text & "','" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "','" & Text7.Text & "','" & Text8.Text & "','" & Text9.Text & "','" & Text10.Text & "','" & Text11.Text & "','" & Text12.Text & "','" & Text13.Text & "','" & Text14.Text & "','" & Text15.Text & "','" & Text16.Text & "','" & Text17.Text & "','" & Text18.Text & "','" & Text19.Text & "','" & Text20.Text & "')"
        koneksi.Execute tambahdata
        MsgBox "TEST 13 Berhasil Di simpan, Silakan lanjutkan ke TEST 14", vbInformation, "Pemberitahuan"
    frmpsi14.Show
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

Private Sub d1_Click()
Text1.Text = "D"
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

Private Sub d2_Click()
Text2.Text = "D"
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






