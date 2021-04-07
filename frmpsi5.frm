VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmpsi5 
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
      DisabledPicture =   "frmpsi5.frx":0000
      DownPicture     =   "frmpsi5.frx":13EA
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
      Picture         =   "frmpsi5.frx":27D4
      TabIndex        =   64
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox id 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1170
      TabIndex        =   63
      Top             =   0
      Width           =   3195
   End
   Begin VB.TextBox nama 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1170
      TabIndex        =   62
      Top             =   390
      Width           =   3195
   End
   Begin VB.TextBox tanggal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1170
      TabIndex        =   61
      Top             =   810
      Width           =   3195
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   19935
      _ExtentX        =   35163
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "SOAL NOMOR 1 - 10"
      TabPicture(0)   =   "frmpsi5.frx":3BBE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label31"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label30"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label29"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label28"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label27"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label26"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label25"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label24"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label23"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label22"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label21"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label20"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label19"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label18"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label17"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label16"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label15"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label14"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label13"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label12"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Text20"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text19"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text18"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text17"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text16"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text15"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text14"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text13"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text12"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text11"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text10"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text9"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text8"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text7"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text6"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text5"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text4"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text3"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text2"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Text1"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).ControlCount=   40
      TabCaption(1)   =   "SOAL NOMOR 11 - 15"
      TabPicture(1)   =   "frmpsi5.frx":3BDA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label41"
      Tab(1).Control(1)=   "Label40"
      Tab(1).Control(2)=   "Label39"
      Tab(1).Control(3)=   "Label38"
      Tab(1).Control(4)=   "Label37"
      Tab(1).Control(5)=   "Label36"
      Tab(1).Control(6)=   "Label35"
      Tab(1).Control(7)=   "Label34"
      Tab(1).Control(8)=   "Label33"
      Tab(1).Control(9)=   "Label32"
      Tab(1).Control(10)=   "Text30"
      Tab(1).Control(11)=   "Text29"
      Tab(1).Control(12)=   "Text28"
      Tab(1).Control(13)=   "Text27"
      Tab(1).Control(14)=   "Text26"
      Tab(1).Control(15)=   "Text25"
      Tab(1).Control(16)=   "Text24"
      Tab(1).Control(17)=   "Text23"
      Tab(1).Control(18)=   "Text22"
      Tab(1).Control(19)=   "Text21"
      Tab(1).ControlCount=   20
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   10560
         TabIndex        =   30
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   12000
         TabIndex        =   29
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   10560
         TabIndex        =   28
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   12000
         TabIndex        =   27
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   10560
         TabIndex        =   26
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   12000
         TabIndex        =   25
         Top             =   2040
         Width           =   615
      End
      Begin VB.TextBox Text7 
         Height          =   495
         Left            =   10560
         TabIndex        =   24
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text8 
         Height          =   495
         Left            =   12000
         TabIndex        =   23
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text9 
         Height          =   495
         Left            =   10560
         TabIndex        =   22
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox Text10 
         Height          =   495
         Left            =   12000
         TabIndex        =   21
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox Text11 
         Height          =   495
         Left            =   10560
         TabIndex        =   20
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox Text12 
         Height          =   495
         Left            =   12000
         TabIndex        =   19
         Top             =   3840
         Width           =   615
      End
      Begin VB.TextBox Text13 
         Height          =   495
         Left            =   10560
         TabIndex        =   18
         Top             =   4440
         Width           =   615
      End
      Begin VB.TextBox Text14 
         Height          =   495
         Left            =   12000
         TabIndex        =   17
         Top             =   4440
         Width           =   615
      End
      Begin VB.TextBox Text15 
         Height          =   495
         Left            =   10560
         TabIndex        =   16
         Top             =   5040
         Width           =   615
      End
      Begin VB.TextBox Text16 
         Height          =   495
         Left            =   12000
         TabIndex        =   15
         Top             =   5040
         Width           =   615
      End
      Begin VB.TextBox Text17 
         Height          =   495
         Left            =   10560
         TabIndex        =   14
         Top             =   5640
         Width           =   615
      End
      Begin VB.TextBox Text18 
         Height          =   495
         Left            =   12000
         TabIndex        =   13
         Top             =   5640
         Width           =   615
      End
      Begin VB.TextBox Text19 
         Height          =   495
         Left            =   10560
         TabIndex        =   12
         Top             =   6240
         Width           =   615
      End
      Begin VB.TextBox Text20 
         Height          =   495
         Left            =   12000
         TabIndex        =   11
         Top             =   6240
         Width           =   615
      End
      Begin VB.TextBox Text21 
         Height          =   495
         Left            =   -64560
         TabIndex        =   10
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text22 
         Height          =   495
         Left            =   -63120
         TabIndex        =   9
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text23 
         Height          =   495
         Left            =   -64560
         TabIndex        =   8
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text24 
         Height          =   495
         Left            =   -63120
         TabIndex        =   7
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text25 
         Height          =   495
         Left            =   -64560
         TabIndex        =   6
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text26 
         Height          =   495
         Left            =   -63120
         TabIndex        =   5
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox Text27 
         Height          =   495
         Left            =   -64560
         TabIndex        =   4
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox Text28 
         Height          =   495
         Left            =   -63120
         TabIndex        =   3
         Top             =   3120
         Width           =   615
      End
      Begin VB.TextBox Text29 
         Height          =   495
         Left            =   -64560
         TabIndex        =   2
         Top             =   3960
         Width           =   615
      End
      Begin VB.TextBox Text30 
         Height          =   495
         Left            =   -63120
         TabIndex        =   1
         Top             =   3960
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "1.   a, b, c, f, e, d, g, h, i, l, k, j, k, j, m, ..., ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   60
         Top             =   960
         Width           =   4575
      End
      Begin VB.Label Label13 
         Caption         =   "2.   a, c, c, e, e, g, g, i, ..., ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   59
         Top             =   1560
         Width           =   4575
      End
      Begin VB.Label Label14 
         Caption         =   "3.   a, c, b, a, d, c, a, e, d, a, ..., ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   58
         Top             =   2160
         Width           =   3975
      End
      Begin VB.Label Label15 
         Caption         =   "4.   a, c, e, g, i, k, ..., ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   57
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Label Label16 
         Caption         =   "5.   a, b, d, b, b, d, c, b, d, d, b, d, ..., ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   56
         Top             =   3360
         Width           =   4575
      End
      Begin VB.Label Label17 
         Caption         =   "6.   c, f, e, h, g, j, i, l, ..., ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   55
         Top             =   3960
         Width           =   2655
      End
      Begin VB.Label Label18 
         Caption         =   "7.  z, w, o, y, t, p, w, q, q, t, n, r, ..., ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   54
         Top             =   4560
         Width           =   4215
      End
      Begin VB.Label Label19 
         Caption         =   "8.   a, b, c, b, c, d, c, d, e, d, e, f, e, ..., ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   53
         Top             =   5160
         Width           =   4095
      End
      Begin VB.Label Label20 
         Caption         =   "9.  k, m, l, l, m, k, n, j, ..., ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   52
         Top             =   5760
         Width           =   3135
      End
      Begin VB.Label Label21 
         Caption         =   "10.   m, l, n, o, n, p, q, p, r, s, r, t, ..., ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   51
         Top             =   6360
         Width           =   3975
      End
      Begin VB.Label Label22 
         Caption         =   "Seri Selanjutnya :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   50
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label23 
         Caption         =   "Seri Selanjutnya :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   49
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label24 
         Caption         =   "Seri Selanjutnya :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   48
         Top             =   2160
         Width           =   2655
      End
      Begin VB.Label Label25 
         Caption         =   "Seri Selanjutnya :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   47
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Label Label26 
         Caption         =   "Seri Selanjutnya :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   46
         Top             =   3360
         Width           =   2655
      End
      Begin VB.Label Label27 
         Caption         =   "Seri Selanjutnya :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   45
         Top             =   3960
         Width           =   2655
      End
      Begin VB.Label Label28 
         Caption         =   "Seri Selanjutnya :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   44
         Top             =   4560
         Width           =   2655
      End
      Begin VB.Label Label29 
         Caption         =   "Seri Selanjutnya :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   43
         Top             =   5160
         Width           =   2655
      End
      Begin VB.Label Label30 
         Caption         =   "Seri Selanjutnya :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   42
         Top             =   5760
         Width           =   2655
      End
      Begin VB.Label Label31 
         Caption         =   "Seri Selanjutnya :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   41
         Top             =   6360
         Width           =   2655
      End
      Begin VB.Label Label32 
         Caption         =   "Seri Selanjutnya :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69120
         TabIndex        =   40
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label33 
         Caption         =   "11.   a, d, d, g, g, j, m, ..., ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74640
         TabIndex        =   39
         Top             =   720
         Width           =   3615
      End
      Begin VB.Label Label34 
         Caption         =   "Seri Selanjutnya :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69120
         TabIndex        =   38
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label35 
         Caption         =   "12.   b, a, p, e, d, r, h, g, t, k, j, v, ..., ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74640
         TabIndex        =   37
         Top             =   1560
         Width           =   4215
      End
      Begin VB.Label Label36 
         Caption         =   "Seri Selanjutnya :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69120
         TabIndex        =   36
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Label Label37 
         Caption         =   "13.   a, x, z, e, x, z, i, x, z, m, x, z, ..., ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74640
         TabIndex        =   35
         Top             =   2400
         Width           =   4575
      End
      Begin VB.Label Label38 
         Caption         =   "Seri Selanjutnya :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69120
         TabIndex        =   34
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label39 
         Caption         =   "14.   abc, abe, abg, ..., ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74640
         TabIndex        =   33
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label40 
         Caption         =   "Seri Selanjutnya :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -69120
         TabIndex        =   32
         Top             =   4080
         Width           =   2655
      End
      Begin VB.Label Label41 
         Caption         =   "15.   mnz, opy, qrx, ..., ..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74640
         TabIndex        =   31
         Top             =   4080
         Width           =   2655
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   20250
      Y1              =   1140
      Y2              =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "TEST V"
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
      TabIndex        =   75
      Top             =   120
      Width           =   2115
   End
   Begin VB.Label Label2 
      Caption         =   "Test Seri Huruf"
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
      Left            =   9360
      TabIndex        =   74
      Top             =   720
      Width           =   1485
   End
   Begin VB.Label Label5 
      Caption         =   "Isilah Titik-Titik Kosong Berikut Sesuai Dengan Contoh."
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
      TabIndex        =   73
      Top             =   1650
      Width           =   19995
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
      Left            =   5040
      TabIndex        =   72
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label Label6 
      Caption         =   "ID                 :"
      Height          =   285
      Left            =   30
      TabIndex        =   71
      Top             =   0
      Width           =   1125
   End
   Begin VB.Label Label7 
      Caption         =   "NAMA           :"
      Height          =   165
      Left            =   0
      TabIndex        =   70
      Top             =   420
      Width           =   1245
   End
   Begin VB.Label Label8 
      Caption         =   "Tanggal Test :"
      Height          =   195
      Left            =   0
      TabIndex        =   69
      Top             =   870
      Width           =   1305
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
      TabIndex        =   68
      Top             =   1200
      Width           =   2085
   End
   Begin VB.Label Label3 
      Caption         =   "Contoh :"
      Height          =   255
      Left            =   120
      TabIndex        =   67
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label10 
      Caption         =   "a, b, c, d, ..., ...    seri selanjutnya e, f"
      Height          =   255
      Left            =   120
      TabIndex        =   66
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label Label11 
      Caption         =   "p, o, n, m, ..., ...    seri selanjutnya l, k"
      Height          =   255
      Left            =   120
      TabIndex        =   65
      Top             =   2640
      Width           =   3255
   End
End
Attribute VB_Name = "frmpsi5"
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
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Or Text11.Text = "" Or Text12.Text = "" Or Text13.Text = "" Or Text14.Text = "" Or Text15.Text = "" Or Text16.Text = "" Or Text17.Text = "" Or Text18.Text = "" Or Text19.Text = "" Or Text20.Text = "" Or Text21.Text = "" Or Text22.Text = "" Or Text23.Text = "" Or Text24.Text = "" Or Text25.Text = "" Or Text26.Text = "" Or Text27.Text = "" Or Text28.Text = "" Or Text29.Text = "" Or Text30.Text = "" Then
    MsgBox "Data Belum Lengkap"
    Else
    Dim tambahdata As String
        tambahdata = "Insert Into hasiltest5 values ('" & id.Text & "','" & nama.Text & "','" & tanggal.Text & "','" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "','" & Text6.Text & "','" & Text7.Text & "','" & Text8.Text & "','" & Text9.Text & "','" & Text10.Text & "','" & Text11.Text & "','" & Text12.Text & "','" & Text13.Text & "','" & Text14.Text & "','" & Text15.Text & "','" & Text16.Text & "','" & Text17.Text & "','" & Text18.Text & "','" & Text19.Text & "','" & Text20.Text & "','" & Text21.Text & "','" & Text22.Text & "','" & Text23.Text & "','" & Text24.Text & "','" & Text25.Text & "','" & Text26.Text & "','" & Text27.Text & "','" & Text28.Text & "','" & Text29.Text & "','" & Text30.Text & "')"
        koneksi.Execute tambahdata
        MsgBox "TEST 5 Berhasil Di simpan, Silakan lanjutkan ke TEST 6", vbInformation, "Pemberitahuan"
    frmpsi6.Show
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


