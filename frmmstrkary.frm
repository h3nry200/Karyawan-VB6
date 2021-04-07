VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmmstrkary 
   Caption         =   "MASTER KARYAWAN"
   ClientHeight    =   10950
   ClientLeft      =   -285
   ClientTop       =   1620
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "PRINT"
      Height          =   555
      Left            =   10800
      TabIndex        =   86
      Top             =   10200
      Width           =   885
   End
   Begin MSDataGridLib.DataGrid DGkary 
      Height          =   3015
      Left            =   240
      TabIndex        =   85
      Top             =   120
      Width           =   19785
      _ExtentX        =   34899
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   34
      Top             =   4560
      Width           =   20055
      _ExtentX        =   35375
      _ExtentY        =   7646
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "DATA PRIBADI"
      TabPicture(0)   =   "frmmstrkary.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label9"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label13"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label14"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label15"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label16"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label24"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label25"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "namatxt"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "alamatktptxt"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "alamattgltxt"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "jeniskeltxt"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "noktptxt"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "pendterakhirtxt"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "tlptxt"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "tmptlahirtxt"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "tgllahirtxt"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "hptxt"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "saudaratxt"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "anakketxt"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).ControlCount=   24
      TabCaption(1)   =   "DATA KELUARGA"
      TabPicture(1)   =   "frmmstrkary.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "jmlhanaktxt"
      Tab(1).Control(1)=   "hpsuamitxt"
      Tab(1).Control(2)=   "namasuamitxt"
      Tab(1).Control(3)=   "statuspertxt"
      Tab(1).Control(4)=   "namaanaktxt"
      Tab(1).Control(5)=   "usiaanaktxt"
      Tab(1).Control(6)=   "Label20"
      Tab(1).Control(7)=   "Label21"
      Tab(1).Control(8)=   "Label26"
      Tab(1).Control(9)=   "Label27"
      Tab(1).Control(10)=   "Label28"
      Tab(1).Control(11)=   "Label29"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "DATA ECON"
      TabPicture(2)   =   "frmmstrkary.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "hpecontxt"
      Tab(2).Control(1)=   "alamatecontxt"
      Tab(2).Control(2)=   "hubecontxt"
      Tab(2).Control(3)=   "namaecontxt"
      Tab(2).Control(4)=   "kerja1txt"
      Tab(2).Control(5)=   "kerja2txt"
      Tab(2).Control(6)=   "kerja3txt"
      Tab(2).Control(7)=   "Label30"
      Tab(2).Control(8)=   "Label31"
      Tab(2).Control(9)=   "Label32"
      Tab(2).Control(10)=   "Label33"
      Tab(2).Control(11)=   "Label34"
      Tab(2).Control(12)=   "Label35"
      Tab(2).Control(13)=   "Label36"
      Tab(2).ControlCount=   14
      Begin VB.TextBox hpecontxt 
         BackColor       =   &H0000FF00&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72840
         TabIndex        =   77
         Top             =   2280
         Width           =   4155
      End
      Begin VB.TextBox alamatecontxt 
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -72780
         MultiLine       =   -1  'True
         TabIndex        =   76
         Top             =   1590
         Width           =   4155
      End
      Begin VB.TextBox hubecontxt 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72840
         TabIndex        =   75
         Top             =   1170
         Width           =   4155
      End
      Begin VB.TextBox namaecontxt 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72840
         TabIndex        =   74
         Top             =   720
         Width           =   4155
      End
      Begin VB.TextBox kerja1txt 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -65280
         MultiLine       =   -1  'True
         TabIndex        =   73
         Top             =   630
         Width           =   4155
      End
      Begin VB.TextBox kerja2txt 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -65280
         MultiLine       =   -1  'True
         TabIndex        =   72
         Top             =   1350
         Width           =   4155
      End
      Begin VB.TextBox kerja3txt 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -65280
         MultiLine       =   -1  'True
         TabIndex        =   71
         Top             =   2070
         Width           =   4155
      End
      Begin VB.TextBox jmlhanaktxt 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72450
         TabIndex        =   64
         Top             =   2400
         Width           =   4155
      End
      Begin VB.TextBox hpsuamitxt 
         BackColor       =   &H0000FF00&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72450
         TabIndex        =   63
         Top             =   1980
         Width           =   4155
      End
      Begin VB.TextBox namasuamitxt 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72450
         TabIndex        =   62
         Top             =   1530
         Width           =   4155
      End
      Begin VB.TextBox statuspertxt 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72450
         TabIndex        =   61
         Top             =   1080
         Width           =   4155
      End
      Begin VB.TextBox namaanaktxt 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72450
         TabIndex        =   60
         Top             =   2820
         Width           =   4155
      End
      Begin VB.TextBox usiaanaktxt 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -72450
         TabIndex        =   59
         Top             =   3210
         Width           =   4155
      End
      Begin VB.TextBox anakketxt 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13710
         TabIndex        =   46
         Top             =   1950
         Width           =   1905
      End
      Begin VB.TextBox saudaratxt 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9720
         TabIndex        =   45
         Top             =   1950
         Width           =   1905
      End
      Begin VB.TextBox hptxt 
         BackColor       =   &H0000FF00&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6060
         TabIndex        =   44
         Top             =   3210
         Width           =   1905
      End
      Begin VB.TextBox tgllahirtxt 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13710
         TabIndex        =   43
         Top             =   1500
         Width           =   1905
      End
      Begin VB.TextBox tmptlahirtxt 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10140
         TabIndex        =   42
         Top             =   1500
         Width           =   1905
      End
      Begin VB.TextBox tlptxt 
         BackColor       =   &H0000FF00&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1710
         TabIndex        =   41
         Top             =   3210
         Width           =   1905
      End
      Begin VB.TextBox pendterakhirtxt 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10830
         TabIndex        =   40
         Top             =   1080
         Width           =   4815
      End
      Begin VB.TextBox noktptxt 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10830
         TabIndex        =   39
         Top             =   630
         Width           =   4815
      End
      Begin VB.TextBox jeniskeltxt 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3150
         TabIndex        =   38
         Top             =   2790
         Width           =   4815
      End
      Begin VB.TextBox alamattgltxt 
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3150
         MultiLine       =   -1  'True
         TabIndex        =   37
         Top             =   1980
         Width           =   4815
      End
      Begin VB.TextBox alamatktptxt 
         BackColor       =   &H0000FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3150
         MultiLine       =   -1  'True
         TabIndex        =   36
         Top             =   1140
         Width           =   4815
      End
      Begin VB.TextBox namatxt 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3150
         TabIndex        =   35
         Top             =   720
         Width           =   4815
      End
      Begin VB.Label Label30 
         Caption         =   "NAMA ECON             :"
         Height          =   315
         Left            =   -74520
         TabIndex        =   84
         Top             =   780
         Width           =   1695
      End
      Begin VB.Label Label31 
         Caption         =   "HUBUNGAN ECON  :"
         Height          =   285
         Left            =   -74490
         TabIndex        =   83
         Top             =   1260
         Width           =   1785
      End
      Begin VB.Label Label32 
         Caption         =   "ALAMAT ECON         :"
         Height          =   315
         Left            =   -74520
         TabIndex        =   82
         Top             =   1710
         Width           =   1725
      End
      Begin VB.Label Label33 
         Caption         =   "HP ECON                   :"
         Height          =   315
         Left            =   -74520
         TabIndex        =   81
         Top             =   2310
         Width           =   1665
      End
      Begin VB.Label Label34 
         Caption         =   "KERJA 1                    :"
         Height          =   255
         Left            =   -66900
         TabIndex        =   80
         Top             =   690
         Width           =   1725
      End
      Begin VB.Label Label35 
         Caption         =   "KERJA 2                    :"
         Height          =   285
         Left            =   -66900
         TabIndex        =   79
         Top             =   1470
         Width           =   1665
      End
      Begin VB.Label Label36 
         Caption         =   "KERJA 3                    :"
         Height          =   345
         Left            =   -66900
         TabIndex        =   78
         Top             =   2190
         Width           =   1785
      End
      Begin VB.Label Label20 
         Caption         =   "STATUS PERKAWINAN  :"
         Height          =   285
         Left            =   -74520
         TabIndex        =   70
         Top             =   1170
         Width           =   1995
      End
      Begin VB.Label Label21 
         Caption         =   "JUMLAH ANAK                :"
         Height          =   285
         Left            =   -74490
         TabIndex        =   69
         Top             =   2490
         Width           =   2085
      End
      Begin VB.Label Label26 
         Caption         =   "NAMA SUAMI / ISTRI      :"
         Height          =   195
         Left            =   -74520
         TabIndex        =   68
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label27 
         Caption         =   "HP SUAMI / ISTRI           :"
         Height          =   255
         Left            =   -74490
         TabIndex        =   67
         Top             =   2040
         Width           =   1995
      End
      Begin VB.Label Label28 
         Caption         =   "NAMA ANAK                    :"
         Height          =   225
         Left            =   -74490
         TabIndex        =   66
         Top             =   2880
         Width           =   1995
      End
      Begin VB.Label Label29 
         Caption         =   "USIA ANAK                     :"
         Height          =   255
         Left            =   -74490
         TabIndex        =   65
         Top             =   3300
         Width           =   2205
      End
      Begin VB.Label Label25 
         Caption         =   "TANGGAL LAHIR  :"
         Height          =   195
         Left            =   12180
         TabIndex        =   58
         Top             =   1560
         Width           =   1665
      End
      Begin VB.Label Label24 
         Caption         =   "TEMPAT LAHIR  :"
         Height          =   315
         Left            =   8640
         TabIndex        =   57
         Top             =   1560
         Width           =   1545
      End
      Begin VB.Label Label16 
         Caption         =   "SAUDARA  :"
         Height          =   195
         Left            =   8670
         TabIndex        =   56
         Top             =   2010
         Width           =   975
      End
      Begin VB.Label Label15 
         Caption         =   "ANAK KE :"
         Height          =   255
         Left            =   12240
         TabIndex        =   55
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label14 
         Caption         =   "PENDIDIKAN TERAKHIR  :"
         Height          =   255
         Left            =   8670
         TabIndex        =   54
         Top             =   1170
         Width           =   2265
      End
      Begin VB.Label Label13 
         Caption         =   "NOMOR KTP                     :"
         Height          =   315
         Left            =   8700
         TabIndex        =   53
         Top             =   720
         Width           =   1995
      End
      Begin VB.Label Label9 
         Caption         =   "HANDPHONE (HP)  :"
         Height          =   285
         Left            =   4170
         TabIndex        =   52
         Top             =   3240
         Width           =   1875
      End
      Begin VB.Label Label8 
         Caption         =   "TELP   :"
         Height          =   285
         Left            =   1020
         TabIndex        =   51
         Top             =   3270
         Width           =   1065
      End
      Begin VB.Label Label7 
         Caption         =   "JENIS KELAMIN               :"
         Height          =   225
         Left            =   1020
         TabIndex        =   50
         Top             =   2790
         Width           =   1935
      End
      Begin VB.Label Label6 
         Caption         =   "ALAMAT TINGGAL           :"
         Height          =   225
         Left            =   1050
         TabIndex        =   49
         Top             =   1980
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "ALAMAT KTP                   :"
         Height          =   225
         Left            =   1020
         TabIndex        =   48
         Top             =   1230
         Width           =   1995
      End
      Begin VB.Label Label4 
         Caption         =   "NAMA                               :"
         Height          =   255
         Left            =   1020
         TabIndex        =   47
         Top             =   780
         Width           =   2085
      End
   End
   Begin MSComCtl2.DTPicker tglkontrak1 
      Height          =   375
      Left            =   11040
      TabIndex        =   33
      Top             =   3240
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   20709379
      CurrentDate     =   42809
   End
   Begin MSComCtl2.DTPicker tglmasuktxt 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "MM-dd-yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   32
      Top             =   4080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   20709379
      CurrentDate     =   42809
   End
   Begin VB.TextBox tglkontrak3 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11040
      TabIndex        =   30
      Top             =   3240
      Width           =   1905
   End
   Begin VB.CommandButton seekbtn 
      Caption         =   "SEEK"
      Height          =   555
      Left            =   11880
      TabIndex        =   29
      Top             =   10200
      Width           =   885
   End
   Begin VB.CommandButton newbtn 
      Caption         =   "NEW"
      Height          =   555
      Left            =   5160
      TabIndex        =   28
      Top             =   10200
      Width           =   1005
   End
   Begin VB.CommandButton editbtn 
      Caption         =   "EDIT"
      Height          =   555
      Left            =   6360
      TabIndex        =   27
      Top             =   10200
      Width           =   975
   End
   Begin VB.CommandButton deletebtn 
      Caption         =   "DELETE"
      Height          =   555
      Left            =   7560
      TabIndex        =   26
      Top             =   10200
      Width           =   885
   End
   Begin VB.CommandButton savebtn 
      Caption         =   "SAVE"
      Height          =   555
      Left            =   8640
      TabIndex        =   25
      Top             =   10200
      Width           =   885
   End
   Begin VB.CommandButton exitbtn 
      Caption         =   "EXIT"
      Height          =   555
      Left            =   12960
      TabIndex        =   24
      Top             =   10200
      Width           =   885
   End
   Begin VB.CommandButton cancelbtn 
      Caption         =   "CANCEL"
      Height          =   555
      Left            =   9720
      TabIndex        =   23
      Top             =   10200
      Width           =   885
   End
   Begin VB.TextBox recordtxt 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   2040
      TabIndex        =   22
      Top             =   9030
      Width           =   18075
   End
   Begin VB.TextBox lastuptxt 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6570
      TabIndex        =   20
      Top             =   3240
      Width           =   2355
   End
   Begin VB.TextBox ketlaintxt 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   15810
      TabIndex        =   19
      Top             =   4050
      Width           =   4485
   End
   Begin VB.TextBox alasankeltxt 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   18
      Top             =   4050
      Width           =   4995
   End
   Begin VB.TextBox tglajurestxt 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   17550
      TabIndex        =   17
      Top             =   3630
      Width           =   2235
   End
   Begin VB.TextBox tglresigntxt 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12570
      TabIndex        =   16
      Top             =   3660
      Width           =   2265
   End
   Begin VB.TextBox jabatantxt 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8490
      TabIndex        =   15
      Top             =   3660
      Width           =   2235
   End
   Begin VB.TextBox tglmasuktxt1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5310
      TabIndex        =   14
      Top             =   4080
      Width           =   1905
   End
   Begin VB.TextBox statkartxt 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5310
      TabIndex        =   13
      Top             =   3660
      Width           =   1905
   End
   Begin VB.TextBox niktxt 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   12
      Top             =   4080
      Width           =   2445
   End
   Begin VB.TextBox tglcretxt 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2010
      TabIndex        =   11
      Top             =   3240
      Width           =   2745
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   825
      Left            =   2940
      Top             =   1920
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   1455
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   960
      Top             =   2520
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label41 
      Caption         =   "TANGGAL KONTRAK 1 :"
      Height          =   225
      Left            =   9120
      TabIndex        =   31
      Top             =   3270
      Width           =   1905
   End
   Begin VB.Label Label40 
      Caption         =   "KETERANGAN  :"
      Height          =   375
      Left            =   480
      TabIndex        =   21
      Top             =   9240
      Width           =   1365
   End
   Begin VB.Label Label37 
      Caption         =   "DATA PRIBADI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   4500
      Width           =   4065
   End
   Begin VB.Label Label23 
      Caption         =   "LAST UPDATE  :"
      Height          =   315
      Left            =   5220
      TabIndex        =   9
      Top             =   3300
      Width           =   1635
   End
   Begin VB.Label Label22 
      Caption         =   "KETERANGAN LAIN  :"
      Height          =   345
      Left            =   14130
      TabIndex        =   8
      Top             =   4080
      Width           =   1725
   End
   Begin VB.Label Label19 
      Caption         =   "ALASAN KELUAR  :"
      Height          =   255
      Left            =   7500
      TabIndex        =   7
      Top             =   4110
      Width           =   1605
   End
   Begin VB.Label Label18 
      Caption         =   "STATUS KARYAWAN  :"
      Height          =   285
      Left            =   3510
      TabIndex        =   6
      Top             =   3720
      Width           =   1905
   End
   Begin VB.Label Label17 
      Caption         =   "JABATAN  :"
      Height          =   315
      Left            =   7470
      TabIndex        =   5
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "TANGGAL PENGAJUAN RESIGN :"
      Height          =   285
      Left            =   14940
      TabIndex        =   4
      Top             =   3690
      Width           =   2595
   End
   Begin VB.Label Label11 
      Caption         =   "TANGGAL RESIGN :"
      Height          =   285
      Left            =   10950
      TabIndex        =   3
      Top             =   3720
      Width           =   1665
   End
   Begin VB.Label Label10 
      Caption         =   "TANGGAL MASUK       :"
      Height          =   225
      Left            =   3510
      TabIndex        =   2
      Top             =   4110
      Width           =   1785
   End
   Begin VB.Label Label3 
      Caption         =   "NIK   :"
      Height          =   270
      Left            =   390
      TabIndex        =   1
      Top             =   4080
      Width           =   705
   End
   Begin VB.Label Label2 
      Caption         =   "TANGGAL CREATE  :"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   3300
      Width           =   1755
   End
End
Attribute VB_Name = "frmmstrkary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim koneksi As New rdoConnection
Dim rQuery As New rdoQuery
Dim rs As rdoResultset

Private Sub Command1_Click()
Datakary2.Show

Set Datakary2.DataSource = Adodc2
Datakary2.Sections("Section1").Controls.Item("namalb").Caption = namatxt.Text
Datakary2.Sections("Section1").Controls.Item("jabatanlb").Caption = jabatantxt.Text
Datakary2.Sections("Section1").Controls.Item("tglmasuklb").Caption = tglmasuktxt.Value
Datakary2.Sections("Section1").Controls.Item("statuskarlb").Caption = statkartxt.Text
Datakary2.Sections("Section1").Controls.Item("tglkontraklb").Caption = tglkontrak1.Value
Datakary2.Sections("Section1").Controls.Item("niklb").Caption = niktxt.Text
Datakary2.Sections("Section1").Controls.Item("jeniskellb").Caption = jeniskeltxt.Text
Datakary2.Sections("Section1").Controls.Item("tglajulb").Caption = tglajurestxt.Text
Datakary2.Sections("Section1").Controls.Item("tglresignlb").Caption = tglresigntxt.Text
Datakary2.Sections("Section1").Controls.Item("alasanlb").Caption = alasankeltxt.Text & " " & "or" & " " & ketlaintxt.Text
Datakary2.Sections("Section1").Controls.Item("rec").Caption = recordtxt.Text

Datakary2.Sections("Section1").Controls.Item("alamatktplb").Caption = alamatktptxt.Text
Datakary2.Sections("Section1").Controls.Item("alamattgllb").Caption = alamattgltxt.Text
Datakary2.Sections("Section1").Controls.Item("pendlb").Caption = pendterakhirtxt.Text
Datakary2.Sections("Section1").Controls.Item("tlprmhlb").Caption = tlptxt.Text
Datakary2.Sections("Section1").Controls.Item("hplb").Caption = hptxt.Text
Datakary2.Sections("Section1").Controls.Item("noktplb").Caption = noktptxt.Text
Datakary2.Sections("Section1").Controls.Item("tmptlahirlb").Caption = tmptlahirtxt.Text
Datakary2.Sections("Section1").Controls.Item("tgllhrlb").Caption = tgllahirtxt.Text
Datakary2.Sections("Section1").Controls.Item("saudaralb").Caption = saudaratxt.Text
Datakary2.Sections("Section1").Controls.Item("anakkelb").Caption = anakketxt.Text

Datakary2.Sections("Section1").Controls.Item("stsperlb").Caption = statuspertxt.Text
Datakary2.Sections("Section1").Controls.Item("nmsuis").Caption = namasuamitxt.Text
Datakary2.Sections("Section1").Controls.Item("hpsuis").Caption = hpsuamitxt.Text
Datakary2.Sections("Section1").Controls.Item("jmlanak").Caption = jmlhanaktxt.Text
Datakary2.Sections("Section1").Controls.Item("nmanak").Caption = namaanaktxt.Text
Datakary2.Sections("Section1").Controls.Item("usiaanak").Caption = namaanaktxt.Text

Datakary2.Sections("Section1").Controls.Item("nmecon").Caption = namaecontxt.Text
Datakary2.Sections("Section1").Controls.Item("hubecon").Caption = hubecontxt.Text
Datakary2.Sections("Section1").Controls.Item("alecon").Caption = alamatecontxt.Text
Datakary2.Sections("Section1").Controls.Item("hpecon").Caption = hpecontxt.Text
Datakary2.Sections("Section1").Controls.Item("kerja1ec").Caption = kerja1txt.Text
Datakary2.Sections("Section1").Controls.Item("kerja2ec").Caption = kerja2txt.Text
Datakary2.Sections("Section1").Controls.Item("kerja3ec").Caption = kerja3txt.Text

Adodc1.Refresh

End Sub

Private Sub deletebtn_Click()
If MsgBox("Yakin Ingin Menghapus Data?", vbQuestion + vbOKCancel, "konfirmasi") = vbOK Then
Dim hapusdata As String
        hapusdata = "delete from namakar where nama =" & namatxt.Text & ""
        koneksi.Execute hapusdata
        MsgBox "Data Berhasil Dihapus", vbInformation, "Pemberitahuan"
    Adodc1.Refresh
DGkary.Refresh
Call Bersih

Else
Call Bersih
End If

End Sub

Private Sub DGkary_Click()
'idtxt.Text = DGkary.Columns(0).Text
tglcretxt.Text = DGkary.Columns(1).Text
niktxt.Text = DGkary.Columns(2).Text
namatxt.Text = DGkary.Columns(3).Text
alamatktptxt.Text = DGkary.Columns(4).Text
alamattgltxt.Text = DGkary.Columns(5).Text
jeniskeltxt.Text = DGkary.Columns(6).Text
tlptxt.Text = DGkary.Columns(7).Text
hptxt.Text = DGkary.Columns(8).Text
tglmasuktxt.Value = DGkary.Columns(9).Text
tglresigntxt.Text = DGkary.Columns(10).Text
tglajurestxt.Text = DGkary.Columns(11).Text
noktptxt.Text = DGkary.Columns(12).Text
pendterakhirtxt.Text = DGkary.Columns(13).Text
anakketxt.Text = DGkary.Columns(14).Text
saudaratxt.Text = DGkary.Columns(15).Text
jabatantxt.Text = DGkary.Columns(16).Text
statkartxt.Text = DGkary.Columns(17).Text
alasankeltxt.Text = DGkary.Columns(18).Text
statuspertxt.Text = DGkary.Columns(19).Text
jmlhanaktxt.Text = DGkary.Columns(20).Text
ketlaintxt.Text = DGkary.Columns(21).Text
lastuptxt.Text = DGkary.Columns(22).Text
tmptlahirtxt.Text = DGkary.Columns(23).Text
tgllahirtxt.Text = DGkary.Columns(24).Text
namasuamitxt.Text = DGkary.Columns(25).Text
hpsuamitxt.Text = DGkary.Columns(26).Text
namaanaktxt.Text = DGkary.Columns(27).Text
usiaanaktxt.Text = DGkary.Columns(28).Text
namaecontxt.Text = DGkary.Columns(29).Text
hubecontxt.Text = DGkary.Columns(30).Text
alamatecontxt.Text = DGkary.Columns(31).Text
hpecontxt.Text = DGkary.Columns(32).Text
kerja1txt.Text = DGkary.Columns(33).Text
kerja2txt.Text = DGkary.Columns(34).Text
kerja3txt.Text = DGkary.Columns(35).Text
recordtxt.Text = DGkary.Columns(36).Text
tglkontrak1.Value = DGkary.Columns(37).Text

If tglajurestxt.Text = "-" Then
tglajurestxt.BackColor = &HFFFF00
Else
tglajurestxt.BackColor = &HFF&
End If

If tglresigntxt.Text = "-" Then
tglresigntxt.BackColor = &HFFFF00
Else
tglresigntxt.BackColor = &HFF&
End If

If alasankeltxt.Text = "-" Then
alasankeltxt.BackColor = &HFFFF00
Else
alasankeltxt.BackColor = &HFF&
End If

If ketlaintxt.Text = "-" Then
ketlaintxt.BackColor = &HFFFF00
Else
ketlaintxt.BackColor = &HFF&
End If

'If DGkary.Columns(37).Text = " " Then
'tglkontrak1.Value = DateAdd("d", 90, tglmasuktxt.Value)
'Else
'tglkontrak1.Value = DGkary.Columns(37).Text
'End If

End Sub

Private Sub DGkary_KeyDown(KeyCode As Integer, Shift As Integer)
'idtxt.Text = DGkary.Columns(0).Text
tglcretxt.Text = DGkary.Columns(1).Text
niktxt.Text = DGkary.Columns(2).Text
namatxt.Text = DGkary.Columns(3).Text
alamatktptxt.Text = DGkary.Columns(4).Text
alamattgltxt.Text = DGkary.Columns(5).Text
jeniskeltxt.Text = DGkary.Columns(6).Text
tlptxt.Text = DGkary.Columns(7).Text
hptxt.Text = DGkary.Columns(8).Text
tglmasuktxt.Value = DGkary.Columns(9).Text
tglresigntxt.Text = DGkary.Columns(10).Text
tglajurestxt.Text = DGkary.Columns(11).Text
noktptxt.Text = DGkary.Columns(12).Text
pendterakhirtxt.Text = DGkary.Columns(13).Text
anakketxt.Text = DGkary.Columns(14).Text
saudaratxt.Text = DGkary.Columns(15).Text
jabatantxt.Text = DGkary.Columns(16).Text
statkartxt.Text = DGkary.Columns(17).Text
alasankeltxt.Text = DGkary.Columns(18).Text
statuspertxt.Text = DGkary.Columns(19).Text
jmlhanaktxt.Text = DGkary.Columns(20).Text
ketlaintxt.Text = DGkary.Columns(21).Text
lastuptxt.Text = DGkary.Columns(22).Text
tmptlahirtxt.Text = DGkary.Columns(23).Text
tgllahirtxt.Text = DGkary.Columns(24).Text
namasuamitxt.Text = DGkary.Columns(25).Text
hpsuamitxt.Text = DGkary.Columns(26).Text
namaanaktxt.Text = DGkary.Columns(27).Text
usiaanaktxt.Text = DGkary.Columns(28).Text
namaecontxt.Text = DGkary.Columns(29).Text
hubecontxt.Text = DGkary.Columns(30).Text
alamatecontxt.Text = DGkary.Columns(31).Text
hpecontxt.Text = DGkary.Columns(32).Text
kerja1txt.Text = DGkary.Columns(33).Text
kerja2txt.Text = DGkary.Columns(34).Text
kerja3txt.Text = DGkary.Columns(35).Text
recordtxt.Text = DGkary.Columns(36).Text
tglkontrak1.Value = DGkary.Columns(37).Text

If tglajurestxt.Text = "-" Then
tglajurestxt.BackColor = &HFFFF00
Else
tglajurestxt.BackColor = &HFF&
End If

If tglresigntxt.Text = "-" Then
tglresigntxt.BackColor = &HFFFF00
Else
tglresigntxt.BackColor = &HFF&
End If

If alasankeltxt.Text = "-" Then
alasankeltxt.BackColor = &HFFFF00
Else
alasankeltxt.BackColor = &HFF&
End If

If ketlaintxt.Text = "-" Then
ketlaintxt.BackColor = &HFFFF00
Else
ketlaintxt.BackColor = &HFF&
End If

'tglkontrak3.Text = DateAdd("d", 90, tglmasuktxt.Value)
'If DGkary.Columns(37).Text = " " Then
'tglkontrak1.Value = DateAdd("d", 90, tglmasuktxt.Value)
'Else
'tglkontrak1.Value = DGkary.Columns(37).Text
'End If

End Sub

Private Sub editbtn_Click()
editbtn.Enabled = False
deletebtn.Enabled = False
savebtn.Enabled = True
newbtn.Enabled = False
seekbtn.Enabled = False
cancelbtn.Enabled = True
'idtxt.Enabled = True
tglcretxt.Enabled = False
niktxt.Enabled = True
namatxt.Enabled = True
alamatktptxt.Enabled = True
alamattgltxt.Enabled = True
jeniskeltxt.Enabled = True
tlptxt.Enabled = True
hptxt.Enabled = True
tglmasuktxt.Enabled = True
tglresigntxt.Enabled = True
tglajurestxt.Enabled = True
noktptxt.Enabled = True
pendterakhirtxt.Enabled = True
anakketxt.Enabled = True
saudaratxt.Enabled = True
jabatantxt.Enabled = True
statkartxt.Enabled = True
alasankeltxt.Enabled = True
statuspertxt.Enabled = True
jmlhanaktxt.Enabled = True
ketlaintxt.Enabled = True
lastuptxt.Enabled = False
tmptlahirtxt.Enabled = True
tgllahirtxt.Enabled = True
namasuamitxt.Enabled = True
hpsuamitxt.Enabled = True
namaanaktxt.Enabled = True
usiaanaktxt.Enabled = True
namaecontxt.Enabled = True
hubecontxt.Enabled = True
alamatecontxt.Enabled = True
hpecontxt.Enabled = True
kerja1txt.Enabled = True
kerja2txt.Enabled = True
kerja3txt.Enabled = True
recordtxt.Enabled = True

'tglcretxt.SetFocus
editbtn.Caption = "EDITDATA"
lastuptxt.Text = login2.lbtanggal.Caption

End Sub

Private Sub Form_Load()
Call Bersih
'If konn.State = adStateOpen Then konn.Close

'Call koneksi
'koneksi.Connect = "Driver={MySQL ODBC 5.1 Driver};SERVER=localhost;PWD=123;UID=root;PORT=3306;DATABASE=karyawan;"
'koneksi.EstablishConnection
    Adodc1.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};SERVER=localhost;PWD=;UID=root;PORT=3306;DATABASE=karyawan;"

'    Adodc1.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};SERVER=192.168.2.11;PWD=123456;UID=BRIWIRA;PORT=3306;DATABASE=karyawan;"
    Adodc1.RecordSource = "data_karyawan"
    Adodc1.Refresh

    Adodc2.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};SERVER=localhost;PWD=;UID=root;PORT=3306;DATABASE=karyawan;"
    Adodc2.RecordSource = "data_karyawan"
    Adodc2.Refresh


Set DGkary.DataSource = Adodc1
editbtn.Enabled = True
deletebtn.Enabled = True
savebtn.Enabled = False
newbtn.Enabled = True
cancelbtn.Enabled = False
'idtxt.Enabled = True
tglcretxt.Enabled = False
niktxt.Enabled = False
namatxt.Enabled = False
alamatktptxt.Enabled = False
alamattgltxt.Enabled = False
jeniskeltxt.Enabled = False
tlptxt.Enabled = False
hptxt.Enabled = False
tglmasuktxt.Enabled = False
tglresigntxt.Enabled = False
tglajurestxt.Enabled = False
noktptxt.Enabled = False
pendterakhirtxt.Enabled = False
anakketxt.Enabled = False
saudaratxt.Enabled = False
jabatantxt.Enabled = False
statkartxt.Enabled = False
alasankeltxt.Enabled = False
statuspertxt.Enabled = False
jmlhanaktxt.Enabled = False
ketlaintxt.Enabled = False
lastuptxt.Enabled = False
tmptlahirtxt.Enabled = False
tgllahirtxt.Enabled = False
namasuamitxt.Enabled = False
hpsuamitxt.Enabled = False
namaanaktxt.Enabled = False
usiaanaktxt.Enabled = False
namaecontxt.Enabled = False
hubecontxt.Enabled = False
alamatecontxt.Enabled = False
hpecontxt.Enabled = False
kerja1txt.Enabled = False
kerja2txt.Enabled = False
kerja3txt.Enabled = False
recordtxt.Enabled = False
tglkontrak1.Enabled = False


cancelbtn.Enabled = False


End Sub

Sub Bersih()
'idtxt.Text = ""
tglcretxt.Text = ""
niktxt.Text = ""
namatxt.Text = ""
alamatktptxt.Text = ""
alamattgltxt.Text = ""
jeniskeltxt.Text = ""
tlptxt.Text = ""
hptxt.Text = ""
'tglmasuktxt.Value = "NULL"
tglresigntxt.Text = ""
tglajurestxt.Text = ""
noktptxt.Text = ""
pendterakhirtxt.Text = ""
anakketxt.Text = ""
saudaratxt.Text = ""
jabatantxt.Text = ""
statkartxt.Text = ""
alasankeltxt.Text = ""
statuspertxt.Text = ""
jmlhanaktxt.Text = ""
ketlaintxt.Text = ""
lastuptxt.Text = ""
tmptlahirtxt.Text = ""
tgllahirtxt.Text = ""
namasuamitxt.Text = ""
hpsuamitxt.Text = ""
namaanaktxt.Text = ""
usiaanaktxt.Text = ""
namaecontxt.Text = ""
hubecontxt.Text = ""
alamatecontxt.Text = ""
hpecontxt.Text = ""
kerja1txt.Text = ""
kerja2txt.Text = ""
kerja3txt.Text = ""
recordtxt.Text = ""
End Sub

Sub KondisiAwal()
'idtxt.Text = ""
tglcretxt.Text = ""
niktxt.Text = ""
namatxt.Text = ""
alamatktptxt.Text = ""
alamattgltxt.Text = ""
jeniskeltxt.Text = ""
tlptxt.Text = ""
hptxt.Text = ""
'tglmasuktxt.Value = ""
tglresigntxt.Text = ""
tglajurestxt.Text = ""
noktptxt.Text = ""
pendterakhirtxt.Text = ""
anakketxt.Text = ""
saudaratxt.Text = ""
jabatantxt.Text = ""
statkartxt.Text = ""
alasankeltxt.Text = ""
statuspertxt.Text = ""
jmlhanaktxt.Text = ""
ketlaintxt.Text = ""
lastuptxt.Text = ""
tmptlahirtxt.Text = ""
tgllahirtxt.Text = ""
namasuamitxt.Text = ""
hpsuamitxt.Text = ""
namaanaktxt.Text = ""
usiaanaktxt.Text = ""
namaecontxt.Text = ""
hubecontxt.Text = ""
alamatecontxt.Text = ""
hpecontxt.Text = ""
kerja1txt.Text = ""
kerja2txt.Text = ""
kerja3txt.Text = ""
recordtxt.Text = ""
newbtn.Enabled = True
editbtn.Enabled = True
deletebtn.Enabled = True
seekbtn.Enabled = True


DGkary.Refresh
End Sub


Private Sub newbtn_Click()
editbtn.Enabled = False
seekbtn.Enabled = False
deletebtn.Enabled = False
savebtn.Enabled = True
newbtn.Enabled = False
cancelbtn.Enabled = True
'idtxt.Enabled = True
tglcretxt.Enabled = False
tglkontrak1.Enabled = True
niktxt.Enabled = True
namatxt.Enabled = True
alamatktptxt.Enabled = True
alamattgltxt.Enabled = True
jeniskeltxt.Enabled = True
tlptxt.Enabled = True
hptxt.Enabled = True
tglmasuktxt.Enabled = True
tglresigntxt.Enabled = True
tglajurestxt.Enabled = True
noktptxt.Enabled = True
pendterakhirtxt.Enabled = True
anakketxt.Enabled = True
saudaratxt.Enabled = True
jabatantxt.Enabled = True
statkartxt.Enabled = True
alasankeltxt.Enabled = True
statuspertxt.Enabled = True
jmlhanaktxt.Enabled = True
ketlaintxt.Enabled = True
lastuptxt.Enabled = False
tmptlahirtxt.Enabled = True
tgllahirtxt.Enabled = True
namasuamitxt.Enabled = True
hpsuamitxt.Enabled = True
namaanaktxt.Enabled = True
usiaanaktxt.Enabled = True
namaecontxt.Enabled = True
hubecontxt.Enabled = True
alamatecontxt.Enabled = True
hpecontxt.Enabled = True
kerja1txt.Enabled = True
kerja2txt.Enabled = True
kerja3txt.Enabled = True
recordtxt.Enabled = True
'tglcretxt.SetFocus
newbtn.Caption = "NEWDATA"
tglcretxt.Text = login2.lbtanggal.Caption
lastuptxt.Text = login2.lbtanggal.Caption

End Sub

Private Sub cancelbtn_Click()
 Call Bersih
 newbtn.Enabled = True
 editbtn.Enabled = True
 deletebtn.Enabled = True
 newbtn.Caption = "NEW"
 editbtn.Caption = "EDIT"

cancelbtn.Enabled = False
seekbtn.Enabled = True

End Sub

Private Sub exitbtn_Click()
If konn.State = adStateOpen Then konn.Close
'koneksi.Close
Unload Me
End Sub

Private Sub savebtn_Click()
If newbtn.Caption = "NEWDATA" Then
If tglcretxt.Text = "" Or niktxt.Text = "" Or namatxt.Text = "" Or alamatktptxt.Text = "" Or alamattgltxt.Text = "" Or jeniskeltxt.Text = "" Or tlptxt.Text = "" Or hptxt.Text = "" Or tglmasuktxt.Value = "" Or tglresigntxt.Text = "" Or tglajurestxt.Text = "" Or noktptxt.Text = "" Or pendterakhirtxt.Text = "" Or anakketxt.Text = "" Or saudaratxt.Text = "" Or jabatantxt.Text = "" Or statkartxt.Text = "" Or alasankeltxt.Text = "" Or statuspertxt.Text = "" Or jmlhanaktxt.Text = "" Or ketlaintxt.Text = "" Or lastuptxt.Text = "" Or tmptlahirtxt.Text = "" Or tgllahirtxt.Text = "" Or namasuamitxt.Text = "" Or hpsuamitxt.Text = "" Or namaanaktxt.Text = "" Or usiaanaktxt.Text = "" Or namaecontxt.Text = "" Or hubecontxt.Text = "" Or alamatecontxt.Text = "" Or hpecontxt.Text = "" Or kerja1txt.Text = "" Or kerja2txt.Text = "" Or kerja3txt.Text = "" Or recordtxt.Text = "" Then
    MsgBox "Data Belum Lengkap"
    Else
    Dim tambahdata As String
        tambahdata = "Insert Into namakar (TGL_CREATE,NIK,NAMA,ALAMAT_KTP,ALAMAT_TIGGAL,JK,Tlp,HP,TGL_MASUK,NO_KTP,PDDK_AKHIR,ANAK_KE,SAUDARA,JABATAN,STS_KARYAWAN,STS_NIKAH,JML_ANAK,KET_LAIN,LAST_UPDATE,TMP_LAHIR,TGL_LAHIR,NM_SMIATAUIST,HP_SMIATAUIST,NM_ANAK,USIA_ANAK,NM_ECON,HUB_ECON,ALMT_ECON,HP_ECON,KERJA1,KERJA2,KERJA3,RECORD,TGLKONTRAK1) values ('" & tglcretxt.Text & "'," _
        & "'" & niktxt.Text & "','" & namatxt.Text & "','" & alamatktptxt.Text & "','" & alamattgltxt.Text & "','" & jeniskeltxt.Text & "','" & tlptxt.Text & "','" & hptxt.Text & "','" & tglmasuktxt.Value & "','" & noktptxt.Text & "','" & pendterakhirtxt.Text & "','" & anakketxt.Text & "','" & saudaratxt.Text & "','" & jabatantxt.Text & "','" & statkartxt.Text & "','" & statuspertxt.Text & "','" & jmlhanaktxt.Text & "','" & ketlaintxt.Text & "','" & lastuptxt.Text & "','" & tmptlahirtxt.Text & "','" & tgllahirtxt.Text & "','" & namasuamitxt.Text & "','" & hpsuamitxt.Text & "','" & namaanaktxt.Text & "','" & usiaanaktxt.Text & "','" & namaecontxt.Text & "','" & hubecontxt.Text & "','" & alamatecontxt.Text & "','" & hpecontxt.Text & "','" & kerja1txt.Text & "','" & kerja2txt.Text & "','" & kerja3txt.Text & "','" & recordtxt.Text & "', " _
        & " '" & tglkontrak1.Value & "')"
        konn.Execute tambahdata
        MsgBox "Data Berhasil Ditambah", vbInformation, "Pemberitahuan"
        newbtn.Caption = "NEW"
        Adodc1.Refresh
        DGkary.Refresh
        Call KondisiAwal
If konn.State = adStateOpen Then konn.Close
    
    End If
Else
If editbtn.Caption = "EDITDATA" Then
    Dim editdata As String
        editdata = "update namakar set TGL_CREATE = '" & tglcretxt.Text & "',NIK = '" & niktxt.Text & "',NAMA = '" & namatxt.Text & "',ALAMAT_KTP = '" & alamatktptxt.Text & "', ALAMAT_TIGGAL ='" & alamattgltxt.Text & "',JK = '" & jeniskeltxt.Text & "',Tlp = '" & tlptxt.Text & "',HP = '" & hptxt.Text & "',TGL_MASUK = '" & tglmasuktxt.Value & "',TGL_KELUAR = '" & tglresigntxt.Text & "',TGL_AJU_KELUAR = '" & tglajurestxt.Text & "',NO_KTP = '" & noktptxt.Text & "',PDDK_AKHIR = '" & pendterakhirtxt.Text & "',ANAK_KE = '" & anakketxt.Text & "',SAUDARA= '" & saudaratxt.Text & "',JABATAN = '" & jabatantxt.Text & "',STS_KARYAWAN = '" & statkartxt.Text & "',ALASAN_KELUAR = '" & alasankeltxt.Text & "',STS_NIKAH = '" & statuspertxt.Text & "',JML_ANAK = '" & jmlhanaktxt.Text & "'," _
& " KET_LAIN = '" & ketlaintxt.Text & " ',LAST_UPDATE = '" & lastuptxt.Text & "',TMP_LAHIR = '" & tmptlahirtxt.Text & "',TGL_LAHIR = '" & tgllahirtxt.Text & "', NM_SMIATAUIST = '" & namasuamitxt.Text & "',HP_SMIATAUIST = '" & hpsuamitxt.Text & "',NM_ANAK = '" & namaanaktxt.Text & "',USIA_ANAK = '" & usiaanaktxt.Text & "',NM_ECON = '" & namaecontxt.Text & "',HUB_ECON = '" & hubecontxt.Text & "',ALMT_ECON = '" & alamatecontxt.Text & "',HP_ECON = '" & hpecontxt.Text & "',KERJA1 = '" & kerja1txt.Text & "',KERJA2 = '" & kerja2txt.Text & "',KERJA3 = '" & kerja3txt.Text & "',RECORD = '" & recordtxt.Text & "', tglkontrak1 = '" & tglkontrak1.Value & "'  where id = '" & DGkary.Columns(0).Text & "'"
        konn.Execute editdata
        MsgBox "Data Berhasil Diedit", vbInformation, "Pemberitahuan"
        editbtn.Caption = "EDIT"
        Adodc1.Refresh
        DGkary.Refresh
 Call KondisiAwal
If konn.State = adStateOpen Then konn.Close

End If
End If

End Sub

Private Sub seekbtn_Click()
frmseekkaryawan.Show

End Sub

