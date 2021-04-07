VERSION 5.00
Begin VB.Form frmpsi15 
   Appearance      =   0  'Flat
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   435
      Left            =   600
      TabIndex        =   38
      Top             =   8400
      Width           =   615
   End
   Begin VB.CommandButton nextbtn 
      Caption         =   "NEXT >>>>>"
      DisabledPicture =   "frmpsi15.frx":0000
      DownPicture     =   "frmpsi15.frx":13EA
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
      Picture         =   "frmpsi15.frx":27D4
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
      Left            =   6000
      TabIndex        =   1
      Top             =   0
      Width           =   3195
   End
   Begin VB.TextBox tanggal 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   3195
   End
   Begin VB.Label Label35 
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   39
      Top             =   840
      Width           =   255
   End
   Begin VB.Line Line23 
      X1              =   960
      X2              =   960
      Y1              =   1200
      Y2              =   8400
   End
   Begin VB.Line Line22 
      X1              =   1320
      X2              =   1320
      Y1              =   1200
      Y2              =   8880
   End
   Begin VB.Label Label34 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   37
      Top             =   8040
      Width           =   255
   End
   Begin VB.Label Label33 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   36
      Top             =   7560
      Width           =   255
   End
   Begin VB.Label Label32 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   35
      Top             =   7080
      Width           =   255
   End
   Begin VB.Label Label31 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   34
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label Label30 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   33
      Top             =   6120
      Width           =   255
   End
   Begin VB.Label Label29 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   32
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label Label28 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   31
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label27 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   30
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label26 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   29
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label25 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   28
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   27
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label24 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   26
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label23 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   25
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label22 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   24
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label21 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   23
      Top             =   1320
      Width           =   255
   End
   Begin VB.Line Line21 
      X1              =   480
      X2              =   18000
      Y1              =   8880
      Y2              =   8880
   End
   Begin VB.Line Line20 
      X1              =   480
      X2              =   18000
      Y1              =   8880
      Y2              =   8880
   End
   Begin VB.Line Line19 
      X1              =   480
      X2              =   18000
      Y1              =   8400
      Y2              =   8400
   End
   Begin VB.Line Line18 
      X1              =   480
      X2              =   18000
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Line Line17 
      X1              =   480
      X2              =   18000
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Line Line16 
      X1              =   480
      X2              =   18000
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line15 
      X1              =   480
      X2              =   18000
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line14 
      X1              =   480
      X2              =   18000
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line13 
      X1              =   480
      X2              =   18000
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Line Line12 
      X1              =   480
      X2              =   18000
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line11 
      X1              =   480
      X2              =   18000
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line Line10 
      X1              =   480
      X2              =   18000
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line9 
      X1              =   480
      X2              =   18000
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line8 
      X1              =   480
      X2              =   18000
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line7 
      X1              =   480
      X2              =   18000
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line6 
      X1              =   480
      X2              =   18000
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Line Line5 
      X1              =   480
      X2              =   18000
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line4 
      X1              =   480
      X2              =   18000
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label20 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   22
      Top             =   8040
      Width           =   255
   End
   Begin VB.Label Label19 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   21
      Top             =   7560
      Width           =   255
   End
   Begin VB.Label Label18 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   20
      Top             =   7080
      Width           =   255
   End
   Begin VB.Label Label17 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   19
      Top             =   6600
      Width           =   255
   End
   Begin VB.Label Label16 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   18
      Top             =   6120
      Width           =   255
   End
   Begin VB.Label Label15 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   17
      Top             =   5640
      Width           =   255
   End
   Begin VB.Label Label14 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   16
      Top             =   5160
      Width           =   255
   End
   Begin VB.Label Label13 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   4680
      Width           =   255
   End
   Begin VB.Label Label12 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   4200
      Width           =   255
   End
   Begin VB.Label Label11 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   2280
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   1800
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   8
      Top             =   1320
      Width           =   255
   End
   Begin VB.Line Line3 
      X1              =   480
      X2              =   480
      Y1              =   1200
      Y2              =   8880
   End
   Begin VB.Line Line2 
      X1              =   480
      X2              =   18000
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label8 
      Caption         =   "Tanggal Test :"
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   360
      Width           =   1305
   End
   Begin VB.Label Label7 
      Caption         =   "NAMA           :"
      Height          =   165
      Left            =   4680
      TabIndex        =   6
      Top             =   0
      Width           =   1245
   End
   Begin VB.Label Label6 
      Caption         =   "ID                 :"
      Height          =   285
      Left            =   30
      TabIndex        =   5
      Top             =   0
      Width           =   1125
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   20370
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Caption         =   "TEST XV"
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
      TabIndex        =   4
      Top             =   120
      Width           =   2355
   End
End
Attribute VB_Name = "frmpsi15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
