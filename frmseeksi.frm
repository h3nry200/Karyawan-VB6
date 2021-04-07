VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmseeksi 
   Caption         =   "FORM SEEK SURAT IJIN"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15090
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   15090
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker tanggalijin 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   3960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   150142977
      CurrentDate     =   42655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SEEK"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   4440
      Width           =   3735
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3855
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   6800
      _Version        =   393216
      Enabled         =   0   'False
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   240
      Top             =   720
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
   Begin VB.Label Label2 
      Caption         =   "NAMA               :"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "TANGGAL IJIN :"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4080
      Width           =   1335
   End
End
Attribute VB_Name = "frmseeksi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim koneksi As New rdoConnection
Dim rQuery As New rdoQuery
Dim rs As rdoResultset

Private Sub Command1_Click()
frmsuratijin.idtxt.Text = DataGrid1.Columns(0).Text
frmsuratijin.namatxt.Text = DataGrid1.Columns(1).Text
frmsuratijin.tglawal.Value = DataGrid1.Columns(2).Text
frmsuratijin.tglakhir.Value = DataGrid1.Columns(3).Text
frmsuratijin.lamatxt.Text = DataGrid1.Columns(4).Text
frmsuratijin.statuscmb.Text = DataGrid1.Columns(5).Text
frmsuratijin.kettxt.Text = DataGrid1.Columns(6).Text
frmsuratijin.appr.Value = DataGrid1.Columns(7).Text
frmsuratijin.Adodc1.CommandType = adCmdText
frmsuratijin.Adodc1.RecordSource = "select * from suratijin where nama like '%" & frmseeksi.Text1.Text & "%'"
frmsuratijin.Adodc1.Refresh
frmsuratijin.DGsrt.Refresh

frmsuratijin.Show
Unload Me
koneksi.Close

End Sub

Private Sub Form_Load()
koneksi.Connect = "Driver={MySQL ODBC 5.1 Driver};SERVER=192.168.2.11;PWD=123456;UID=BRIWIRA;PORT=3306;DATABASE=karyawan;"
koneksi.EstablishConnection
    Adodc1.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};SERVER=192.168.2.11;PWD=123456;UID=BRIWIRA;PORT=3306;DATABASE=karyawan;"
    Adodc1.RecordSource = "suratijin"
    Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1

End Sub

Private Sub Command2_Click()
Unload Me
koneksi.Close
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
With Adodc1
.CommandType = adCmdText
.RecordSource = "select * from suratijin where nama like '%" & Text1 & "%'"
.RecordSource = "select * from suratijin where tgl_awal_ijin = '" & tanggalijin.Value & "'"
.Refresh
End With
DataGrid1.Refresh
If KeyAscii = 13 Then
frmsuratijin.idtxt.Text = DataGrid1.Columns(0).Text
frmsuratijin.namatxt.Text = DataGrid1.Columns(1).Text
frmsuratijin.tglawal.Value = DataGrid1.Columns(2).Text
frmsuratijin.tglakhir.Value = DataGrid1.Columns(3).Text
frmsuratijin.lamatxt.Text = DataGrid1.Columns(4).Text
frmsuratijin.statuscmb.Text = DataGrid1.Columns(5).Text
frmsuratijin.kettxt.Text = DataGrid1.Columns(6).Text
frmsuratijin.appr.Value = DataGrid1.Columns(7).Text
frmsuratijin.Adodc1.CommandType = adCmdText
frmsuratijin.Adodc1.RecordSource = "select * from suratijin where nama like '%" & frmseeksi.Text1.Text & "%'"
frmsuratijin.Adodc1.Refresh
frmsuratijin.DGsrt.Refresh
frmsuratijin.Show
Unload Me
koneksi.Close

End If
End Sub

