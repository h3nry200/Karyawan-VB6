VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form formseekabsen 
   Caption         =   "SEEK ABSEN"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   6690
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox thncmb 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2160
      TabIndex        =   5
      Top             =   0
      Width           =   1575
   End
   Begin VB.ComboBox blncmb 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4440
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SEEK"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   4440
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   6375
      _ExtentX        =   11245
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
      Left            =   360
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
End
Attribute VB_Name = "formseekabsen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim koneksi As New rdoConnection
Dim rQuery As New rdoQuery
Dim rs As rdoResultset

Private Sub Command1_Click()

frmabsensi.idtxt.Text = DataGrid1.Columns(0).Text
frmabsensi.tgltxt.Text = DataGrid1.Columns(1).Text
frmabsensi.namatxt.Text = DataGrid1.Columns(2).Text
frmabsensi.Combo1.Text = DataGrid1.Columns(3).Text
frmabsensi.Combo2.Text = DataGrid1.Columns(4).Text
frmabsensi.jmktxt.Text = DataGrid1.Columns(5).Text
frmabsensi.jmtxt.Text = DataGrid1.Columns(6).Text
frmabsensi.jpktxt.Text = DataGrid1.Columns(7).Text
frmabsensi.jptxt.Text = DataGrid1.Columns(8).Text
frmabsensi.atttxt.Text = DataGrid1.Columns(9).Text
frmabsensi.bln2cmb.Text = blncmb.Text
frmabsensi.thn2cmb.Text = thncmb.Text
frmabsensi.Adodc1.CommandType = adCmdText
frmabsensi.Adodc1.RecordSource = "select * from absensi where nama like '%" & formseekabsen.Text1.Text & "%'"
frmabsensi.Adodc1.Refresh
frmabsensi.DGkary.Refresh

frmabsensi.Show
Unload Me
koneksi.Close

End Sub

Private Sub Command2_Click()
Unload Me
koneksi.Close
End Sub


Private Sub Form_Load()
koneksi.Connect = "Driver={MySQL ODBC 5.1 Driver};SERVER=192.168.2.11;PWD=123456;UID=BRIWIRA;PORT=3306;DATABASE=karyawan;"
koneksi.EstablishConnection
    Adodc1.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};SERVER=192.168.2.11;PWD=123456;UID=BRIWIRA;PORT=3306;DATABASE=karyawan;"
    Adodc1.RecordSource = "absensi"
    Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
    blncmb.AddItem "JANUARY"
    blncmb.AddItem "FEBRUARY"
    blncmb.AddItem "MARCH"
    blncmb.AddItem "APRIL"
    blncmb.AddItem "MAY"
    blncmb.AddItem "JUNE"
    blncmb.AddItem "JULY"
    blncmb.AddItem "AUGUST"
    blncmb.AddItem "SEPTEMBER"
    blncmb.AddItem "OCTOBER"
    blncmb.AddItem "NOVEMBER"
    blncmb.AddItem "DECEMBER"
    thncmb.AddItem "2016"
    thncmb.AddItem "2017"
    thncmb.AddItem "2018"
    thncmb.AddItem "2019"
    thncmb.AddItem "2020"
    thncmb.AddItem "2021"
    thncmb.AddItem "2022"
    thncmb.AddItem "2023"
    thncmb.AddItem "2024"
    thncmb.AddItem "2025"
    thncmb.AddItem "2026"
    thncmb.AddItem "2027"
    thncmb.AddItem "2028"
    thncmb.AddItem "2029"
    thncmb.AddItem "2030"
    thncmb.AddItem "2031"
    thncmb.AddItem "2032"
    thncmb.AddItem "2033"
    thncmb.AddItem "2034"
    thncmb.AddItem "2035"
    thncmb.AddItem "2036"
    thncmb.AddItem "2037"
    thncmb.AddItem "2038"
    thncmb.AddItem "2039"
    thncmb.AddItem "2040"
blncmb.Text = "AUGUST"
thncmb.Text = "2016"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
With Adodc1
.CommandType = adCmdText
.RecordSource = "select * from absensi where nama like '%" & Text1 & "%'"
.RecordSource = "select * from absensi where tanggal like '%" & blncmb & "%' and tanggal like '%" & thncmb & "%'"
.Refresh
End With
DataGrid1.Refresh
If KeyAscii = 13 Then
frmabsensi.idtxt.Text = DataGrid1.Columns(0).Text
frmabsensi.tgltxt.Text = DataGrid1.Columns(1).Text
frmabsensi.namatxt.Text = DataGrid1.Columns(2).Text
frmabsensi.Combo1.Text = DataGrid1.Columns(3).Text
frmabsensi.Combo2.Text = DataGrid1.Columns(4).Text
frmabsensi.jmktxt.Text = DataGrid1.Columns(5).Text
frmabsensi.jmtxt.Text = DataGrid1.Columns(6).Text
frmabsensi.jpktxt.Text = DataGrid1.Columns(7).Text
frmabsensi.jptxt.Text = DataGrid1.Columns(8).Text
frmabsensi.atttxt.Text = DataGrid1.Columns(9).Text
frmabsensi.bln2cmb.Text = blncmb.Text
frmabsensi.thn2cmb.Text = thncmb.Text
frmabsensi.Adodc1.CommandType = adCmdText
frmabsensi.Adodc1.RecordSource = "select * from absensi where nama like '%" & formseekabsen.Text1.Text & "%'"
frmabsensi.Adodc1.Refresh
frmabsensi.DGkary.Refresh

frmabsensi.Show
Unload Me
koneksi.Close
End If



End Sub

