VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmseekkaryawan 
   Caption         =   "SEEK KARYAWAN"
   ClientHeight    =   4980
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SEEK"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   3735
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7435
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
      Height          =   375
      Left            =   360
      Top             =   3480
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
Attribute VB_Name = "frmseekkaryawan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim koneksi As New rdoConnection
Dim rQuery As New rdoQuery
Dim rs As rdoResultset

Private Sub Command1_Click()

'frmmstrkary.idtxt.Text = DataGrid1.Columns(0).Text
frmmstrkary.tglcretxt.Text = DataGrid1.Columns(1).Text
frmmstrkary.niktxt.Text = DataGrid1.Columns(2).Text
frmmstrkary.namatxt.Text = DataGrid1.Columns(3).Text
frmmstrkary.alamatktptxt.Text = DataGrid1.Columns(4).Text
frmmstrkary.alamattgltxt.Text = DataGrid1.Columns(5).Text
frmmstrkary.jeniskeltxt.Text = DataGrid1.Columns(6).Text
frmmstrkary.tlptxt.Text = DataGrid1.Columns(7).Text
frmmstrkary.hptxt.Text = DataGrid1.Columns(8).Text
frmmstrkary.tglmasuktxt.Value = DataGrid1.Columns(9).Text
frmmstrkary.tglresigntxt.Text = DataGrid1.Columns(10).Text
frmmstrkary.tglajurestxt.Text = DataGrid1.Columns(11).Text
frmmstrkary.noktptxt.Text = DataGrid1.Columns(12).Text
frmmstrkary.pendterakhirtxt.Text = DataGrid1.Columns(13).Text
frmmstrkary.anakketxt.Text = DataGrid1.Columns(14).Text
frmmstrkary.saudaratxt.Text = DataGrid1.Columns(15).Text
frmmstrkary.jabatantxt.Text = DataGrid1.Columns(16).Text
frmmstrkary.statkartxt.Text = DataGrid1.Columns(17).Text
frmmstrkary.alasankeltxt.Text = DataGrid1.Columns(18).Text
frmmstrkary.statuspertxt.Text = DataGrid1.Columns(19).Text
frmmstrkary.jmlhanaktxt.Text = DataGrid1.Columns(20).Text
frmmstrkary.ketlaintxt.Text = DataGrid1.Columns(21).Text
frmmstrkary.lastuptxt.Text = DataGrid1.Columns(22).Text
frmmstrkary.tmptlahirtxt.Text = DataGrid1.Columns(23).Text
frmmstrkary.tgllahirtxt.Text = DataGrid1.Columns(24).Text
frmmstrkary.namasuamitxt.Text = DataGrid1.Columns(25).Text
frmmstrkary.hpsuamitxt.Text = DataGrid1.Columns(26).Text
frmmstrkary.namaanaktxt.Text = DataGrid1.Columns(27).Text
frmmstrkary.usiaanaktxt.Text = DataGrid1.Columns(28).Text
frmmstrkary.namaecontxt.Text = DataGrid1.Columns(29).Text
frmmstrkary.hubecontxt.Text = DataGrid1.Columns(30).Text
frmmstrkary.alamatecontxt.Text = DataGrid1.Columns(31).Text
frmmstrkary.hpecontxt.Text = DataGrid1.Columns(32).Text
frmmstrkary.kerja1txt.Text = DataGrid1.Columns(33).Text
frmmstrkary.kerja2txt.Text = DataGrid1.Columns(34).Text
frmmstrkary.kerja3txt.Text = DataGrid1.Columns(35).Text
frmmstrkary.recordtxt.Text = DataGrid1.Columns(36).Text

frmmstrkary.Adodc1.CommandType = adCmdText
frmmstrkary.Adodc1.RecordSource = "select * from namakar where nama like '%" & frmseekkaryawan.Text1.Text & "%'"
frmmstrkary.Adodc1.Refresh
frmmstrkary.DGkary.Refresh


frmmstrkary.Show
Unload Me
koneksi.Close

End Sub


Private Sub Command2_Click()
koneksi.Close
Unload Me
End Sub


Private Sub Form_Load()
koneksi.Connect = "Driver={MySQL ODBC 5.1 Driver};SERVER=192.168.2.11;PWD=123456;UID=BRIWIRA;PORT=3306;DATABASE=karyawan;"
koneksi.EstablishConnection
    Adodc1.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};SERVER=192.168.2.11;PWD=123456;UID=BRIWIRA;PORT=3306;DATABASE=karyawan;"
    Adodc1.RecordSource = "namakar"
    Adodc1.Refresh
Set DataGrid1.DataSource = Adodc1
'Text1.SetFocus

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
With Adodc1
.CommandType = adCmdText
.RecordSource = "select * from namakar where nama like '%" & Text1 & "%'"
.Refresh
End With
DataGrid1.Refresh
If KeyAscii = 13 Then
'frmmstrkary.idtxt.Text = DataGrid1.Columns(0).Text
frmmstrkary.tglcretxt.Text = DataGrid1.Columns(1).Text
frmmstrkary.niktxt.Text = DataGrid1.Columns(2).Text
frmmstrkary.namatxt.Text = DataGrid1.Columns(3).Text
frmmstrkary.alamatktptxt.Text = DataGrid1.Columns(4).Text
frmmstrkary.alamattgltxt.Text = DataGrid1.Columns(5).Text
frmmstrkary.jeniskeltxt.Text = DataGrid1.Columns(6).Text
frmmstrkary.tlptxt.Text = DataGrid1.Columns(7).Text
frmmstrkary.hptxt.Text = DataGrid1.Columns(8).Text
frmmstrkary.tglmasuktxt.Value = DataGrid1.Columns(9).Text
frmmstrkary.tglresigntxt.Text = DataGrid1.Columns(10).Text
frmmstrkary.tglajurestxt.Text = DataGrid1.Columns(11).Text
frmmstrkary.noktptxt.Text = DataGrid1.Columns(12).Text
frmmstrkary.pendterakhirtxt.Text = DataGrid1.Columns(13).Text
frmmstrkary.anakketxt.Text = DataGrid1.Columns(14).Text
frmmstrkary.saudaratxt.Text = DataGrid1.Columns(15).Text
frmmstrkary.jabatantxt.Text = DataGrid1.Columns(16).Text
frmmstrkary.statkartxt.Text = DataGrid1.Columns(17).Text
frmmstrkary.alasankeltxt.Text = DataGrid1.Columns(18).Text
frmmstrkary.statuspertxt.Text = DataGrid1.Columns(19).Text
frmmstrkary.jmlhanaktxt.Text = DataGrid1.Columns(20).Text
frmmstrkary.ketlaintxt.Text = DataGrid1.Columns(21).Text
frmmstrkary.lastuptxt.Text = DataGrid1.Columns(22).Text
frmmstrkary.tmptlahirtxt.Text = DataGrid1.Columns(23).Text
frmmstrkary.tgllahirtxt.Text = DataGrid1.Columns(24).Text
frmmstrkary.namasuamitxt.Text = DataGrid1.Columns(25).Text
frmmstrkary.hpsuamitxt.Text = DataGrid1.Columns(26).Text
frmmstrkary.namaanaktxt.Text = DataGrid1.Columns(27).Text
frmmstrkary.usiaanaktxt.Text = DataGrid1.Columns(28).Text
frmmstrkary.namaecontxt.Text = DataGrid1.Columns(29).Text
frmmstrkary.hubecontxt.Text = DataGrid1.Columns(30).Text
frmmstrkary.alamatecontxt.Text = DataGrid1.Columns(31).Text
frmmstrkary.hpecontxt.Text = DataGrid1.Columns(32).Text
frmmstrkary.kerja1txt.Text = DataGrid1.Columns(33).Text
frmmstrkary.kerja2txt.Text = DataGrid1.Columns(34).Text
frmmstrkary.kerja3txt.Text = DataGrid1.Columns(35).Text
frmmstrkary.recordtxt.Text = DataGrid1.Columns(36).Text

frmmstrkary.Adodc1.CommandType = adCmdText
frmmstrkary.Adodc1.RecordSource = "select * from namakar where nama like '%" & frmseekkaryawan.Text1.Text & "%'"
frmmstrkary.Adodc1.Refresh
frmmstrkary.DGkary.Refresh

frmmstrkary.Show
Unload Me
koneksi.Close

End If
End Sub
