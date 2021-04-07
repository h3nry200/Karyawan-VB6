VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmseekreportadm 
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   11205
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   3960
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SEEK"
      Height          =   375
      Left            =   8520
      TabIndex        =   1
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   9960
      TabIndex        =   0
      Top             =   3960
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid dgkar 
      Height          =   3855
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
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
      Left            =   120
      Top             =   120
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
Attribute VB_Name = "frmseekreportadm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim koneksi As New rdoConnection
Dim rQuery As New rdoQuery
Dim rs As rdoResultset

Private Sub Command1_Click()
frmreportadm.tglcretxt.Text = dgkar.Columns(22).Text
frmreportadm.jamcrttxt.Text = dgkar.Columns(23).Text
frmreportadm.lastuptxt.Text = dgkar.Columns(24).Text
frmreportadm.jamlasttxt.Text = dgkar.Columns(25).Text
frmreportadm.usertxt.Text = dgkar.Columns(21).Text
frmreportadm.idtxt.Text = dgkar.Columns(0).Text
frmreportadm.DTPicker1.Value = dgkar.Columns(1).Text
frmreportadm.jamincome = dgkar.Columns(2).Text
frmreportadm.DTPicker2.Value = dgkar.Columns(6).Text
frmreportadm.jamfinishhm = dgkar.Columns(7).Text
frmreportadm.DTPicker3.Value = dgkar.Columns(8).Text
frmreportadm.jamfinishoff = dgkar.Columns(9).Text
frmreportadm.DTPicker4.Value = dgkar.Columns(10).Text
frmreportadm.jamfinishhmoff = dgkar.Columns(11).Text
frmreportadm.producttxt.Text = dgkar.Columns(3).Text
frmreportadm.noapltxt.Text = dgkar.Columns(4).Text
frmreportadm.namaaplikantxt.Text = dgkar.Columns(5).Text
frmreportadm.surhmtxt.Text = dgkar.Columns(12).Text
frmreportadm.backupsurhmtxt = dgkar.Columns(13).Text
frmreportadm.surofftxt.Text = dgkar.Columns(14).Text
frmreportadm.backupsurofftxt = dgkar.Columns(15).Text
frmreportadm.surhmofftxt.Text = dgkar.Columns(16).Text
frmreportadm.backupsurhmofftxt = dgkar.Columns(17).Text
frmreportadm.ordertxt.Text = dgkar.Columns(13).Text
frmreportadm.alasantxt.Text = dgkar.Columns(14).Text
frmreportadm.statushmcmb.Text = dgkar.Columns(18).Text
frmreportadm.statusoffcmb.Text = dgkar.Columns(19).Text
frmreportadm.statushmoffcmb.Text = dgkar.Columns(20).Text
frmreportadm.Show
Unload Me
koneksi.Close

End Sub

Private Sub Command2_Click()
Unload Me
koneksi.Close
End Sub

Private Sub Form_Load()
koneksi.Connect = "Driver={MySQL ODBC 5.1 Driver};SERVER=localhost;PWD=123;UID=root;PORT=3306;DATABASE=karyawan;"
koneksi.EstablishConnection
    Adodc1.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};SERVER=localhost;PWD=123;UID=root;PORT=3306;DATABASE=karyawan;"
    Adodc1.RecordSource = "tbl_dataadmin"
    Adodc1.Refresh
Set dgkar.DataSource = Adodc1

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
With Adodc1
.CommandType = adCmdText
.RecordSource = "select * from tbl_dataadmin where noaplikasi like '%" & Text1 & "%'"
'.RecordSource = "select * from absensi where tanggal like '%" & blncmb & "%' and tanggal like '%" & thncmb & "%'"
.Refresh
End With
dgkar.Refresh
If KeyAscii = 13 Then
frmreportadm.tglcretxt.Text = dgkar.Columns(22).Text
frmreportadm.jamcrttxt.Text = dgkar.Columns(23).Text
frmreportadm.lastuptxt.Text = dgkar.Columns(24).Text
frmreportadm.jamlasttxt.Text = dgkar.Columns(25).Text
frmreportadm.usertxt.Text = dgkar.Columns(21).Text
frmreportadm.idtxt.Text = dgkar.Columns(0).Text
frmreportadm.DTPicker1.Value = dgkar.Columns(1).Text
frmreportadm.jamincome = dgkar.Columns(2).Text
frmreportadm.DTPicker2.Value = dgkar.Columns(6).Text
frmreportadm.jamfinishhm = dgkar.Columns(7).Text
frmreportadm.DTPicker3.Value = dgkar.Columns(8).Text
frmreportadm.jamfinishoff = dgkar.Columns(9).Text
frmreportadm.DTPicker4.Value = dgkar.Columns(10).Text
frmreportadm.jamfinishhmoff = dgkar.Columns(11).Text
frmreportadm.producttxt.Text = dgkar.Columns(3).Text
frmreportadm.noapltxt.Text = dgkar.Columns(4).Text
frmreportadm.namaaplikantxt.Text = dgkar.Columns(5).Text
frmreportadm.surhmtxt.Text = dgkar.Columns(12).Text
frmreportadm.backupsurhmtxt = dgkar.Columns(13).Text
frmreportadm.surofftxt.Text = dgkar.Columns(14).Text
frmreportadm.backupsurofftxt = dgkar.Columns(15).Text
frmreportadm.surhmofftxt.Text = dgkar.Columns(16).Text
frmreportadm.backupsurhmofftxt = dgkar.Columns(17).Text
frmreportadm.ordertxt.Text = dgkar.Columns(13).Text
frmreportadm.alasantxt.Text = dgkar.Columns(14).Text
frmreportadm.statushmcmb.Text = dgkar.Columns(18).Text
frmreportadm.statusoffcmb.Text = dgkar.Columns(19).Text
frmreportadm.statushmoffcmb.Text = dgkar.Columns(20).Text
frmreportadm.Show
Unload Me
koneksi.Close
End If

End Sub

