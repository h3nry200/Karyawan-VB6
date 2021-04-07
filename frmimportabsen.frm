VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmimportdata 
   Caption         =   "IMPORT DATAENTRY"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TxtnamaFile 
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   30
      Width           =   7545
   End
   Begin VB.CommandButton CmdBuka 
      Caption         =   "OPEN FILE"
      Height          =   405
      Left            =   7890
      TabIndex        =   3
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton CmdImport 
      Caption         =   "IMPORT"
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   3210
      Width           =   1155
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   7650
      TabIndex        =   0
      Top             =   3210
      Width           =   1155
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2415
      Left            =   0
      TabIndex        =   2
      Top             =   540
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   4260
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3270
      Top             =   3330
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmimportdata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsExcel As ADODB.Recordset
Dim strSql As String
Dim Baris As Long
Dim SQL As String
Dim koneksi As New rdoConnection
Dim rQuery As New rdoQuery
Dim rs As rdoResultset




Private Sub Form_Load()
koneksi.Connect = "Driver={MySQL ODBC 5.1 Driver};SERVER=localhost;PWD=123;UID=root;PORT=3306;DATABASE=db_qnb;"
koneksi.EstablishConnection
'    Adodc1.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};SERVER=192.168.2.11;PWD=123456;UID=BRIWIRA;PORT=3306;DATABASE=karyawan;"
 '   Adodc1.RecordSource = "tbl_mhs"
  '  Adodc1.Refresh
Call AktifMSFlexGrid1

End Sub

'Untuk Mengatur Tampilan MSFlexGrid1
Sub AktifMSFlexGrid1()

    MSFlexGrid1.Cols = 171
    MSFlexGrid1.RowHeightMin = 300
    '-------------------------------------------------
    MSFlexGrid1.Col = 0
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "ID"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(0) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '-------------------------------------------------
    MSFlexGrid1.Col = 1
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "USER ID"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(1) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '-------------------------------------------------
    MSFlexGrid1.Col = 2
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "STATUS"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(2) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '-------------------------------------------------
    MSFlexGrid1.Col = 3
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TANGGAL REGISTER"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(3) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 4
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "JAM REGISTER"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 5
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NOMOR APLIKASI"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 6
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "AREA MARKET ID"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 7
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NIP"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 8
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NAMA SALES"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 9
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "PRODUCT ID"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 10
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "CREDIT FACTORY"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter

'---------------------------------------------------
    MSFlexGrid1.Col = 11
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TUJUAN PEMINJAMAN"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 12
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "AGAMA"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 13
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "STATUS PEMOHON ID"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 14
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "STATUS PINJAMAN"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 15
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KODE UNIT SALES"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 16
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KODE SALES"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 17
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "JUMLAH PINJAMAN"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 18
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TENOR ID"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 19
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NAMA KTP APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 20
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NAMA KTP MIDDLE APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 21
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NAMA KTP LAST APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 22
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NAMA LENGKAP APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 23
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NO KTP APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 24
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KTP BERLAKU APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 25
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KTP HINGGA APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 26
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "ALAMAT SESUAI KTP APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 27
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TGL LAHIR APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 28
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TEMPAT LAHIR APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 29
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "JENIS KELAMIN APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 30
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "PENDIDIKAN APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 31
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "STATUS PERKAWINAN"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 32
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "WARGA NEGARA APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 33
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NAMA SUAMI / ISTRI"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 34
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "JUMLAH TANGGUNGAN APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 35
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "EMAIL APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 36
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NAMA IBU KANDUNG APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 37
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "ALAMAT SAAT INI APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 38
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "RT APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 39
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "RW APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 40
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KODE POS APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 41
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KELURAHAN APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 42
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KECAMATAN APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 43
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KOTA APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 44
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "PROVINSI APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 45
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "STATUS TEMPAT TINGGAL APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 46
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "LAMA MENEMPATI (TAHUN)"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 47
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "LAMA MENEMPATI (BULAN)"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 48
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "RT KTP APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 49
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "RW KTP APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 50
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KODE POS KTP APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 51
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KELURAHAN KTP APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 52
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KECAMATAN KTP APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 53
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KOTA KTP APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 54
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "PROVINSI KTP APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 55
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TELP RUMAH 1 KODE AREA APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 56
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TELP RUMAH 1 APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 57
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TELP RUMAH 1 EXT APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 58
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TELP RUMAH 2 KODE AREA APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 59
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TELP RUMAH 2 APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 60
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TELP RUMAH 2 EXT APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 61
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NOMOR HP APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 62
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "ALAMAT SURAT APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 63
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "INFO APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 64
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "JENIS PEKERJAAN"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 65
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NAMA PERUSAHAAN"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 66
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KATEGORI INDUSTRI"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 67
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NPWP JOB"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 68
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "ALAMAT PEKERJAAN"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 69
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "RT JOB"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 70
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "RW JOB"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 71
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KODE POS JOB"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 72
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KELURAHAN JOB"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 73
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KECAMATAN JOB"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 74
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KOTA JOB"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 75
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "PROVINSI JOB"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 76
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NOTELP KODE AREA JOB"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 77
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NO TELP JOB"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 78
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NO TELP EXT JOB"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 79
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NO FAX JOB"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 80
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "STATUS PEKERJAAN JOB"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 81
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "OCUPATION JOB"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 82
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "UNIT KERJA JOB"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 83
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "URAIAN JABATAN"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 84
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "JUMLAH KARYAWAN"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 85
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "LAMA BEKERJA (TAHUN)"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 86
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "LAMA BEKERJA (BULAN)"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 87
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "PENGHASILAN BULAN"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 88
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NAMA PERUSAHAAN SEBELUMNYA"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 89
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "LAMA BEKERJA SEBELUMNYA (TAHUN)"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 90
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "LAMA BEKERJA SEBELUMNYA (BULAN)"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 91
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "INFO JOB"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 92
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NAMA LENGKAP ECON"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 93
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "HUBUNGAN ECON"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 94
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "ALAMAT ECON"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 95
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KODE POS ECON"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 96
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "RT ECON"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 97
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "RW ECON"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 98
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KELURAHAN ECON"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 99
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KECAMATAN ECON"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 100
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KOTA ECON"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 101
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "PROVINSI ECON"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 102
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TELP RUMAH ECON KODE AREA"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 103
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TELP RUMAH ECON"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 104
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TELP RUMAH ECON EXT"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 105
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NOMOR HP ECON"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 106
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "INFO ECON"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 107
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KODE BANK CC1"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 108
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NOMOR CC1 "
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 109
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "LIMIT CC1"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 110
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KODE BANK CC2"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 111
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NOMOR CC2"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 112
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "LIMIT CC2"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 113
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NAMA REKENING"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 114
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NOMOR REKENING"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 115
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KODE BANK PENERIMA"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 116
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "CABANG REKENING"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 117
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KOTA REKENING"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 118
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "PILIH ASURANSI"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 119
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "ASURANSI ID"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 120
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NAMA SESUAI KTP (ASURANSI)"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 121
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TANGGAL LAHIR KTP (ASURANSI)"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 122
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "INFO ADD"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 123
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "PEKERJAAN SID"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 124
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "PROFESIONAL ADD"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 125
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "SEKTOR EKONOMI ID"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 126
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KODE LHVP"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 127
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NOMOR REGISTER"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 128
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KODE PRODUCT"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 129
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KETERANGAN PRODUCT"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 130
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NAMA LENGKAP MIDDLE APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 131
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NAMA LENGKAP LAST APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 132
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "MEMO DEDUP 1"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 133
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "MEMO DEDUP 2"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 134
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "MEMO DEDUP 3"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 135
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "MEMO DEDUP 4"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 136
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "URAIAN BIDANG USAHA"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 137
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NPWP APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 138
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NO FAX KODE AREA JOB"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 139
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TUNJANGAN BULAN"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 140
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "OMZET BULANAN"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 141
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NAMA REKENING LOAN"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 142
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NOMOR REKENING LOAN"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 143
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KODE BANK PENERIMA LOAN"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 144
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "CABANG REKENING LOAN"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 145
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KOTA REKENING LOAN"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 146
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KONFIRMASI ALAMAT SESUAI KTP"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 147
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KONFIRMASI KTP SEUMUR HIDUP"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 148
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NAMA KTP FIRST APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 149
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NAMA FIRST APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 150
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "STATUS BI CHECKING"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 151
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "STATUS CAC CHECKING"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 152
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "MEMO CAC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 153
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TGL TTD APLIKASI"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 154
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NO HP APC LAINNYA"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 155
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NO TLP JOB KODE AREA"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 156
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NO TLP JOB LAIN"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 157
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NO TLP JOB EXT LAIN"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 158
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TLP RUMAH ECON KODE AREA LAINNYA"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 159
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TLP RUMAH ECON LAINNYA"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 160
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TLP RUMAH ECON EXT LAINNYA"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 161
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NO HP ECON LAINNYA"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 162
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "ALAMAT SAAT INI1 APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 163
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "ALAMAT SAAT INI2 APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 164
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "ALAMAT SAAT INI3 APC"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 165
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "ALAMAT1 JOB"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 166
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "ALAMAT2 JOB"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 167
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "ALAMAT3 JOB"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 168
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "ALAMAT1 ECON"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 169
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "ALAMAT2 ECON"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 170
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "ALAMAT3 ECON"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter

End Sub

Private Sub CmdBuka_Click()
MSFlexGrid1.Clear

    Call AktifMSFlexGrid1
    Baris = 0
    'Memilih File Excel
    With CommonDialog1
        .DialogTitle = "Pilih File Excelnya (.xls)"
        .InitDir = App.Path
        .Filter = "SQL Files (*.xls)|*.xls"
        'jika menggunakan file excel 2007 keatas
        'untuk .Filter = "SQL Files (*.xls)|*.xls" '
        'Ganti dengan .Filter = "SQL Files (*.xlsx)|*.xlsx"
        .ShowOpen
    End With

   'menampilkan nama filenya di textbox
    TxtnamaFile.Text = CommonDialog1.FileName
    'Membuka File Excel
    If openExcelFile(CommonDialog1.FileName) Then
        'selanjutnya data yg diambil ada di sheet1,
        'sheet disini sama seperti tabel yang ada di database
        strSql = "SELECT * FROM [Sheet1$]" ' penting !!!, jangan lupa menambahkan karakter $
        Set rsExcel = New ADODB.Recordset
        rsExcel.Open strSql, conXls, adOpenForwardOnly, adLockReadOnly
        'tampilkan data yg ada sheet1 ke MSFlexGrid1
        If Not rsExcel.EOF Then
            Do While Not rsExcel.EOF
                Baris = Baris + 1
                MSFlexGrid1.Rows = Baris + 1
                MSFlexGrid1.TextMatrix(Baris, 0) = rsExcel(0).Value
                MSFlexGrid1.TextMatrix(Baris, 1) = "0" 'rsExcel(1).Value
                MSFlexGrid1.TextMatrix(Baris, 2) = "0" 'rsExcel(178).Value
                MSFlexGrid1.TextMatrix(Baris, 3) = Format(rsExcel(1).Value, "ddMMyyyy")
                MSFlexGrid1.TextMatrix(Baris, 4) = Format(rsExcel(1).Value, "yyyy-MM-dd 23:59:00")
                MSFlexGrid1.TextMatrix(Baris, 5) = rsExcel(3).Value
                MSFlexGrid1.TextMatrix(Baris, 6) = "0" 'rsExcel(179).Value ' 164
                MSFlexGrid1.TextMatrix(Baris, 7) = rsExcel(4).Value
                MSFlexGrid1.TextMatrix(Baris, 8) = rsExcel(5).Value
                MSFlexGrid1.TextMatrix(Baris, 9) = "1" 'rsExcel(179).Value '6
                MSFlexGrid1.TextMatrix(Baris, 10) = rsExcel(8).Value
                MSFlexGrid1.TextMatrix(Baris, 11) = "0" 'rsExcel(179).Value '9
                MSFlexGrid1.TextMatrix(Baris, 12) = "0" 'rsExcel(179).Value '10
                MSFlexGrid1.TextMatrix(Baris, 13) = "1" 'rsExcel(179).Value '4
                MSFlexGrid1.TextMatrix(Baris, 14) = "1" 'rsExcel(179).Value '11
                MSFlexGrid1.TextMatrix(Baris, 15) = rsExcel(13).Value
                MSFlexGrid1.TextMatrix(Baris, 16) = rsExcel(14).Value
                MSFlexGrid1.TextMatrix(Baris, 17) = Format(rsExcel(15).Value, "0.00")
                MSFlexGrid1.TextMatrix(Baris, 18) = "0" 'rsExcel(16).Value
                MSFlexGrid1.TextMatrix(Baris, 19) = rsExcel(17).Value
                MSFlexGrid1.TextMatrix(Baris, 20) = rsExcel(143).Value
                MSFlexGrid1.TextMatrix(Baris, 21) = rsExcel(144).Value
                MSFlexGrid1.TextMatrix(Baris, 22) = rsExcel(136).Value
                MSFlexGrid1.TextMatrix(Baris, 23) = rsExcel(19).Value
                MSFlexGrid1.TextMatrix(Baris, 24) = rsExcel(4).Value
                MSFlexGrid1.TextMatrix(Baris, 25) = rsExcel(20).Value
                MSFlexGrid1.TextMatrix(Baris, 26) = rsExcel(39).Value
                MSFlexGrid1.TextMatrix(Baris, 27) = rsExcel(21).Value
                MSFlexGrid1.TextMatrix(Baris, 28) = rsExcel(22).Value
                MSFlexGrid1.TextMatrix(Baris, 29) = "0" 'rsExcel(179).Value '18
                MSFlexGrid1.TextMatrix(Baris, 30) = "0" 'rsExcel(179).Value '23
                MSFlexGrid1.TextMatrix(Baris, 31) = "0" 'rsExcel(179).Value '24
                MSFlexGrid1.TextMatrix(Baris, 32) = rsExcel(4).Value
                MSFlexGrid1.TextMatrix(Baris, 33) = rsExcel(25).Value
                MSFlexGrid1.TextMatrix(Baris, 34) = rsExcel(26).Value
                MSFlexGrid1.TextMatrix(Baris, 35) = Format(rsExcel(139).Value, "0.00")
                MSFlexGrid1.TextMatrix(Baris, 36) = Format(rsExcel(27).Value, "0.00")
                MSFlexGrid1.TextMatrix(Baris, 37) = rsExcel(28).Value
                MSFlexGrid1.TextMatrix(Baris, 38) = rsExcel(29).Value
                MSFlexGrid1.TextMatrix(Baris, 39) = rsExcel(30).Value
                MSFlexGrid1.TextMatrix(Baris, 40) = rsExcel(31).Value
                MSFlexGrid1.TextMatrix(Baris, 41) = rsExcel(32).Value
                MSFlexGrid1.TextMatrix(Baris, 42) = rsExcel(33).Value
                MSFlexGrid1.TextMatrix(Baris, 43) = rsExcel(34).Value
                MSFlexGrid1.TextMatrix(Baris, 44) = rsExcel(4).Value
                MSFlexGrid1.TextMatrix(Baris, 45) = "0" 'rsExcel(179).Value '37
                MSFlexGrid1.TextMatrix(Baris, 46) = rsExcel(38).Value
                MSFlexGrid1.TextMatrix(Baris, 47) = rsExcel(4).Value
                MSFlexGrid1.TextMatrix(Baris, 48) = rsExcel(40).Value
                MSFlexGrid1.TextMatrix(Baris, 49) = rsExcel(41).Value
                MSFlexGrid1.TextMatrix(Baris, 50) = rsExcel(42).Value
                MSFlexGrid1.TextMatrix(Baris, 51) = rsExcel(43).Value
                MSFlexGrid1.TextMatrix(Baris, 52) = rsExcel(44).Value
                MSFlexGrid1.TextMatrix(Baris, 53) = rsExcel(45).Value
                MSFlexGrid1.TextMatrix(Baris, 54) = rsExcel(4).Value
                MSFlexGrid1.TextMatrix(Baris, 55) = rsExcel(4).Value
                MSFlexGrid1.TextMatrix(Baris, 56) = rsExcel(35).Value
                MSFlexGrid1.TextMatrix(Baris, 57) = rsExcel(133).Value
                MSFlexGrid1.TextMatrix(Baris, 58) = rsExcel(4).Value
                MSFlexGrid1.TextMatrix(Baris, 59) = rsExcel(132).Value
                MSFlexGrid1.TextMatrix(Baris, 60) = rsExcel(134).Value
                MSFlexGrid1.TextMatrix(Baris, 61) = rsExcel(36).Value
                MSFlexGrid1.TextMatrix(Baris, 62) = "0" 'rsExcel(179).Value '138
                MSFlexGrid1.TextMatrix(Baris, 63) = rsExcel(4).Value
                MSFlexGrid1.TextMatrix(Baris, 64) = "0" 'rsExcel(179).Value '129
                MSFlexGrid1.TextMatrix(Baris, 65) = rsExcel(54).Value
                MSFlexGrid1.TextMatrix(Baris, 66) = "0" 'rsExcel(179).Value '56
                MSFlexGrid1.TextMatrix(Baris, 67) = rsExcel(59).Value
                MSFlexGrid1.TextMatrix(Baris, 68) = rsExcel(57).Value
                MSFlexGrid1.TextMatrix(Baris, 69) = rsExcel(155).Value
                MSFlexGrid1.TextMatrix(Baris, 70) = rsExcel(156).Value
                MSFlexGrid1.TextMatrix(Baris, 71) = rsExcel(160).Value
                MSFlexGrid1.TextMatrix(Baris, 72) = rsExcel(157).Value
                MSFlexGrid1.TextMatrix(Baris, 73) = rsExcel(158).Value
                MSFlexGrid1.TextMatrix(Baris, 74) = rsExcel(159).Value
                MSFlexGrid1.TextMatrix(Baris, 75) = rsExcel(4).Value
                MSFlexGrid1.TextMatrix(Baris, 76) = rsExcel(4).Value
                MSFlexGrid1.TextMatrix(Baris, 77) = rsExcel(58).Value
                MSFlexGrid1.TextMatrix(Baris, 78) = rsExcel(131).Value
                MSFlexGrid1.TextMatrix(Baris, 79) = rsExcel(130).Value
                MSFlexGrid1.TextMatrix(Baris, 80) = "0" 'rsExcel(179).Value '62
                MSFlexGrid1.TextMatrix(Baris, 81) = "0" 'rsExcel(179).Value '129
                MSFlexGrid1.TextMatrix(Baris, 82) = rsExcel(61).Value
                MSFlexGrid1.TextMatrix(Baris, 83) = rsExcel(60).Value
                MSFlexGrid1.TextMatrix(Baris, 84) = "0" 'rsExcel(179).Value '64
                MSFlexGrid1.TextMatrix(Baris, 85) = rsExcel(65).Value
                MSFlexGrid1.TextMatrix(Baris, 86) = rsExcel(140).Value
                MSFlexGrid1.TextMatrix(Baris, 87) = Format(rsExcel(68).Value, "0.00")
                MSFlexGrid1.TextMatrix(Baris, 88) = rsExcel(67).Value
                MSFlexGrid1.TextMatrix(Baris, 89) = rsExcel(66).Value
                MSFlexGrid1.TextMatrix(Baris, 90) = rsExcel(141).Value
                MSFlexGrid1.TextMatrix(Baris, 91) = rsExcel(4).Value
                MSFlexGrid1.TextMatrix(Baris, 92) = rsExcel(46).Value
                MSFlexGrid1.TextMatrix(Baris, 93) = rsExcel(47).Value
                MSFlexGrid1.TextMatrix(Baris, 94) = rsExcel(48).Value
                MSFlexGrid1.TextMatrix(Baris, 95) = rsExcel(154).Value
                MSFlexGrid1.TextMatrix(Baris, 96) = rsExcel(49).Value
                MSFlexGrid1.TextMatrix(Baris, 97) = rsExcel(50).Value
                MSFlexGrid1.TextMatrix(Baris, 98) = rsExcel(152).Value
                MSFlexGrid1.TextMatrix(Baris, 99) = rsExcel(153).Value
                MSFlexGrid1.TextMatrix(Baris, 100) = rsExcel(51).Value
                MSFlexGrid1.TextMatrix(Baris, 101) = rsExcel(4).Value
                MSFlexGrid1.TextMatrix(Baris, 102) = rsExcel(4).Value
                MSFlexGrid1.TextMatrix(Baris, 103) = rsExcel(52).Value
                MSFlexGrid1.TextMatrix(Baris, 104) = rsExcel(135).Value
                MSFlexGrid1.TextMatrix(Baris, 105) = rsExcel(53).Value
                MSFlexGrid1.TextMatrix(Baris, 106) = rsExcel(4).Value
                MSFlexGrid1.TextMatrix(Baris, 107) = rsExcel(121).Value
                MSFlexGrid1.TextMatrix(Baris, 108) = rsExcel(122).Value
                MSFlexGrid1.TextMatrix(Baris, 109) = rsExcel(123).Value
                MSFlexGrid1.TextMatrix(Baris, 110) = "0" 'rsExcel(124).Value
                MSFlexGrid1.TextMatrix(Baris, 111) = rsExcel(125).Value
                MSFlexGrid1.TextMatrix(Baris, 112) = rsExcel(126).Value
                MSFlexGrid1.TextMatrix(Baris, 113) = rsExcel(98).Value
                MSFlexGrid1.TextMatrix(Baris, 114) = rsExcel(99).Value
                MSFlexGrid1.TextMatrix(Baris, 115) = "1" 'rsExcel(179).Value '100
                MSFlexGrid1.TextMatrix(Baris, 116) = rsExcel(101).Value
                MSFlexGrid1.TextMatrix(Baris, 117) = rsExcel(102).Value
                MSFlexGrid1.TextMatrix(Baris, 118) = "0" 'rsExcel(112).Value
                MSFlexGrid1.TextMatrix(Baris, 119) = "0" 'rsExcel(115).Value
                MSFlexGrid1.TextMatrix(Baris, 120) = rsExcel(113).Value
                MSFlexGrid1.TextMatrix(Baris, 121) = rsExcel(114).Value
                MSFlexGrid1.TextMatrix(Baris, 122) = rsExcel(4).Value
                MSFlexGrid1.TextMatrix(Baris, 123) = rsExcel(165).Value
                MSFlexGrid1.TextMatrix(Baris, 124) = "1" 'rsExcel(179).Value '166
                MSFlexGrid1.TextMatrix(Baris, 125) = "1" 'rsExcel(179).Value '167
                MSFlexGrid1.TextMatrix(Baris, 126) = rsExcel(55).Value
                MSFlexGrid1.TextMatrix(Baris, 127) = rsExcel(4).Value
                MSFlexGrid1.TextMatrix(Baris, 128) = rsExcel(6).Value
                MSFlexGrid1.TextMatrix(Baris, 129) = rsExcel(7).Value
                MSFlexGrid1.TextMatrix(Baris, 130) = rsExcel(146).Value
                MSFlexGrid1.TextMatrix(Baris, 131) = rsExcel(147).Value
                MSFlexGrid1.TextMatrix(Baris, 132) = rsExcel(148).Value
                MSFlexGrid1.TextMatrix(Baris, 133) = rsExcel(4).Value
                MSFlexGrid1.TextMatrix(Baris, 134) = rsExcel(4).Value
                MSFlexGrid1.TextMatrix(Baris, 135) = rsExcel(4).Value
                MSFlexGrid1.TextMatrix(Baris, 136) = rsExcel(4).Value
                MSFlexGrid1.TextMatrix(Baris, 137) = rsExcel(59).Value
                MSFlexGrid1.TextMatrix(Baris, 138) = rsExcel(4).Value
                MSFlexGrid1.TextMatrix(Baris, 139) = rsExcel(127).Value
                MSFlexGrid1.TextMatrix(Baris, 140) = rsExcel(128).Value
                MSFlexGrid1.TextMatrix(Baris, 141) = rsExcel(116).Value
                MSFlexGrid1.TextMatrix(Baris, 142) = rsExcel(117).Value
                MSFlexGrid1.TextMatrix(Baris, 143) = rsExcel(118).Value
                MSFlexGrid1.TextMatrix(Baris, 144) = rsExcel(119).Value
                MSFlexGrid1.TextMatrix(Baris, 145) = rsExcel(120).Value
                MSFlexGrid1.TextMatrix(Baris, 146) = "0" 'rsExcel(150).Value
                MSFlexGrid1.TextMatrix(Baris, 147) = "0" 'rsExcel(4).Value
                MSFlexGrid1.TextMatrix(Baris, 148) = rsExcel(142).Value
                MSFlexGrid1.TextMatrix(Baris, 149) = rsExcel(145).Value
                MSFlexGrid1.TextMatrix(Baris, 150) = "1" 'rsExcel(178).Value '4
                MSFlexGrid1.TextMatrix(Baris, 151) = "1" 'rsExcel(178).Value '4
                MSFlexGrid1.TextMatrix(Baris, 152) = "1" 'rsExcel(178).Value '4
                MSFlexGrid1.TextMatrix(Baris, 153) = Format("2016-11-24", "yyyy-MM-dd")  'rsExcel(178).Value '4
                MSFlexGrid1.TextMatrix(Baris, 154) = "1" 'rsExcel(178).Value '4
                MSFlexGrid1.TextMatrix(Baris, 155) = "1" 'rsExcel(178).Value '4
                MSFlexGrid1.TextMatrix(Baris, 156) = "1" 'rsExcel(178).Value '4
                MSFlexGrid1.TextMatrix(Baris, 157) = "1" 'rsExcel(178).Value '4
                MSFlexGrid1.TextMatrix(Baris, 158) = "1" 'rsExcel(178).Value '4
                MSFlexGrid1.TextMatrix(Baris, 159) = "1" 'rsExcel(178).Value '4
                MSFlexGrid1.TextMatrix(Baris, 160) = "1" 'rsExcel(178).Value '4
                MSFlexGrid1.TextMatrix(Baris, 161) = "1" 'rsExcel(178).Value '4
                MSFlexGrid1.TextMatrix(Baris, 162) = "1" 'rsExcel(178).Value '4
                MSFlexGrid1.TextMatrix(Baris, 163) = "1" 'rsExcel(178).Value '4
                MSFlexGrid1.TextMatrix(Baris, 164) = "1" 'rsExcel(178).Value '4
                MSFlexGrid1.TextMatrix(Baris, 165) = "1" 'rsExcel(178).Value '4
                MSFlexGrid1.TextMatrix(Baris, 166) = "1" 'rsExcel(178).Value '4
                MSFlexGrid1.TextMatrix(Baris, 167) = "1" 'rsExcel(178).Value '4
                MSFlexGrid1.TextMatrix(Baris, 168) = "1" 'rsExcel(178).Value '4
                MSFlexGrid1.TextMatrix(Baris, 169) = "1" 'rsExcel(178).Value '4
                MSFlexGrid1.TextMatrix(Baris, 170) = "1" 'rsExcel(178).Value '4

                
                '  MSFlexGrid1.TextMatrix(Baris, 9) = rsExcel(8).Value
                rsExcel.MoveNext
                DoEvents
            Loop
        End If
        rsExcel.Close
        Set rsExcel = Nothing
    End If

End Sub

Private Sub CmdClose_Click()
koneksi.Close
Unload Me
End Sub

Private Sub CmdImport_Click()
On Error GoTo AdaError

    Dim i As Integer
    Dim tambahdata As String
    

    For i = 1 To MSFlexGrid1.Rows - 1
        tambahdata = ""
        tambahdata = "INSERT INTO tbldataentry " _
        & " VALUES ('" & MSFlexGrid1.TextMatrix(i, 0) & "','" & MSFlexGrid1.TextMatrix(i, 1) & "','" & MSFlexGrid1.TextMatrix(i, 2) & "','" & MSFlexGrid1.TextMatrix(i, 3) & "','" & MSFlexGrid1.TextMatrix(i, 4) & "','" & MSFlexGrid1.TextMatrix(i, 5) & "','" & MSFlexGrid1.TextMatrix(i, 6) & "','" & MSFlexGrid1.TextMatrix(i, 7) & "','" & MSFlexGrid1.TextMatrix(i, 8) & "','" & MSFlexGrid1.TextMatrix(i, 9) & "','" & MSFlexGrid1.TextMatrix(i, 10) & "','" & MSFlexGrid1.TextMatrix(i, 11) & "','" & MSFlexGrid1.TextMatrix(i, 12) & "','" & MSFlexGrid1.TextMatrix(i, 13) & "','" & MSFlexGrid1.TextMatrix(i, 14) & "','" & MSFlexGrid1.TextMatrix(i, 15) & "','" & MSFlexGrid1.TextMatrix(i, 16) & "','" & MSFlexGrid1.TextMatrix(i, 17) & "','" & MSFlexGrid1.TextMatrix(i, 18) & "','" & MSFlexGrid1.TextMatrix(i, 19) & "','" & MSFlexGrid1.TextMatrix(i, 20) & "','" & MSFlexGrid1.TextMatrix(i, 21) & "','" & MSFlexGrid1.TextMatrix(i, 22) & "','" & MSFlexGrid1.TextMatrix(i, 23) & "','" & MSFlexGrid1.TextMatrix(i, 24) & "'," _
        & "'" & MSFlexGrid1.TextMatrix(i, 25) & "','" & MSFlexGrid1.TextMatrix(i, 26) & "','" & MSFlexGrid1.TextMatrix(i, 27) & "','" & MSFlexGrid1.TextMatrix(i, 28) & "','" & MSFlexGrid1.TextMatrix(i, 29) & "','" & MSFlexGrid1.TextMatrix(i, 30) & "','" & MSFlexGrid1.TextMatrix(i, 31) & "','" & MSFlexGrid1.TextMatrix(i, 32) & "','" & MSFlexGrid1.TextMatrix(i, 33) & " ','" & MSFlexGrid1.TextMatrix(i, 34) & " ','" & MSFlexGrid1.TextMatrix(i, 35) & " ','" & MSFlexGrid1.TextMatrix(i, 36) & " ','" & MSFlexGrid1.TextMatrix(i, 37) & " ','" & MSFlexGrid1.TextMatrix(i, 38) & " ','" & MSFlexGrid1.TextMatrix(i, 39) & " ','" & MSFlexGrid1.TextMatrix(i, 40) & " ','" & MSFlexGrid1.TextMatrix(i, 41) & " ','" & MSFlexGrid1.TextMatrix(i, 42) & " ','" & MSFlexGrid1.TextMatrix(i, 43) & " ','" & MSFlexGrid1.TextMatrix(i, 44) & " ','" & MSFlexGrid1.TextMatrix(i, 45) & " ','" & MSFlexGrid1.TextMatrix(i, 46) & " ','" & MSFlexGrid1.TextMatrix(i, 47) & " ','" & MSFlexGrid1.TextMatrix(i, 48) & " ', " _
        & "'" & MSFlexGrid1.TextMatrix(i, 49) & "', '" & MSFlexGrid1.TextMatrix(i, 50) & " ', '" & MSFlexGrid1.TextMatrix(i, 51) & " ', '" & MSFlexGrid1.TextMatrix(i, 52) & " ', '" & MSFlexGrid1.TextMatrix(i, 53) & " ','" & MSFlexGrid1.TextMatrix(i, 54) & " ','" & MSFlexGrid1.TextMatrix(i, 55) & "','" & MSFlexGrid1.TextMatrix(i, 56) & "','" & MSFlexGrid1.TextMatrix(i, 57) & "','" & MSFlexGrid1.TextMatrix(i, 58) & "','" & MSFlexGrid1.TextMatrix(i, 59) & "','" & MSFlexGrid1.TextMatrix(i, 60) & "','" & MSFlexGrid1.TextMatrix(i, 61) & "','" & MSFlexGrid1.TextMatrix(i, 62) & "','" & MSFlexGrid1.TextMatrix(i, 63) & "','" & MSFlexGrid1.TextMatrix(i, 64) & "','" & MSFlexGrid1.TextMatrix(i, 65) & "','" & MSFlexGrid1.TextMatrix(i, 66) & "','" & MSFlexGrid1.TextMatrix(i, 67) & "','" & MSFlexGrid1.TextMatrix(i, 68) & "','" & MSFlexGrid1.TextMatrix(i, 69) & "','" & MSFlexGrid1.TextMatrix(i, 70) & "','" & MSFlexGrid1.TextMatrix(i, 71) & "','" & MSFlexGrid1.TextMatrix(i, 72) & "','" & MSFlexGrid1.TextMatrix(i, 73) & "'," _
        & "'" & MSFlexGrid1.TextMatrix(i, 74) & "','" & MSFlexGrid1.TextMatrix(i, 75) & "','" & MSFlexGrid1.TextMatrix(i, 76) & "','" & MSFlexGrid1.TextMatrix(i, 77) & "','" & MSFlexGrid1.TextMatrix(i, 78) & "','" & MSFlexGrid1.TextMatrix(i, 79) & "','" & MSFlexGrid1.TextMatrix(i, 80) & "','" & MSFlexGrid1.TextMatrix(i, 81) & "','" & MSFlexGrid1.TextMatrix(i, 82) & "','" & MSFlexGrid1.TextMatrix(i, 83) & "','" & MSFlexGrid1.TextMatrix(i, 84) & "','" & MSFlexGrid1.TextMatrix(i, 85) & "','" & MSFlexGrid1.TextMatrix(i, 86) & "','" & MSFlexGrid1.TextMatrix(i, 87) & "','" & MSFlexGrid1.TextMatrix(i, 88) & "','" & MSFlexGrid1.TextMatrix(i, 89) & "','" & MSFlexGrid1.TextMatrix(i, 90) & "','" & MSFlexGrid1.TextMatrix(i, 91) & "','" & MSFlexGrid1.TextMatrix(i, 92) & "','" & MSFlexGrid1.TextMatrix(i, 93) & "','" & MSFlexGrid1.TextMatrix(i, 94) & "','" & MSFlexGrid1.TextMatrix(i, 95) & "','" & MSFlexGrid1.TextMatrix(i, 96) & "','" & MSFlexGrid1.TextMatrix(i, 97) & "','" & MSFlexGrid1.TextMatrix(i, 98) & "'," _
        & "'" & MSFlexGrid1.TextMatrix(i, 99) & "','" & MSFlexGrid1.TextMatrix(i, 100) & "','" & MSFlexGrid1.TextMatrix(i, 101) & "','" & MSFlexGrid1.TextMatrix(i, 102) & "','" & MSFlexGrid1.TextMatrix(i, 103) & "','" & MSFlexGrid1.TextMatrix(i, 104) & "','" & MSFlexGrid1.TextMatrix(i, 105) & "','" & MSFlexGrid1.TextMatrix(i, 106) & "','" & MSFlexGrid1.TextMatrix(i, 107) & "','" & MSFlexGrid1.TextMatrix(i, 108) & "','" & MSFlexGrid1.TextMatrix(i, 109) & "','" & MSFlexGrid1.TextMatrix(i, 110) & "','" & MSFlexGrid1.TextMatrix(i, 111) & "','" & MSFlexGrid1.TextMatrix(i, 112) & "','" & MSFlexGrid1.TextMatrix(i, 113) & "','" & MSFlexGrid1.TextMatrix(i, 114) & "','" & MSFlexGrid1.TextMatrix(i, 115) & "','" & MSFlexGrid1.TextMatrix(i, 116) & "','" & MSFlexGrid1.TextMatrix(i, 117) & "','" & MSFlexGrid1.TextMatrix(i, 118) & "','" & MSFlexGrid1.TextMatrix(i, 119) & "','" & MSFlexGrid1.TextMatrix(i, 120) & "','" & MSFlexGrid1.TextMatrix(i, 121) & "'," _
        & "'" & MSFlexGrid1.TextMatrix(i, 122) & "','" & MSFlexGrid1.TextMatrix(i, 123) & "','" & MSFlexGrid1.TextMatrix(i, 124) & "','" & MSFlexGrid1.TextMatrix(i, 125) & "','" & MSFlexGrid1.TextMatrix(i, 126) & "','" & MSFlexGrid1.TextMatrix(i, 127) & "','" & MSFlexGrid1.TextMatrix(i, 128) & "','" & MSFlexGrid1.TextMatrix(i, 129) & "','" & MSFlexGrid1.TextMatrix(i, 130) & "','" & MSFlexGrid1.TextMatrix(i, 131) & "','" & MSFlexGrid1.TextMatrix(i, 132) & "','" & MSFlexGrid1.TextMatrix(i, 133) & "','" & MSFlexGrid1.TextMatrix(i, 134) & "','" & MSFlexGrid1.TextMatrix(i, 135) & "','" & MSFlexGrid1.TextMatrix(i, 136) & "','" & MSFlexGrid1.TextMatrix(i, 137) & "','" & MSFlexGrid1.TextMatrix(i, 138) & "','" & MSFlexGrid1.TextMatrix(i, 139) & "','" & MSFlexGrid1.TextMatrix(i, 140) & "','" & MSFlexGrid1.TextMatrix(i, 141) & "','" & MSFlexGrid1.TextMatrix(i, 142) & "','" & MSFlexGrid1.TextMatrix(i, 143) & "','" & MSFlexGrid1.TextMatrix(i, 144) & "','" & MSFlexGrid1.TextMatrix(i, 145) & "'," _
        & "'" & MSFlexGrid1.TextMatrix(i, 146) & "','" & MSFlexGrid1.TextMatrix(i, 147) & "','" & MSFlexGrid1.TextMatrix(i, 148) & "','" & MSFlexGrid1.TextMatrix(i, 149) & "','" & MSFlexGrid1.TextMatrix(i, 150) & "','" & MSFlexGrid1.TextMatrix(i, 151) & "','" & MSFlexGrid1.TextMatrix(i, 152) & "','" & MSFlexGrid1.TextMatrix(i, 153) & "','" & MSFlexGrid1.TextMatrix(i, 154) & "','" & MSFlexGrid1.TextMatrix(i, 155) & "','" & MSFlexGrid1.TextMatrix(i, 156) & "','" & MSFlexGrid1.TextMatrix(i, 157) & "','" & MSFlexGrid1.TextMatrix(i, 158) & "','" & MSFlexGrid1.TextMatrix(i, 159) & "','" & MSFlexGrid1.TextMatrix(i, 160) & "','" & MSFlexGrid1.TextMatrix(i, 161) & "','" & MSFlexGrid1.TextMatrix(i, 162) & "','" & MSFlexGrid1.TextMatrix(i, 163) & "','" & MSFlexGrid1.TextMatrix(i, 164) & "','" & MSFlexGrid1.TextMatrix(i, 165) & "','" & MSFlexGrid1.TextMatrix(i, 166) & "','" & MSFlexGrid1.TextMatrix(i, 167) & "','" & MSFlexGrid1.TextMatrix(i, 168) & "','" & MSFlexGrid1.TextMatrix(i, 169) & "'," _
        & "'" & MSFlexGrid1.TextMatrix(i, 170) & "')"

'$query = "INSERT INTO users_tb SET
'          user_status    = '". mysql_real_escape_string($status) ."',
'          user_gender    = '". mysql_real_escape_string($gender) ."',
'          user_firstname = '". mysql_real_escape_string($firstname) ."',
'          user_surname   = '". mysql_real_escape_string($surname) ."',
'          student_number = '". mysql_real_escape_string($hnumber) ."',
'          user_email     = '". mysql_real_escape_string($email) ."',
'          user_dob       = '". mysql_real_escape_string($dob) ."',
'          user_name      = '". mysql_real_escape_string($username) ."',
'          user_pass      = '". mysql_real_escape_string($password) ."'";
'mysql_query($query) or die(mysql_error());
'mysql_close();
        
'        "&rw_ktp_apc ,kodepos_ktp_apc,kalurahan_ktp_apc,kota_ktp_apc,propinsi_ktp_apc,tlprumah1_apc,tlprumah1_ext_apc,tlprumah2_apc,tlprumah2_ext_apc,nohp_apc,alamatsurat_id_apc,info_apc,jenispekerjaan_id_job,namaperusahaan_job,kategoriindustri_id_job,npwp_job,alamat_job,rt_job,rw_job,kodepos_job,kalurahan_job,kecamatan_job,kota_job,propinsi_job,notelp_job,notelp_ext_job,nofax_job,statuspekerjaan_id_job,ocupation_id_job,unitkerja_job,uraianjabatan_job,jmlkaryawan_id_job,lamabekerja_tahun_job,lamabekerja_bulan_job,penghasilan_bulan_job,namaperusahaan_sebelum_job,lamabekerja_sebelum_tahun_job,lamabekerja_sebelum_bulan_job,info_job,namalengkap_econ,hubecon_id_econ,alamat_econ,kodepos_econ &_"

        koneksi.Execute tambahdata
        DoEvents
    Next i
    MsgBox "Import data berhasil, Silahkan di cek...", vbInformation, "... Sukses..."
    Exit Sub
AdaError:
If Err.Number = -2147467259 Then
    MsgBox "id " & MSFlexGrid1.TextMatrix(i, 3) & " sudah ada dalam database." & vbCrLf & _
    "Pada file excelnya di baris " & i + 1 & " ,silahkan hapus terlebih dahulu lalu ulangi.", vbCritical, ".:: Gagal...!!!"
    Exit Sub
Else
    MsgBox "Error No : " & Err.Number & vbCrLf & _
    Err.Description, vbCritical + vbOKOnly, "Error......"

End If

End Sub



