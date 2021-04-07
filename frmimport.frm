VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmimportkar 
   Caption         =   "INPORT NAMA KARYAWAN"
   ClientHeight    =   4005
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdClose 
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   7830
      TabIndex        =   4
      Top             =   3360
      Width           =   1155
   End
   Begin VB.CommandButton CmdImport 
      Caption         =   "IMPORT"
      Height          =   375
      Left            =   6660
      TabIndex        =   3
      Top             =   3360
      Width           =   1155
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2415
      Left            =   180
      TabIndex        =   2
      Top             =   690
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   4260
      _Version        =   393216
   End
   Begin VB.CommandButton CmdBuka 
      Caption         =   "OPEN FILE"
      Height          =   405
      Left            =   8070
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox TxtnamaFile 
      Height          =   375
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   7545
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3450
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   450
      Left            =   150
      Top             =   1410
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   794
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
Attribute VB_Name = "frmimportkar"
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


'Untuk Mengatur Tampilan MSFlexGrid1
Sub AktifMSFlexGrid1()

    MSFlexGrid1.Cols = 38
    MSFlexGrid1.RowHeightMin = 300
    '-------------------------------------------------
    MSFlexGrid1.Col = 0
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NO"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(0) = 500
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '-------------------------------------------------
    MSFlexGrid1.Col = 1
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "ID"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(1) = 900
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '-------------------------------------------------
    MSFlexGrid1.Col = 2
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TANGGAL CREATE"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(2) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '-------------------------------------------------
    MSFlexGrid1.Col = 3
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NIK"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(3) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 4
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NAMA"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 5
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "ALAMAT KTP"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 6
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "ALAMAT TINGGAL"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 7
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "JK"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 8
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TELP"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 9
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "HP"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 10
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TANGGAL MASUK"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 11
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TANGGAL KELUAR"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 12
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TANGGAL AJU KELUAR"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 13
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NO KTP"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 14
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "PEND TERAKHIR"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 15
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "ANAK KE"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 16
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "SAUDARA"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 17
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "JABATAN"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 18
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "STATUS KARY"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 19
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "ALASAN KELUAR"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 20
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "STATUS NIKAH"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 21
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "JUMLAH ANAK"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 22
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KETERANGAN LAIN"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 23
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "LAST UPDATE"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 24
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TEMPAT LAHIR"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 25
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "TANGGAL LAHIR"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 26
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NAMA SMI/ISTRI"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 27
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NAMA SMI/ISTRI"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 28
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NAMA ANAK"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 29
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "USIA ANAK"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 30
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NAMA ECON"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 31
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "HUB ECON"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 32
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "ALAMAT ECON"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 33
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "HP ECON"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 34
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KERJA1"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 35
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KERJA2"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 36
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "KERJA3"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 37
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "RECORD"
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
                MSFlexGrid1.TextMatrix(Baris, 0) = Baris
                MSFlexGrid1.TextMatrix(Baris, 1) = rsExcel(0).Value
                MSFlexGrid1.TextMatrix(Baris, 2) = rsExcel(1).Value
                MSFlexGrid1.TextMatrix(Baris, 3) = rsExcel(2).Value
                MSFlexGrid1.TextMatrix(Baris, 4) = rsExcel(3).Value
                MSFlexGrid1.TextMatrix(Baris, 5) = rsExcel(4).Value
                MSFlexGrid1.TextMatrix(Baris, 6) = rsExcel(5).Value
                MSFlexGrid1.TextMatrix(Baris, 7) = rsExcel(6).Value
                MSFlexGrid1.TextMatrix(Baris, 8) = rsExcel(7).Value
                MSFlexGrid1.TextMatrix(Baris, 9) = rsExcel(8).Value
                MSFlexGrid1.TextMatrix(Baris, 10) = rsExcel(9).Value
                MSFlexGrid1.TextMatrix(Baris, 11) = rsExcel(10).Value
                MSFlexGrid1.TextMatrix(Baris, 12) = rsExcel(11).Value
                MSFlexGrid1.TextMatrix(Baris, 13) = rsExcel(12).Value
                MSFlexGrid1.TextMatrix(Baris, 14) = rsExcel(13).Value
                MSFlexGrid1.TextMatrix(Baris, 15) = rsExcel(14).Value
                MSFlexGrid1.TextMatrix(Baris, 16) = rsExcel(15).Value
                MSFlexGrid1.TextMatrix(Baris, 17) = rsExcel(16).Value
                MSFlexGrid1.TextMatrix(Baris, 18) = rsExcel(17).Value
                MSFlexGrid1.TextMatrix(Baris, 19) = rsExcel(18).Value
                MSFlexGrid1.TextMatrix(Baris, 20) = rsExcel(19).Value
                MSFlexGrid1.TextMatrix(Baris, 21) = rsExcel(20).Value
                MSFlexGrid1.TextMatrix(Baris, 22) = rsExcel(21).Value
                MSFlexGrid1.TextMatrix(Baris, 23) = rsExcel(22).Value
                MSFlexGrid1.TextMatrix(Baris, 24) = rsExcel(23).Value
                MSFlexGrid1.TextMatrix(Baris, 25) = rsExcel(24).Value
                MSFlexGrid1.TextMatrix(Baris, 26) = rsExcel(25).Value
                MSFlexGrid1.TextMatrix(Baris, 27) = rsExcel(26).Value
                MSFlexGrid1.TextMatrix(Baris, 28) = rsExcel(27).Value
                MSFlexGrid1.TextMatrix(Baris, 29) = rsExcel(28).Value
                MSFlexGrid1.TextMatrix(Baris, 30) = rsExcel(29).Value
                MSFlexGrid1.TextMatrix(Baris, 31) = rsExcel(30).Value
                MSFlexGrid1.TextMatrix(Baris, 32) = rsExcel(31).Value
                MSFlexGrid1.TextMatrix(Baris, 33) = rsExcel(32).Value
                MSFlexGrid1.TextMatrix(Baris, 34) = rsExcel(33).Value
                MSFlexGrid1.TextMatrix(Baris, 35) = rsExcel(34).Value
                MSFlexGrid1.TextMatrix(Baris, 36) = rsExcel(35).Value
                'MSFlexGrid1.TextMatrix(Baris, 37) = rsExcel(36).value
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
        tambahdata = "INSERT INTO namakar(id,TGL_CREATE,NIK,NAMA,ALAMAT_KTP,ALAMAT_TIGGAL,JK,Tlp,HP,TGL_MASUK,TGL_KELUAR,TGL_AJU_KELUAR,NO_KTP,PDDK_AKHIR,ANAK_KE,SAUDARA,JABATAN,STS_KARYAWAN,ALASAN_KELUAR,STS_NIKAH,JML_ANAK,KET_LAIN,LAST_UPDATE,TMP_LAHIR,TGL_LAHIR,NM_SMIATAUIST,HP_SMIATAUIST,NM_ANAK,USIA_ANAK,NM_ECON,HUB_ECON,ALMT_ECON,HP_ECON,KERJA1,KERJA2,KERJA3,RECORD) " _
            & "VALUES ('" & MSFlexGrid1.TextMatrix(i, 1) & "'," & "'" & MSFlexGrid1.TextMatrix(i, 2) & "'," _
            & "'" & MSFlexGrid1.TextMatrix(i, 3) & "'," & "'" & MSFlexGrid1.TextMatrix(i, 4) & "'," _
            & "'" & MSFlexGrid1.TextMatrix(i, 5) & "'," & "'" & MSFlexGrid1.TextMatrix(i, 6) & "'," _
            & "'" & MSFlexGrid1.TextMatrix(i, 7) & "'," & "'" & MSFlexGrid1.TextMatrix(i, 8) & "'," _
            & "'" & MSFlexGrid1.TextMatrix(i, 9) & "'," & "'" & MSFlexGrid1.TextMatrix(i, 10) & "'," _
            & "'" & MSFlexGrid1.TextMatrix(i, 11) & "'," & "'" & MSFlexGrid1.TextMatrix(i, 12) & "'," _
            & "'" & MSFlexGrid1.TextMatrix(i, 13) & "'," & "'" & MSFlexGrid1.TextMatrix(i, 14) & "'," _
            & "'" & MSFlexGrid1.TextMatrix(i, 15) & "'," & "'" & MSFlexGrid1.TextMatrix(i, 16) & "'," _
            & "'" & MSFlexGrid1.TextMatrix(i, 17) & "'," & "'" & MSFlexGrid1.TextMatrix(i, 18) & "'," _
            & "'" & MSFlexGrid1.TextMatrix(i, 19) & "'," & "'" & MSFlexGrid1.TextMatrix(i, 20) & "'," _
            & "'" & MSFlexGrid1.TextMatrix(i, 21) & "'," & "'" & MSFlexGrid1.TextMatrix(i, 22) & "'," _
            & "'" & MSFlexGrid1.TextMatrix(i, 23) & "'," & "'" & MSFlexGrid1.TextMatrix(i, 24) & "'," _
            & "'" & MSFlexGrid1.TextMatrix(i, 25) & "'," & "'" & MSFlexGrid1.TextMatrix(i, 26) & "'," _
            & "'" & MSFlexGrid1.TextMatrix(i, 27) & "'," & "'" & MSFlexGrid1.TextMatrix(i, 28) & "'," _
            & "'" & MSFlexGrid1.TextMatrix(i, 29) & "'," & "'" & MSFlexGrid1.TextMatrix(i, 30) & "'," _
            & "'" & MSFlexGrid1.TextMatrix(i, 31) & "'," & "'" & MSFlexGrid1.TextMatrix(i, 32) & "'," _
            & "'" & MSFlexGrid1.TextMatrix(i, 33) & "'," _
            & "'" & MSFlexGrid1.TextMatrix(i, 34) & "'," _
            & "'" & MSFlexGrid1.TextMatrix(i, 35) & "'," _
            & "'" & MSFlexGrid1.TextMatrix(i, 36) & "'," _
            & "'" & MSFlexGrid1.TextMatrix(i, 37) & "')"
        koneksi.Execute tambahdata
        DoEvents
    Next i
    MsgBox "Import data berhasil, Silahkan di cek...", vbInformation, ".:: Sukses..."
    Exit Sub
AdaError:
If Err.Number = -2147467259 Then
    MsgBox "nama " & MSFlexGrid1.TextMatrix(i, 4) & " sudah ada dalam database." & vbCrLf & _
    "Pada file excelnya di baris " & i + 1 & " ,silahkan hapus terlebih dahulu lalu ulangi.", vbCritical, ".:: Gagal...!!!"
    Exit Sub
Else
    MsgBox "Error No : " & Err.Number & vbCrLf & _
    Err.Description, vbCritical + vbOKOnly, "Error......"

End If

End Sub

Private Sub Form_Load()
koneksi.Connect = "Driver={MySQL ODBC 5.1 Driver};SERVER=192.168.2.11;PWD=123456;UID=BRIWIRA;PORT=3306;DATABASE=karyawan;"
koneksi.EstablishConnection
'    Adodc1.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};SERVER=192.168.2.11;PWD=123456;UID=BRIWIRA;PORT=3306;DATABASE=karyawan;"
 '   Adodc1.RecordSource = "tbl_mhs"
  '  Adodc1.Refresh
Call AktifMSFlexGrid1
End Sub

