VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmimportreportadm 
   Caption         =   "Form1"
   ClientHeight    =   4080
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   9135
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   2640
      Top             =   3360
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "CLOSE"
      Height          =   375
      Left            =   7650
      TabIndex        =   3
      Top             =   3570
      Width           =   1155
   End
   Begin VB.CommandButton CmdImport 
      Caption         =   "IMPORT"
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   3570
      Width           =   1155
   End
   Begin VB.CommandButton CmdBuka 
      Caption         =   "OPEN FILE"
      Height          =   405
      Left            =   7890
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox TxtnamaFile 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   390
      Width           =   7545
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2415
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   4260
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbwaktu 
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label lbuser 
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5895
   End
   Begin VB.Label lbjam 
      Height          =   375
      Left            =   8280
      TabIndex        =   5
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frmimportreportadm"
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
Dim shari As String
Dim ahari

Private Sub Form_Load()
koneksi.Connect = "Driver={MySQL ODBC 5.1 Driver};SERVER=localhost;PWD=123;UID=root;PORT=3306;DATABASE=karyawan;"
koneksi.EstablishConnection
Call AktifMSFlexGrid1

  ahari = Array("Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu")
  Timer1.Interval = 500
  Timer1.Enabled = True
'lbuser.Caption = "' & usertxt.Text & '"
lbuser.Caption = mdihalutama.Caption

End Sub

Private Sub Timer1_Timer()
 shari = ahari(Abs(Weekday(Date) - 1))
  lbwaktu.Caption = "" & shari & ", " _
                   & Format(Date, "dd mmmm yyyy")
lbjam.Caption = Format(Time, "hh:mm:ss")
End Sub

'Untuk Mengatur Tampilan MSFlexGrid1
Sub AktifMSFlexGrid1()

    MSFlexGrid1.Cols = 21
    MSFlexGrid1.RowHeightMin = 300
    '-------------------------------------------------
    MSFlexGrid1.Col = 0
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NO"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(0) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '-------------------------------------------------
    MSFlexGrid1.Col = 1
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "INCOMING DATE"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(1) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '-------------------------------------------------
    MSFlexGrid1.Col = 2
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "JAM INCOMING DATE"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(2) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '-------------------------------------------------
    MSFlexGrid1.Col = 3
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "PRODUCT"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(3) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 4
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NOMOR APLIKASI"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(4) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 5
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "NAMA APLIKAN"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(5) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 6
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "FINISH HOME"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(6) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
    '---------------------------------------------------
    MSFlexGrid1.Col = 7
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "JAM FINISH HOME"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(7) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 8
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "FINISH OFFICE"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(8) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 9
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "JAM FINISH OFFICE"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(9) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 10
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "FINISH HOME + OFFICE"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(10) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 11
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "JAM FINISH HOME + OFFICE"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(11) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 12
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "SURVEYOR HOME"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(12) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 13
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "BACK UP HOME"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(13) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 14
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "SURVEYOR OFFICE"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(14) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 15
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "BACK UP OFFICE"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(15) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 16
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "SURVEYOR HOME + OFFICE"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(16) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 17
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "BACK UP HOME + OFFICE"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(17) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 18
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "STATUS HOME"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(18) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 19
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "STATUS OFFICE"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(19) = 2000
    MSFlexGrid1.AllowUserResizing = flexResizeColumns
    MSFlexGrid1.CellAlignment = flexAlignCenterCenter
'---------------------------------------------------
    MSFlexGrid1.Col = 20
    MSFlexGrid1.Row = 0
    MSFlexGrid1.Text = "STATUS HOME + OFFICE"
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColWidth(20) = 2000
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
'mengambil tgl crt------------------------------------------
                MSFlexGrid1.TextMatrix(Baris, 1) = Left(rsExcel(4).Value, 10)
'mengambil jam crt------------------------------------------
                MSFlexGrid1.TextMatrix(Baris, 2) = Right(rsExcel(4).Value, 8)
                MSFlexGrid1.TextMatrix(Baris, 3) = rsExcel(23).Value
                MSFlexGrid1.TextMatrix(Baris, 4) = rsExcel(2).Value
                MSFlexGrid1.TextMatrix(Baris, 5) = rsExcel(3).Value
'mengambil tgl finish home------------------------------------------
                MSFlexGrid1.TextMatrix(Baris, 6) = Left(rsExcel(5).Value, 10)
'mengambil jam finish home-------------------------------------------
                MSFlexGrid1.TextMatrix(Baris, 7) = Right(rsExcel(5).Value, 8)
'mengambil tgl finish office------------------------------------------
                MSFlexGrid1.TextMatrix(Baris, 8) = Left(rsExcel(6).Value, 10)
'mengambil jam finish office------------------------------------------
                MSFlexGrid1.TextMatrix(Baris, 9) = Right(rsExcel(6).Value, 8)
'mengambil tgl finish home+office--------------------------------------
                MSFlexGrid1.TextMatrix(Baris, 10) = Left(rsExcel(7).Value, 10)
'mengambil jam finish home+office--------------------------------------
                MSFlexGrid1.TextMatrix(Baris, 11) = Right(rsExcel(7).Value, 8)
'mengambil survey home-------------------------------------------------
                If IsNull(rsExcel(9).Value) Then
                MSFlexGrid1.TextMatrix(Baris, 12) = "0"
                Else
                MSFlexGrid1.TextMatrix(Baris, 12) = rsExcel(9).Value
                End If
'mengambil back up home------------------------------------------------
                If IsNull(rsExcel(10).Value) Then
                MSFlexGrid1.TextMatrix(Baris, 13) = "0"
                Else
                MSFlexGrid1.TextMatrix(Baris, 13) = rsExcel(10).Value
                End If
'mengambil survey office-----------------------------------------------
                If IsNull(rsExcel(12).Value) Then
                MSFlexGrid1.TextMatrix(Baris, 14) = "0"
                Else
                MSFlexGrid1.TextMatrix(Baris, 14) = rsExcel(12).Value
                End If
'mengambil back up office---------------------------------------------
                If IsNull(rsExcel(13).Value) Then
                MSFlexGrid1.TextMatrix(Baris, 15) = "0"
                Else
                MSFlexGrid1.TextMatrix(Baris, 15) = rsExcel(13).Value
                End If
'mengambil survey home+off---------------------------------------------
                If IsNull(rsExcel(15).Value) Then
                MSFlexGrid1.TextMatrix(Baris, 16) = "0"
                Else
                MSFlexGrid1.TextMatrix(Baris, 16) = rsExcel(15).Value
                End If
'mengambil back up home+off---------------------------------------------
                If IsNull(rsExcel(16).Value) Then
                MSFlexGrid1.TextMatrix(Baris, 17) = "0"
                Else
                MSFlexGrid1.TextMatrix(Baris, 17) = rsExcel(16).Value
                End If
'menentukan status home---------------------------------------------
                If IsNull(rsExcel(24).Value) Then
                MSFlexGrid1.TextMatrix(Baris, 18) = "NONE"
                Else
                MSFlexGrid1.TextMatrix(Baris, 18) = rsExcel(24).Value
                End If
                If MSFlexGrid1.TextMatrix(Baris, 12) = "0" And MSFlexGrid1.TextMatrix(Baris, 13) = "0" Then
                MSFlexGrid1.TextMatrix(Baris, 18) = "NONE"
                Else
'                MSFlexGrid1.TextMatrix(Baris, 18) = rsExcel(24).Value
                End If
'menentukan status office---------------------------------------------------

                If IsNull(rsExcel(25).Value) Then
                MSFlexGrid1.TextMatrix(Baris, 19) = "NONE"
                Else
                MSFlexGrid1.TextMatrix(Baris, 19) = rsExcel(25).Value
                End If
                If MSFlexGrid1.TextMatrix(Baris, 14) = "0" And MSFlexGrid1.TextMatrix(Baris, 15) = "0" Then
                MSFlexGrid1.TextMatrix(Baris, 19) = "NONE"
                Else
'                MSFlexGrid1.TextMatrix(Baris, 19) = rsExcel(25).Value
                End If
'menentukan status home=off------------------------------------------------

                If IsNull(rsExcel(26).Value) Then
                MSFlexGrid1.TextMatrix(Baris, 20) = "NONE"
                Else
                MSFlexGrid1.TextMatrix(Baris, 20) = rsExcel(26).Value
                End If
                If MSFlexGrid1.TextMatrix(Baris, 16) = "0" And MSFlexGrid1.TextMatrix(Baris, 17) = "0" Then
                MSFlexGrid1.TextMatrix(Baris, 20) = "NONE"
                Else
'                MSFlexGrid1.TextMatrix(Baris, 20) = rsExcel(26).Value
                End If
'------------------------------------------------------------------------

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
        tambahdata = "INSERT INTO tbl_dataadmin (incomingdate, jamincome, product, noaplikasi, namaaplikan, finishhm, jamfinishhm, finishoff, jamfinishoff, finishhmoff, jamfinishhmoff, surhome, backuphm, suroffice, backupoff, surhomeandoff, backuphmoff, statushome, statusoffice, statushomeandoff, user, tglinput, jaminput, lastupdate, jamlastup) " _
        & " VALUES ('" & MSFlexGrid1.TextMatrix(i, 1) & "','" & MSFlexGrid1.TextMatrix(i, 2) & "','" & MSFlexGrid1.TextMatrix(i, 3) & "','" & MSFlexGrid1.TextMatrix(i, 4) & "','" & MSFlexGrid1.TextMatrix(i, 5) & "','" & MSFlexGrid1.TextMatrix(i, 6) & "','" & MSFlexGrid1.TextMatrix(i, 7) & "','" & MSFlexGrid1.TextMatrix(i, 8) & "','" & MSFlexGrid1.TextMatrix(i, 9) & "','" & MSFlexGrid1.TextMatrix(i, 10) & "','" & MSFlexGrid1.TextMatrix(i, 11) & "','" & MSFlexGrid1.TextMatrix(i, 12) & "','" & MSFlexGrid1.TextMatrix(i, 13) & "','" & MSFlexGrid1.TextMatrix(i, 14) & "','" & MSFlexGrid1.TextMatrix(i, 15) & "','" & MSFlexGrid1.TextMatrix(i, 16) & "','" & MSFlexGrid1.TextMatrix(i, 17) & "','" & MSFlexGrid1.TextMatrix(i, 18) & "','" & MSFlexGrid1.TextMatrix(i, 19) & "','" & MSFlexGrid1.TextMatrix(i, 20) & "','" & lbuser.Caption & "','" & lbwaktu.Caption & "','" & lbjam.Caption & "','" & lbwaktu.Caption & "','" & lbjam.Caption & "') on duplicate key update incomingdate = '" & MSFlexGrid1.TextMatrix(i, 1) & "' , " _
        & " jamincome = '" & MSFlexGrid1.TextMatrix(i, 2) & "', product = '" & MSFlexGrid1.TextMatrix(i, 3) & "', noaplikasi = '" & MSFlexGrid1.TextMatrix(i, 4) & "', namaaplikan = '" & MSFlexGrid1.TextMatrix(i, 5) & "',finishhm = '" & MSFlexGrid1.TextMatrix(i, 6) & "',jamfinishhm = '" & MSFlexGrid1.TextMatrix(i, 7) & "',finishoff = '" & MSFlexGrid1.TextMatrix(i, 8) & "',jamfinishoff = '" & MSFlexGrid1.TextMatrix(i, 9) & "',finishhmoff = '" & MSFlexGrid1.TextMatrix(i, 10) & "',jamfinishhmoff = '" & MSFlexGrid1.TextMatrix(i, 11) & "',surhome = '" & MSFlexGrid1.TextMatrix(i, 12) & "',backuphm = '" & MSFlexGrid1.TextMatrix(i, 13) & "',suroffice = '" & MSFlexGrid1.TextMatrix(i, 14) & "',backupoff = '" & MSFlexGrid1.TextMatrix(i, 15) & "', surhomeandoff = '" & MSFlexGrid1.TextMatrix(i, 16) & "', backuphmoff = '" & MSFlexGrid1.TextMatrix(i, 17) & "', " _
        & " statushome = '" & MSFlexGrid1.TextMatrix(i, 18) & "', statusoffice = '" & MSFlexGrid1.TextMatrix(i, 19) & "', statushomeandoff = '" & MSFlexGrid1.TextMatrix(i, 20) & "', user = '" & lbuser.Caption & "', tglinput = '" & lbwaktu.Caption & "', jaminput = '" & lbjam.Caption & "', lastupdate = '" & lbwaktu.Caption & "', jamlastup = '" & lbjam.Caption & "'"
        koneksi.Execute tambahdata
        DoEvents
    Next i
    MsgBox "Import data berhasil, Silahkan di cek...", vbInformation, "... Sukses..."
    Exit Sub
AdaError:
If Err.Number = -2147467259 Then
    MsgBox "Nomor aplikasi " & MSFlexGrid1.TextMatrix(i, 4) & " sudah ada dalam database." & vbCrLf & _
    "Pada file excelnya di baris " & i + 1 & " ,silahkan hapus terlebih dahulu lalu ulangi.", vbCritical, ".:: Gagal...!!!"
    Exit Sub
Else
    MsgBox "Error No : " & Err.Number & vbCrLf & _
    Err.Description, vbCritical + vbOKOnly, "Error......"

End If

End Sub

