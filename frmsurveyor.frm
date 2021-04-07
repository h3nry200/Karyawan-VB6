VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmsurveyor 
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   Tag             =   "mdihalutama.mnkaryawan.Enabled = False"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "REPORT"
      Height          =   555
      Left            =   12720
      TabIndex        =   59
      Top             =   10200
      Width           =   885
   End
   Begin VB.Frame Frame1 
      Height          =   9855
      Left            =   10320
      TabIndex        =   11
      Top             =   240
      Width           =   9855
      Begin VB.TextBox tgllahirtxt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   67
         Top             =   5160
         Width           =   3045
      End
      Begin VB.TextBox ket2txt 
         Enabled         =   0   'False
         Height          =   855
         Left            =   1950
         TabIndex        =   65
         Top             =   8280
         Width           =   7845
      End
      Begin VB.TextBox ket1txt 
         Enabled         =   0   'False
         Height          =   855
         Left            =   1950
         TabIndex        =   63
         Top             =   7320
         Width           =   7845
      End
      Begin VB.TextBox lastuptxt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   6390
         TabIndex        =   47
         Top             =   120
         Width           =   2355
      End
      Begin VB.ComboBox cmbstatus 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmsurveyor.frx":0000
         Left            =   6990
         List            =   "frmsurveyor.frx":0002
         TabIndex        =   46
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox norektxt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   6990
         TabIndex        =   45
         Top             =   720
         Width           =   2685
      End
      Begin VB.TextBox skcktxt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1950
         TabIndex        =   44
         Top             =   6360
         Width           =   3045
      End
      Begin VB.TextBox mitratxt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1950
         TabIndex        =   43
         Top             =   6840
         Width           =   3045
      End
      Begin VB.TextBox rahasiatxt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   6750
         TabIndex        =   42
         Top             =   6360
         Width           =   3045
      End
      Begin VB.TextBox namaibutxt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   41
         Top             =   5640
         Width           =   3045
      End
      Begin VB.TextBox anaktxt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   6720
         TabIndex        =   40
         Top             =   4680
         Width           =   3045
      End
      Begin VB.TextBox tlptxt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   6720
         TabIndex        =   39
         Top             =   5160
         Width           =   3045
      End
      Begin VB.TextBox hptxt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   6750
         TabIndex        =   38
         Top             =   5640
         Width           =   3045
      End
      Begin VB.TextBox jamlasttxt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   8760
         TabIndex        =   37
         Top             =   120
         Width           =   1035
      End
      Begin VB.TextBox areatxt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   6990
         TabIndex        =   36
         Top             =   1080
         Width           =   2685
      End
      Begin VB.TextBox alasankeltxt 
         Enabled         =   0   'False
         Height          =   735
         Left            =   1950
         TabIndex        =   31
         Top             =   3240
         Width           =   7755
      End
      Begin VB.TextBox ktptxt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   30
         Top             =   4680
         Width           =   3045
      End
      Begin VB.TextBox addtxt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   29
         Top             =   4200
         Width           =   7845
      End
      Begin VB.TextBox kodetxt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   21
         Top             =   1800
         Width           =   3045
      End
      Begin VB.TextBox namasvrtxt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   20
         Top             =   1440
         Width           =   3045
      End
      Begin VB.TextBox vendortxt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   19
         Top             =   1080
         Width           =   3045
      End
      Begin VB.TextBox masatxt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   6720
         TabIndex        =   18
         Top             =   2280
         Width           =   3045
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Selisih"
         Height          =   375
         Left            =   4200
         TabIndex        =   17
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox idtxt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   15
         Top             =   720
         Width           =   2445
      End
      Begin VB.TextBox tglcretxt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Top             =   120
         Width           =   2265
      End
      Begin VB.TextBox jamcrttxt 
         Enabled         =   0   'False
         Height          =   375
         Left            =   4080
         TabIndex        =   12
         Top             =   120
         Width           =   945
      End
      Begin MSComCtl2.DTPicker DTjoin 
         Height          =   375
         Left            =   1920
         TabIndex        =   22
         Top             =   2280
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd/M/yyyy"
         Format          =   58589187
         CurrentDate     =   43171.3538078704
      End
      Begin MSComCtl2.DTPicker DTresign 
         Height          =   375
         Left            =   6750
         TabIndex        =   28
         Top             =   2760
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/M/yyyy"
         Format          =   58589185
         CurrentDate     =   43195.3637268519
      End
      Begin MSComCtl2.DTPicker DTkontrak 
         Height          =   375
         Left            =   1920
         TabIndex        =   60
         Top             =   2760
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd/M/yyyy"
         Format          =   58589187
         CurrentDate     =   43171.3538078704
      End
      Begin VB.Label Label26 
         Caption         =   "KETERANGAN 2  :"
         Height          =   225
         Left            =   480
         TabIndex        =   66
         Top             =   8400
         Width           =   1395
      End
      Begin VB.Label Label25 
         Caption         =   "KETERANGAN 1  :"
         Height          =   225
         Left            =   480
         TabIndex        =   64
         Top             =   7440
         Width           =   1395
      End
      Begin VB.Line Line2 
         X1              =   0
         X2              =   9840
         Y1              =   6120
         Y2              =   6120
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   9840
         Y1              =   4080
         Y2              =   4080
      End
      Begin VB.Label Label24 
         Caption         =   "TANGGAL LAHIR       :"
         Height          =   285
         Left            =   240
         TabIndex        =   62
         Top             =   5160
         Width           =   1665
      End
      Begin VB.Label Label22 
         Caption         =   "TANGGAL CONTRACT :"
         Height          =   285
         Left            =   120
         TabIndex        =   61
         Top             =   2760
         Width           =   1785
      End
      Begin VB.Label Label23 
         Caption         =   "LAST UPDATE  :"
         Height          =   315
         Left            =   5070
         TabIndex        =   58
         Top             =   180
         Width           =   1635
      End
      Begin VB.Label Label18 
         Caption         =   "STATUS KARYAWAN  :"
         Height          =   285
         Left            =   5160
         TabIndex        =   57
         Top             =   1560
         Width           =   1905
      End
      Begin VB.Label Label4 
         Caption         =   "NO REKENING     :"
         Height          =   225
         Left            =   5520
         TabIndex        =   56
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label Label6 
         Caption         =   "SKCK         :"
         Height          =   225
         Left            =   960
         TabIndex        =   55
         Top             =   6480
         Width           =   1395
      End
      Begin VB.Label Label7 
         Caption         =   "KEMITRAAN      :"
         Height          =   225
         Left            =   600
         TabIndex        =   54
         Top             =   6960
         Width           =   1395
      End
      Begin VB.Label Label12 
         Caption         =   "KERAHASIAAN    :"
         Height          =   225
         Left            =   5310
         TabIndex        =   53
         Top             =   6480
         Width           =   1395
      End
      Begin VB.Label Label13 
         Caption         =   "NAMA IBU         :"
         Height          =   225
         Left            =   600
         TabIndex        =   52
         Top             =   5640
         Width           =   1395
      End
      Begin VB.Label Label14 
         Caption         =   "JUMLAH ANAK     :"
         Height          =   225
         Left            =   5190
         TabIndex        =   51
         Top             =   4800
         Width           =   1395
      End
      Begin VB.Label Label15 
         Caption         =   "TLP RUMAH        :"
         Height          =   225
         Left            =   5280
         TabIndex        =   50
         Top             =   5280
         Width           =   1395
      End
      Begin VB.Label Label16 
         Caption         =   "HP               :"
         Height          =   225
         Left            =   5640
         TabIndex        =   49
         Top             =   5760
         Width           =   1395
      End
      Begin VB.Label Label17 
         Caption         =   "AREA SURVEY       :"
         Height          =   225
         Left            =   5400
         TabIndex        =   48
         Top             =   1200
         Width           =   1515
      End
      Begin VB.Label Label19 
         Caption         =   "ALASAN KELUAR  :"
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   3360
         Width           =   1605
      End
      Begin VB.Label Label11 
         Caption         =   "TANGGAL RESIGN :"
         Height          =   285
         Left            =   5160
         TabIndex        =   34
         Top             =   2760
         Width           =   1665
      End
      Begin VB.Label Label5 
         Caption         =   "ID  NUMBER        :"
         Height          =   225
         Left            =   480
         TabIndex        =   33
         Top             =   4800
         Width           =   1395
      End
      Begin VB.Label Label9 
         Caption         =   "ALAMAT         :"
         Height          =   225
         Left            =   720
         TabIndex        =   32
         Top             =   4320
         Width           =   1395
      End
      Begin VB.Label Label3 
         Caption         =   "KODE SURV        :"
         Height          =   225
         Left            =   480
         TabIndex        =   27
         Top             =   1920
         Width           =   1425
      End
      Begin VB.Label Label8 
         Caption         =   "NAMA SURV          :"
         Height          =   225
         Left            =   360
         TabIndex        =   26
         Top             =   1560
         Width           =   1515
      End
      Begin VB.Label Label10 
         Caption         =   "VENDOR NAME :"
         Height          =   225
         Left            =   480
         TabIndex        =   25
         Top             =   1200
         Width           =   1305
      End
      Begin VB.Label Label20 
         Caption         =   "TANGGAL JOIN         :"
         Height          =   285
         Left            =   240
         TabIndex        =   24
         Top             =   2280
         Width           =   1665
      End
      Begin VB.Label Label21 
         Caption         =   "MASA KERJA          :"
         Height          =   285
         Left            =   5160
         TabIndex        =   23
         Top             =   2400
         Width           =   1665
      End
      Begin VB.Label Label1 
         Caption         =   "ID      :"
         Height          =   225
         Left            =   1320
         TabIndex        =   16
         Top             =   840
         Width           =   585
      End
      Begin VB.Label Label2 
         Caption         =   "TANGGAL CREATE  :"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1755
      End
   End
   Begin VB.Timer Timer1 
      Left            =   5040
      Top             =   120
   End
   Begin VB.CommandButton cancelbtn 
      Caption         =   "CANCEL"
      Height          =   555
      Left            =   9540
      TabIndex        =   6
      Top             =   10200
      Width           =   885
   End
   Begin VB.CommandButton exitbtn 
      Caption         =   "EXIT"
      Height          =   555
      Left            =   11700
      TabIndex        =   5
      Top             =   10200
      Width           =   885
   End
   Begin VB.CommandButton savebtn 
      Caption         =   "SAVE"
      Height          =   555
      Left            =   8460
      TabIndex        =   4
      Top             =   10200
      Width           =   885
   End
   Begin VB.CommandButton deletebtn 
      Caption         =   "DELETE"
      Height          =   555
      Left            =   7380
      TabIndex        =   3
      Top             =   10200
      Width           =   885
   End
   Begin VB.CommandButton editbtn 
      Caption         =   "EDIT"
      Height          =   555
      Left            =   6180
      TabIndex        =   2
      Top             =   10200
      Width           =   975
   End
   Begin VB.CommandButton newbtn 
      Caption         =   "NEW"
      Height          =   555
      Left            =   4980
      TabIndex        =   1
      Top             =   10200
      Width           =   1005
   End
   Begin VB.CommandButton seekbtn 
      Caption         =   "SEEK"
      Height          =   555
      Left            =   10620
      TabIndex        =   0
      Top             =   10200
      Width           =   885
   End
   Begin MSDataGridLib.DataGrid DGkary 
      Height          =   9735
      Left            =   150
      TabIndex        =   7
      Top             =   330
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   17171
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   825
      Left            =   2880
      Top             =   1800
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
   Begin VB.Label lbjam 
      Height          =   375
      Left            =   18360
      TabIndex        =   10
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label lbuser 
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   0
      Width           =   5895
   End
   Begin VB.Label lbwaktu 
      Height          =   375
      Left            =   16200
      TabIndex        =   8
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmsurveyor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim koneksi As New rdoConnection
Dim rQuery As New rdoQuery
Dim rs As rdoResultset

Dim shari As String
Dim ahari

Private Sub Command1_Click()
Dim Tahun As Integer, Sisa As Integer
Dim SelisihBulan As Integer
On Error GoTo Pesan
lbwaktu.Caption = Format(Date, "d/m/yyyy")
'DTjoin.Value = Format(Date, "d/m/yyyy")

SelisihBulan = DateDiff("m", DTjoin.Value, lbwaktu.Caption)
Tahun = SelisihBulan \ 12
Sisa = SelisihBulan Mod 12
masatxt.Text = "" & Tahun & " Tahun " & Sisa & " Bulan"
Exit Sub
Pesan:
MsgBox "Tipe tanggal salah!", vbCritical, "Error Tanggal"
'lbwaktu.Caption = Format(Date, "d/m/yyyy")
'DTjoin.Value = Format(Date, "d/m/yyyy")

'masatxt.Text = lbwaktu.Caption - jointxt.Text


End Sub

Private Sub Command2_Click()
Dim appexcel As Excel.Application
Dim excelWBk As Excel.Workbook
Dim ExcelWS As Excel.Worksheet
Dim dbCon As New ADODB.Connection

dbCon.Open "Driver={MySQL ODBC 5.1 Driver};SERVER=192.168.2.11;PWD=123456;UID=BRIWIRA;PORT=3306;DATABASE=karyawan;"
On Error Resume Next
        
    Set appexcel = New Excel.Application
    Set excelWBk = appexcel.Workbooks.Add
    
    Dim dbRec As New ADODB.Recordset
    Dim colfield As Collection
    Dim jmltabel As Integer, namatable As String
    Dim brother As Boolean
    
    Me.Enabled = False
        If dbRec.State = 1 Then dbRec.Close
        dbRec.Open "select id,vendor,nama_svr,kode_srv,no_rek,area,status_kary,tgl_join,masakerja,tgl_kontrak,addr,id_ktp,tgl_lahir,nama_ibu,jml_anak,tlp_rmh,hp,skck,tgl_mitra,tgl_janji,tgl_resign,alasan_kel,ket1,ket2 from tbl_surveyor", dbCon, adOpenDynamic, adLockOptimistic
        
        Set ExcelWS = excelWBk.Worksheets.Add
        
        Dim jmlfield As Integer
        Set colfield = New Collection
        
        For jmlfield = 0 To dbRec.Fields.Count - 1
            ExcelWS.Cells(1, jmlfield + 1) = dbRec.Fields(jmlfield).Name
            colfield.Add dbRec.Fields(jmlfield).Name
            DoEvents
        Next jmlfield
        
        Dim Pos As Long
        Pos = 2
        If Not dbRec.EOF Then
            dbRec.MoveFirst
            While Not dbRec.EOF
                For jmlfield = 1 To colfield.Count
                ExcelWS.Cells(Pos, jmlfield) = dbRec(colfield(jmlfield))
                Next jmlfield
            Pos = Pos + 1
            dbRec.MoveNext
            DoEvents
            Wend
        End If
        dbRec.Close
        
        If Err <> 0 Then
        brother = True
        Err.Clear
        End If
        
    excelWBk.SaveAs "C:\ReportHRD\D-base srv Update" & Format(Date, "YYMMDD") & Format(Time, "HHMMSS") & ".xlsx"
    excelWBk.Saved = True
    
    Me.Enabled = True
    If MsgBox("Data tersimpan di C:\ReportHRD." & IIf(brother, " Tapi ada sebagian data yg ga bisa di export", "") & ". Mau melihat hasilnya sekarang?", 32 + vbYesNo) = vbYes Then
        appexcel.Visible = True
    Else
        appexcel.Quit
    End If
    
End Sub

Private Sub Form_Load()

Call Bersih
    
    koneksi.Connect = "Driver={MySQL ODBC 5.1 Driver};SERVER=192.168.2.11;PWD=123456;UID=BRIWIRA;PORT=3306;DATABASE=karyawan;"
    koneksi.EstablishConnection
    Adodc1.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};SERVER=192.168.2.11;PWD=123456;UID=BRIWIRA;PORT=3306;DATABASE=karyawan;"
    Adodc1.RecordSource = "tbl_surveyor"
    Adodc1.Refresh
Set DGkary.DataSource = Adodc1
cancelbtn.Enabled = False

  ahari = Array("Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu")
  Timer1.Interval = 500
  Timer1.Enabled = True
'lbuser.Caption = "' & usertxt.Text & '"
lbuser.Caption = mdihalutama.Caption

    cmbstatus.AddItem "Kemitraan"
    cmbstatus.AddItem "Karyawan Tetap"
    cmbstatus.AddItem "AKTIF"
    cmbstatus.AddItem "NON AKTIF"

End Sub

Private Sub seekbtn_Click()
frmseeksurv.Show
frmseeksurv.Text1.SetFocus
End Sub

Private Sub Timer1_Timer()
 shari = ahari(Abs(Weekday(Date) - 1))
  lbwaktu.Caption = "" & shari & ", " _
                   & Format(Date, "dd mmmm yyyy")
lbjam.Caption = Format(Time, "hh:mm:ss")
End Sub

Sub Bersih()
idtxt.Text = ""
tglcretxt.Text = ""
jamcrttxt.Text = ""
lastuptxt.Text = ""
jamlasttxt.Text = ""
idtxt.Text = ""
vendortxt.Text = ""
namasvrtxt.Text = ""
kodetxt.Text = ""
norektxt.Text = ""
areatxt.Text = ""
cmbstatus.Text = ""
DTjoin.Value = "01/01/0001"
masatxt.Text = ""
DTkontrak.Value = "01/01/0001"
DTresign.Value = "01/01/0001"
alasankeltxt.Text = ""
addtxt.Text = ""
ktptxt.Text = ""
tgllahirtxt.Text = ""
namaibutxt.Text = ""
anaktxt.Text = ""
tlptxt.Text = ""
hptxt.Text = ""
skcktxt.Text = ""
mitratxt.Text = ""
rahasiatxt.Text = ""
ket1txt.Text = ""
ket2txt.Text = ""

End Sub

Sub KondisiAwal()
idtxt.Text = ""
tglcretxt.Text = ""
jamcrttxt.Text = ""
lastuptxt.Text = ""
jamlasttxt.Text = ""
idtxt.Text = ""
vendortxt.Text = ""
namasvrtxt.Text = ""
kodetxt.Text = ""
norektxt.Text = ""
areatxt.Text = ""
cmbstatus.Text = ""
DTjoin.Value = "01/01/0001"
masatxt.Text = ""
DTkontrak.Value = "01/01/0001"
DTresign.Value = "01/01/0001"
alasankeltxt.Text = ""
addtxt.Text = ""
ktptxt.Text = ""
tgllahirtxt.Text = ""
namaibutxt.Text = ""
anaktxt.Text = ""
tlptxt.Text = ""
hptxt.Text = ""
skcktxt.Text = ""
mitratxt.Text = ""
rahasiatxt.Text = ""
ket1txt.Text = ""
ket2txt.Text = ""

idtxt.Enabled = False
tglcretxt.Enabled = False
jamcrttxt.Enabled = False
lastuptxt.Enabled = False
jamlasttxt.Enabled = False
idtxt.Enabled = False
vendortxt.Enabled = False
namasvrtxt.Enabled = False
kodetxt.Enabled = False
norektxt.Enabled = False
areatxt.Enabled = False
cmbstatus.Enabled = False
DTjoin.Enabled = False
masatxt.Enabled = False
DTkontrak.Enabled = False
DTresign.Enabled = False
alasankeltxt.Enabled = False
addtxt.Enabled = False
ktptxt.Enabled = False
tgllahirtxt.Enabled = False
namaibutxt.Enabled = False
anaktxt.Enabled = False
tlptxt.Enabled = False
hptxt.Enabled = False
skcktxt.Enabled = False
mitratxt.Enabled = False
rahasiatxt.Enabled = False
ket1txt.Enabled = False
ket2txt.Enabled = False


newbtn.Enabled = True
editbtn.Enabled = True
deletebtn.Enabled = True
savebtn.Enabled = False
cancelbtn.Enabled = False
seekbtn.Enabled = True
exitbtn.Enabled = True
Command2.Enabled = True

DGkary.Refresh
End Sub

Private Sub DGkary_Click()
idtxt.Text = DGkary.Columns(0).Text
tglcretxt.Text = DGkary.Columns(1).Text
jamcrttxt.Text = DGkary.Columns(2).Text
lastuptxt.Text = DGkary.Columns(3).Text
jamlasttxt.Text = DGkary.Columns(4).Text
vendortxt.Text = DGkary.Columns(5).Text
namasvrtxt.Text = DGkary.Columns(6).Text
kodetxt.Text = DGkary.Columns(7).Text
norektxt.Text = DGkary.Columns(8).Text
areatxt.Text = DGkary.Columns(9).Text
cmbstatus.Text = DGkary.Columns(10).Text
masatxt.Text = DGkary.Columns(12).Text
alasankeltxt.Text = DGkary.Columns(25).Text
addtxt.Text = DGkary.Columns(14).Text
ktptxt.Text = DGkary.Columns(15).Text
tgllahirtxt.Text = DGkary.Columns(16).Text
namaibutxt.Text = DGkary.Columns(17).Text
anaktxt.Text = DGkary.Columns(18).Text
tlptxt.Text = DGkary.Columns(19).Text
hptxt.Text = DGkary.Columns(20).Text
skcktxt.Text = DGkary.Columns(21).Text
mitratxt.Text = DGkary.Columns(22).Text
rahasiatxt.Text = DGkary.Columns(23).Text
ket1txt.Text = DGkary.Columns(26).Text
ket2txt.Text = DGkary.Columns(27).Text

If DGkary.Columns(11).Text = "0" Then
DTjoin.Value = "1/1/0001"
Else
DTjoin.Value = DGkary.Columns(11).Text
End If

If DGkary.Columns(13).Text = "0" Then
DTkontrak.Value = "1/1/0001"
Else
DTkontrak.Value = DGkary.Columns(13).Text
End If

If DGkary.Columns(24).Text = "0" Then
DTresign.Value = "1/1/0001"
Else
DTresign.Value = DGkary.Columns(24).Text
End If


End Sub

Private Sub DGkary_change()
idtxt.Text = DGkary.Columns(0).Text
tglcretxt.Text = DGkary.Columns(1).Text
jamcrttxt.Text = DGkary.Columns(2).Text
lastuptxt.Text = DGkary.Columns(3).Text
jamlasttxt.Text = DGkary.Columns(4).Text
vendortxt.Text = DGkary.Columns(5).Text
namasvrtxt.Text = DGkary.Columns(6).Text
kodetxt.Text = DGkary.Columns(7).Text
norektxt.Text = DGkary.Columns(8).Text
areatxt.Text = DGkary.Columns(9).Text
cmbstatus.Text = DGkary.Columns(10).Text
masatxt.Text = DGkary.Columns(12).Text
alasankeltxt.Text = DGkary.Columns(25).Text
addtxt.Text = DGkary.Columns(14).Text
ktptxt.Text = DGkary.Columns(15).Text
tgllahirtxt.Text = DGkary.Columns(16).Text
namaibutxt.Text = DGkary.Columns(17).Text
anaktxt.Text = DGkary.Columns(18).Text
tlptxt.Text = DGkary.Columns(19).Text
hptxt.Text = DGkary.Columns(20).Text
skcktxt.Text = DGkary.Columns(21).Text
mitratxt.Text = DGkary.Columns(22).Text
rahasiatxt.Text = DGkary.Columns(23).Text
ket1txt.Text = DGkary.Columns(26).Text
ket2txt.Text = DGkary.Columns(27).Text

If DGkary.Columns(11).Text = "0" Then
DTjoin.Value = "1/1/0001"
Else
DTjoin.Value = DGkary.Columns(11).Text
End If

If DGkary.Columns(13).Text = "0" Then
DTkontrak.Value = "1/1/0001"
Else
DTkontrak.Value = DGkary.Columns(13).Text
End If

If DGkary.Columns(24).Text = "0" Then
DTresign.Value = "1/1/0001"
Else
DTresign.Value = DGkary.Columns(24).Text
End If


End Sub

Private Sub DGkary_KeyDown(KeyCode As Integer, Shift As Integer)
idtxt.Text = DGkary.Columns(0).Text
tglcretxt.Text = DGkary.Columns(1).Text
jamcrttxt.Text = DGkary.Columns(2).Text
lastuptxt.Text = DGkary.Columns(3).Text
jamlasttxt.Text = DGkary.Columns(4).Text
vendortxt.Text = DGkary.Columns(5).Text
namasvrtxt.Text = DGkary.Columns(6).Text
kodetxt.Text = DGkary.Columns(7).Text
norektxt.Text = DGkary.Columns(8).Text
areatxt.Text = DGkary.Columns(9).Text
cmbstatus.Text = DGkary.Columns(10).Text
masatxt.Text = DGkary.Columns(12).Text
alasankeltxt.Text = DGkary.Columns(25).Text
addtxt.Text = DGkary.Columns(14).Text
ktptxt.Text = DGkary.Columns(15).Text
tgllahirtxt.Text = DGkary.Columns(16).Text
namaibutxt.Text = DGkary.Columns(17).Text
anaktxt.Text = DGkary.Columns(18).Text
tlptxt.Text = DGkary.Columns(19).Text
hptxt.Text = DGkary.Columns(20).Text
skcktxt.Text = DGkary.Columns(21).Text
mitratxt.Text = DGkary.Columns(22).Text
rahasiatxt.Text = DGkary.Columns(23).Text
ket1txt.Text = DGkary.Columns(26).Text
ket2txt.Text = DGkary.Columns(27).Text

If DGkary.Columns(11).Text = "0" Then
DTjoin.Value = "1/1/0001"
Else
DTjoin.Value = DGkary.Columns(11).Text
End If

If DGkary.Columns(13).Text = "0" Then
DTkontrak.Value = "1/1/0001"
Else
DTkontrak.Value = DGkary.Columns(13).Text
End If

If DGkary.Columns(24).Text = "0" Then
DTresign.Value = "1/1/0001"
Else
DTresign.Value = DGkary.Columns(24).Text
End If

End Sub

Private Sub newbtn_Click()
Call KondisiAwal
editbtn.Enabled = False
deletebtn.Enabled = False
savebtn.Enabled = True
newbtn.Enabled = False
cancelbtn.Enabled = True
Command2.Enabled = False

idtxt.Enabled = True
tglcretxt.Enabled = True
jamcrttxt.Enabled = True
lastuptxt.Enabled = True
jamlasttxt.Enabled = True
idtxt.Enabled = True
vendortxt.Enabled = True
namasvrtxt.Enabled = True
kodetxt.Enabled = True
norektxt.Enabled = True
areatxt.Enabled = True
cmbstatus.Enabled = True
DTjoin.Enabled = True
masatxt.Enabled = True
DTkontrak.Enabled = True
DTresign.Enabled = True
alasankeltxt.Enabled = True
addtxt.Enabled = True
ktptxt.Enabled = True
tgllahirtxt.Enabled = True
namaibutxt.Enabled = True
anaktxt.Enabled = True
tlptxt.Enabled = True
hptxt.Enabled = True
skcktxt.Enabled = True
mitratxt.Enabled = True
rahasiatxt.Enabled = True
ket1txt.Enabled = True
ket2txt.Enabled = True

tglcretxt.Text = lbwaktu.Caption
jamcrttxt.Text = lbjam.Caption
lastuptxt.Text = lbwaktu.Caption
jamlasttxt.Text = lbjam.Caption

idtxt.SetFocus
newbtn.Caption = "NEWDATA"
End Sub

Private Sub editbtn_Click()
editbtn.Enabled = False
deletebtn.Enabled = False
savebtn.Enabled = True
newbtn.Enabled = False
cancelbtn.Enabled = True
Command2.Enabled = False

idtxt.Enabled = True
tglcretxt.Enabled = True
jamcrttxt.Enabled = True
lastuptxt.Enabled = True
jamlasttxt.Enabled = True
idtxt.Enabled = True
vendortxt.Enabled = True
namasvrtxt.Enabled = True
kodetxt.Enabled = True
norektxt.Enabled = True
areatxt.Enabled = True
cmbstatus.Enabled = True
DTjoin.Enabled = True
masatxt.Enabled = True
DTkontrak.Enabled = True
DTresign.Enabled = True
alasankeltxt.Enabled = True
addtxt.Enabled = True
ktptxt.Enabled = True
tgllahirtxt.Enabled = True
namaibutxt.Enabled = True
anaktxt.Enabled = True
tlptxt.Enabled = True
hptxt.Enabled = True
skcktxt.Enabled = True
mitratxt.Enabled = True
rahasiatxt.Enabled = True
ket1txt.Enabled = True
ket2txt.Enabled = True

lastuptxt.Text = lbwaktu.Caption
jamlasttxt.Text = lbjam.Caption

idtxt.SetFocus
editbtn.Caption = "EDITDATA"

End Sub

Private Sub savebtn_Click()
If newbtn.Caption = "NEWDATA" Then
If idtxt.Text = "" Or kodetxt.Text = "" Or namasvrtxt.Text = "" Then
    MsgBox "Data Belum Lengkap"
    Else
    Dim tambahdata As String
        tambahdata = "Insert Into tbl_surveyor values ('" & idtxt.Text & "', '" & tglcretxt.Text & "','" & jamcrttxt.Text & "','" & lastuptxt.Text & "','" & jamlasttxt.Text & "','" & vendortxt.Text & "','" & namasvrtxt.Text & "','" & kodetxt.Text & "','" & norektxt.Text & "','" & areatxt.Text & "','" & cmbstatus.Text & "','" & DTjoin.Value & "','" & masatxt.Text & "','" & DTkontrak.Value & "','" & addtxt.Text & "','" & ktptxt.Text & "','" & tgllahirtxt.Text & "','" & namaibutxt.Text & "','" & anaktxt.Text & "','" & tlptxt.Text & "','" & hptxt.Text & "','" & skcktxt.Text & "','" & mitratxt.Text & "','" & rahasiatxt.Text & "','" & DTresign.Value & "','" & alasankeltxt.Text & "','" & ket1txt.Text & "','" & ket2txt.Text & "','" & lbuser.Caption & "')"
        koneksi.Execute tambahdata
        MsgBox "Data Berhasil Ditambah", vbInformation, "Pemberitahuan"
        newbtn.Caption = "NEW"
        Adodc1.Refresh
        DGkary.Refresh
        Call KondisiAwal
    
    End If
Else
If editbtn.Caption = "EDITDATA" Then
    Dim editdata As String
        editdata = "update tbl_surveyor set id = '" & idtxt.Text & "',  last_update = '" & lastuptxt.Text & "', jam_lastupt = '" & jamlasttxt.Text & "', vendor = '" & vendortxt.Text & "', nama_svr = '" & namasvrtxt.Text & "', kode_srv = '" & kodetxt.Text & "', no_rek = '" & norektxt.Text & "', area = '" & areatxt.Text & "', status_kary = '" & cmbstatus.Text & "', tgl_join = '" & DTjoin.Value & "', masakerja = '" & masatxt.Text & "', tgl_kontrak = '" & DTkontrak.Value & "', addr = '" & addtxt.Text & "', id_ktp = '" & ktptxt.Text & "', tgl_lahir = '" & tgllahirtxt.Text & "', nama_ibu = '" & namaibutxt.Text & "', jml_anak = '" & anaktxt.Text & "', tlp_rmh = '" & tlptxt.Text & "',hp = '" & hptxt.Text & "',skck = '" & skcktxt.Text & "',tgl_mitra = '" & mitratxt.Text & "',tgl_janji = '" & rahasiatxt.Text & "',tgl_resign = '" & DTresign.Value & "', " _
        & " alasan_kel = '" & alasankeltxt.Text & "', ket1 = '" & ket1txt.Text & "',ket2 = '" & ket2txt.Text & "', user = '" & lbuser.Caption & "' where id = '" & idtxt.Text & "'"
        koneksi.Execute editdata
        MsgBox "Data Berhasil Diedit", vbInformation, "Pemberitahuan"
        editbtn.Caption = "EDIT"
        Adodc1.Refresh
        DGkary.Refresh
 Call KondisiAwal
End If
End If
End Sub

Private Sub cancelbtn_Click()
 Call Bersih
 Call KondisiAwal
 newbtn.Enabled = True
 editbtn.Enabled = True
 deletebtn.Enabled = True
 Command2.Enabled = True

 newbtn.Caption = "NEW"
 editbtn.Caption = "EDIT"
End Sub

Private Sub deletebtn_Click()
If MsgBox("Yakin Ingin Menghapus Data?", vbQuestion + vbOKCancel, "konfirmasi") = vbOK Then
Dim hapusdata As String
        hapusdata = "delete from tbl_surveyor where id =" & idtxt.Text & ""
        koneksi.Execute hapusdata
        MsgBox "Data Berhasil Dihapus", vbInformation, "Pemberitahuan"
    Adodc1.Refresh
DGkary.Refresh
Call Bersih

Else
Call Bersih
End If

End Sub

Private Sub exitbtn_Click()
koneksi.Close
Unload Me
End Sub

