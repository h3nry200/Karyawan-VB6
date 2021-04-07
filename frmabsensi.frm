VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmabsensi 
   Caption         =   "ABSENSI KARYAWAN"
   ClientHeight    =   10305
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10305
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.ComboBox bln2cmb 
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
      ItemData        =   "frmabsensi.frx":0000
      Left            =   120
      List            =   "frmabsensi.frx":0002
      TabIndex        =   29
      Top             =   240
      Width           =   1935
   End
   Begin VB.ComboBox thn2cmb 
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
      Left            =   2280
      TabIndex        =   28
      Top             =   240
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   9000
      TabIndex        =   27
      Top             =   8280
      Width           =   3735
   End
   Begin VB.ComboBox Combo1 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmabsensi.frx":0004
      Left            =   2040
      List            =   "frmabsensi.frx":0006
      TabIndex        =   25
      Top             =   8280
      Width           =   3975
   End
   Begin VB.CommandButton exitbtn 
      Caption         =   "EXIT"
      Height          =   375
      Left            =   13200
      TabIndex        =   23
      Top             =   9240
      Width           =   1095
   End
   Begin VB.CommandButton deletebtn 
      Caption         =   "DELETE"
      Height          =   375
      Left            =   14760
      TabIndex        =   22
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cancelbtn 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   14760
      TabIndex        =   21
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton savebtn 
      Caption         =   "SAVE"
      Height          =   375
      Left            =   14760
      TabIndex        =   20
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton editbtn 
      Caption         =   "EDIT"
      Height          =   375
      Left            =   13200
      TabIndex        =   19
      Top             =   8520
      Width           =   1095
   End
   Begin VB.CommandButton newbtn 
      Caption         =   "NEW"
      Height          =   375
      Left            =   13200
      TabIndex        =   18
      Top             =   7800
      Width           =   1095
   End
   Begin VB.CommandButton seekbtn 
      Caption         =   "SEEK"
      Height          =   375
      Left            =   13200
      TabIndex        =   17
      Top             =   7080
      Width           =   1095
   End
   Begin VB.TextBox atttxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   9000
      TabIndex        =   16
      Top             =   7680
      Width           =   3735
   End
   Begin VB.TextBox jptxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   9000
      TabIndex        =   14
      Top             =   9720
      Width           =   3735
   End
   Begin VB.TextBox jmtxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   9000
      TabIndex        =   13
      Top             =   9120
      Width           =   3735
   End
   Begin VB.TextBox jpktxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   9720
      Width           =   3975
   End
   Begin VB.TextBox jmktxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   9120
      Width           =   3975
   End
   Begin VB.TextBox namatxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   7680
      Width           =   3975
   End
   Begin VB.TextBox tgltxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   9000
      TabIndex        =   7
      Top             =   7080
      Width           =   3735
   End
   Begin VB.TextBox idtxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   7080
      Width           =   3975
   End
   Begin MSDataGridLib.DataGrid DGkary 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   20055
      _ExtentX        =   35375
      _ExtentY        =   10610
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   29
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
         Size            =   13.5
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
      Left            =   960
      Top             =   1440
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
   Begin VB.Label Label10 
      Caption         =   "Status Lembur                :"
      Height          =   375
      Left            =   6960
      TabIndex        =   26
      Top             =   8280
      Width           =   1815
   End
   Begin VB.Label Label9 
      Caption         =   "Status Kehadiran :"
      Height          =   375
      Left            =   360
      TabIndex        =   24
      Top             =   8280
      Width           =   1335
   End
   Begin VB.Line Line5 
      X1              =   120
      X2              =   120
      Y1              =   6720
      Y2              =   10200
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   12960
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line3 
      X1              =   12960
      X2              =   12960
      Y1              =   6720
      Y2              =   10200
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   12960
      Y1              =   10200
      Y2              =   10200
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   12960
      Y1              =   8760
      Y2              =   8760
   End
   Begin VB.Label Label8 
      Caption         =   "Att Time                          :"
      Height          =   375
      Left            =   6960
      TabIndex        =   15
      Top             =   7680
      Width           =   1935
   End
   Begin VB.Label Label7 
      Caption         =   "Jam Masuk Karyawan   :"
      Height          =   375
      Left            =   6960
      TabIndex        =   12
      Top             =   9120
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Jam Pulang Karyawan     :"
      Height          =   375
      Left            =   6960
      TabIndex        =   11
      Top             =   9720
      Width           =   1935
   End
   Begin VB.Label label5 
      Caption         =   "Jam Pulang Kerja :"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   9720
      Width           =   1335
   End
   Begin VB.Label label4 
      Caption         =   "Jam Masuk Kerja :"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   9120
      Width           =   1335
   End
   Begin VB.Label label3 
      Caption         =   "Nama Karyawan  :"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label label2 
      Caption         =   "Tanggal Hadir                 :"
      Height          =   375
      Left            =   6960
      TabIndex        =   2
      Top             =   7080
      Width           =   1935
   End
   Begin VB.Label label1 
      Caption         =   "ID                         :"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   7080
      Width           =   1335
   End
End
Attribute VB_Name = "frmabsensi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim koneksi As New rdoConnection
Dim rQuery As New rdoQuery
Dim rs As rdoResultset

Private Sub bln2cmb_Click()
With Adodc1
.CommandType = adCmdText
.RecordSource = "select * from absensi where tanggal like '%" & bln2cmb.Text & "%' and tanggal like '%" & thn2cmb.Text & "%'"
.Refresh
End With
DGkary.Refresh
End Sub

Private Sub thn2cmb_Click()
With Adodc1
.CommandType = adCmdText
.RecordSource = "select * from absensi where tanggal like '%" & bln2cmb.Text & "%' and tanggal like '%" & thn2cmb.Text & "%'"
.Refresh
End With
DGkary.Refresh
End Sub

Private Sub cancelbtn_Click()
 Call Bersih
 newbtn.Enabled = True
 editbtn.Enabled = True
 deletebtn.Enabled = True
 newbtn.Caption = "NEW"
 editbtn.Caption = "EDIT"
seekbtn.Enabled = True
End Sub

Private Sub Command1_Click()
Dim caridata As String
        caridata = "select * from absensi where tanggal like '%" & bln2cmb & "%'"
        koneksi.Execute caridata
    Adodc1.Refresh
DGkary.Refresh

End Sub

Private Sub deletebtn_Click()
If MsgBox("Yakin Ingin Menghapus Data?", vbQuestion + vbOKCancel, "konfirmasi") = vbOK Then
Dim hapusdata As String
        hapusdata = "delete from absensi where id =" & idtxt.Text & ""
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
idtxt.Text = DGkary.Columns(0).Text
tgltxt.Text = DGkary.Columns(1).Text
namatxt.Text = DGkary.Columns(2).Text
Combo1.Text = DGkary.Columns(3).Text
Combo2.Text = DGkary.Columns(4).Text
jmktxt.Text = DGkary.Columns(5).Text
jmtxt.Text = DGkary.Columns(6).Text
jpktxt.Text = DGkary.Columns(7).Text
jptxt.Text = DGkary.Columns(8).Text
atttxt.Text = DGkary.Columns(9).Text

End Sub


Private Sub DGkary_DblClick()
idtxt.Text = DGkary.Columns(0).Text
tgltxt.Text = DGkary.Columns(1).Text
namatxt.Text = DGkary.Columns(2).Text
Combo1.Text = DGkary.Columns(3).Text
Combo2.Text = DGkary.Columns(4).Text
jmktxt.Text = DGkary.Columns(5).Text
jmtxt.Text = DGkary.Columns(6).Text
jpktxt.Text = DGkary.Columns(7).Text
jptxt.Text = DGkary.Columns(8).Text
atttxt.Text = DGkary.Columns(9).Text

End Sub

Private Sub DGkary_KeyDown(KeyCode As Integer, Shift As Integer)
idtxt.Text = DGkary.Columns(0).Text
tgltxt.Text = DGkary.Columns(1).Text
namatxt.Text = DGkary.Columns(2).Text
Combo1.Text = DGkary.Columns(3).Text
Combo2.Text = DGkary.Columns(4).Text
jmktxt.Text = DGkary.Columns(5).Text
jmtxt.Text = DGkary.Columns(6).Text
jpktxt.Text = DGkary.Columns(7).Text
jptxt.Text = DGkary.Columns(8).Text
atttxt.Text = DGkary.Columns(9).Text
End Sub

Private Sub exitbtn_Click()
koneksi.Close
Unload Me
End Sub

Private Sub Form_Load()
koneksi.Connect = "Driver={MySQL ODBC 5.1 Driver};SERVER=192.168.2.11;PWD=123456;UID=BRIWIRA;PORT=3306;DATABASE=karyawan;"
koneksi.EstablishConnection
    Adodc1.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};SERVER=192.168.2.11;PWD=123456;UID=BRIWIRA;PORT=3306;DATABASE=karyawan;"
    Adodc1.RecordSource = "absensi"
    Adodc1.Refresh
Set DGkary.DataSource = Adodc1
    Combo1.AddItem "MASUK"
    Combo1.AddItem "IJIN"
    Combo1.AddItem "SAKIT"
    Combo1.AddItem "CUTI"
    Combo1.AddItem "ALPHA"
    Combo2.AddItem "LEMBUR"
    Combo2.AddItem "TIDAK LEMBUR"
    bln2cmb.AddItem "JANUARY"
    bln2cmb.AddItem "FEBRUARY"
    bln2cmb.AddItem "MARCH"
    bln2cmb.AddItem "APRIL"
    bln2cmb.AddItem "MAY"
    bln2cmb.AddItem "JUNE"
    bln2cmb.AddItem "JULY"
    bln2cmb.AddItem "AUGUST"
    bln2cmb.AddItem "SEPTEMBER"
    bln2cmb.AddItem "OCTOBER"
    bln2cmb.AddItem "NOVEMBER"
    bln2cmb.AddItem "DECEMBER"
    thn2cmb.AddItem "2016"
    thn2cmb.AddItem "2017"
    thn2cmb.AddItem "2018"
    thn2cmb.AddItem "2019"
    thn2cmb.AddItem "2020"
    thn2cmb.AddItem "2021"
    thn2cmb.AddItem "2022"
    thn2cmb.AddItem "2023"
    thn2cmb.AddItem "2024"
    thn2cmb.AddItem "2025"
    thn2cmb.AddItem "2026"
    thn2cmb.AddItem "2027"
    thn2cmb.AddItem "2028"
    thn2cmb.AddItem "2029"
    thn2cmb.AddItem "2030"
    thn2cmb.AddItem "2031"
    thn2cmb.AddItem "2032"
    thn2cmb.AddItem "2033"
    thn2cmb.AddItem "2034"
    thn2cmb.AddItem "2035"
    thn2cmb.AddItem "2036"
    thn2cmb.AddItem "2037"
    thn2cmb.AddItem "2038"
    thn2cmb.AddItem "2039"
    thn2cmb.AddItem "2040"
bln2cmb.Text = "AUGUST"
thn2cmb.Text = "2016"
cancelbtn.Enabled = False

Call Bersih
End Sub

Sub Bersih()
idtxt.Text = ""
tgltxt.Text = ""
namatxt.Text = ""
Combo1.Text = ""
Combo2.Text = ""
jmktxt.Text = ""
jmtxt.Text = ""
jpktxt.Text = ""
jptxt.Text = ""
atttxt.Text = ""
End Sub

Private Sub newbtn_Click()
editbtn.Enabled = False
deletebtn.Enabled = False
savebtn.Enabled = True
newbtn.Enabled = False
seekbtn.Enabled = False
cancelbtn.Enabled = True
idtxt.Enabled = True
tgltxt.Enabled = True
namatxt.Enabled = True
jmktxt.Enabled = True
jmtxt.Enabled = True
jpktxt.Enabled = True
jptxt.Enabled = True
atttxt.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
idtxt.SetFocus
Call Bersih
newbtn.Caption = "NEWDATA"
End Sub

Private Sub editbtn_Click()
editbtn.Enabled = False
deletebtn.Enabled = False
savebtn.Enabled = True
newbtn.Enabled = False
cancelbtn.Enabled = True
seekbtn.Enabled = False
idtxt.Enabled = True
tgltxt.Enabled = True
namatxt.Enabled = True
jmktxt.Enabled = True
jmtxt.Enabled = True
jpktxt.Enabled = True
jptxt.Enabled = True
atttxt.Enabled = True
Combo1.Enabled = True
Combo2.Enabled = True
idtxt.SetFocus
editbtn.Caption = "EDITDATA"
End Sub

Private Sub savebtn_Click()
If newbtn.Caption = "NEWDATA" Then
If idtxt.Text = "" Or tgltxt.Text = "" Or namatxt.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Or jmktxt.Text = "" Or jmtxt.Text = "" Or jpktxt.Text = "" Or jptxt.Text = "" Or atttxt.Text = "" Then
    MsgBox "Data Belum Lengkap"
    Else
    Dim tambahdata As String
        tambahdata = "Insert Into absensi values ('" & idtxt.Text & "','" & tgltxt.Text & "','" & namatxt.Text & "','" & Combo1.Text & "','" & Combo2.Text & "','" & jmktxt.Text & "','" & jmtxt.Text & "','" & jpktxt.Text & "','" & jptxt.Text & "','" & atttxt.Text & "')"
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
        editdata = "update absensi set id= '" & idtxt.Text & "',tanggal = '" & tgltxt.Text & "',Nama = '" & namatxt.Text & "',statushadir = '" & Combo1.Text & "',statuslembur = '" & Combo2.Text & "',Jam_Masuk_Kerja = '" & jmktxt.Text & "',Jam_Masuk = '" & jmtxt.Text & "',Jam_Pulang_Kerja = '" & jpktxt.Text & "',Jam_Pulang = '" & jptxt.Text & "',Att_Time = '" & atttxt.Text & "' where id = '" & idtxt.Text & "'"
        koneksi.Execute editdata
        MsgBox "Data Berhasil Diedit", vbInformation, "Pemberitahuan"
        editbtn.Caption = "EDIT"
        Adodc1.Refresh
        DGkary.Refresh
 Call KondisiAwal
End If
End If

End Sub

Private Sub seekbtn_Click()
formseekabsen.Show
End Sub

Sub KondisiAwal()
idtxt.Text = ""
tgltxt.Text = ""
namatxt.Text = ""
jmktxt.Text = ""
jmtxt.Text = ""
jpktxt.Text = ""
jptxt.Text = ""
atttxt.Text = ""
Combo1.Text = ""
Combo2.Text = ""
newbtn.Enabled = True
editbtn.Enabled = True
deletebtn.Enabled = True
DGkary.Refresh
End Sub

