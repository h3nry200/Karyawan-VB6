VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmsuratijin 
   Caption         =   "FORM SURAT IJIN"
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox idtxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   23
      Top             =   4920
      Width           =   855
   End
   Begin VB.CommandButton printbtn 
      Caption         =   "PRINT"
      Height          =   375
      Left            =   11760
      TabIndex        =   22
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton seek2 
      Caption         =   "SEEK"
      Height          =   375
      Left            =   10200
      TabIndex        =   21
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton savebtn 
      Caption         =   "SAVE"
      Height          =   375
      Left            =   11760
      TabIndex        =   20
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton newbtn 
      Caption         =   "NEW"
      Height          =   375
      Left            =   10200
      TabIndex        =   19
      Top             =   5640
      Width           =   1095
   End
   Begin VB.CommandButton editbtn 
      Caption         =   "EDIT"
      Height          =   375
      Left            =   10200
      TabIndex        =   18
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton cancelbtn 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   11760
      TabIndex        =   17
      Top             =   6360
      Width           =   1095
   End
   Begin VB.CommandButton deletebtn 
      Caption         =   "DELETE"
      Height          =   375
      Left            =   11760
      TabIndex        =   16
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton exitbtn 
      Caption         =   "EXIT"
      Height          =   375
      Left            =   10200
      TabIndex        =   15
      Top             =   7080
      Width           =   1095
   End
   Begin VB.TextBox kettxt 
      Enabled         =   0   'False
      Height          =   1335
      Left            =   360
      TabIndex        =   14
      Top             =   8280
      Width           =   6855
   End
   Begin VB.CheckBox appr 
      Caption         =   "APPROVED"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8160
      TabIndex        =   12
      Top             =   4920
      Width           =   1335
   End
   Begin VB.TextBox lamatxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   11
      Top             =   6120
      Width           =   3495
   End
   Begin VB.ComboBox statuscmb 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2520
      TabIndex        =   9
      Top             =   7320
      Width           =   3495
   End
   Begin VB.CommandButton seekbtn 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   6720
      Width           =   375
   End
   Begin VB.TextBox namatxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   6720
      Width           =   3495
   End
   Begin MSComCtl2.DTPicker tglakhir 
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   5520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   154075137
      CurrentDate     =   42654
   End
   Begin MSDataGridLib.DataGrid DGsrt 
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   19935
      _ExtentX        =   35163
      _ExtentY        =   8281
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
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
         Size            =   9.75
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
      Left            =   1440
      Top             =   2400
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
   Begin MSComCtl2.DTPicker tglawal 
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   5520
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   154075137
      CurrentDate     =   42654
   End
   Begin VB.Label Label7 
      Caption         =   "ID                                    :"
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Line Line3 
      X1              =   9840
      X2              =   9840
      Y1              =   9840
      Y2              =   4800
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   9840
      Y1              =   9840
      Y2              =   9840
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   120
      Y1              =   4800
      Y2              =   9840
   End
   Begin VB.Label Label6 
      Caption         =   "LAMA IJIN                       :"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   6120
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "KETERANGAN               :"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   7920
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "STATUS IJIN                  :"
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "NAMA KARYAWAN        :"
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "TO"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   5520
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "PERIODE                        :"
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   5520
      Width           =   1935
   End
End
Attribute VB_Name = "frmsuratijin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim koneksi As New rdoConnection
Dim rQuery As New rdoQuery
Dim rs As rdoResultset


Private Sub Form_Load()
koneksi.Connect = "Driver={MySQL ODBC 5.1 Driver};SERVER=192.168.2.11;PWD=123456;UID=BRIWIRA;PORT=3306;DATABASE=karyawan;"
koneksi.EstablishConnection
    Adodc1.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};SERVER=192.168.2.11;PWD=123456;UID=BRIWIRA;PORT=3306;DATABASE=karyawan;"
    Adodc1.RecordSource = "suratijin"
    Adodc1.Refresh
Set DGsrt.DataSource = Adodc1
    statuscmb.AddItem "IJIN"
    statuscmb.AddItem "SAKIT"
    statuscmb.AddItem "CUTI"
DGsrt.Refresh
cancelbtn.Enabled = False
Call Bersih

End Sub

Private Sub DGsrt_Click()
'    Adodc1.Refresh
idtxt.Text = DGsrt.Columns(0).Text
namatxt.Text = DGsrt.Columns(1).Text
tglawal.Value = DGsrt.Columns(2).Text
tglakhir.Value = DGsrt.Columns(3).Text
lamatxt.Text = DGsrt.Columns(4).Text
statuscmb.Text = DGsrt.Columns(5).Text
kettxt.Text = DGsrt.Columns(6).Text
appr.Value = DGsrt.Columns(7).Text
End Sub

Sub Bersih()
idtxt.Text = ""
lamatxt.Text = ""
namatxt.Text = ""
statuscmb.Text = ""
kettxt.Text = ""
appr.Value = 0
End Sub

Sub KondisiAwal()
idtxt.Text = ""
lamatxt.Text = ""
namatxt.Text = ""
statuscmb.Text = ""
kettxt.Text = ""
appr.Value = 0
seek2.Enabled = True
newbtn.Enabled = True
editbtn.Enabled = True
deletebtn.Enabled = True
cancelbtn.Enabled = False
DGsrt.Refresh
End Sub

Private Sub deletebtn_Click()
If MsgBox("Yakin Ingin Menghapus Data?", vbQuestion + vbOKCancel, "konfirmasi") = vbOK Then
Dim hapusdata As String
        hapusdata = "delete from suratijin where id =" & idtxt.Text & ""
        koneksi.Execute hapusdata
        MsgBox "Data Berhasil Dihapus", vbInformation, "Pemberitahuan"
    Adodc1.Refresh
DGsrt.Refresh
Call Bersih

Else
Call Bersih
End If

End Sub

Private Sub cancelbtn_Click()
 Call Bersih
Call KondisiAwal
 newbtn.Enabled = True
 editbtn.Enabled = True
 deletebtn.Enabled = True
seek2.Enabled = True
idtxt.Enabled = False
lamatxt.Enabled = False
namatxt.Enabled = False
statuscmb.Enabled = False
kettxt.Enabled = False
tglawal.Enabled = False
tglakhir.Enabled = False
appr.Enabled = False
seekbtn.Enabled = False
 newbtn.Caption = "NEW"
 editbtn.Caption = "EDIT"
End Sub

Private Sub exitbtn_Click()
koneksi.Close
Unload Me
End Sub


Private Sub newbtn_Click()
editbtn.Enabled = False
deletebtn.Enabled = False
savebtn.Enabled = True
newbtn.Enabled = False
seek2.Enabled = False
cancelbtn.Enabled = True
idtxt.Enabled = True
lamatxt.Enabled = False
namatxt.Enabled = True
statuscmb.Enabled = True
kettxt.Enabled = True
tglawal.Enabled = True
tglakhir.Enabled = True
appr.Value = 0
appr.Enabled = True
seekbtn.Enabled = True
idtxt.SetFocus
Call Bersih
newbtn.Caption = "NEWDATA"
End Sub

Private Sub editbtn_Click()
editbtn.Enabled = False
deletebtn.Enabled = False
savebtn.Enabled = True
newbtn.Enabled = False
seek2.Enabled = False
cancelbtn.Enabled = True
idtxt.Enabled = False
lamatxt.Enabled = False
namatxt.Enabled = True
statuscmb.Enabled = True
kettxt.Enabled = True
tglawal.Enabled = True
tglakhir.Enabled = True
appr.Enabled = True
seekbtn.Enabled = True

editbtn.Caption = "EDITDATA"
End Sub


Private Sub printbtn_Click()

DataReport1.Show
Set DataReport1.DataSource = Adodc1

'koneksi
DataReport1.Sections("Section4").Controls.Item("LABEL4").Caption = tglawal.Value
DataReport1.Sections("Section4").Controls.Item("LABEL7").Caption = tglakhir.Value
DataReport1.Sections("Section4").Controls.Item("LABEL2").Caption = idtxt.Text
DataReport1.Sections("Section4").Controls.Item("LABEL11").Caption = namatxt.Text
DataReport1.Sections("Section4").Controls.Item("LABEL9").Caption = lamatxt.Text
DataReport1.Sections("Section4").Controls.Item("LABEL13").Caption = statuscmb.Text
DataReport1.Sections("Section4").Controls.Item("LABEL15").Caption = kettxt.Text
DataReport1.Sections("Section4").Controls.Item("LABEL17").Caption = namatxt.Text

End Sub

Private Sub savebtn_Click()
If newbtn.Caption = "NEWDATA" Then
If tglawal.Value = "" Or tglakhir.Value = "" Or namatxt.Text = "" Or lamatxt.Text = "" Or statuscmb.Text = "" Or kettxt.Text = "" Then
    MsgBox "Data Belum Lengkap"
    Else
    Dim tambahdata As String
        tambahdata = "Insert Into suratijin values ('" & idtxt.Text & "', '" & namatxt.Text & "','" & tglawal.Value & "','" & tglakhir.Value & "','" & lamatxt.Text & "','" & statuscmb.Text & "','" & kettxt.Text & "','" & appr.Value & "')"
        koneksi.Execute tambahdata
        MsgBox "Data Berhasil Ditambah", vbInformation, "Pemberitahuan"
        newbtn.Caption = "NEW"
        Adodc1.Refresh
        DGsrt.Refresh
        Call KondisiAwal
    
    End If
Else
If editbtn.Caption = "EDITDATA" Then
    Dim editdata As String
        editdata = "update suratijin set id = '" & idtxt.Text & "',nama = '" & namatxt.Text & "',tgl_awal_ijin = '" & tglawal.Value & "',tgl_akhir_ijin = '" & tglakhir.Value & "',lama_ijin = '" & lamatxt.Text & "',status = '" & statuscmb.Text & "',keterangan = '" & kettxt.Text & "',appr = '" & appr.Value & "' where id = '" & idtxt.Text & "'"
        koneksi.Execute editdata
        MsgBox "Data Berhasil Diedit", vbInformation, "Pemberitahuan"
        editbtn.Caption = "EDIT"
        Adodc1.Refresh
        DGsrt.Refresh
 Call KondisiAwal
End If
End If
End Sub

Private Sub seek2_Click()
frmseeksi.Show
End Sub

Private Sub seekbtn_Click()
frmseekkar.Show
End Sub


Private Sub tglawal_Change()
Dim Tahun As Integer, Sisa As Integer
Dim SelisihBulan As Integer
Dim Selisihhari As Integer
Selisihhari = DateDiff("d", tglawal.Value, tglakhir.Value)
Sisa = Selisihhari Mod 12
lamatxt.Text = "" & Sisa & ""
End Sub

Private Sub tglakhir_Change()
Dim Tahun As Integer, Sisa As Integer
Dim Selisihhari As Integer
Selisihhari = DateDiff("d", tglawal.Value, tglakhir.Value)
Sisa = Selisihhari Mod 12
lamatxt.Text = "" & Sisa & ""
End Sub

