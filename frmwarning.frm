VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmwarning 
   Caption         =   "WARNING KARYAWAN "
   ClientHeight    =   6240
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16080
   LinkTopic       =   "Form2"
   ScaleHeight     =   6240
   ScaleWidth      =   16080
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1200
      Top             =   480
   End
   Begin VB.CommandButton save 
      Caption         =   "SAVE"
      Height          =   495
      Left            =   8520
      TabIndex        =   18
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox keterangantxt 
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   10320
      TabIndex        =   16
      Top             =   4320
      Width           =   5655
   End
   Begin VB.TextBox statuskartxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   13
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton exit 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   10680
      TabIndex        =   12
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox tglmasuktxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   3360
      Width           =   2175
   End
   Begin VB.TextBox jabatantxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   4800
      Width           =   2175
   End
   Begin VB.TextBox namatxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   6
      Top             =   4080
      Width           =   2175
   End
   Begin VB.TextBox idtxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   270
      Top             =   240
   End
   Begin MSDataGridLib.DataGrid DGkary 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   3836
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
      Height          =   495
      Left            =   720
      Top             =   1800
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   873
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
   Begin MSComCtl2.DTPicker tglkontrak1txt 
      Height          =   375
      Left            =   10320
      TabIndex        =   19
      Top             =   3360
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   0
      CalendarTitleForeColor=   0
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   180289539
      CurrentDate     =   42809
   End
   Begin VB.Label Label8 
      Caption         =   "KETERANGAN        :"
      Height          =   255
      Left            =   8400
      TabIndex        =   17
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "TGL KONTRAK        :"
      Height          =   255
      Left            =   8400
      TabIndex        =   15
      Top             =   3360
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "STATUS KARYAWAN :"
      Height          =   255
      Left            =   3840
      TabIndex        =   14
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "TGL MASUK :"
      Height          =   255
      Left            =   3840
      TabIndex        =   11
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "JABATAN :"
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "NAMA      :"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "ID             :"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "WARNING!!!  ADA KARYAWAN YANG TELAH LEWAT TGL KONTRAK DAN BELUM DI PERPANJANG"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   15735
   End
   Begin VB.Label lbjam 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   14640
      TabIndex        =   2
      Top             =   0
      Width           =   1275
   End
   Begin VB.Label lbtanggal 
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "frmwarning"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim shari As String
Dim ahari


Private Sub Exit_Click()

If DGkary.Columns(0).Text = " " Then
mdihalutama.Caption = Me.Caption
Call cekuser
Unload Me
mdihalutama.Show

Else
    MsgBox "HARAP GANTI TGL KONTRAK DAN ISI KETERANGAN"
Call cekuser
End If

End Sub

Private Sub Form_Load()
  
  ahari = Array("Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu")
  Timer1.Interval = 500
  Timer1.Enabled = True

'Call koneksi

    Adodc1.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};SERVER=192.168.2.11;PWD=123456;UID=BRIWIRA;PORT=3306;DATABASE=karyawan;"
    Adodc1.RecordSource = "select * from namakar where tglkontrak1 >= '" & lbtanggal.Caption & "' and tgl_masuk <> tglkontrak1 and sts_karyawan = 'KONTRAK' and status = 0 and tgl_keluar = '-' and tgl_aju_keluar = '-'"
    
    Adodc1.Refresh
    
Set DGkary.DataSource = Adodc1

End Sub

Private Sub save_Click()
Dim editdata As String
Call koneksi
        editdata = "update namakar set tglkontrak1 = '" & tglkontrak1txt.Value & "',status = '1',userdt = '" & frmwarning.Caption & "', keteranganlain2 = '" & keterangantxt.Text & "'  where id = '" & idtxt.Text & "'"
        konn.Execute editdata
        MsgBox "Data Berhasil Diedit", vbInformation, "Pemberitahuan"
'        editbtn.Caption = "EDIT"
        Adodc1.Refresh
        DGkary.Refresh
End Sub

Private Sub Timer1_Timer()
 
 shari = ahari(Abs(Weekday(Date) - 1))
  lbtanggal.Caption = Format(Date, "yyyy-MM-dd")
lbjam.Caption = Format(Time, "hh:mm:ss")

'If usertxt.Text = "" Then
'usertxt.SetFocus
'End If

End Sub


Private Sub DGkary_Click()
idtxt.Text = DGkary.Columns(0).Text
namatxt.Text = DGkary.Columns(3).Text
jabatantxt.Text = DGkary.Columns(16).Text
tglmasuktxt.Text = DGkary.Columns(9).Text
statuskartxt.Text = DGkary.Columns(17).Text
tglkontrak1txt.Value = DGkary.Columns(37).Text
keterangantxt.Text = DGkary.Columns(41).Text

End Sub

Private Sub DGkary_KeyDown(KeyCode As Integer, Shift As Integer)
idtxt.Text = DGkary.Columns(0).Text
namatxt.Text = DGkary.Columns(3).Text
jabatantxt.Text = DGkary.Columns(16).Text
tglmasuktxt.Text = DGkary.Columns(9).Text
statuskartxt.Text = DGkary.Columns(17).Text
tglkontrak1txt.Value = DGkary.Columns(37).Text
keterangantxt.Text = DGkary.Columns(41).Text

End Sub

Private Sub cekuser()
If Me.Caption = "IT SPV" Then
mdihalutama.Caption = UCase(Me.Caption)
Unload Me
mdihalutama.Show
Else
If Me.Caption = "PSIKOTES" Then
frmpsi1.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : PESERTA " & UCase(Me.Caption)
Unload Me
frmpsi1.Show
Else
If Me.Caption = "HRD SPV" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(Me.Caption)
Unload Me
mdihalutama.Show
Else
If Me.Caption = "HRD" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(Me.Caption)
Unload Me
mdihalutama.Show
mdihalutama.mnuser.Enabled = False

Else
If Me.Caption = "IT" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(Me.Caption)
Unload Me
mdihalutama.Show
Else
If Me.Caption = "ADMIN SPV" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(Me.Caption)
Unload Me
mdihalutama.Show
mdihalutama.mnkaryawan.Enabled = False
mdihalutama.absen.Enabled = False
mdihalutama.surat_ijin.Enabled = False
mdihalutama.importnamakary.Enabled = False
mdihalutama.Importabsen.Enabled = False

Else
If Me.Caption = "ADMIN" Then
mdihalutama.Caption = " " & UCase(Me.Caption)
Unload Me
mdihalutama.Show
mdihalutama.mnuser.Enabled = False
mdihalutama.mnkaryawan.Enabled = False
mdihalutama.absen.Enabled = False
mdihalutama.surat_ijin.Enabled = False
mdihalutama.importnamakary.Enabled = False
mdihalutama.Importabsen.Enabled = False
'frmsurveyor.lbuser.Caption = "'" & usertxt.Text & "'"
Else
Unload Me
mdihalutama.Show

End If
End If
End If
End If
End If
End If
End If

End Sub

Private Sub Timer2_Timer()
Label1.ForeColor = RGB(Rnd * 250, Rnd * 250, Rnd * 250)
    If (Label1.Left + Label1.Width) <= 0 Then
        Label1.Left = Me.Width
    End If
    Label1.Left = Label1.Left - 100
End Sub
