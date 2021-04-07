VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form login 
   Caption         =   "LOGIN HRD DAN PSIKOTES"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox lvltxt 
      BackColor       =   &H00FFFF00&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   405
      Left            =   1920
      TabIndex        =   11
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   270
      Top             =   240
   End
   Begin VB.CommandButton exitbtn 
      Caption         =   "EXIT"
      Height          =   465
      Left            =   3120
      TabIndex        =   6
      Top             =   2160
      Width           =   1185
   End
   Begin VB.CommandButton loginbtn 
      Caption         =   "LOGIN"
      Height          =   465
      Left            =   480
      TabIndex        =   5
      Top             =   2160
      Width           =   1185
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   2160
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
   Begin VB.TextBox passtxt 
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1200
      Width           =   1845
   End
   Begin VB.TextBox usertxt 
      Height          =   405
      Left            =   1920
      TabIndex        =   3
      Top             =   720
      Width           =   1845
   End
   Begin VB.Label Label4 
      Caption         =   "Level Admin  :"
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   1800
      Width           =   1125
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
      Left            =   3360
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label lokasi 
      BackColor       =   &H8000000E&
      Caption         =   "jkt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4200
      TabIndex        =   7
      Top             =   360
      Width           =   525
   End
   Begin VB.Label Label3 
      Caption         =   "Password    :"
      Height          =   225
      Left            =   720
      TabIndex        =   2
      Top             =   1320
      Width           =   1125
   End
   Begin VB.Label Label2 
      Caption         =   "User Name  :"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   840
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   945
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   0
      Picture         =   "login.frx":324A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4560
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Koneksi As New rdoConnection
Dim rQuery As New rdoQuery
Dim rs As rdoResultset
Dim shari As String
Dim ahari


Private Sub Form_Load()

Koneksi.Connect = "Driver={MySQL ODBC 5.1 Driver};SERVER=192.168.2.11;PWD=123456;UID=BRIWIRA;PORT=3306;DATABASE=karyawan;"
Koneksi.EstablishConnection
ahari = Array("Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu")
  Timer1.Interval = 500
  Timer1.Enabled = True
End Sub

Private Sub loginbtn_Click()
rQuery.SQL = "select * from userid where usernames='" & usertxt.Text & "' and pids='" & passtxt.Text & "'"
rQuery.RowsetSize = 1
Set rQuery.ActiveConnection = Koneksi
Set rs = rQuery.OpenResultset(rdOpenKeyset, rdConcurRowVer)
If rs.RowCount > 0 Then
MsgBox "login sukses "
MsgBox "WELCOME " & UCase(usertxt.Text) & " "
Else
MsgBox "username atau password salah"
End If

If usertxt.Text = "HENRY" And lvltxt.Text = "IT SPV" And passtxt.Text = "12345" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(usertxt.Text)
Unload Me
mdihalutama.Show
Else
If usertxt.Text = "PSIKOTES" And lvltxt.Text = "PSIKOTES" And passtxt.Text = "12345" Then
frmpsi1.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : PESERTA " & UCase(usertxt.Text)
Unload Me
frmpsi1.Show
Else
If usertxt.Text = "henry" And lvltxt.Text = "IT SPV" And passtxt.Text = "12345" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(usertxt.Text)
Unload Me
mdihalutama.Show
Else
If usertxt.Text = "psikotes" And lvltxt.Text = "PSIKOTES" And passtxt.Text = "12345" Then
frmpsi1.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : PESERTA " & UCase(usertxt.Text)
Unload Me
frmpsi1.Show
Else
If usertxt.Text = "defi" And lvltxt.Text = "HRD" And passtxt.Text = "12345" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(usertxt.Text)
Unload Me
mdihalutama.Show
Else
If usertxt.Text = "ismi" And lvltxt.Text = "HRD" And passtxt.Text = "12345" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(usertxt.Text)
Unload Me
mdihalutama.Show
Else
If usertxt.Text = "nur" And lvltxt.Text = "HRD" And passtxt.Text = "12345" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(usertxt.Text)
Unload Me
mdihalutama.Show
Else
If usertxt.Text = "andreas" And lvltxt.Text = "IT" And passtxt.Text = "12345" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(usertxt.Text)
Unload Me
mdihalutama.Show
Else
If usertxt.Text = "DEFI" And lvltxt.Text = "HRD" And passtxt.Text = "12345" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(usertxt.Text)
Unload Me
mdihalutama.Show
If usertxt.Text = "ISMI" And lvltxt.Text = "HRD" And passtxt.Text = "12345" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(usertxt.Text)
Unload Me
mdihalutama.Show
If usertxt.Text = "NUR" And lvltxt.Text = "HRD" And passtxt.Text = "12345" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(usertxt.Text)
Unload Me
mdihalutama.Show
If usertxt.Text = "ANDREAS" And lvltxt.Text = "IT" And passtxt.Text = "12345" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(usertxt.Text)
Unload Me
mdihalutama.Show

End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If

rs.Close
Koneksi.Close
End Sub

Private Sub exitbtn_Click()
End
End Sub


Private Sub passtxt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
rQuery.SQL = "select * from userid where usernames='" & usertxt.Text & "' and pids='" & passtxt.Text & "'"
rQuery.RowsetSize = 1
Set rQuery.ActiveConnection = Koneksi
Set rs = rQuery.OpenResultset(rdOpenKeyset, rdConcurRowVer)
If rs.RowCount > 0 Then
MsgBox "login sukses "
MsgBox "WELCOME " & UCase(usertxt.Text) & " "
Else
MsgBox "username atau password salah"
End If


If usertxt.Text = "HENRY" And lvltxt.Text = "IT SPV" And passtxt.Text = "12345" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(usertxt.Text)
Unload Me
mdihalutama.Show
Else
If usertxt.Text = "PSIKOTES" And lvltxt.Text = "PSIKOTES" And passtxt.Text = "12345" Then
frmpsi1.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : PESERTA " & UCase(usertxt.Text)
Unload Me
frmpsi1.Show
Else
If usertxt.Text = "henry" And lvltxt.Text = "IT SPV" And passtxt.Text = "12345" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(usertxt.Text)
Unload Me
mdihalutama.Show
Else
If usertxt.Text = "psikotes" And lvltxt.Text = "PSIKOTES" And passtxt.Text = "12345" Then
frmpsi1.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : PESERTA " & UCase(usertxt.Text)
Unload Me
frmpsi1.Show
Else
If usertxt.Text = "defi" And lvltxt.Text = "HRD" And passtxt.Text = "12345" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(usertxt.Text)
Unload Me
mdihalutama.Show
Else
If usertxt.Text = "ismi" And lvltxt.Text = "HRD" And passtxt.Text = "12345" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(usertxt.Text)
Unload Me
mdihalutama.Show
Else
If usertxt.Text = "nur" And lvltxt.Text = "HRD" And passtxt.Text = "12345" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(usertxt.Text)
Unload Me
mdihalutama.Show
Else
If usertxt.Text = "andreas" And lvltxt.Text = "IT" And passtxt.Text = "12345" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(usertxt.Text)
Unload Me
mdihalutama.Show
Else
If usertxt.Text = "DEFI" And lvltxt.Text = "HRD" And passtxt.Text = "12345" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(usertxt.Text)
Unload Me
mdihalutama.Show
Else
If usertxt.Text = "ISMI" And lvltxt.Text = "HRD" And passtxt.Text = "12345" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(usertxt.Text)
Unload Me
mdihalutama.Show
Else
If usertxt.Text = "NUR" And lvltxt.Text = "HRD" And passtxt.Text = "12345" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(usertxt.Text)
Unload Me
mdihalutama.Show
Else
If usertxt.Text = "ANDREAS" And lvltxt.Text = "IT" And passtxt.Text = "12345" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(usertxt.Text)
Unload Me
mdihalutama.Show
Else

End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If


rs.Close
Koneksi.Close
End If
End Sub

Private Sub Timer1_Timer()
 shari = ahari(Abs(Weekday(Date) - 1))
  lbtanggal.Caption = "" & shari & ", " _
                   & Format(Date, "dd mmmm yyyy")
lbjam.Caption = Format(Time, "hh:mm:ss")
End Sub

Private Sub usertxt_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
If usertxt.Text = "HENRY" Then
lvltxt.Text = "IT SPV"
passtxt.SetFocus
Else
If usertxt.Text = "henry" Then
lvltxt.Text = "IT SPV"
passtxt.SetFocus
Else
If usertxt.Text = "PSIKOTES" Then
lvltxt.Text = "PSIKOTES"
passtxt.SetFocus
Else
If usertxt.Text = "psikotes" Then
lvltxt.Text = "PSIKOTES"
passtxt.SetFocus
Else
If usertxt.Text = "DEFI" Then
lvltxt.Text = "HRD"
passtxt.SetFocus
Else
If usertxt.Text = "NUR" Then
lvltxt.Text = "HRD"
passtxt.SetFocus
Else
If usertxt.Text = "ISMI" Then
lvltxt.Text = "HRD"
passtxt.SetFocus
Else
If usertxt.Text = "ANDREAS" Then
lvltxt.Text = "IT"
passtxt.SetFocus
Else
If usertxt.Text = "defi" Then
lvltxt.Text = "HRD"
passtxt.SetFocus
Else
If usertxt.Text = "nur" Then
lvltxt.Text = "HRD"
passtxt.SetFocus
Else
If usertxt.Text = "ismi" Then
lvltxt.Text = "HRD"
passtxt.SetFocus
Else
If usertxt.Text = "andreas" Then
lvltxt.Text = "IT"
passtxt.SetFocus
Else
lvltxt.Text = ""
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If
End If

End Sub
