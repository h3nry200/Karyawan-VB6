VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form login2 
   Caption         =   "KARYAWAN"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "login2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H80000003&
      Height          =   375
      Left            =   1920
      TabIndex        =   16
      Top             =   1680
      Width           =   1875
   End
   Begin VB.TextBox usertxt 
      Height          =   405
      Left            =   1920
      TabIndex        =   4
      Top             =   720
      Width           =   1845
   End
   Begin VB.TextBox passtxt 
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1200
      Width           =   1845
   End
   Begin VB.CommandButton loginbtn 
      Caption         =   "LOGIN"
      Height          =   465
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   1185
   End
   Begin VB.CommandButton exitbtn 
      Caption         =   "EXIT"
      Height          =   465
      Left            =   3120
      TabIndex        =   1
      Top             =   2160
      Width           =   1185
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   270
      Top             =   240
   End
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
      Left            =   2490
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
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
      TabIndex        =   7
      Top             =   0
      Width           =   2535
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
      TabIndex        =   6
      Top             =   0
      Width           =   1275
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
      TabIndex        =   8
      Top             =   360
      Width           =   525
   End
   Begin VB.Label Label9 
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
      TabIndex        =   15
      Top             =   120
      Width           =   945
   End
   Begin VB.Label Label8 
      Caption         =   "User Name  :"
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   840
      Width           =   1125
   End
   Begin VB.Label Label7 
      Caption         =   "Password    :"
      Height          =   225
      Left            =   600
      TabIndex        =   13
      Top             =   1320
      Width           =   1125
   End
   Begin VB.Label Label5 
      Caption         =   "Level Admin  :"
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   1800
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
      TabIndex        =   11
      Top             =   240
      Width           =   945
   End
   Begin VB.Label Label2 
      Caption         =   "User Name  :"
      Height          =   255
      Left            =   720
      TabIndex        =   10
      Top             =   840
      Width           =   1125
   End
   Begin VB.Label Label3 
      Caption         =   "Password    :"
      Height          =   225
      Left            =   720
      TabIndex        =   9
      Top             =   1320
      Width           =   1125
   End
   Begin VB.Label Label4 
      Caption         =   "Level Admin  :"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1800
      Width           =   1125
   End
End
Attribute VB_Name = "login2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim shari As String
Dim ahari


Private Sub Form_Load()

  ahari = Array("Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu")
  Timer1.Interval = 500
  Timer1.Enabled = True
End Sub



Private Sub Timer1_Timer()
 
 shari = ahari(Abs(Weekday(Date) - 1))
  lbtanggal.Caption = "" & shari & ", " _
                   & Format(Date, "dd mmmm yyyy")
lbjam.Caption = Format(Time, "hh:mm:ss")


'If usertxt.Text = "" Then
'usertxt.SetFocus
'End If

End Sub

Private Sub passtxt_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Call koneksi
    If usertxt.Text = "" Then
        MsgBox "NAMA USER MASIH KOSONG !", vbCritical + vbOKOnly, "Error"
        usertxt.SetFocus
    ElseIf passtxt.Text = "" Then
        MsgBox "PASSWORD MASIH KOSONG !", vbCritical + vbOKOnly, "Error"
        passtxt.SetFocus
    ElseIf lvltxt.Text = "" Then
        MsgBox "LEVEL MASIH KOSONG !", vbCritical + vbOKOnly, "Error"
        lvltxt.SetFocus
ElseIf lvltxt.Text = "PSIKOTES" Then
frmpsi1.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : PESERTA " & UCase(lvltxt.Text)
Unload Me
frmpsi1.Show
ElseIf lvltxt.Text = "1" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(usertxt.Text)

MsgBox "login sukses "
            MsgBox "WELCOME " & UCase(usertxt.Text) & " "
Unload Me
mdihalutama.Show
ElseIf lvltxt.Text = "2" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(usertxt.Text)
mdihalutama.mnuser.Enabled = False
            MsgBox "WELCOME " & UCase(usertxt.Text) & " "
Unload Me
mdihalutama.Show
ElseIf lvltxt.Text = "ADMIN SPV" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(lvltxt.Text)
Unload Me
mdihalutama.Show
ElseIf lvltxt.Text = "ADMIN" Then
mdihalutama.Caption = " " & UCase(lvltxt.Text)
Unload Me
mdihalutama.Show
'frmsurveyor.lbuser.Caption = "'" & usertxt.Text & "'"
    
    Else
        SQL = ""
        SQL = "SELECT * FROM master_users " _
            & "WHERE usernames='" & usertxt.Text & "' " _
            & " AND password='" & passtxt.Text & "' and role_id='" & lvltxt.Text & "'"
            Set rsPeriksa = konn.Execute(SQL)
                   
        If Not rsPeriksa.BOF Then
MsgBox "login sukses "
            MsgBox "WELCOME " & UCase(usertxt.Text) & " "
    
    
            Unload Me
konn.Close

 Call cekdata

'Call cekuser
        Else
                MsgBox "ANDA BUKAN USER YANG BERHAK!", vbCritical + vbOKOnly, "Error"
        End If
    End If
End If
'End If
'End If

End Sub

Private Sub loginbtn_Click()
Call koneksi
    If usertxt.Text = "" Then
        MsgBox "NAMA USER MASIH KOSONG !", vbCritical + vbOKOnly, "Error"
        usertxt.SetFocus
    ElseIf passtxt.Text = "" Then
        MsgBox "PASSWORD MASIH KOSONG !", vbCritical + vbOKOnly, "Error"
        passtxt.SetFocus
    ElseIf lvltxt.Text = "" Then
        MsgBox "LEVEL MASIH KOSONG !", vbCritical + vbOKOnly, "Error"
        lvltxt.SetFocus
    Else
        SQL = ""
        SQL = "SELECT * FROM master_users " _
            & "WHERE usernames='" & usertxt.Text & "' " _
            & " AND pids='" & passtxt.Text & "' and level='" & lvltxt.Text & "'"
            Set rsPeriksa = konn.Execute(SQL)
                   
        If Not rsPeriksa.BOF Then
            MsgBox "login sukses "
            MsgBox "WELCOME " & UCase(usertxt.Text) & " "
'untuk memunculkan warning
frmwarning.Caption = UCase(lvltxt.Text)
'    Call cekuser
            Unload Me
konn.Close
 Call cekdata
konn.Close
'            Call cekuser
'            mdihalutama.Show
        Else
                MsgBox "ANDA BUKAN USER YANG BERHAK!", vbCritical + vbOKOnly, "Error"
        End If
    End If

End Sub

Private Sub usertxt_keypress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If usertxt.Text = "HENRY" Or usertxt.Text = "henry" Or usertxt.Text = "Henry" Or usertxt.Text = "ANDREAS" Then
    lvltxt.Text = "IT SPV"
    passtxt.SetFocus
Else
If usertxt.Text = "Andreas" Or usertxt.Text = "andreas" Then
    lvltxt.Text = "IT"
    passtxt.SetFocus
Else
If usertxt.Text = "PSIKOTES" Or usertxt.Text = "Psikotes" Or usertxt.Text = "psikotes" Then
    lvltxt.Text = "PSIKOTES"
    passtxt.SetFocus
Else
If usertxt.Text = "SUHARYADI" Or usertxt.Text = "suharyadi" Or usertxt.Text = "Suharyadi" Or usertxt.Text = "EVI" Or usertxt.Text = "evi" Or usertxt.Text = "Evi" Then
    lvltxt.Text = "HRD"
    passtxt.SetFocus
Else
If usertxt.Text = "ISMI" Or usertxt.Text = "Ismi" Or usertxt.Text = "ismi" Then
    lvltxt.Text = "HRD SPV"
    passtxt.SetFocus
Else
If usertxt.Text = "rizki" Or usertxt.Text = "Rizki" Or usertxt.Text = "RIZKI" Then
    lvltxt.Text = "ADMIN"
    passtxt.SetFocus
Else
If usertxt.Text = "yuni" Or usertxt.Text = "Yuni" Or usertxt.Text = "YUNI" Then
    lvltxt.Text = "ADMIN SPV"
    passtxt.SetFocus
Else
If usertxt.Text = "admin@admin.co.id" Or usertxt.Text = "Admin@admin.co.id" Or usertxt.Text = "Admin@Admin.co.id" Then
    lvltxt.Text = "1"
        Text1.Text = "Admin"

    passtxt.SetFocus
Else
If usertxt.Text = "Karyawan@Karyawan.co.id" Or usertxt.Text = "karyawan@karyawan.co.id" Or usertxt.Text = "Karyawan@karyawan.co.id" Then
    lvltxt.Text = "2"
        Text1.Text = "Karyawan"

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

End Sub

Private Sub exitbtn_Click()
    End
End Sub

Private Sub cekdata()

Call koneksi
 SQL = ""
        SQL = "SELECT * FROM namakar " _
            & "WHERE tglkontrak1 < '" & lbtanggal.Caption & "' "
            Set rsPeriksa = konn.Execute(SQL)

        If rsPeriksa.BOF Then

frmwarning.Show
konn.Close
Else
Unload Me
Call cekuser
'            MsgBox "login sukses "
 '           MsgBox "WELCOME " & UCase(usertxt.Text) & " "
'konn.Close
End If

End Sub

Private Sub cekuser()

If lvltxt.Text = "IT SPV" Then
mdihalutama.Caption = UCase(usertxt.Text)
Unload Me
mdihalutama.Show
Else
If lvltxt.Text = "PSIKOTES" Then
frmpsi1.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : PESERTA " & UCase(lvltxt.Text)
Unload Me
frmpsi1.Show
Else
If lvltxt.Text = "HRD SPV" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(lvltxt.Text)
Unload Me
mdihalutama.Show
Else
If lvltxt.Text = "HRD" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(lvltxt.Text)
Unload Me
mdihalutama.Show
mdihalutama.mnuser.Enabled = False
Else
If lvltxt.Text = "IT" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(lvltxt.Text)
Unload Me
mdihalutama.Show
Else
If lvltxt.Text = "ADMIN SPV" Then
mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : " & UCase(lvltxt.Text)
Unload Me
mdihalutama.Show
mdihalutama.mnkaryawan.Enabled = False
mdihalutama.mnuser.Enabled = False
Else
If lvltxt.Text = "ADMIN" Then
mdihalutama.Caption = " " & UCase(usertxt.Text)
Unload Me
mdihalutama.Show
mdihalutama.mnuser.Enabled = False
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
