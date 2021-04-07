VERSION 5.00
Begin VB.MDIForm mdihalutama 
   BackColor       =   &H8000000C&
   Caption         =   "HALAMAN UTAMA"
   ClientHeight    =   9015
   ClientLeft      =   -75
   ClientTop       =   555
   ClientWidth     =   13425
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Menu mntabel 
      Caption         =   "Master Table"
      Begin VB.Menu mnkaryawan 
         Caption         =   "Tabel Karyawan"
      End
      Begin VB.Menu logout 
         Caption         =   "Log Out"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnsetting 
      Caption         =   "Setting"
      Begin VB.Menu mnuser 
         Caption         =   "User"
      End
   End
End
Attribute VB_Name = "mdihalutama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim koneksi As New rdoConnection
Dim rQuery As New rdoQuery
Dim rs As rdoResultset

Private Sub Form_Load()
'koneksi.Connect = "Driver={MySQL ODBC 5.1 Driver};SERVER=192.168.2.11;PWD=123456;UID=BRIWIRA;PORT=3306;DATABASE=karyawan;"
'koneksi.EstablishConnection
If mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : admin@admin.co.id" Then
Else
If mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : ADMIN@ADMIN.CO.ID" Then
Else
If mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : Admin@admin.co.id" Then
Else
If mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : KARYAWAN@KARYAWAN.CO.ID" Then


Else
If mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : Karyawan@karyawan.co.id" Then
Else
If mdihalutama.Caption = "HALAMAN UTAMA. LOGIN SEBAGAI : karyawan@karyawan.co.id" Then
End If
End If
End If
End If
End If
End If
End Sub
    
Private Sub absen_Click()
frmabsensi.Show
End Sub

Private Sub Exit_Click()
Dim psn As Integer
psn = MsgBox("Are you sure want exit", vbYesNo, "Exit")
If psn = vbYes Then
End
Else
End If
End Sub

Private Sub Importabsen_Click()
'frmimportabsen.Show
End Sub

Private Sub importadmin_Click()
frmimportreportadm.Show
End Sub

Private Sub importnamakary_Click()
frmimportkar.Show
End Sub

Private Sub logout_Click()
login2.Show
Unload Me
'koneksi.Close
If konn.State = adStateOpen Then konn.Close
'konn.Close
End Sub




Private Sub mnkaryawan_Click()
frmmstrkary.Show
End Sub

Private Sub mnsurveyor_Click()
frmsurveyor.Show
End Sub

Private Sub mnuser_Click()
frmuser.Show
End Sub

Private Sub reportadm_Click()
frmreportadm.Show
End Sub

Private Sub surat_ijin_Click()
frmsuratijin.Show
End Sub
