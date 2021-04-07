VERSION 5.00
Begin VB.Form frmpsi1 
   Caption         =   "SELAMAT DATANG DI PSIKOTES"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8370
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   8370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Save 
      Caption         =   "GO !!!!!"
      Height          =   495
      Left            =   2670
      TabIndex        =   9
      Top             =   3300
      Width           =   2115
   End
   Begin VB.TextBox tgltxt 
      Height          =   375
      Left            =   1950
      TabIndex        =   8
      Top             =   2730
      Width           =   6405
   End
   Begin VB.TextBox lamartxt 
      Height          =   375
      Left            =   1950
      TabIndex        =   7
      Top             =   2100
      Width           =   6405
   End
   Begin VB.TextBox namatxt 
      Height          =   375
      Left            =   1950
      TabIndex        =   6
      Top             =   1500
      Width           =   6405
   End
   Begin VB.Timer Timer1 
      Left            =   210
      Top             =   360
   End
   Begin VB.TextBox idtxt 
      Height          =   375
      Left            =   1950
      TabIndex        =   4
      Top             =   930
      Width           =   6405
   End
   Begin VB.Label Label5 
      Caption         =   "ID                          :"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1020
      Width           =   1515
   End
   Begin VB.Label Label4 
      Caption         =   "TANGGAL TEST   :"
      Height          =   165
      Left            =   240
      TabIndex        =   3
      Top             =   2820
      Width           =   1605
   End
   Begin VB.Label Label3 
      Caption         =   "LAMAR SEBAGAI :"
      Height          =   405
      Left            =   270
      TabIndex        =   2
      Top             =   2220
      Width           =   1485
   End
   Begin VB.Label Label2 
      Caption         =   "NAMA                :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   270
      TabIndex        =   1
      Top             =   1530
      Width           =   1425
   End
   Begin VB.Label Label1 
      Caption         =   "SILAKAN MASUKKAN DATA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2160
      TabIndex        =   0
      Top             =   480
      Width           =   4635
   End
End
Attribute VB_Name = "frmpsi1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim koneksi As New rdoConnection
Dim rQuery As New rdoQuery
Dim rs As rdoResultset


Private Sub Form_Load()
Call Bersih
koneksi.Connect = "Driver={MySQL ODBC 5.1 Driver};SERVER=192.168.2.11;PWD=123456;UID=BRIWIRA;PORT=3306;DATABASE=karyawan;"
koneksi.EstablishConnection
End Sub
Sub Bersih()
namatxt.Text = ""
lamartxt.Text = ""
tgltxt.Text = ""
End Sub


'Private Sub Command1_Click()
'If idtxt.Text = "" Or namatxt.Text = "" Or lamartxt.Text = "" Or tgltxt.Text = "" Then
'    MsgBox "Data Belum Lengkap"
'    Else
'    Dim tambahdata As String
'        tambahdata = "Insert Into usertest values ('" & idtxt.Text & "','" & namatxt.Text & "','" & lamartxt.Text & "','" & tgltxt.Text & "')"
'        koneksi.Execute tambahdata
'frmpsi2.id.Text = idtxt.Text
'frmpsi2.nama.Text = namatxt.Text
'frmpsi2.tanggal.Text = tgltxt.Text
'frmpsi2.Show
'koneksi.Close
'Unload Me

'        MsgBox "Data Berhasil Ditambah", vbInformation, "Pemberitahuan"
 '       newbtn.Caption = "NEW"
  '      Adodc1.Refresh
   '     DGuser.Refresh
    '    Call KondisiAwal
    
'    End If

'End Sub

Private Sub Save_Click()
If idtxt.Text = "" Or namatxt.Text = "" Or lamartxt.Text = "" Or tgltxt.Text = "" Then
    MsgBox "Data Belum Lengkap"
    Else
    Dim tambahdata As String
        tambahdata = "Insert Into usertest values ('" & idtxt.Text & "','" & namatxt.Text & "','" & lamartxt.Text & "','" & tgltxt.Text & "')"
        koneksi.Execute tambahdata
frmpsi2.id.Text = idtxt.Text
frmpsi2.nama.Text = namatxt.Text
frmpsi2.tanggal.Text = tgltxt.Text
frmpsi3.id.Text = idtxt.Text
frmpsi3.nama.Text = namatxt.Text
frmpsi3.tanggal.Text = tgltxt.Text
frmpsi4.id.Text = idtxt.Text
frmpsi4.nama.Text = namatxt.Text
frmpsi4.tanggal.Text = tgltxt.Text
frmpsi5.id.Text = idtxt.Text
frmpsi5.nama.Text = namatxt.Text
frmpsi5.tanggal.Text = tgltxt.Text
frmpsi6.id.Text = idtxt.Text
frmpsi6.nama.Text = namatxt.Text
frmpsi6.tanggal.Text = tgltxt.Text
frmpsi7.id.Text = idtxt.Text
frmpsi7.nama.Text = namatxt.Text
frmpsi7.tanggal.Text = tgltxt.Text
frmpsi8.id.Text = idtxt.Text
frmpsi8.nama.Text = namatxt.Text
frmpsi8.tanggal.Text = tgltxt.Text
frmpsi9.id.Text = idtxt.Text
frmpsi9.nama.Text = namatxt.Text
frmpsi9.tanggal.Text = tgltxt.Text
frmpsi10.id.Text = idtxt.Text
frmpsi10.nama.Text = namatxt.Text
frmpsi10.tanggal.Text = tgltxt.Text
frmpsi11.id.Text = idtxt.Text
frmpsi11.nama.Text = namatxt.Text
frmpsi11.tanggal.Text = tgltxt.Text
frmpsi12.id.Text = idtxt.Text
frmpsi12.nama.Text = namatxt.Text
frmpsi12.tanggal.Text = tgltxt.Text
frmpsi13.id.Text = idtxt.Text
frmpsi13.nama.Text = namatxt.Text
frmpsi13.tanggal.Text = tgltxt.Text
frmpsi14.id.Text = idtxt.Text
frmpsi14.nama.Text = namatxt.Text
frmpsi14.tanggal.Text = tgltxt.Text
frmpsi15.id.Text = idtxt.Text
frmpsi15.nama.Text = namatxt.Text
frmpsi15.tanggal.Text = tgltxt.Text
frmpsi2.Show
koneksi.Close
Unload Me

'        MsgBox "Data Berhasil Ditambah", vbInformation, "Pemberitahuan"
 '       newbtn.Caption = "NEW"
  '      Adodc1.Refresh
   '     DGuser.Refresh
    '    Call KondisiAwal
    
    End If


End Sub
