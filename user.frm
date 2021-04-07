VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmuser 
   Caption         =   "USER SETTING"
   ClientHeight    =   11010
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19350
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11010
   ScaleWidth      =   19350
   WindowState     =   2  'Maximized
   Begin VB.TextBox createdbytxt 
      Enabled         =   0   'False
      Height          =   465
      Left            =   12660
      TabIndex        =   23
      Top             =   6960
      Width           =   2805
   End
   Begin VB.TextBox activetxt 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   465
      Left            =   6690
      TabIndex        =   18
      Top             =   7020
      Width           =   2805
   End
   Begin VB.TextBox blocktxt 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   6690
      TabIndex        =   17
      Top             =   7590
      Width           =   2805
   End
   Begin VB.TextBox createtxt 
      Enabled         =   0   'False
      Height          =   495
      Left            =   6720
      TabIndex        =   16
      Top             =   8280
      Width           =   2805
   End
   Begin VB.TextBox updatetxt 
      Enabled         =   0   'False
      Height          =   465
      Left            =   6750
      TabIndex        =   15
      Top             =   8880
      Width           =   2805
   End
   Begin MSDataGridLib.DataGrid DGuser 
      Height          =   2505
      Left            =   60
      TabIndex        =   14
      Top             =   450
      Width           =   19155
      _ExtentX        =   33787
      _ExtentY        =   4419
      _Version        =   393216
      AllowUpdate     =   0   'False
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
      Height          =   375
      Left            =   210
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.CommandButton cancelbtn 
      Caption         =   "CANCEL"
      Height          =   555
      Left            =   8580
      TabIndex        =   13
      Top             =   10080
      Width           =   885
   End
   Begin VB.TextBox leveltxt 
      Enabled         =   0   'False
      Height          =   465
      Left            =   1200
      TabIndex        =   11
      Top             =   8970
      Width           =   2805
   End
   Begin VB.CommandButton exitbtn 
      Caption         =   "EXIT"
      Height          =   555
      Left            =   9450
      TabIndex        =   10
      Top             =   10080
      Width           =   885
   End
   Begin VB.CommandButton savebtn 
      Caption         =   "SAVE"
      Height          =   555
      Left            =   7710
      TabIndex        =   9
      Top             =   10080
      Width           =   885
   End
   Begin VB.CommandButton deletebtn 
      Caption         =   "DELETE"
      Height          =   555
      Left            =   6840
      TabIndex        =   8
      Top             =   10080
      Width           =   885
   End
   Begin VB.CommandButton editbtn 
      Caption         =   "EDIT"
      Height          =   555
      Left            =   5880
      TabIndex        =   7
      Top             =   10080
      Width           =   975
   End
   Begin VB.CommandButton newbtn 
      Caption         =   "NEW"
      Height          =   555
      Left            =   4890
      TabIndex        =   6
      Top             =   10080
      Width           =   1005
   End
   Begin VB.TextBox passtxt 
      Enabled         =   0   'False
      Height          =   495
      Left            =   1230
      TabIndex        =   5
      Top             =   8310
      Width           =   2805
   End
   Begin VB.TextBox usertxt 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   7620
      Width           =   2805
   End
   Begin VB.TextBox idtxt 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   465
      Left            =   1200
      TabIndex        =   3
      Top             =   7050
      Width           =   2805
   End
   Begin MSDataGridLib.DataGrid DGuser2 
      Height          =   2595
      Left            =   60
      TabIndex        =   25
      Top             =   3690
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   4577
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   3870
      Top             =   1680
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSDataGridLib.DataGrid DGuser3 
      Height          =   2595
      Left            =   8430
      TabIndex        =   26
      Top             =   3690
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   4577
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   8220
      Top             =   1530
      Width           =   1335
      _ExtentX        =   2355
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
   Begin VB.Label Label12 
      Caption         =   "LIST MENU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   7680
      TabIndex        =   29
      Top             =   3060
      Width           =   5475
   End
   Begin VB.Label Label11 
      Caption         =   "ROLES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   90
      TabIndex        =   28
      Top             =   3090
      Width           =   5475
   End
   Begin VB.Label Label10 
      Caption         =   "USER ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   360
      TabIndex        =   27
      Top             =   -30
      Width           =   5475
   End
   Begin VB.Label Label9 
      Caption         =   "CREATED BY        :"
      Height          =   225
      Left            =   11220
      TabIndex        =   24
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "ACTIVE       :"
      Height          =   225
      Left            =   5640
      TabIndex        =   22
      Top             =   7170
      Width           =   945
   End
   Begin VB.Label Label7 
      Caption         =   "BLOCK        :"
      Height          =   285
      Left            =   5640
      TabIndex        =   21
      Top             =   7770
      Width           =   1005
   End
   Begin VB.Label Label6 
      Caption         =   "CREATE        : "
      Height          =   255
      Left            =   5640
      TabIndex        =   20
      Top             =   8370
      Width           =   1125
   End
   Begin VB.Label Label5 
      Caption         =   "UPDATE        :"
      Height          =   225
      Left            =   5610
      TabIndex        =   19
      Top             =   9000
      Width           =   1065
   End
   Begin VB.Label label4 
      Caption         =   "LEVEL           :"
      Height          =   225
      Left            =   120
      TabIndex        =   12
      Top             =   9030
      Width           =   1035
   End
   Begin VB.Label Label3 
      Caption         =   "PASS            : "
      Height          =   255
      Left            =   150
      TabIndex        =   2
      Top             =   8400
      Width           =   1125
   End
   Begin VB.Label Label2 
      Caption         =   "USERNAME :"
      Height          =   285
      Left            =   150
      TabIndex        =   1
      Top             =   7800
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "ID                :"
      Height          =   225
      Left            =   150
      TabIndex        =   0
      Top             =   7200
      Width           =   945
   End
End
Attribute VB_Name = "frmuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim koneksi As New rdoConnection
Dim rQuery As New rdoQuery
Dim rs As rdoResultset
'Dim rinput As New rdorecordset


Private Sub DGuser_Click()
idtxt.Text = DGuser.Columns(0).Text
usertxt.Text = DGuser.Columns(1).Text
passtxt.Text = DGuser.Columns(3).Text
leveltxt.Text = DGuser.Columns(2).Text
activetxt.Text = DGuser.Columns(4).Text
blocktxt.Text = DGuser.Columns(5).Text
createtxt.Text = DGuser.Columns(6).Text
updatetxt.Text = DGuser.Columns(7).Text
createdbytxt.Text = DGuser.Columns(8).Text
End Sub


Private Sub deletebtn_Click()
If MsgBox("Yakin Ingin Menghapus Data?", vbQuestion + vbOKCancel, "konfirmasi") = vbOK Then
Dim hapusdata As String
        hapusdata = "delete from userid where id_user =" & idtxt.Text & ""
        koneksi.Execute hapusdata
        MsgBox "Data Berhasil Dihapus", vbInformation, "Pemberitahuan"
    Adodc1.Refresh
DGuser.Refresh
Call Bersih

Else
Call Bersih
End If

End Sub

Private Sub DGuser_KeyPress(KeyAscii As Integer)
idtxt.Text = DGuser.Columns(0).Text
usertxt.Text = DGuser.Columns(1).Text
passtxt.Text = DGuser.Columns(3).Text
leveltxt.Text = DGuser.Columns(2).Text
activetxt.Text = DGuser.Columns(4).Text
blocktxt.Text = DGuser.Columns(5).Text
createtxt.Text = DGuser.Columns(6).Text
updatetxt.Text = DGuser.Columns(7).Text
createdbytxt.Text = DGuser.Columns(8).Text
End Sub

Private Sub Form_Load()

Call Bersih

koneksi.Connect = "Driver={MySQL ODBC 5.1 Driver};SERVER=localhost;PWD=;UID=root;PORT=3306;DATABASE=karyawan;"
koneksi.EstablishConnection
    Adodc1.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};SERVER=localhost;PWD=;UID=root;PORT=3306;DATABASE=karyawan;"
    Adodc1.RecordSource = "master_users"
    Adodc1.Refresh
Set DGuser.DataSource = Adodc1

    Adodc2.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};SERVER=localhost;PWD=;UID=root;PORT=3306;DATABASE=karyawan;"
    Adodc2.RecordSource = "master_role"
    Adodc2.Refresh
Set DGuser2.DataSource = Adodc2

   Adodc3.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};SERVER=localhost;PWD=;UID=root;PORT=3306;DATABASE=karyawan;"
    Adodc3.RecordSource = "master_list_menu"
    Adodc3.Refresh
Set DGuser3.DataSource = Adodc3

idtxt.Enabled = False
usertxt.Enabled = False
passtxt.Enabled = False
leveltxt.Enabled = False
activetxt.Enabled = False
blocktxt.Enabled = False
createtxt.Enabled = False
updatetxt.Enabled = False
createdbytxt.Enabled = False
cancelbtn.Enabled = False
savebtn.Enabled = False

End Sub

Sub Bersih()
idtxt.Text = ""
usertxt.Text = ""
passtxt.Text = ""
leveltxt.Text = ""
activetxt.Text = ""
blocktxt.Text = ""
createtxt.Text = ""
updatetxt.Text = ""
createdbytxt.Text = ""
End Sub

Private Sub newbtn_Click()
editbtn.Enabled = False
deletebtn.Enabled = False
savebtn.Enabled = True
newbtn.Enabled = False
cancelbtn.Enabled = True
idtxt.Enabled = True
usertxt.Enabled = True
passtxt.Enabled = True
leveltxt.Enabled = True
activetxt.Enabled = True
blocktxt.Enabled = True
createtxt.Enabled = True
updatetxt.Enabled = True
createdbytxt.Enabled = True
idtxt.SetFocus
newbtn.Caption = "NEWDATA"
End Sub

Private Sub savebtn_Click()
If newbtn.Caption = "NEWDATA" Then
If idtxt.Text = "" Or usertxt.Text = "" Or passtxt.Text = "" Or leveltxt.Text = "" Then
    MsgBox "Data Belum Lengkap"
    Else
    Dim tambahdata As String
        tambahdata = "Insert Into userid values ('" & idtxt.Text & "','" & usertxt.Text & "','" & leveltxt.Text & "','" & passtxt.Text & "','" & activetxt.Text & "' ,'" & blocktxt.Text & "' ,'" & createtxt.Text & "','" & updatetxt.Text & "','" & createdbytxt.Text & "')"
        koneksi.Execute tambahdata
        MsgBox "Data Berhasil Ditambah", vbInformation, "Pemberitahuan"
        newbtn.Caption = "NEW"
        Adodc1.Refresh
        DGuser.Refresh
        Call KondisiAwal
    
    End If
Else
If editbtn.Caption = "EDITDATA" Then
    Dim editdata As String
        editdata = "update userid set id_user = '" & idtxt.Text & "',usernames = '" & usertxt.Text & "',role_id = '" & leveltxt.Text & "',password = '" & passtxt.Text & "', active ='" & activetxt.Text & "' ,block = '" & blocktxt.Text & "' ,created_date = '" & createtxt.Text & "', updated_date = '" & updatetxt.Text & "',created_by = '" & createdbytxt.Text & "' where id_user = '" & idtxt.Text & "'"
        koneksi.Execute editdata
        MsgBox "Data Berhasil Diedit", vbInformation, "Pemberitahuan"
        editbtn.Caption = "EDIT"
        Adodc1.Refresh
        DGuser.Refresh
 Call KondisiAwal
End If
End If
End Sub

Private Sub cancelbtn_Click()
 Call Bersih
 newbtn.Enabled = True
 editbtn.Enabled = True
 deletebtn.Enabled = True
 newbtn.Caption = "NEW"
 editbtn.Caption = "EDIT"
cancelbtn.Enabled = False
savebtn.Enabled = False
End Sub

Private Sub editbtn_Click()
editbtn.Enabled = False
deletebtn.Enabled = False
savebtn.Enabled = True
newbtn.Enabled = False
cancelbtn.Enabled = True
idtxt.Enabled = True
usertxt.Enabled = True
passtxt.Enabled = True
leveltxt.Enabled = True
editbtn.Caption = "EDITDATA"
activetxt.Enabled = True
blocktxt.Enabled = True
createtxt.Enabled = True
updatetxt.Enabled = True
createdbytxt.Enabled = True
idtxt.Enabled = False
usertxt.Enabled = True
passtxt.Enabled = True
leveltxt.Enabled = True
activetxt.Enabled = True
blocktxt.Enabled = True
createtxt.Enabled = False
updatetxt.Enabled = True
createdbytxt.Enabled = False



End Sub

Private Sub exitbtn_Click()
koneksi.Close
Unload Me
End Sub
Sub KondisiAwal()
    idtxt.Text = ""
    passtxt.Text = ""
    usertxt.Text = ""
    leveltxt.Text = ""
    activetxt.Text = ""
    blocktxt.Text = ""
    createtxt.Text = ""
    updatetxt.Text = ""
    createdbytxt.Text = ""
    newbtn.Enabled = True
    editbtn.Enabled = True
    deletebtn.Enabled = True
DGuser.Refresh
End Sub



Private Sub Text1_Change()

End Sub
