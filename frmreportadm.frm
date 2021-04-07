VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmreportadm 
   ClientHeight    =   10950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   WindowState     =   2  'Maximized
   Begin VB.TextBox jamfinishhmoff 
      Enabled         =   0   'False
      Height          =   375
      Left            =   18720
      TabIndex        =   68
      Top             =   8400
      Width           =   975
   End
   Begin VB.TextBox jamfinishoff 
      Enabled         =   0   'False
      Height          =   375
      Left            =   18720
      TabIndex        =   65
      Top             =   7800
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "....."
      Enabled         =   0   'False
      Height          =   375
      Left            =   9120
      TabIndex        =   63
      Top             =   8520
      Width           =   495
   End
   Begin VB.TextBox backupsurhmofftxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   62
      Top             =   8520
      Width           =   2205
   End
   Begin VB.TextBox backupsurofftxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   60
      Top             =   7920
      Width           =   2205
   End
   Begin VB.CommandButton Command5 
      Caption         =   "....."
      Enabled         =   0   'False
      Height          =   375
      Left            =   9120
      TabIndex        =   59
      Top             =   7920
      Width           =   495
   End
   Begin VB.TextBox jamfinishhm 
      Enabled         =   0   'False
      Height          =   375
      Left            =   18720
      TabIndex        =   58
      Top             =   7200
      Width           =   975
   End
   Begin VB.TextBox jamincome 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   57
      Top             =   4920
      Width           =   975
   End
   Begin VB.CommandButton print3 
      Caption         =   "Print Per surveyor"
      Height          =   375
      Left            =   10920
      TabIndex        =   56
      Top             =   9960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton print2 
      Caption         =   "Print Per Tanggal"
      Height          =   375
      Left            =   12360
      TabIndex        =   55
      Top             =   10560
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton print1 
      Caption         =   "Print Satuan"
      Height          =   375
      Left            =   11040
      TabIndex        =   54
      Top             =   10560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox usertxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   17760
      TabIndex        =   53
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton printbtn 
      Caption         =   "PRINT"
      Height          =   555
      Left            =   11760
      TabIndex        =   51
      Top             =   10200
      Width           =   885
   End
   Begin VB.CommandButton Command4 
      Caption         =   "....."
      Enabled         =   0   'False
      Height          =   375
      Left            =   9120
      TabIndex        =   50
      Top             =   7320
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "....."
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   49
      Top             =   8520
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "....."
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   48
      Top             =   7920
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "....."
      Enabled         =   0   'False
      Height          =   375
      Left            =   4440
      TabIndex        =   47
      Top             =   7320
      Width           =   495
   End
   Begin VB.TextBox ordertxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2040
      TabIndex        =   45
      Top             =   9720
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.ComboBox statushmoffcmb 
      Enabled         =   0   'False
      Height          =   315
      Left            =   12120
      TabIndex        =   44
      Top             =   8520
      Width           =   2535
   End
   Begin VB.ComboBox statusoffcmb 
      Enabled         =   0   'False
      Height          =   315
      Left            =   12120
      TabIndex        =   43
      Top             =   7920
      Width           =   2535
   End
   Begin VB.ComboBox statushmcmb 
      Enabled         =   0   'False
      Height          =   315
      Left            =   12120
      TabIndex        =   42
      Top             =   7320
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2160
      TabIndex        =   36
      Top             =   4920
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   202309633
      CurrentDate     =   42723
   End
   Begin VB.CommandButton seekbtn 
      Caption         =   "SEEK"
      Height          =   555
      Left            =   10680
      TabIndex        =   35
      Top             =   10200
      Width           =   885
   End
   Begin VB.CommandButton newbtn 
      Caption         =   "NEW"
      Height          =   555
      Left            =   5040
      TabIndex        =   34
      Top             =   10200
      Width           =   1005
   End
   Begin VB.CommandButton editbtn 
      Caption         =   "EDIT"
      Height          =   555
      Left            =   6240
      TabIndex        =   33
      Top             =   10200
      Width           =   975
   End
   Begin VB.CommandButton deletebtn 
      Caption         =   "DELETE"
      Height          =   555
      Left            =   7440
      TabIndex        =   32
      Top             =   10200
      Width           =   885
   End
   Begin VB.CommandButton savebtn 
      Caption         =   "SAVE"
      Height          =   555
      Left            =   8520
      TabIndex        =   31
      Top             =   10200
      Width           =   885
   End
   Begin VB.CommandButton exitbtn 
      Caption         =   "EXIT"
      Height          =   555
      Left            =   12840
      TabIndex        =   30
      Top             =   10200
      Width           =   885
   End
   Begin VB.CommandButton cancelbtn 
      Caption         =   "CANCEL"
      Height          =   555
      Left            =   9600
      TabIndex        =   29
      Top             =   10200
      Width           =   885
   End
   Begin VB.TextBox lastuptxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   12600
      TabIndex        =   12
      Top             =   3480
      Width           =   2835
   End
   Begin VB.TextBox idtxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2100
      TabIndex        =   11
      Top             =   4380
      Width           =   2445
   End
   Begin VB.TextBox tglcretxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1650
      TabIndex        =   10
      Top             =   3480
      Width           =   2745
   End
   Begin VB.Timer Timer1 
      Left            =   4860
      Top             =   120
   End
   Begin VB.TextBox namaaplikantxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2100
      TabIndex        =   9
      Top             =   6720
      Width           =   3045
   End
   Begin VB.TextBox alasantxt 
      Enabled         =   0   'False
      Height          =   855
      Left            =   11340
      TabIndex        =   8
      Top             =   9120
      Visible         =   0   'False
      Width           =   6525
   End
   Begin VB.TextBox producttxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2100
      TabIndex        =   7
      Top             =   5520
      Width           =   3045
   End
   Begin VB.TextBox noapltxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2100
      TabIndex        =   6
      Top             =   6120
      Width           =   3045
   End
   Begin VB.TextBox surhmtxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2100
      TabIndex        =   5
      Top             =   7320
      Width           =   2205
   End
   Begin VB.TextBox surofftxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2100
      TabIndex        =   4
      Top             =   7920
      Width           =   2205
   End
   Begin VB.TextBox surhmofftxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2100
      TabIndex        =   3
      Top             =   8520
      Width           =   2205
   End
   Begin VB.TextBox backupsurhmtxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Top             =   7320
      Width           =   2205
   End
   Begin VB.TextBox jamcrttxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4380
      TabIndex        =   1
      Top             =   3480
      Width           =   945
   End
   Begin VB.TextBox jamlasttxt 
      Enabled         =   0   'False
      Height          =   375
      Left            =   15480
      TabIndex        =   0
      Top             =   3480
      Width           =   1065
   End
   Begin MSDataGridLib.DataGrid DGkary 
      Height          =   3015
      Left            =   0
      TabIndex        =   13
      Top             =   360
      Width           =   19785
      _ExtentX        =   34899
      _ExtentY        =   5318
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
      Left            =   2700
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
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   16320
      TabIndex        =   37
      Top             =   7200
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   73793537
      CurrentDate     =   42723
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   375
      Left            =   16320
      TabIndex        =   66
      Top             =   7800
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   73793537
      CurrentDate     =   42723
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   375
      Left            =   16320
      TabIndex        =   69
      Top             =   8400
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   73793537
      CurrentDate     =   42723
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   3240
      Top             =   2760
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
      Caption         =   "Adodc2"
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
   Begin VB.Label Label21 
      Caption         =   "Finish DateHm+Off :"
      Height          =   225
      Left            =   14760
      TabIndex        =   70
      Top             =   8520
      Width           =   1515
   End
   Begin VB.Label Label20 
      Caption         =   "Finish DateOff        :"
      Height          =   225
      Left            =   14760
      TabIndex        =   67
      Top             =   7920
      Width           =   1515
   End
   Begin VB.Label Label19 
      Caption         =   "Back up Sur Hm + Off   :"
      Height          =   225
      Left            =   5040
      TabIndex        =   64
      Top             =   8640
      Width           =   1755
   End
   Begin VB.Label Label18 
      Caption         =   "Back up Sur Office        :"
      Height          =   225
      Left            =   5040
      TabIndex        =   61
      Top             =   8040
      Width           =   1755
   End
   Begin VB.Label Label17 
      Caption         =   "USER  :"
      Height          =   375
      Left            =   17040
      TabIndex        =   52
      Top             =   3480
      Width           =   735
   End
   Begin VB.Label Label16 
      Caption         =   "Order By Branch            :"
      Height          =   225
      Left            =   240
      TabIndex        =   46
      Top             =   9840
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Label Label11 
      Caption         =   "Status Home = Office        :"
      Height          =   225
      Left            =   10080
      TabIndex        =   41
      Top             =   8640
      Width           =   1995
   End
   Begin VB.Label Label10 
      Caption         =   "Status Office                 :"
      Height          =   225
      Left            =   10320
      TabIndex        =   40
      Top             =   8040
      Width           =   1755
   End
   Begin VB.Label Label4 
      Caption         =   "Status Home                 :"
      Height          =   225
      Left            =   10320
      TabIndex        =   39
      Top             =   7440
      Width           =   1755
   End
   Begin VB.Label Label14 
      Caption         =   "Surveyor Home = Office    :"
      Height          =   225
      Left            =   60
      TabIndex        =   38
      Top             =   8640
      Width           =   1995
   End
   Begin VB.Label Label23 
      Caption         =   "LAST UPDATE  :"
      Height          =   315
      Left            =   11340
      TabIndex        =   28
      Top             =   3540
      Width           =   1635
   End
   Begin VB.Label Label2 
      Caption         =   "TANGGAL CREATE  :"
      Height          =   255
      Left            =   0
      TabIndex        =   27
      Top             =   3540
      Width           =   1755
   End
   Begin VB.Label Label1 
      Caption         =   "ID      :"
      Height          =   225
      Left            =   1380
      TabIndex        =   26
      Top             =   4440
      Width           =   585
   End
   Begin VB.Label Label3 
      Caption         =   "Incoming Date     :"
      Height          =   225
      Left            =   540
      TabIndex        =   25
      Top             =   5040
      Width           =   1425
   End
   Begin VB.Label lbwaktu 
      Height          =   375
      Left            =   16020
      TabIndex        =   24
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label lbuser 
      Height          =   255
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   5895
   End
   Begin VB.Label lbjam 
      Height          =   375
      Left            =   18180
      TabIndex        =   22
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Finish Date Home   :"
      Height          =   225
      Left            =   14760
      TabIndex        =   21
      Top             =   7320
      Width           =   1515
   End
   Begin VB.Label Label6 
      Caption         =   "Nama Aplikan     :"
      Height          =   225
      Left            =   660
      TabIndex        =   20
      Top             =   6840
      Width           =   1395
   End
   Begin VB.Label Label7 
      Caption         =   "Alasan                 :"
      Height          =   225
      Left            =   9900
      TabIndex        =   19
      Top             =   9240
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Label8 
      Caption         =   "Product                :"
      Height          =   225
      Left            =   540
      TabIndex        =   18
      Top             =   5640
      Width           =   1515
   End
   Begin VB.Label Label9 
      Caption         =   "Nomor Aplikasi     :"
      Height          =   225
      Left            =   600
      TabIndex        =   17
      Top             =   6240
      Width           =   1395
   End
   Begin VB.Label Label12 
      Caption         =   "Surveyor Home    :"
      Height          =   225
      Left            =   660
      TabIndex        =   16
      Top             =   7440
      Width           =   1395
   End
   Begin VB.Label Label13 
      Caption         =   "Surveyor Office    :"
      Height          =   225
      Left            =   660
      TabIndex        =   15
      Top             =   8040
      Width           =   1395
   End
   Begin VB.Label Label15 
      Caption         =   "Back up Sur Home        :"
      Height          =   225
      Left            =   5040
      TabIndex        =   14
      Top             =   7440
      Width           =   1755
   End
End
Attribute VB_Name = "frmreportadm"
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
frmseeksur1.Show
frmseeksur1.Text1.SetFocus
End Sub

Private Sub Command2_Click()
frmseeksur2.Show
frmseeksur2.Text1.SetFocus
End Sub

Private Sub Command3_Click()
frmseeksur3.Show
frmseeksur3.Text1.SetFocus
End Sub

Private Sub Command4_Click()
frmseeksur4.Show
frmseeksur4.Text1.SetFocus
End Sub

Private Sub Command6_Click()
'Frame1.Visible = False
End Sub


Private Sub Form_Load()

Call Bersih
    
    koneksi.Connect = "Driver={MySQL ODBC 5.1 Driver};SERVER=localhost;PWD=123;UID=root;PORT=3306;DATABASE=karyawan;"
    koneksi.EstablishConnection
    Adodc1.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};SERVER=localhost;PWD=123;UID=root;PORT=3306;DATABASE=karyawan;"
    Adodc1.RecordSource = "tbl_dataadmin"
    Adodc1.Refresh
    Adodc2.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};SERVER=localhost;PWD=123;UID=root;PORT=3306;DATABASE=karyawan;"
    Adodc2.RecordSource = "tbl_dataadmin"
    Adodc2.Refresh


Set DGkary.DataSource = Adodc1
cancelbtn.Enabled = False

  ahari = Array("Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu")
  Timer1.Interval = 500
  Timer1.Enabled = True
'lbuser.Caption = "' & usertxt.Text & '"
lbuser.Caption = mdihalutama.Caption

    statushmcmb.AddItem "SUKSESS"
    statushmcmb.AddItem "PENDING"
    statushmcmb.AddItem "RETURN"
    statushmcmb.AddItem "NONE"
    
    statusoffcmb.AddItem "SUKSESS"
    statusoffcmb.AddItem "PENDING"
    statusoffcmb.AddItem "RETURN"
    statusoffcmb.AddItem "NONE"
    
    statushmoffcmb.AddItem "SUKSESS"
    statushmoffcmb.AddItem "PENDING"
    statushmoffcmb.AddItem "RETURN"
    statushmoffcmb.AddItem "NONE"

End Sub

Private Sub print1_Click()
printdatareportadm.Show
print1.Visible = False
print2.Visible = False
print3.Visible = False
Set printdatareportadm.DataSource = Adodc2

'koneksi
printdatareportadm.Sections("Section2").Controls.Item("LABEL5").Caption = usertxt.Text
printdatareportadm.Sections("Section2").Controls.Item("LABEL6").Caption = tglcretxt.Text
printdatareportadm.Sections("Section2").Controls.Item("LABEL8").Caption = jamcrttxt.Text
printdatareportadm.Sections("Section2").Controls.Item("LABEL7").Caption = lastuptxt.Text
printdatareportadm.Sections("Section2").Controls.Item("LABEL9").Caption = jamlasttxt.Text
printdatareportadm.Sections("Section2").Controls.Item("LABEL11").Caption = DTPicker1.Value

'printdatareportadm.Sections("Section2").Controls.Item("LABEL26").Caption = "1"
printdatareportadm.Sections("Section2").Controls.Item("LABEL27").Caption = producttxt.Text
'printdatareportadm.Sections("Section2").Controls.Item("LABEL28").Caption = DTPicker2.Value
printdatareportadm.Sections("Section2").Controls.Item("LABEL29").Caption = noapltxt.Text
printdatareportadm.Sections("Section2").Controls.Item("LABEL30").Caption = namaaplikantxt.Text
printdatareportadm.Sections("Section2").Controls.Item("LABEL31").Caption = surhmtxt.Text
printdatareportadm.Sections("Section2").Controls.Item("LABEL32").Caption = surofftxt.Text
printdatareportadm.Sections("Section2").Controls.Item("LABEL33").Caption = surhmofftxt.Text

printdatareportadm.Sections("Section2").Controls.Item("LABEL26").Caption = backupsurhmtxt.Text
printdatareportadm.Sections("Section2").Controls.Item("LABEL28").Caption = backupsurofftxt.Text
printdatareportadm.Sections("Section2").Controls.Item("LABEL34").Caption = backupsurhmofftxt.Text

printdatareportadm.Sections("Section2").Controls.Item("LABEL42").Caption = DTPicker2.Value
printdatareportadm.Sections("Section2").Controls.Item("LABEL43").Caption = DTPicker3.Value
printdatareportadm.Sections("Section2").Controls.Item("LABEL44").Caption = DTPicker4.Value

printdatareportadm.Sections("Section2").Controls.Item("LABEL45").Caption = jamfinishhm.Text
printdatareportadm.Sections("Section2").Controls.Item("LABEL46").Caption = jamfinishoff.Text
printdatareportadm.Sections("Section2").Controls.Item("LABEL47").Caption = jamfinishhmoff.Text

'printdatareportadm.Sections("Section2").Controls.Item("LABEL34").Caption = backupsurtxt.Text
printdatareportadm.Sections("Section2").Controls.Item("LABEL35").Caption = statushmcmb.Text
printdatareportadm.Sections("Section2").Controls.Item("LABEL36").Caption = statusoffcmb.Text
printdatareportadm.Sections("Section2").Controls.Item("LABEL37").Caption = statushmoffcmb.Text
'printdatareportadm.Sections("Section2").Controls.Item("LABEL40").Caption = alasantxt.Text

'printdatareportadm.Sections("Section4").Controls.Item("LABEL13").Caption = statuscmb.Text
'printdatareportadm.Sections("Section4").Controls.Item("LABEL15").Caption = kettxt.Text
'printdatareportadm.Sections("Section4").Controls.Item("LABEL17").Caption = namatxt.Text

End Sub

Private Sub print2_Click()
frmprint2.Show
End Sub

Private Sub print3_Click()
frmprint3.Show
End Sub

Private Sub printbtn_Click()
print1.Visible = True
print2.Visible = True
print3.Visible = True
End Sub

Private Sub seekbtn_Click()
frmseekreportadm.Show
frmseekreportadm.Text1.SetFocus
End Sub

Private Sub Timer1_Timer()
 shari = ahari(Abs(Weekday(Date) - 1))
  lbwaktu.Caption = "" & shari & ", " _
                   & Format(Date, "dd mmmm yyyy")
lbjam.Caption = Format(Time, "hh:mm:ss")
End Sub

Sub Bersih()
tglcretxt.Text = ""
jamcrttxt.Text = ""
lastuptxt.Text = ""
jamlasttxt.Text = ""
idtxt.Text = ""
jamincome.Text = ""
producttxt.Text = ""
noapltxt.Text = ""
namaaplikantxt.Text = ""
surhmtxt.Text = ""
surofftxt.Text = ""
surhmofftxt.Text = ""
backupsurhmtxt.Text = ""
backupsurofftxt.Text = ""
backupsurhmofftxt.Text = ""
ordertxt.Text = ""
alasantxt.Text = ""
usertxt.Text = ""
jamfinishhm.Text = ""
jamfinishoff.Text = ""
jamfinishhmoff.Text = ""
statushmcmb.Text = ""
statusoffcmb.Text = ""
statushmoffcmb.Text = ""

End Sub

Sub KondisiAwal()

tglcretxt.Text = ""
jamcrttxt.Text = ""
lastuptxt.Text = ""
jamlasttxt.Text = ""
idtxt.Text = ""
jamincome.Text = ""
producttxt.Text = ""
noapltxt.Text = ""
namaaplikantxt.Text = ""
surhmtxt.Text = ""
surofftxt.Text = ""
surhmofftxt.Text = ""
backupsurhmtxt.Text = ""
backupsurofftxt.Text = ""
backupsurhmofftxt.Text = ""
ordertxt.Text = ""
alasantxt.Text = ""
usertxt.Text = ""
jamfinishhm.Text = ""
jamfinishoff.Text = ""
jamfinishhmoff.Text = ""
statushmcmb.Text = ""
statusoffcmb.Text = ""
statushmoffcmb.Text = ""

idtxt.Enabled = False
tglcretxt.Enabled = False
jamcrttxt.Enabled = False
lastuptxt.Enabled = False
jamlasttxt.Enabled = False
DTPicker1.Enabled = False
DTPicker2.Enabled = False
DTPicker3.Enabled = False
DTPicker4.Enabled = False
producttxt.Enabled = False
noapltxt.Enabled = False
namaaplikantxt.Enabled = False
surhmtxt.Enabled = False
surofftxt.Enabled = False
surhmofftxt.Enabled = False
backupsurhmtxt.Enabled = False
backupsurofftxt.Enabled = False
backupsurhmofftxt.Enabled = False
ordertxt.Enabled = False
alasantxt.Enabled = False
statushmcmb.Enabled = False
statusoffcmb.Enabled = False
statushmoffcmb.Enabled = False
jamincome.Enabled = False
jamfinishhm.Enabled = False
jamfinishoff.Enabled = False
jamfinishhmoff.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
usertxt.Enabled = False

newbtn.Enabled = True
editbtn.Enabled = True
deletebtn.Enabled = True
savebtn.Enabled = False
cancelbtn.Enabled = False
seekbtn.Enabled = True
exitbtn.Enabled = True
DGkary.Refresh
End Sub

Private Sub DGkary_Click()

tglcretxt.Text = DGkary.Columns(22).Text
jamcrttxt.Text = DGkary.Columns(23).Text
lastuptxt.Text = DGkary.Columns(24).Text
jamlasttxt.Text = DGkary.Columns(25).Text
usertxt.Text = DGkary.Columns(21).Text
idtxt.Text = DGkary.Columns(0).Text
DTPicker1.Value = DGkary.Columns(1).Text
jamincome = DGkary.Columns(2).Text
DTPicker2.Value = DGkary.Columns(6).Text
jamfinishhm = DGkary.Columns(7).Text
DTPicker3.Value = DGkary.Columns(8).Text
jamfinishoff = DGkary.Columns(9).Text
DTPicker4.Value = DGkary.Columns(10).Text
jamfinishhmoff = DGkary.Columns(11).Text
producttxt.Text = DGkary.Columns(3).Text
noapltxt.Text = DGkary.Columns(4).Text
namaaplikantxt.Text = DGkary.Columns(5).Text
surhmtxt.Text = DGkary.Columns(12).Text
backupsurhmtxt = DGkary.Columns(13).Text
surofftxt.Text = DGkary.Columns(14).Text
backupsurofftxt = DGkary.Columns(15).Text
surhmofftxt.Text = DGkary.Columns(16).Text
backupsurhmofftxt = DGkary.Columns(17).Text
ordertxt.Text = DGkary.Columns(13).Text
alasantxt.Text = DGkary.Columns(14).Text
statushmcmb.Text = DGkary.Columns(18).Text
statusoffcmb.Text = DGkary.Columns(19).Text
statushmoffcmb.Text = DGkary.Columns(20).Text
End Sub

Private Sub DGkary_KeyDown(KeyCode As Integer, Shift As Integer)
tglcretxt.Text = DGkary.Columns(22).Text
jamcrttxt.Text = DGkary.Columns(23).Text
lastuptxt.Text = DGkary.Columns(24).Text
jamlasttxt.Text = DGkary.Columns(25).Text
usertxt.Text = DGkary.Columns(21).Text
idtxt.Text = DGkary.Columns(0).Text
DTPicker1.Value = DGkary.Columns(1).Text
jamincome = DGkary.Columns(2).Text
DTPicker2.Value = DGkary.Columns(6).Text
jamfinishhm = DGkary.Columns(7).Text
DTPicker3.Value = DGkary.Columns(8).Text
jamfinishoff = DGkary.Columns(9).Text
DTPicker4.Value = DGkary.Columns(10).Text
jamfinishhmoff = DGkary.Columns(11).Text
producttxt.Text = DGkary.Columns(3).Text
noapltxt.Text = DGkary.Columns(4).Text
namaaplikantxt.Text = DGkary.Columns(5).Text
surhmtxt.Text = DGkary.Columns(12).Text
backupsurhmtxt = DGkary.Columns(13).Text
surofftxt.Text = DGkary.Columns(14).Text
backupsurofftxt = DGkary.Columns(15).Text
surhmofftxt.Text = DGkary.Columns(16).Text
backupsurhmofftxt = DGkary.Columns(17).Text
ordertxt.Text = DGkary.Columns(13).Text
alasantxt.Text = DGkary.Columns(14).Text
statushmcmb.Text = DGkary.Columns(18).Text
statusoffcmb.Text = DGkary.Columns(19).Text
statushmoffcmb.Text = DGkary.Columns(20).Text
End Sub

Private Sub newbtn_Click()
Call KondisiAwal
editbtn.Enabled = False
deletebtn.Enabled = False
savebtn.Enabled = True
newbtn.Enabled = False
seekbtn.Enabled = False
cancelbtn.Enabled = True
idtxt.Enabled = True
DTPicker1.Enabled = True
DTPicker2.Enabled = True
DTPicker3.Enabled = True
DTPicker4.Enabled = True
producttxt.Enabled = True
noapltxt.Enabled = True
namaaplikantxt.Enabled = True
surhmtxt.Enabled = True
surofftxt.Enabled = True
surhmofftxt.Enabled = True
backupsurhmtxt.Enabled = True
backupsurofftxt.Enabled = True
backupsurhmofftxt.Enabled = True
ordertxt.Enabled = True
alasantxt.Enabled = True
statushmcmb.Enabled = True
statusoffcmb.Enabled = True
statushmoffcmb.Enabled = True
jamincome.Enabled = True
jamfinishhm.Enabled = True
jamfinishoff.Enabled = True
jamfinishhmoff.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True
tglcretxt.Text = lbwaktu.Caption
jamcrttxt.Text = lbjam.Caption
lastuptxt.Text = lbwaktu.Caption
jamlasttxt.Text = lbjam.Caption
usertxt.Text = lbuser.Caption
idtxt.SetFocus
newbtn.Caption = "NEWDATA"
End Sub

Private Sub editbtn_Click()
editbtn.Enabled = False
deletebtn.Enabled = False
savebtn.Enabled = True
newbtn.Enabled = False
cancelbtn.Enabled = True
seekbtn.Enabled = False
editbtn.Enabled = False
deletebtn.Enabled = False
savebtn.Enabled = True
newbtn.Enabled = False
cancelbtn.Enabled = True
idtxt.Enabled = True
DTPicker1.Enabled = True
DTPicker2.Enabled = True
DTPicker3.Enabled = True
DTPicker4.Enabled = True
producttxt.Enabled = True
noapltxt.Enabled = True
namaaplikantxt.Enabled = True
surhmtxt.Enabled = True
surofftxt.Enabled = True
surhmofftxt.Enabled = True
backupsurhmtxt.Enabled = True
backupsurofftxt.Enabled = True
backupsurhmofftxt.Enabled = True
ordertxt.Enabled = True
alasantxt.Enabled = True
statushmcmb.Enabled = True
statusoffcmb.Enabled = True
statushmoffcmb.Enabled = True
jamincome.Enabled = True
jamfinishhm.Enabled = True
jamfinishoff.Enabled = True
jamfinishhmoff.Enabled = True
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
Command6.Enabled = True

'idtxt.SetFocus
lastuptxt.Text = lbwaktu.Caption
jamlasttxt.Text = lbjam.Caption
usertxt.Text = lbuser.Caption
editbtn.Caption = "EDITDATA"

End Sub

Private Sub savebtn_Click()
If newbtn.Caption = "NEWDATA" Then
If idtxt.Text = "" Or producttxt.Text = "" Or noapltxt.Text = "" Or namaaplikantxt.Text = "" Or surhmtxt.Text = "" Or surofftxt.Text = "" Or surhmofftxt.Text = "" Then
    MsgBox "Data Belum Lengkap"
    Else
    Dim tambahdata As String
        tambahdata = "Insert Into tbl_dataadmin values ('" & idtxt.Text & "', '" & DTPicker1.Value & "','" & jamincome.Text & "','" & producttxt.Text & "','" & noapltxt.Text & "','" & namaaplikantxt.Text & "','" & DTPicker2.Value & "','" & jamfinishhm.Text & "','" & DTPicker3.Value & "','" & jamfinishoff.Text & "','" & DTPicker4.Value & "','" & jamfinishhmoff.Text & "','" & surhmtxt.Text & "','" & backupsurhmtxt.Text & "','" & surofftxt.Text & "','" & backupsurofftxt.Text & "','" & surhmofftxt.Text & "','" & backupsurhmofftxt.Text & "','" & statushmcmb.Text & "','" & statusoffcmb.Text & "','" & statushmoffcmb.Text & "','" & usertxt.Text & "','" & tglcretxt.Text & "','" & jamcrttxt.Text & "','" & lastuptxt.Text & "','" & jamlasttxt.Text & "')"
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
        editdata = "update tbl_dataadmin set id = '" & idtxt.Text & "', incomingdate = '" & DTPicker1.Value & "', jamincome = '" & jamincome.Text & "', product = '" & producttxt.Text & "', noaplikasi = '" & noapltxt.Text & "', namaaplikan = '" & namaaplikantxt.Text & "',finishhm = '" & DTPicker2.Value & "',jamfinishhm = '" & jamfinishhm.Text & "',finishoff = '" & DTPicker3.Value & "',jamfinishoff = '" & jamfinishoff.Text & "',finishhmoff = '" & DTPicker4.Value & "', jamfinishhmoff ='" & jamfinishhmoff.Text & "',surhome = '" & surhmtxt.Text & "',backuphm = '" & backupsurhmtxt.Text & "', " _
        & " suroffice = '" & surofftxt.Text & "', backupoff = '" & backupsurofftxt.Text & "',surhomeandoff = '" & surhmofftxt.Text & "',backuphmoff = '" & backupsurhmofftxt.Text & "',statushome = '" & statushmcmb.Text & "',statusoffice = '" & statusoffcmb.Text & "', statushomeandoff = '" & statushmoffcmb.Text & "',user = '" & usertxt.Text & "',tglinput = '" & tglcretxt.Text & "', jaminput = '" & jamcrttxt.Text & "',lastupdate = '" & lastuptxt.Text & "', jamlastup = '" & jamlasttxt.Text & "' where id = '" & idtxt.Text & "'"
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
print1.Visible = False
print2.Visible = False
print3.Visible = False
 
 newbtn.Caption = "NEW"
 editbtn.Caption = "EDIT"
End Sub

Private Sub deletebtn_Click()
If MsgBox("Yakin Ingin Menghapus Data?", vbQuestion + vbOKCancel, "konfirmasi") = vbOK Then
Dim hapusdata As String
        hapusdata = "delete from tbl_dataadmin where id =" & idtxt.Text & ""
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



