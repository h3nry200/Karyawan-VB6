VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmprint2 
   ClientHeight    =   960
   ClientLeft      =   11580
   ClientTop       =   10050
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   960
   ScaleWidth      =   4215
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      Format          =   132317185
      CurrentDate     =   42731
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CANCEL"
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   480
      Top             =   600
      Visible         =   0   'False
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1720
      _Version        =   393216
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
   Begin VB.Label Label1 
      Caption         =   "TGL INCOMING :"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmprint2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim koneksi As New rdoConnection

Private Sub Command1_Click()
If DataEnvironment2.rsCommand1.State = adStateOpen Then DataEnvironment2.rsCommand1.Close
DataEnvironment2.rsCommand1.Open "select * from tbl_dataadmin where incomingdate = '" & DTPicker1 & "'"
printdatareportadm3.Show
Unload Me

End Sub

Private Sub Form_Load()

With DataEnvironment2
End With

End Sub


'------------------------------------------------------------------------------------------------
'kl mau pakai adodc silakan apus tanda "'"
'Option Explicit

'Dim koneksi As New rdoConnection
'Dim rQuery As New rdoQuery
'Dim rs As rdoResultset



'Private Sub Form_Load()
'koneksi.Connect = "Driver={MySQL ODBC 5.1 Driver};SERVER=localhost;PWD=123;UID=root;PORT=3306;DATABASE=karyawan;"
'koneksi.EstablishConnection
'    Adodc1.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};SERVER=localhost;PWD=123;UID=root;PORT=3306;DATABASE=karyawan;"
'    Adodc1.RecordSource = "tbl_dataadmin"
'    Adodc1.Refresh
'Set DataGrid1.DataSource = Adodc1

'End Sub
'Private Sub Command1_Click()
'With Adodc1
'.CommandType = adCmdText
'.RecordSource = "select * from tbl_dataadmin where incomingdate = '" & DTPicker1 & "' "
''.RecordSource = "select * from absensi where tanggal like '%" & blncmb & "%' and tanggal like '%" & thncmb & "%'"
'.Refresh
'End With
'DataGrid1.Refresh
'printdatareportadm2.Show

'Set printdatareportadm2.DataSource = Adodc1

''koneksi
'printdatareportadm2.Sections("Section1").Controls.Item("LABEL5").Caption = DataGrid1.Columns(15).Text
'printdatareportadm2.Sections("Section2").Controls.Item("LABEL6").Caption = DataGrid1.Columns(16).Text
'printdatareportadm2.Sections("Section2").Controls.Item("LABEL8").Caption = DataGrid1.Columns(17).Text
'printdatareportadm2.Sections("Section2").Controls.Item("LABEL7").Caption = DataGrid1.Columns(18).Text
'printdatareportadm2.Sections("Section2").Controls.Item("LABEL9").Caption = DataGrid1.Columns(19).Text
'printdatareportadm2.Sections("Section2").Controls.Item("LABEL11").Caption = DTPicker1.Value
''printdatareportadm2.Sections("Section2").Controls.Item("LABEL41").Caption = DTPicker2.Value
'printdatareportadm2.Sections("Section1").Controls.Item("LABEL26").Caption = DataGrid1.Columns(0).Text
'printdatareportadm2.Sections("Section1").Controls.Item("LABEL27").Caption = DataGrid1.Columns(3).Text
'printdatareportadm2.Sections("Section1").Controls.Item("LABEL28").Caption = DataGrid1.Columns(2).Text
'printdatareportadm2.Sections("Section1").Controls.Item("LABEL29").Caption = DataGrid1.Columns(4).Text
'printdatareportadm2.Sections("Section1").Controls.Item("LABEL30").Caption = DataGrid1.Columns(5).Text
'printdatareportadm2.Sections("Section1").Controls.Item("LABEL31").Caption = DataGrid1.Columns(6).Text
'printdatareportadm2.Sections("Section1").Controls.Item("LABEL32").Caption = DataGrid1.Columns(7).Text
'printdatareportadm2.Sections("Section1").Controls.Item("LABEL33").Caption = DataGrid1.Columns(8).Text
'printdatareportadm2.Sections("Section1").Controls.Item("LABEL34").Caption = DataGrid1.Columns(9).Text
'printdatareportadm2.Sections("Section1").Controls.Item("LABEL35").Caption = DataGrid1.Columns(10).Text
'printdatareportadm2.Sections("Section1").Controls.Item("LABEL36").Caption = DataGrid1.Columns(11).Text
'printdatareportadm2.Sections("Section1").Controls.Item("LABEL37").Caption = DataGrid1.Columns(12).Text
'printdatareportadm2.Sections("Section1").Controls.Item("LABEL40").Caption = DataGrid1.Columns(14).Text

'End Sub
Private Sub Command2_Click()
'koneksi.Close
frmreportadm.print1.Visible = False
frmreportadm.print2.Visible = False
frmreportadm.print3.Visible = False

'DataEnvironment2.Connection1.Close

Unload Me

End Sub
