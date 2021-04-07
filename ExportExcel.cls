VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExportExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Public Property Let RecordSource(m_rs As Recordset)
Set rs = m_rs
End Property
Public Sub ExportToExcel()
Dim App As Application
Dim Wb As Workbook
Dim Wk As Worksheet
Dim Baris As Long, Kolom As Integer
Set App = Excel.Application
Set Wb = App.Workbooks.Add
Set Wk = Wb.Worksheets(1)
With rs
Baris = 1
Kolom = 0
'mencetak Header
For Kolom = 1 To rs.Fields.Count
Wk.Cells(1, Kolom) = rs.Fields(Kolom–1).Name
Next Kolom
'mencetak semua data
.MoveFirst
While Not .EOF
Baris = Baris + 1
For Kolom = 1 To rs.Fields.Count
Wk.Cells(Baris, Kolom) = rs.Fields(Kolom–1)
Next Kolom
.MoveNext
Wend
End With
App.Visible = True
Set Wk = Nothing
Set Wb = Nothing
Set App = Nothing
End Sub
