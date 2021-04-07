Attribute VB_Name = "conn"

Option Explicit
'Global CNDB, sqldb, sqldbX As New ADODB.Connection
'Public CNDB As ADODB.Connection
'Public CNDB2 As ADODB.Connection
'Public sqldb As ADODB.Connection
'Public sqldbX As ADODB.Connection
Public conXls As ADODB.Connection
Public Conn As New ADODB.Connection
Public strkoneksi As String
Public konn As New ADODB.Connection
Public rsUser As New ADODB.Recordset

 
Public Function koneksi() As Boolean
'settingan koneksi
On Error GoTo er
'koneksi string ke mysql konektor
strkoneksi = "DRIVER={MySQL ODBC 5.1 Driver};SERVER=localhost;PWD=;UID=root;PORT=3306;DATABASE=karyawan;"

If konn.State = adStateOpen Then Conn.Close

konn.Open strkoneksi
konn.CursorLocation = adUseClient

'buka tabel database
 '   rsUser.Open "SELECT usernames,level,pids FROM userid", strkoneksi, adOpenKeyset, adLockOptimistic

If konn.State = adStateOpen Then
    koneksi = True
    Exit Function
Else
    koneksi = False
    Exit Function

End If

Exit Function

er:
 koneksi = False
 MsgBox "Gagal Loading Database", vbInformation, "Database Error"
End Function
 
'Function Ini di gunakan untuk koneksi ke file excel
Public Function openExcelFile(ByVal excelFile As String) As Boolean

    On Error GoTo errHandle
    '-----------------------
    'Jika menggunakan Office 2007 ke atas ganti Provider=Microsoft.Jet.OLEDB.4.0;
    'menjadi Provider=Microsoft.ACE.OLEDB.12.0;
   
    Set conXls = New ADODB.Connection
    conXls.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" _
                            & "Data Source=" & Replace(excelFile, Chr$(0), "") & ";" _
                            & "Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""
    conXls.Open
    '-------------------
    openExcelFile = True
    Exit Function
errHandle:
    openExcelFile = False
End Function

'-------------------------------------------------------------------------------------------
