Attribute VB_Name = "Module1"
Option Explicit

Private Const ODBC_ADD_SYS_DSN = 4      'Constant for Adding the DSN
Private Const ODBC_REMOVE_SYS_DSN = 6   'Constant for removing the DSN

Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hwndParent As Long, ByVal fRequest As Long, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
   
'Below is the Function for Adding or Deleting the DSN
'These are it's parameters

'DoWhat
    'If you Choose "Add" it Adds the DSN
    'If you Choose "Del" it Deletes the DSN

'DriverName
    'The DataBase Drivers Name
'Note: You can find more in the ODBC/DSN Just copy and past into the Combo1 List

'DSNName
    'What you will be calling on later on in your program
    'EXAMPLE:

    'Public rs As New ADODB.Recordset
    'Public cnn As New ADODB.Connection

    'Function OpenDB(DSNName As String, RSetName As String)

        'If cnn.State = 1 Then cnn.Close
            
            'cnn.Open "DSN=  DSNName  ;UID=Admin;PWD="
            'sqlstmt = "SELECT * from " & RSetName
            'rs.Open sqlstmt, cnn, 3, 3

    'End Function

'NOTE: RSetName = The Records Set Name In the DataBase

'DataBasesName = The Full Path and Name of your database file
    'In this case I have supplied a Demo.mdb

'Description
    'A Brief Description of your program

Public Function DSN(DoWhat As String, DriverName As String, DSNName As String, DataBasesName As String, Description As String)
Dim Added As Byte

If DoWhat = "Add" Then
    Added = SQLConfigDataSource(0, ODBC_ADD_SYS_DSN, DriverName & Chr(0), "DSN=" & DSNName & Chr(0) & "Uid=Admin" & Chr(0) & "pwd=" & Chr(0) & "DBQ=" & DataBasesName & Chr(0) & "Description=" & Description & Chr(0))

If Added = 0 Then     'Error code 0 means something went wrong and the SQLConfigDataSource Never finished
    MsgBox "This program may not work on this PC. Contact the vendor with Error Code: 275", vbOKOnly, "Error"
ElseIf Added = 1 Then 'Error code 1 means SQLConfigDataSource finished successfully
    MsgBox "DSN Added Successfully!", vbOKOnly, "Success!"
End If

ElseIf DoWhat = "Del" Then
    Added = SQLConfigDataSource(0, ODBC_REMOVE_SYS_DSN, DriverName & Chr(0), "DSN=" & DSNName & Chr(0) & "Uid=Admin" & Chr(0) & "pwd=" & Chr(0) & "DBQ=" & DataBasesName & Chr(0) & "Description=" & Description & Chr(0))

If Added = 0 Then
    MsgBox "Error Deleting DSN", vbOKOnly, "Error"
ElseIf Added = 1 Then
    MsgBox "DSN Deleted Successfully!", vbOKOnly, "Success!"
End If

End If

End Function
