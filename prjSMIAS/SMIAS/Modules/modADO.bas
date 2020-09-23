Attribute VB_Name = "modADO"
''*****************************************************************
'' File Name:
'' Purpose:
'' Required Files:
''
'' Programmer: Philip V. Naparan   E-mail: philipnaparan@yahoo.com
'' Date Created:
'' Last Modified:
'' Modified By:
'' Credits: NONE, ALL CODES ARE CODED BY Philip V. Naparan
''*****************************************************************

Option Explicit


'Function used to connect to database
'I created and use this function with SQL Server or Oracle only in client/server application together
'with CloseDB procedure.I use this function and procedure to connect only to
'the server if neccessary to save server resources (ex. When updating or displaying I use
'OpenDB and after the record displayed or update I use the closeDB).
'
'I DID NOT USE THE MAIN PURPOSE OF THIS CODDE WITH CloseDB BECAUSE
'THIS SYSTEM IS A STAND ALONE SYSTEM.
'
'--> This code is also available in VB.NET,J# and C# using ADO.NET. If you want it just e-mail me.
'
Public Function OpenDB() As Boolean
    Dim isOpen      As Boolean
    Dim ANS         As VbMsgBoxResult
    isOpen = False
    On Error GoTo err
        Do Until isOpen = True
                CN.CursorLocation = adUseClient
                CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DBPath & ";Persist Security Info=False;Jet OLEDB:Database Password=philiprj"
            isOpen = True
        Loop
        OpenDB = isOpen
    Exit Function
err:
    ANS = MsgBox("Error # " & err.Number & vbCrLf & "Description: " & err.Description, vbCritical + vbRetryCancel)
    If ANS = vbCancel Then
        OpenDB = False
    ElseIf ANS = vbRetry Then
        Resume
    End If
End Function

Public Sub CloseDB()
    'Close the connection
    CN.Close
    Set CN = Nothing
End Sub

'Function that return the current index for a certain table
Public Function getIndex(ByVal srcTable As String) As Long
    On Error GoTo err
    Dim rs As New Recordset
    Dim RI As Long
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM TBL_GENERATOR WHERE TableName = '" & srcTable & "'", CN, adOpenStatic, adLockOptimistic
    
    RI = rs.Fields("NextNo")
    rs.Fields("NextNo") = RI + 1
    rs.Update
    
    getIndex = RI
    
    srcTable = ""
    RI = 0
    Set rs = Nothing
    Exit Function
err:
        ''Error when incounter a null value
        If err.Number = 94 Then getIndex = 1: Resume Next
End Function

'Function used to get the sum  of fields
Public Function getSumOfFields(ByVal sTable As String, ByVal sField As String, ByRef sCN As ADODB.Connection, Optional inclField As String, Optional sCondition As String) As Double
    On Error GoTo err
    Dim rs As New ADODB.Recordset

    rs.CursorLocation = adUseClient
    If sCondition <> "" Then sCondition = " GROUP BY " & inclField & " HAVING(" & sCondition & ")"
    If inclField <> "" Then inclField = "," & inclField
    rs.Open "SELECT Sum(" & sTable & "." & sField & ") AS fTotal" & inclField & " FROM " & sTable & sCondition, sCN, adOpenStatic, adLockOptimistic
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Do While Not rs.EOF
            getSumOfFields = getSumOfFields + rs.Fields("fTotal")
            rs.MoveNext
        Loop
    Else
        getSumOfFields = 0
    End If
    
    Set rs = Nothing
    Exit Function
err:
        'Error when incounter a null value
        If err.Number = 94 Then getSumOfFields = 0: Resume Next
End Function

'Procedure used to generate DSN
Public Sub GenerateDSN()
Open App.Path & "\rptCN.dsn" For Output As #1
    Print #1, "[ODBC]"
    Print #1, "DRIVER=Microsoft Access Driver (*.mdb)"
    Print #1, "UID=admin"
    Print #1, "UserCommitSync=Yes"
    Print #1, "Threads=3"
    Print #1, "SafeTransactions=0"
    Print #1, "PageTimeout=5"
    Print #1, "MaxScanRows=8"
    Print #1, "MaxBufferSize=2048"
    Print #1, "FIL=MS Access"
    Print #1, "DriverId=25"
    Print #1, "DefaultDir=" & App.Path
    Print #1, "DBQ=" & App.Path & "\Db.mdb"
Close #1
End Sub

'Procedure used to remove DSN
Public Sub RemoveDSN()
On Error Resume Next
Kill App.Path & "\rptCN.dsn"
End Sub
