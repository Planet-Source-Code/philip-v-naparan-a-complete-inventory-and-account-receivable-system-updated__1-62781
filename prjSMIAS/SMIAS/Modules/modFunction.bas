Attribute VB_Name = "modFunction"
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




'Function used to format recordset
Public Function FormatRS(ByVal srcField As Field, Optional AllowNewLine As Boolean) As String
    Dim strRet As String
    
    With srcField
        If AllowNewLine = True Then
            strRet = srcField
        Else
            strRet = Replace(srcField, vbCrLf, " ", , , vbTextCompare)
        End If
        
        If srcField.Type = adCurrency Or srcField.Type = adDouble Then
            strRet = Format$(srcField, "#,##0.00")
        ElseIf srcField.Type = adDate Then
            strRet = Format$(srcField, "MMM-dd-yyyy")
        Else
            strRet = srcField
        End If
    End With
    
    FormatRS = strRet
    
    strRet = vbNullString
End Function

'Function that will format return a generated id
Public Function GenerateID(ByVal srcNo As String, ByVal src1stStr As String, ByVal src2ndStr As String) As String
    If Len(src2ndStr) <= Len(srcNo) Then
        GenerateID = src1stStr & srcNo
    Else
        GenerateID = src1stStr & Left$(src2ndStr, Len(src2ndStr) - Len(srcNo)) & srcNo
    End If
End Function

'Function used to check if the record exit or not.
Public Function isRecordExist(ByVal sTable As String, ByVal sField As String, ByVal sStr As String, Optional isString As Boolean) As Boolean
    Dim rs As New Recordset

    rs.CursorLocation = adUseClient
    If isString = False Then
        rs.Open "Select * From " & sTable & " Where " & sField & " = " & sStr, CN, adOpenStatic, adLockOptimistic
    Else
        rs.Open "Select * From " & sTable & " Where " & sField & " = '" & sStr & "'", CN, adOpenStatic, adLockOptimistic
    End If
    If rs.RecordCount < 1 Then
        isRecordExist = False
    Else
        isRecordExist = True
    End If
    Set rs = Nothing
End Function

'Function used to check if the Ascii is a number or not (return 0 if number)
Public Function isNumber(ByVal sKeyAscii) As Integer
    If Not ((sKeyAscii >= 48 And sKeyAscii <= 57) Or sKeyAscii = 8 Or sKeyAscii = 46) Then
        isNumber = 0
    Else
        isNumber = sKeyAscii
    End If
End Function

'Function used to check if the record exist in Flex grid
Public Function isRecExistInFlex(ByVal srcFlexGrd As MSHFlexGrid, ByVal srcWhatCol As Integer, ByVal srcFindWhat As String) As Boolean
    isRecExistInFlex = False
    Dim i As Long
    For i = 1 To srcFlexGrd.Rows - 1
        If srcFlexGrd.TextMatrix(i, srcWhatCol) = srcFindWhat Then isRecExistInFlex = True: Exit For
    Next i
    i = 0
End Function

'Function used to check if the record exist in Flex grid
Public Function getFlexPos(ByVal srcFlexGrd As MSHFlexGrid, ByVal srcWhatCol As Integer, ByVal srcFindWhat As String) As Integer
    Dim r As Long, ret As Integer
    
    ret = -1 'Means not found
    For r = 0 To srcFlexGrd.Rows - 1
        If srcFlexGrd.TextMatrix(r, srcWhatCol) = srcFindWhat Then ret = r: Exit For
    Next r
    
    getFlexPos = ret
    r = 0: ret = 0
End Function

'Function used to left split user fields
Public Function LeftSplitUF(ByVal srcUF As String) As String
    If srcUF = "*~~~~~*" Then LeftSplitUF = "": Exit Function
    Dim i As Integer
    Dim t As String
    For i = 1 To Len(srcUF)
        If Mid$(srcUF, i, 7) = "*~~~~~*" Then
            Exit For
        Else
            t = t & Mid$(srcUF, i, 1)
        End If
    Next i
    LeftSplitUF = t
    i = 0
    t = ""
End Function

'Function used to right split user fields
Public Function RightSplitUF(ByVal srcUF As String) As String
    If srcUF = "*~~~~~*" Then RightSplitUF = "": Exit Function
    Dim i As Integer
    Dim t As String
    For i = (InStr(1, srcUF, "*~~~~~*", vbTextCompare) + 7) To Len(srcUF)
        t = t & Mid$(srcUF, i, 1)
    Next i
    RightSplitUF = t
    i = 0
    t = ""
End Function

'Function that return true if the control is empty
Public Function is_empty(ByRef sText As Variant, Optional UseTagValue As Boolean) As Boolean
    On Error Resume Next
    If sText.Text = "" Then
        is_empty = True
        If UseTagValue = True Then
            MsgBox "The field '" & sText.Tag & "' is required.Please check it!", vbExclamation
        Else
            MsgBox "The field is required.Please check it!", vbExclamation
        End If
        sText.SetFocus
    Else
        is_empty = False
    End If
End Function


'Function used to change the yes/no value
Public Function changeYNValue(ByVal srcStr As String) As String
    Select Case srcStr
        Case "Y": changeYNValue = "1"
        Case "N": changeYNValue = "0"
        Case "1": changeYNValue = "Y"
        Case "0": changeYNValue = "N"
    End Select
End Function

'Function that return true if the control is numeric
Public Function is_numeric(ByRef sText As String) As Boolean
    If IsNumeric(sText) = False Then
        is_numeric = False
        MsgBox "The field required a numeric input.Please check it!", vbExclamation
    Else
        is_numeric = True
    End If
End Function

'Function that return the value of a certain field
Public Function getValueAt(ByVal srcSQL As String, ByVal whichField As String) As String
    Dim rs As New Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open srcSQL, CN, adOpenStatic, adLockReadOnly
    If rs.RecordCount > 0 Then getValueAt = rs.Fields(whichField)
    
    Set rs = Nothing
End Function

'Convert string to number
'I create this istead of val() co'z val return incorrect value
'ex. Try to see the output of val("3,800")
'It did not support characters like , and etc.
Public Function toNumber(ByVal srcCurrency As String, Optional RetZeroIfNegative As Boolean) As Double
    If srcCurrency = "" Then
        toNumber = 0
    Else
        Dim retValue As Double
        If InStr(1, srcCurrency, ",") > 0 Then
            retValue = Val(Replace(srcCurrency, ",", "", , , vbTextCompare))
        Else
            retValue = Val(srcCurrency)
        End If
        If RetZeroIfNegative = True Then
            If retValue < 1 Then retValue = 0
        End If
        toNumber = retValue
        retValue = 0
    End If
End Function

'Function that return the count of the rows in the table
Public Function getRecordCount(ByVal srcTable As String, Optional srcCondition As String, Optional isFormatted As Boolean) As String
    If srcCondition <> "" Then srcCondition = " " & srcCondition
    Dim rs As New Recordset
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT COUNT(PK) as TCount FROM " & srcTable & srcCondition, CN, adOpenStatic, adLockReadOnly
    If isFormatted = True Then
        getRecordCount = Format$(rs![TCount], "#,##0")
    Else
        getRecordCount = rs![TCount]
    End If
    Set rs = Nothing
End Function

'Function that will return a currenct format
Public Function toMoney(ByVal srcCurr As String) As String
   toMoney = Format$(srcCurr, "#,##0.00")
End Function

'Function used to determine if the object has been set
Public Function isObjectSet(srcObject As Object) As Boolean
    On Error GoTo err
    'I use tag because almost all controls have this
    srcObject.Tag = srcObject.Tag
    isObjectSet = True
    
    Exit Function
err:
    isObjectSet = False
End Function

'Function used to get the end day number of a cetain month
Public Function getEndDay(ByVal srcDate As Date) As Byte
    Dim h1 As String
    h1 = Format(srcDate, "mm")
    On Error GoTo err
    Select Case h1
        Case Is = "01": getEndDay = 31
        Case Is = "02": getEndDay = Day(h1 & "/29/" & Format(srcDate, "yy"))
        Case Is = "03": getEndDay = 31
        Case Is = "04": getEndDay = 30
        Case Is = "05": getEndDay = 31
        Case Is = "06": getEndDay = 30
        Case Is = "07": getEndDay = 31
        Case Is = "08": getEndDay = 31
        Case Is = "09": getEndDay = 30
        Case Is = "10": getEndDay = 31
        Case Is = "11": getEndDay = 30
        Case Is = "12": getEndDay = 31
    End Select
    h1 = ""
    Exit Function
err:
        If err.Number = 13 Then getEndDay = 28: h1 = "" 'Day if encounter not a left-year
End Function

