Public Connection       As Object
Public Recordset        As Object

Public companyArr       As Variant
Public categoryArr      As Variant
Public shipToArr        As Variant
Public poNumberArr      As Variant
Public buyerNameArr     As Variant
Public holdNameArr      As Variant

Public Sub testSQL()

Dim lastRow As Long
Dim SQL     As String

'===================SQL===================
lastRow = 2

Call OpenConnection(ThisWorkbook.FullName)


'Step 1
Call arrayParams("A", companyArr, False)
Call arrayParams("B", categoryArr, True)
Call arrayParams("C", shipToArr, False)
Call arrayParams("D", poNumberArr, True)
Call arrayParams("E", buyerNameArr, False)

SQL = "SELECT * FROM [" & Sheet2.Name & "$]" _
    & " WHERE [COMPANY] IN (" & Join(companyArr, ",") & ")" _
    & " AND [CATEGORY] NOT LIKE " & Join(categoryArr, " ") & "" _
    & " AND [SHIP TO ORG NAME] NOT IN ('" & Join(shipToArr, "','") & "')" _
    & " AND [PO NUMBER] LIKE " & Join(poNumberArr, " ") & "" _
    & " AND [BUYER NAME] IN ('" & Join(buyerNameArr, "','") & "')"
    
lastRow = runQueryAndCopy(SQL, lastRow)


'Step 2
Call arrayParams("G", companyArr, False)
Call arrayParams("H", categoryArr, True)
Call arrayParams("I", shipToArr, False)
Call arrayParams("J", poNumberArr, True)
Call arrayParams("K", buyerNameArr, False)
Call arrayParams("L", holdNameArr, False)

SQL = "SELECT * FROM [" & Sheet2.Name & "$]" _
    & " WHERE [COMPANY] IN (" & Join(companyArr, ",") & ")" _
    & " AND [CATEGORY] NOT LIKE " & Join(categoryArr, " ") & "" _
    & " AND [SHIP TO ORG NAME] NOT IN ('" & Join(shipToArr, "','") & "')" _
    & " AND [PO NUMBER] LIKE " & Join(poNumberArr, " ") & "" _
    & " AND [BUYER NAME] NOT IN ('" & Join(buyerNameArr, "','") & "')" _
    & " AND [HOLD NAME] IN ('" & Join(holdNameArr, "','") & "')" _
    & " AND [ITEM NUMBER] IS NULL"
    
lastRow = runQueryAndCopy(SQL, lastRow)


'Step 3
Call arrayParams("O", companyArr, False)
Call arrayParams("P", categoryArr, True)
Call arrayParams("Q", poNumberArr, True)
Call arrayParams("R", buyerNameArr, False)

SQL = "SELECT * FROM [" & Sheet2.Name & "$]" _
    & " WHERE [COMPANY] IN (" & Join(companyArr, ",") & ")" _
    & " AND [CATEGORY] NOT LIKE " & Join(categoryArr, " ") & "" _
    & " AND [PO NUMBER] LIKE " & Join(poNumberArr, " ") & "" _
    & " AND [BUYER NAME] IN ('" & Join(buyerNameArr, "','") & "')"
    
lastRow = runQueryAndCopy(SQL, lastRow)


'Step 4
Call arrayParams("T", companyArr, False)
Call arrayParams("U", categoryArr, True)
Call arrayParams("V", poNumberArr, True)
Call arrayParams("W", buyerNameArr, False)
Call arrayParams("X", holdNameArr, False)

SQL = "SELECT * FROM [" & Sheet2.Name & "$]" _
    & " WHERE [COMPANY] IN (" & Join(companyArr, ",") & ")" _
    & " AND [CATEGORY] NOT LIKE " & Join(categoryArr, " ") & "" _
    & " AND [PO NUMBER] LIKE " & Join(poNumberArr, " ") & "" _
    & " AND [BUYER NAME] NOT IN ('" & Join(buyerNameArr, "','") & "')" _
    & " AND [HOLD NAME] IN ('" & Join(holdNameArr, "','") & "')"
    
lastRow = runQueryAndCopy(SQL, lastRow)

Call CloseConnection(ThisWorkbook.FullName)
'=========================================

End Sub

Private Function runQueryAndCopy(ByVal SQL As String, ByVal lastRow As Long) As Long

    Set Recordset = Connection.Execute(SQL)
    Sheet5.Cells(lastRow, 1).CopyFromRecordset Recordset
    lastRow = Sheet5.Cells(1, 1).CurrentRegion.Rows.Count + 1
    runQueryAndCopy = lastRow

End Function
Private Function arrayParams(ByVal column As String, ByRef targetArr As Variant, ByVal startsWith As Boolean)

    'resize the array for the new parameters
    lastRow = Sheet1.Cells(Sheet1.Rows.Count, column).End(xlUp).Row
    ReDim targetArr(lastRow - 2)
    
    'if the parameter is a partial parameter (eg: 'starts with'), then concat the "%" sign to work in SQL
    If startsWith = True Then
        For i = LBound(targetArr) To UBound(targetArr)
            If i = 0 Then
                targetArr(i) = "'" & Sheet1.Cells(i + 2, column).Value & "%'"
            Else
                targetArr(i) = "AND [CATEGORY] NOT LIKE '" & Sheet1.Cells(i + 2, column).Value & "%'"
            End If
        Next i
    Else
        For i = LBound(targetArr) To UBound(targetArr)
            targetArr(i) = Sheet1.Cells(i + 2, column).Value
        Next i
    End If

End Function

Private Function OpenConnection(ByVal workbookName As String)

    Dim conn_str As String
    Set Connection = CreateObject("ADODB.Connection")
    Set Recordset = CreateObject("ADODB.Recordset")
    conn_str = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & workbookName & ";Extended Properties=""Excel 12.0 Macro;HDR=YES"";"
    Connection.Open conn_str
    
End Function

Private Function CloseConnection(ByVal workbookName As String)

    Connection.Close
    Set Connection = Nothing
    Set Recordset = Nothing
    
End Function
