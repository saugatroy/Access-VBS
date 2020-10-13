Option Compare Database

'------------------------------------------------------------
' OpenWindows
'
'------------------------------------------------------------
Function OpenWindows()
On Error GoTo OpenWindows_Err

    DoCmd.OpenForm "Bill Updates", acNormal, "", "", , acNormal
    DoCmd.OpenForm "Credit Updates", acNormal, "", "", , acNormal


OpenWindows_Exit:
    Exit Function

OpenWindows_Err:
    MsgBox Error$
    Resume OpenWindows_Exit

End Function

Function NullOrEmpty(strTextToTest As Variant) As Boolean
    If IsNull(strTextToTest) Then
        NullOrEmpty = True
        Exit Function
    End If
    If Trim(strTextToTest & "") = "" Then
        NullOrEmpty = True
        Exit Function
    End If
    NullOrEmpty = False
End Function

Public Function QueryExists(ByVal pName As String) As Boolean
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim blnReturn As Boolean
    Dim strMsg As String

On Error GoTo ErrorHandler

    blnReturn = False ' make it explicit
    Set db = CurrentDb
    Set qdf = db.QueryDefs(pName)
    blnReturn = True

ExitHere:
    Set qdf = Nothing
    Set db = Nothing
    QueryExists = blnReturn
    Exit Function

ErrorHandler:
    Select Case Err.Number
    Case 3265 ' Item not found in this collection.
    Case Else
        strMsg = "Error " & Err.Number & " (" & Err.Description _
            & ") in procedure QueryExists"
        MsgBox strMsg
    End Select
    GoTo ExitHere

End Function

Public Sub cmdCloseQuery(pQueryName As String)
On Error GoTo Err_cmdCloseForm_Click

    DoCmd.Close acQuery, pQueryName, acSaveNo
 
Exit_cmdCloseForm_Click:
    Exit Sub
 
Err_cmdCloseForm_Click:
    'The query may already be closed, in final run ignore
    'MsgBox Err.Description
    Resume Exit_cmdCloseForm_Click

End Sub

'------------------------------------------------------------
' Clear all tables in the DB
'------------------------------------------------------------
Function ClearData()
On Error GoTo ClearData_Err

    Dim db As DAO.Database
    Dim td As DAO.TableDefs
    Dim lTableNames() As String
    
    'Set Function to false (explicitly)
    ClearData = False
    
    Set db = CurrentDb()
    Set td = db.TableDefs
    
    For Each t In td    'loop through all the fields of the tables
        If Left(t.Name, 4) <> "MSys" And t.Name <> "Entity" And Left(t.Name, 1) <> "~" Then
            
            lSQL = "DELETE * from [" & t.Name & "]"
            'MsgBox lSQL, vbOKOnly
            
            'Run the delete command
            DoCmd.RunSQL lSQL
            
        End If
        'Debug.Print t.Name
        Next
    
    'Set function to true
    ClearData = True


ClearData_Exit:
    Exit Function

ClearData_Err:
    MsgBox Error$
    Resume ClearData_Exit

End Function


Public Sub UpdateQuery(varEntity As String, VarType As String)

    Dim cstrQueryName As String
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim sqlStr As String
    'Dim varEntity As String
    Dim varSQLUpdates As String
    Dim rstData As DAO.Recordset

    Set db = CurrentDb
    
    'varEntity = Me.lUpdateEntity.Value
        
    If VarType = "Bill" Then
        cstrQueryName = "Bill-Updates"
        varSQLUpdates = "SELECT A.* FROM [Current-" & varEntity & "-Bill] AS A LEFT JOIN [Previous-" & varEntity & "-Bill] AS B ON (A.BillLineItemAmount=B.BillLineItemAmount) AND (A.BillLineItemOrder = B.BillLineItemOrder)  AND (A.InvoiceNum = B.InvoiceNum) AND (A.VendorName = B.VendorName) WHERE B.VendorName is null"
    Else
        cstrQueryName = "Credit-Updates"
        varSQLUpdates = "SELECT A.* FROM [Current-" & varEntity & "-Credit] AS A LEFT JOIN [Previous-" & varEntity & "-Credit] AS B ON (A.VendorName = B.VendorName) AND (A.RefNum = B.RefNum) AND (A.VendorCreditLineItemAmount=B.VendorCreditLineItemAmount) WHERE B.VendorName is null"
    End If
        
    If Not QueryExists(cstrQueryName) Then
        Set qdf = db.CreateQueryDef(cstrQueryName)
    Else
        Set qdf = db.QueryDefs(cstrQueryName)
        'Close the box in case its oopen
        cmdCloseQuery (cstrQueryName)
    End If
    qdf.SQL = varSQLUpdates
    Set qdf = Nothing
    Set db = Nothing
    
    'Open the Query Box
    DoCmd.OpenQuery cstrQueryName
    'Close the query window so as to not cause confusion
    DoCmd.Close acForm, cstrQueryName, acSaveYes
    
    'Requery the form in the child

End Sub

Public Sub PinWindows(varForm As Form)

    'On multiple screens the windows tend to pop all over the place, this pins the windows to 0,0 - so within the current window
    varForm.Move 0, 0
    'MsgBox "Moving to default"

End Sub