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
    MsgBox Err.Description
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