' NOTE: THIS IS ACCESS VBA, NOT VB

Option Compare Database
Option Explicit

' DEV NOTES:
' NOTE: Max address space = charset size ^ ID length
' TODO: warn the user if they request more IDs than the address space can handle?
' TODO: have some sort of timeout, if there there are too many collisions (i.e., address space is running very low)

' Last edit 2018-03-29 by JG
Public Const DB_Error_Duplicate = 3022 ' Insert failed due to anti-duplicate constraint

Function ID_New(forcer As Variant, length As Long, Optional charSet As String = "", Optional upCase = True) As String
    ' Generates a random ID using length and charset; not guaranteed to be unique (must be tested against existing IDs)
    ' Argument Forcer will cause the Access query to re-run the f() for each row; use something that varies by record from one record to the next)
    ' Last edit 2018-04-23 by JG
    
    Const defaultCharset = "abcdefghjkmnpqrstuvwxyz23456789" ' leaving out easy-to-confuse chars ilo and digits 01 (leaving 31 total)
    
    Dim i As Long
    Dim tmpID As String
    Dim charsetLen As Long
    Dim CharPos As Long
    
    If charSet = "" Then charSet = defaultCharset
    
    charsetLen = Len(charSet)
    
    
    For i = 1 To length
        CharPos = Int(Rnd * charsetLen) + 1
        tmpID = tmpID & Mid(charSet, CharPos, 1)
    Next i
    
    ID_New = IIf(upCase, UCase(tmpID), tmpID)
End Function

Function IDs_Generate(tableName As String, fieldName As String, length As Long, cnt As Long) As Boolean
    ' Generate n IDs and append them to the indicated table
    ' Last edit 2019-05-31 by JG
    
    Dim rs As DAO.Recordset2
    Dim i As Long
    
    Dim dupCnt As Long
    
    Randomize
    
    Set rs = CurrentDb.OpenRecordset(tableName)
    
    Debug.Print Now()
    
    For i = 1 To cnt
        rs.AddNew
        rs(fieldName) = ID_New(Null, length)
        On Error GoTo ErrorHandler
        rs.Update
    Next i
    
    rs.Close
    
    Debug.Print Now()

    Debug.Print "Duplicates re-generated: ", dupCnt
    
    IDs_Generate = True
    
    Exit Function
    
ErrorHandler:
    If Err.Number = DB_Error_Duplicate Then ' if duplicate ID (INSERT will fail), need to try another one
        dupCnt = dupCnt + 1
        rs(fieldName) = ID_New(Null, length)
        Resume ' try the update (INSERT) again
    Else
        MsgBox "Error: " & Err.Description
        Stop
    End If
End Function

Function IDs_Generate_TEST() As Boolean
    Dim i As Long
    
    ' Test: Insert 1000 IDs into test table.
    IDs_Generate_TEST = IDs_Generate("TEST_IDs", "IDField", 3, 1000) ' ID space = 29,791
End Function

Function IDs_GenerateInPlace(tableName As String, fieldName As String, length As Long, Optional charSet As String = "") As Boolean
    ' For each record with blank (NULL or empty string) ID field, gnerate ID and update field
    ' REQUIRES: table field for IDs must have a unique constraint to detect duplicate IDs
    ' Last edit 2021-08-25 by JG
    
    Dim rs As DAO.Recordset2
    Dim recordCnt As Long
    Dim newCnt As Long
    
    Dim dupCnt As Long
    
    Randomize
    
    ' TODO: could return only records with blank target fields
    Set rs = CurrentDb.OpenRecordset(tableName)
    
    rs.MoveFirst
    
    Debug.Print "Scanning table [" & tableName & "] for blanks in ID field [" & fieldName & "]"
    Debug.Print "Start time: ", Now()
    
    Do While Not rs.EOF
        If IsNull(rs(fieldName).Value) Or rs(fieldName).Value = "" Then
            rs.Edit
            rs(fieldName) = ID_New(Null, length, charSet)
            On Error GoTo ErrorHandler
            rs.Update
            newCnt = newCnt + 1
        End If ' else, record already has an ID, skip it
        rs.MoveNext
        recordCnt = recordCnt + 1
    Loop
            
    rs.Close
    
    Debug.Print "End time:", Now()

    Debug.Print "Records scanned: ", recordCnt
    Debug.Print "Duplicates re-generated: ", dupCnt
    Debug.Print "New unique IDs added: ", newCnt
    
    IDs_GenerateInPlace = True
    
    Exit Function
    
ErrorHandler:
    If Err.Number = DB_Error_Duplicate Then ' if duplicate ID (UPDATE will fail), need to try another one
        dupCnt = dupCnt + 1
        rs(fieldName) = ID_New(Null, length, charSet)
        Resume ' try the update again
    Else
        MsgBox "Error: " & Err.Description
        Stop
    End If
End Function