Attribute VB_Name = "Creator"
Sub createJSON()

    Dim fileName, filePath As String
    Dim currentKey As String
    Dim dataColumn, dataRow As Integer
    
    ' Don't reference the sheet so that the worksheet can be copied and still work
    ' This macro will only be activated on the sheet itself so the current sheet will be the one used
    
    ' Get the file name and path
    fileName = Cells(3, 3)
    filePath = Cells(3, 6)
    
    ' Check that a file name was entered
    If fileName = "" Then
        MsgBox ("You must enter a file name")
        Exit Sub
    End If
    
    ' Check that a file path was entered
    If filePath = "" Then
        MsgBox ("You must enter a file path")
        Exit Sub
    End If
    
    ' Make sure the last character is a "/" or "\"
    If Not Right(filePath, 1) = "/" And Not Right(filePath, 1) = "\" Then
        filePath = filePath & "\"
    End If
    
    ' Check that the file path is vaild
    If Not FolderExists(filePath) Then
        MsgBox ("The specified file path does not exist")
        Exit Sub
    End If

    ' Figure out how many columns and rows there are
    headerRow = 5
    
    dataStartColumn = 2
    dataStartRow = 6
    
    dataLastColumn = 2
    dataLastRow = 6
    
    For currentColumn = dataStartColumn To 1000
        If Cells(headerRow, currentColumn) = "" Then
            dataLastColumn = currentColumn - 1
            Exit For
        End If
    Next currentColumn
    
    For currentRow = dataStartRow To 10000
        If Cells(currentRow, dataStartColumn) = "" Then
            dataLastRow = currentRow - 1
            Exit For
        End If
    Next currentRow
    
    ' Make sure there's at least one column and row
    If dataLastColumn < dataStartColumn Then
        MsgBox ("Must have at least one column of data")
        Exit Sub
    End If
    If dataLastRow < dataStartRow Then
        MsgBox ("Must have at least one row of data")
        Exit Sub
    End If
    
    
    ' Create / Open the JSON file
    Dim strFile_Path As String
    strFile_Path = filePath & fileName & ".json"
    Open strFile_Path For Output As #1
    
    Dim lineString, key, value As String
    
    ' Print the opening string
    lineString = "["
    Print #1, lineString
    
    ' Cycle through each of the data rows
    For dataRow = dataStartRow To dataLastRow
    
        ' Each row represents an entry in an array
        ' Each entry is a dictionary of key value pairs
    
        ' Print the dictionary start string
        lineString = "{"
        Print #1, lineString
        
        ' Cycle through each column (or key value pair) and print the line
        For dataColumn = dataStartColumn To dataLastColumn
        
            ' Grab the key and value
            key = Cells(headerRow, dataColumn)
            value = Cells(dataRow, dataColumn)
            
            ' Create the line string
            lineString = """" & key & """: "  ' key within quotes plus semicolon
            
            If value <> "" Then
                lineString = lineString & """" & value & """"
            Else
                lineString = lineString & "null"
            End If
            
            ' Add a comma if it's not the last column
            If dataColumn <> dataLastColumn Then
                lineString = lineString & ","
            End If
            
            ' Print the line
            Print #1, lineString
        
        Next dataColumn
        
        ' Print the dictionary end string
        lineString = "}"
        ' Add a comma if it's not the last row
        If dataRow <> dataLastRow Then
            lineString = lineString & ","
        End If
        Print #1, lineString
            
    Next dataRow
    
    ' Print the closing string
    lineString = "]"
    Print #1, lineString
     
    
    ' Close the JSON document
    Close #1
    
End Sub

Public Function FolderExists(strFolderPath As String) As Boolean
    On Error Resume Next
    FolderExists = (GetAttr(strFolderPath) And vbDirectory) = vbDirectory
    On Error GoTo 0
End Function
