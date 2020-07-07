Option Explicit

Public Sub CreateDummyFile()

    Dim oTxt As New clsTextFileGenerator
    
    oTxt.Delimiter = ","
    oTxt.FieldCount = 12
    oTxt.RowCount = 25
    oTxt.IncludeHeader = True
    oTxt.FileType = ".csv"
    oTxt.FileNameDateStamp = True
    oTxt.Filename = "Test"
    
    If oTxt.GenerateTextFile Then
        If oTxt.CreateSQLFile Then
            MsgBox "File has been created!"
        End If
    Else
        MsgBox "File not created!"
    End If
    Set oTxt = Nothing

End Sub