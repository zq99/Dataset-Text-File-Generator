# Dataset Text File Generator
A class to generate a dummy text file in VBA with randomly populated data.

The output file has mixed data types. 

The purpose for this class was to be able to create dummy files to help build import proceses for EUC applications,
when the actual source file required is not available yet.

## Implementation

This is the code required to create a simple text file:

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

## Output

The class has two outputs. 

The first is a text file with a dataset populated with random information, based on the parameters you have specified.

The second output is a text file, which contains the necessary SQL statement to create a table in SQL Server based on the dimensions and data types found in the text file.

