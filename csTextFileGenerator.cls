Option Explicit

Private mblnIncludeHeader               As Boolean
Private mblnAddFileDateStamp            As Boolean
Private mstrPath                        As String
Private mstrFileName                    As String
Private mstrFileType                    As String
Private mstrDelim                       As String
Private mstrSQLTable                    As String
Private mlngMaxRows                     As Long
Private mintMaxFields                   As Long

Private Const CINT_MAX_COLUMN_TYPES     As Integer = 7
Private Const CINT_MAX_ASCII            As Integer = 90
Private Const CINT_MIN_ASCII            As Integer = 65
Private Const CINT_MAX_FIELDS           As Integer = 12
Private Const CLNG_MAX_ROWS             As Long = 10
Private Const CSTR_DEFAULT_DELIM        As String = ","
Private Const CSTR_DEFAULT_FILE_EXT     As String = ".csv"
Private Const CSTR_DEFAULT_FILE_NAME    As String = "DummyData"
Private Const CSTR_DEFAULT_TABLE_NAME   As String = "DummyDataImportTable"
Private Const CSTR_COLUMN_NAME_PREFIX   As String = "Header"

Private Sub Class_Initialize()
    mlngMaxRows = CLNG_MAX_ROWS
    mintMaxFields = CINT_MAX_FIELDS
    mstrDelim = CSTR_DEFAULT_DELIM
    mstrFileName = CSTR_DEFAULT_FILE_NAME
    mstrFileType = CSTR_DEFAULT_FILE_EXT
    mstrSQLTable = CSTR_DEFAULT_TABLE_NAME
    mblnIncludeHeader = True
    mblnAddFileDateStamp = True
    mstrPath = ThisWorkbook.Path
End Sub

Public Property Let RowCount(ByVal Value As Long)
    mlngMaxRows = Value
End Property

Public Property Let FieldCount(ByVal Value As Integer)
    If Value >= 1 Then
        mintMaxFields = Value
    End If
End Property

Public Property Let Delimiter(ByVal Value As String)
    mstrDelim = Value
End Property

Public Property Let IncludeHeader(ByVal Value As String)
    mblnIncludeHeader = Value
End Property

Public Property Let FileType(ByVal Value As String)
    mstrFileType = Value
End Property

Public Property Let Filename(ByVal Value As String)
    mstrFileName = Value
End Property

Public Property Let Filepath(ByVal Value As String)
    mstrPath = Value
End Property

Public Property Let FileNameDateStamp(ByVal Value As Boolean)
    mblnAddFileDateStamp = Value
End Property

Public Property Let SQLTableName(ByVal Value As String)
    mstrSQLTable = Value
End Property

Public Function GenerateTextFile() As Boolean

On Error GoTo ERR_HANDLER:

    Dim strFile         As String
    Dim strHeader       As String
    Dim strRow          As String
    Dim strFieldValue   As String
    Dim intTypeCnt      As Integer
    Dim intFieldCnt     As Integer
    Dim lngRow          As Long
    Dim fso             As New FileSystemObject
    Dim ts              As TextStream
    
    If mblnAddFileDateStamp Then
        mstrFileName = mstrFileName & "_" & Format(Now(), "YYYYMMDD")
    End If
    
    strFile = mstrPath & Chr(92) & mstrFileName & mstrFileType
    Set ts = fso.CreateTextFile(strFile, True)
    
    If mblnIncludeHeader Then
        For intFieldCnt = 1 To mintMaxFields
            strHeader = strHeader & IIf(strHeader <> Empty, mstrDelim, Empty) & _
                        CSTR_COLUMN_NAME_PREFIX & intFieldCnt
        Next
        ts.WriteLine (strHeader)
    End If

    For lngRow = 1 To mlngMaxRows
        DoEvents
        strRow = Empty
        intTypeCnt = 0
        For intFieldCnt = 1 To mintMaxFields
            DoEvents
            strFieldValue = Empty
            Select Case intTypeCnt
            Case 0
                strFieldValue = lngRow
            Case 1
                strFieldValue = GetRandomString(25)
            Case 2
                strFieldValue = GetRandomDate(True)
            Case 3
                strFieldValue = GetRandomNumber(1, 0, 2)
            Case 4
                strFieldValue = PickRandomFromArray(Array("A", "B", "C", "D", "E"))
            Case 5
                strFieldValue = GetRandomString(4, True, True, 0.8)
            Case 6
                strFieldValue = PickRandomFromArray(Array(1, 0))
            Case 7
                strFieldValue = GetRandomNumber(10000, 1, 10)
            End Select
            strRow = strRow & IIf(strRow <> Empty, mstrDelim, Empty) & strFieldValue
            intTypeCnt = intTypeCnt + 1
            If intTypeCnt > CINT_MAX_COLUMN_TYPES Then intTypeCnt = 1
        Next
        ts.WriteLine (strRow)
    Next
    GenerateTextFile = True
    
EXIT_HERE:

    If Not ts Is Nothing Then
        ts.Close
    End If
    Set ts = Nothing
    Set fso = Nothing
    Exit Function

ERR_HANDLER:

    Debug.Print Err.Description
    GenerateTextFile = False
    GoTo EXIT_HERE
    
End Function

Public Function CreateSQLFile() As Boolean

On Error GoTo ERR_HANDLER:
    Dim strSQLFile               As String
    Dim strSQL                   As String
    Dim fso                      As FileSystemObject
    Dim ts                       As TextStream
    Dim intTypeCnt               As Integer
    Dim intFieldCnt              As Integer
    Dim arrSQLServerDataTypes()  As Variant
    
    arrSQLServerDataTypes = Array("INT", "varchar(50)", "Date", "Decimal", "Char(1)", "varchar(4)", "bit", "Decimal")
    strSQLFile = mstrPath & Chr(92) & mstrFileName & "_SQL.txt"
    Set fso = New FileSystemObject
    Set ts = fso.CreateTextFile(strSQLFile, True)
    
    strSQL = "CREATE TABLE " & mstrSQLTable & "( "
    intTypeCnt = LBound(arrSQLServerDataTypes)
    For intFieldCnt = 1 To mintMaxFields
        DoEvents
        strSQL = strSQL & IIf(intFieldCnt = 1, Empty, ",") & _
                 CSTR_COLUMN_NAME_PREFIX & intFieldCnt & vbTab & arrSQLServerDataTypes(intFieldCnt)
        intTypeCnt = intTypeCnt + 1
        If intTypeCnt > UBound(arrSQLServerDataTypes) Then intTypeCnt = LBound(arrSQLServerDataTypes) + 1
    Next
    strSQL = strSQL & " )"
    
    CreateSQLFile = True

EXIT_HERE:
    ts.WriteLine (strSQL)
    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    Exit Function
ERR_HANDLER:
    CreateSQLFile = False
    GoTo EXIT_HERE
End Function

Private Function PickRandomFromArray(ByVal vntArray) As Variant
    Dim intMax As Integer
    Dim intMin As Integer
    
    intMax = UBound(vntArray)
    intMin = LBound(vntArray)
    PickRandomFromArray = vntArray(GetRandomNumber(intMax, intMin))
End Function

Private Function GetRandomString(ByVal intMaxLength As Integer, _
                        Optional ByVal blnFixedLength As Boolean = False, _
                        Optional ByVal blnAllowNulls As Boolean = False, _
                        Optional ByVal dblNullChance As Double = 0.5) As String
    Dim intLen  As Integer
    Dim dblNull As Double

    If blnAllowNulls = True Then
        dblNull = GetRandomNumber(1, 0, 2)
        If dblNull <= dblNullChance Then
            GetRandomString = Empty
        Else
            If blnFixedLength Then intLen = intMaxLength Else intLen = GetRandomNumber(intMaxLength, 1)
            GetRandomString = BuildCharacterString(intLen)
        End If
    Else
        If blnFixedLength Then intLen = intMaxLength Else intLen = GetRandomNumber(intMaxLength, 1)
        GetRandomString = BuildCharacterString(intLen)
    End If
End Function

Private Function BuildCharacterString(ByVal intLength) As String
    Dim strNewString As String
    Dim intCnt       As Integer
        
    For intCnt = 1 To intLength
        strNewString = strNewString & Chr(GetRandomNumber(CINT_MAX_ASCII, CINT_MIN_ASCII))
    Next
    BuildCharacterString = strNewString
End Function

Private Function GetRandomDate(Optional ByVal blnAsString As Boolean = True) As Variant
    Dim intYear   As Integer
    Dim intMonth  As Integer
    Dim intDay    As Integer
    Dim intMaxDay As Integer

    intYear = GetRandomNumber(Year(Now()), 1930)
    intMonth = GetRandomNumber(12, 1)
    Select Case intMonth
    Case 2
        intMaxDay = 28
    Case 9, 4, 6, 11
        intMaxDay = 30
    Case Else
        intMaxDay = 31
    End Select
    
    intDay = GetRandomNumber(intMaxDay, 1)
    If blnAsString Then
         GetRandomDate = CStr(intYear & "-" & intMonth & "-" & intMaxDay)
    Else
         GetRandomDate = CDate(intMaxDay & "/" & intMonth & "/" & intYear)
    End If
End Function

Private Function GetRandomNumber(ByVal lngMax As Long, ByVal lngMin As Long, Optional ByVal intDecimals As Integer)
   If IsMissing(intDecimals) Or intDecimals = 0 Then
      Randomize
      GetRandomNumber = Int((lngMax + 1 - lngMin) * Rnd + lngMin)
   Else
      Randomize
      GetRandomNumber = Round((lngMax - lngMin) * Rnd + lngMin, intDecimals)
   End If
End Function