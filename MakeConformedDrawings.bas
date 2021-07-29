Attribute VB_Name = "MakeConformedDrawings"

Sub getConformedList()

    'Goal: get only the current pdfs in the Conformed Drawing sheet
    'Problem 1, filtering a table makes this glacial, so let's copy the whole table into a 'bucket' sheet, _
        plus we can get rid of formulas that way and only work with values
    'Problem 2, we only want Current marked rows, aka "!!!", but holding shift+crtl+down is selecting the _
        whole table column, so we need to find the last row of "!!!"

    'first, turn off stuff to make this run faster
    Application.ScreenUpdating = False
    Application.Calculation = xlManual

    'Clear out the old list in Conformed Drawings
    Sheets("ConformedDrawings").Select
    Range("A6:D6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Clear
    
    'Clear out Bucket where we'll be pasting our new data
    Sheets("Bucket").Select
    Range("A3:V3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Clear
    
    'Go into Drawing Data, copy the entire table, and paste it into Bucket as just values
    Sheets("DrawingData").Select
    Range("B5:W5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Bucket").Select
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    Sheets("DrawingData").Select
    Range("B5").Select 'just putting things visually back to normal here
    
    'now in Bucket, sort Current to have all the "!!!" at the top, aka column F
    Sheets("Bucket").Select
    Worksheets("Bucket").AutoFilter.Sort.SortFields.Clear
    Worksheets("Bucket").AutoFilter.Sort.SortFields.Add2 Key:= _
    Range("F2"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Bucket").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    'Build the range to copy - we know the columns and starting row, just need to figure out the last row
    Dim countCurrent As String
    lastrowCurrent = Worksheets("Bucket").Range("currentCount") 'get the total count of "!!!" from the cell E1, named currentCount
    lastrowCurrent = lastrowCurrent * 1 ' convert it to a number
    lastrowCurrent = lastrowCurrent + 2 ' add 2 to account for our Header rows --> now we know the last row!
    
    
    'copy everything (As Values Only) we need into the Conformed Drawing tab
        'row A, Original PDF File List
        Worksheets("Bucket").Select
        Sheets("Bucket").Range("A3:A" & lastrowCurrent).Select
        Selection.Copy
        Sheets("ConformedDrawings").Select
        Range("A6").Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        'row T, SubVolume
        Worksheets("Bucket").Select
        Sheets("Bucket").Range("T3:T" & lastrowCurrent).Select
        Selection.Copy
        Sheets("ConformedDrawings").Select
        Range("B6").Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        'row U, File Name
        Worksheets("Bucket").Select
        Sheets("Bucket").Range("U3:U" & lastrowCurrent).Select
        Selection.Copy
        Sheets("ConformedDrawings").Select
        Range("C6").Select
        Selection.PasteSpecial Paste:=xlPasteValues
        
        'row V, File Path
        Worksheets("Bucket").Select
        Sheets("Bucket").Range("V3:V" & lastrowCurrent).Select
        Selection.Copy
        Sheets("ConformedDrawings").Select
        Range("D6").Select
        Selection.PasteSpecial Paste:=xlPasteValues
     
    'Make it visually nice
    Sheets("ConformedDrawings").Select
    Range("A4").Select
     
    'turn stuff back on
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True

End Sub
    
    
Sub copyConformed()

     'Goal:  go into the All_Drawings folder, copy the files and paste them into their designated volume folder
    ' https://www.rondebruin.nl/win/s3/win026.htm

    Dim rng As Range
    Dim row As Range
    Dim wsConformed As Worksheet
    Set wsConformed = Sheets("ConformedDrawings")
    
    'figure out our range first, aka every drawing row on the sheet
    Set rng = wsConformed.Range(wsConformed.Range("A6"), wsConformed.Range("A" & Rows.Count).End(xlUp))

    Dim fso As Object
    Dim FromPath As String
    Dim ToPath As String
    Dim FromPathFile As String
    Dim ToPathFile As String

    '1. set the "copy from" folder path and "paste into" folder path
        'my FromPath will be where I store all the drawings, aka my ALL_DRAWINGS folder.
        'I put the folder path in a cell and named it "ALL_DRAWINGS"
    FromPath = Range("ALL_DRAWINGS").Value
    ToPath = "" 'this is just a hold

    Set fso = CreateObject("scripting.filesystemobject")

    'calling another macro to clear the old files within the Conformed Folders
    Call cleanOutFolders

    For Each row In rng 'for each drawing...
        FromPathFile = FromPath & "\" & row.Value 'building the file original location
        ToPathFile = row.Offset(, 3).Value 'look 3 cells over to find the new file home
        fso.CopyFile Source:=FromPathFile, Destination:=ToPathFile 'now that we have both paths, perform the copy paste

        On Error Resume Next 'only here because sometimes the range extends beyond actual text and the code 'freaks out' with a blank cell

    Next row
    


End Sub

Sub cleanOutFolders()

    Dim ws As Worksheet
    Dim volumePath As String
    Dim volumeRange As Range
    Set volumeRange = Application.Range("volumePathRange")

    volumePath = ""

    'Goal: clear out all the old PDFs
    'I made a range of the volume paths, so for each row, go into that folder:
        '1) check if the row is empty, if so skip it
        '2) if not empty, clear out/kill all PDFs
        
    'POSSIBLE ERROR: if the folder has no PDFs to delete, it throws an error

    For Each row In volumeRange
        With row
            If row = "" Then
                volumePath = "skip"
            Else
                volumePath = row & "\*pdf"
                Kill volumePath
            End If
            
            On Error Resume Next 'here because if the folder is empty, it throws an error
            
        End With
    Next row
    

End Sub


Sub getConformedList_SSI()

    Application.ScreenUpdating = False

    'first clear out the old list
    Sheets("ConformedDrawings_SSI").Select
    Range("A6:C5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Clear
    
    'go into Drawing Data_SSI, filter for current, copy here
    Sheets("DrawingData_SSI").Select
    ActiveSheet.ListObjects("DrawingData_SSI").Range.AutoFilter Field:=7, Criteria1 _
        :="<>"
    
    Range("DrawingData_SSI[[#Headers],[Original PDF File List]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("ConformedDrawings_SSI").Select
    Range("A5").Select
    ActiveSheet.Paste
    
    Sheets("DrawingData_SSI").Select
    Range("DrawingData_SSI[[#Headers],[Volume]:[File Name]]").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("ConformedDrawings_SSI").Select
    Range("B5").Select
    ActiveSheet.Paste
    
    'put things back to normal in DrawingData
    Sheets("DrawingData_SSI").Select
    ActiveSheet.ListObjects("DrawingData_SSI").Range.AutoFilter Field:=7

    Sheets("ConformedDrawings_SSI").Select
    Range("A4").Select

    Application.ScreenUpdating = True

End Sub


Sub copyConformed_SSI()

     'Goal:  go into the All_Drawings SSI folder, copy the files and paste them into their designated volume folder
    ' https://www.rondebruin.nl/win/s3/win026.htm

    'first, clean out the SSI folder
            Dim Volume1_6_SSI As String
            
            'set the folder path in tab Doc Info
            Volume1_6_SSI = Range("Vol1.6_SSI") & "\*pdf"
        
            'Delete every PDF in the folders designated below
            Kill Volume1_6_SSI

    'now copy and past the files
    Dim rng As Range
    Dim row As Range
    Dim wsConformedSSI As Worksheet
    Set wsConformedSSI = Sheets("ConformedDrawings_SSI")
    
    Set rng = wsConformedSSI.Range(wsConformedSSI.Range("A6"), wsConformedSSI.Range("A" & Rows.Count).End(xlUp))

    Dim fso As Object
    Dim FromPath As String
    Dim ToPath As String
    Dim FromPathFile As String
    Dim ToPathFile As String

    '1. set the "copy from" folder path and "paste into" folder path
    FromPath = Range("ALL_DRAWINGS_SSI").Value
    ToPath = "" 'this is just a hold

    Set fso = CreateObject("scripting.filesystemobject")
    
    For Each row In rng
        FromPathFile = FromPath & "\" & row.Value 'building the file original location
        ToPathFile = row.Offset(, 3).Value 'in the new volume folder, place this specific file
        fso.CopyFile Source:=FromPathFile, Destination:=ToPathFile 'now that we have both paths, perform the copy paste
        
        On Error Resume Next
        
    Next row

End Sub
