Attribute VB_Name = "Module1"
'Sub ��έp����()
'    ' Get region
'    Dim region As String
'    region = ActiveSheet.Range("A1").Value
'
'    ' Get end day of the month
'    Dim year As Integer
'    Dim month As Integer
'    year = Application.InputBox(Prompt:="�п�J�Ӧ褸�~�G", Type:=1)
'    month = Application.InputBox(Prompt:="�п�J�Ӥ���G", Type:=1)
'    endDate = WorksheetFunction.EoMonth(DateSerial(year, month, 1), 0)
'    endDay = Day(endDate)
'
'    ' Set up dictionary
'    Set columnDict = CreateObject("Scripting.Dictionary")
'    columnDict.Add "����`��", "T"
'    columnDict.Add "�а��H��", "X"
'    columnDict.Add "��|�H��", "Y"
'    columnDict.Add "�`�H���", "AG"
'    columnDict.Add "�s�i�i�@", "U"
'    columnDict.Add "�s�i����", "V"
'    columnDict.Add "�s�i����", "W"
'    columnDict.Add "�X�|�H��", "Z"
'
'    Set outputcolumndict = CreateObject("Scripting.Dictionary")
'    outputcolumndict.Add "����`��", "B"
'    outputcolumndict.Add "�а��H��", "C"
'    outputcolumndict.Add "��|�H��", "D"
'    outputcolumndict.Add "�`�H���", "E"
'    outputcolumndict.Add "�s�i�i�@", "F"
'    outputcolumndict.Add "�s�i����", "F"
'    outputcolumndict.Add "�s�i����", "F"
'    outputcolumndict.Add "�X�|�H��", "G"
'
'    ActiveSheet.Range("B3:G33").ClearContents
'
'    For d = 1 To endDay
'        ' Get row column for the region
'        regionRow = Application.Match(region, Worksheets(CStr(d)).Columns(1))
'        ' Get statics
'        For Each Key In columnDict.Keys
'            tempPopulation = Worksheets(CStr(d)).Range(columnDict(Key) & regionRow).Value
'            ' Write to active sheet
'            ActiveSheet.Range(outputcolumndict(Key) & CStr(d + 2)).Value = tempPopulation + _
'            ActiveSheet.Range(outputcolumndict(Key) & CStr(d + 2)).Value
'
'            If Key = "�X�|�H��" Then
'                ' Get imported population
'                importedPopulation = ActiveSheet.Range(outputcolumndict("�s�i�i�@") & CStr(d + 2)).Value
'                ActiveSheet.Range(outputcolumndict(Key) & CStr(d + 2)).Value = _
'                ActiveSheet.Range(outputcolumndict(Key) & CStr(d + 2)).Value + _
'                importedPopulation
'            End If
'        Next Key
'    Next d
'End Sub

Sub �h�Ϥ�έp����()
    ' Set current Workbook
    Dim thisWb As Workbook
    Set thisWb = ThisWorkbook
    
    ' Create new Workbook
    Dim countWb As Workbook
    Set countWb = Workbooks.Add
        
    ' Get end day of the month
    Dim year As Integer
    Dim month As Integer
    year = Application.InputBox(Prompt:="�п�J�Ӧ褸�~�G", Type:=1)
    month = Application.InputBox(Prompt:="�п�J�Ӥ���G", Type:=1)
    endDate = WorksheetFunction.EoMonth(DateSerial(year, month, 1), 0)
    endDay = Day(endDate)
    
    ' Set up dictionary
    Set columnDict = CreateObject("Scripting.Dictionary")
    columnDict.Add "����`��", "T"
    columnDict.Add "�а��H��", "X"
    columnDict.Add "��|�H��", "Y"
    columnDict.Add "�`�H���", "AG"
    columnDict.Add "�s�i�i�@", "U"
    columnDict.Add "�s�i����", "V"
    columnDict.Add "�s�i����", "W"
    columnDict.Add "�X�|�H��", "Z"
    
    Set outputcolumndict = CreateObject("Scripting.Dictionary")
    outputcolumndict.Add "����`��", "B"
    outputcolumndict.Add "�а��H��", "C"
    outputcolumndict.Add "��|�H��", "D"
    outputcolumndict.Add "�`�H���", "E"
    outputcolumndict.Add "�s�i�i�@", "F"
    outputcolumndict.Add "�s�i����", "F"
    outputcolumndict.Add "�s�i����", "F"
    outputcolumndict.Add "�X�|�H��", "G"
    
    ' add Worksheets
    Dim regionArr As Variant
    regionArr = Array("1C", "1D", "1E", "2C", "2D", "2E", "3C", "3D", "3E")
    For Each region In regionArr
        countWb.Worksheets.Add(After:=countWb.Worksheets(countWb.Worksheets.Count)).Name = region
        countWb.Worksheets(region).Range("A1").Value = "���"
        countWb.Worksheets(region).Range("B1").Value = "����`��"
        countWb.Worksheets(region).Range("C1").Value = "�а��H��"
        countWb.Worksheets(region).Range("D1").Value = "��|�H��"
        countWb.Worksheets(region).Range("E1").Value = "�`�H���"
        countWb.Worksheets(region).Range("F1").Value = "�s�i�H��"
        countWb.Worksheets(region).Range("G1").Value = "�X�|�[�s�i�H��"
    Next region
    
    For d = 1 To endDay
        For Each region In regionArr
            ' Get row for the region
            regionRow = Application.Match(region, thisWb.Worksheets(CStr(d)).Columns(1))
            ' Write dates
            countWb.Worksheets(region).Range("A" & CStr(d + 1)).Value = d
            ' Get statics
            For Each Key In columnDict.Keys
                tempPopulation = thisWb.Worksheets(CStr(d)).Range(columnDict(Key) & regionRow).Value
                ' Write to active sheet
                countWb.Worksheets(region).Range(outputcolumndict(Key) & CStr(d + 1)).Value = tempPopulation + _
                countWb.Worksheets(region).Range(outputcolumndict(Key) & CStr(d + 1)).Value
                
                If Key = "�X�|�H��" Then
                    ' Get imported population
                    importedPopulation = countWb.Worksheets(region).Range(outputcolumndict("�s�i�i�@") & CStr(d + 2)).Value
                    countWb.Worksheets(region).Range(outputcolumndict(Key) & CStr(d + 1)).Value = _
                    countWb.Worksheets(region).Range(outputcolumndict(Key) & CStr(d + 1)).Value + _
                    importedPopulation
                End If
            Next Key
        Next region
    Next d
End Sub
