Attribute VB_Name = "Module1"
'Sub る参璸厨()
'    ' Get region
'    Dim region As String
'    region = ActiveSheet.Range("A1").Value
'
'    ' Get end day of the month
'    Dim year As Integer
'    Dim month As Integer
'    year = Application.InputBox(Prompt:="叫块赣﹁じ", Type:=1)
'    month = Application.InputBox(Prompt:="叫块赣る", Type:=1)
'    endDate = WorksheetFunction.EoMonth(DateSerial(year, month, 1), 0)
'    endDay = Day(endDate)
'
'    ' Set up dictionary
'    Set columnDict = CreateObject("Scripting.Dictionary")
'    columnDict.Add "チ羆计", "T"
'    columnDict.Add "叫安计", "X"
'    columnDict.Add "皘计", "Y"
'    columnDict.Add "羆ら计", "AG"
'    columnDict.Add "穝秈緄臔", "U"
'    columnDict.Add "穝秈酚", "V"
'    columnDict.Add "穝秈ア醇", "W"
'    columnDict.Add "皘计", "Z"
'
'    Set outputcolumndict = CreateObject("Scripting.Dictionary")
'    outputcolumndict.Add "チ羆计", "B"
'    outputcolumndict.Add "叫安计", "C"
'    outputcolumndict.Add "皘计", "D"
'    outputcolumndict.Add "羆ら计", "E"
'    outputcolumndict.Add "穝秈緄臔", "F"
'    outputcolumndict.Add "穝秈酚", "F"
'    outputcolumndict.Add "穝秈ア醇", "F"
'    outputcolumndict.Add "皘计", "G"
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
'            If Key = "皘计" Then
'                ' Get imported population
'                importedPopulation = ActiveSheet.Range(outputcolumndict("穝秈緄臔") & CStr(d + 2)).Value
'                ActiveSheet.Range(outputcolumndict(Key) & CStr(d + 2)).Value = _
'                ActiveSheet.Range(outputcolumndict(Key) & CStr(d + 2)).Value + _
'                importedPopulation
'            End If
'        Next Key
'    Next d
'End Sub

Sub 跋る参璸厨()
    ' Set current Workbook
    Dim thisWb As Workbook
    Set thisWb = ThisWorkbook
    
    ' Create new Workbook
    Dim countWb As Workbook
    Set countWb = Workbooks.Add
        
    ' Get end day of the month
    Dim year As Integer
    Dim month As Integer
    year = Application.InputBox(Prompt:="叫块赣﹁じ", Type:=1)
    month = Application.InputBox(Prompt:="叫块赣る", Type:=1)
    endDate = WorksheetFunction.EoMonth(DateSerial(year, month, 1), 0)
    endDay = Day(endDate)
    
    ' Set up dictionary
    Set columnDict = CreateObject("Scripting.Dictionary")
    columnDict.Add "チ羆计", "T"
    columnDict.Add "叫安计", "X"
    columnDict.Add "皘计", "Y"
    columnDict.Add "羆ら计", "AG"
    columnDict.Add "穝秈緄臔", "U"
    columnDict.Add "穝秈酚", "V"
    columnDict.Add "穝秈ア醇", "W"
    columnDict.Add "皘计", "Z"
    
    Set outputcolumndict = CreateObject("Scripting.Dictionary")
    outputcolumndict.Add "チ羆计", "B"
    outputcolumndict.Add "叫安计", "C"
    outputcolumndict.Add "皘计", "D"
    outputcolumndict.Add "羆ら计", "E"
    outputcolumndict.Add "穝秈緄臔", "F"
    outputcolumndict.Add "穝秈酚", "F"
    outputcolumndict.Add "穝秈ア醇", "F"
    outputcolumndict.Add "皘计", "G"
    
    ' add Worksheets
    Dim regionArr As Variant
    regionArr = Array("1C", "1D", "1E", "2C", "2D", "2E", "3C", "3D", "3E")
    For Each region In regionArr
        countWb.Worksheets.Add(After:=countWb.Worksheets(countWb.Worksheets.Count)).Name = region
        countWb.Worksheets(region).Range("A1").Value = "ら戳"
        countWb.Worksheets(region).Range("B1").Value = "チ羆计"
        countWb.Worksheets(region).Range("C1").Value = "叫安计"
        countWb.Worksheets(region).Range("D1").Value = "皘计"
        countWb.Worksheets(region).Range("E1").Value = "羆ら计"
        countWb.Worksheets(region).Range("F1").Value = "穝秈计"
        countWb.Worksheets(region).Range("G1").Value = "皘穝秈计"
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
                
                If Key = "皘计" Then
                    ' Get imported population
                    importedPopulation = countWb.Worksheets(region).Range(outputcolumndict("穝秈緄臔") & CStr(d + 2)).Value
                    countWb.Worksheets(region).Range(outputcolumndict(Key) & CStr(d + 1)).Value = _
                    countWb.Worksheets(region).Range(outputcolumndict(Key) & CStr(d + 1)).Value + _
                    importedPopulation
                End If
            Next Key
        Next region
    Next d
End Sub
