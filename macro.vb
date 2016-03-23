Rem Attribute VBA_ModuleType=VBAFormModule
Option VBASupport 1


Private Sub ComboBox1_Click()
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets(ComboBox1.Value)
    ws.Activate
'AdvancedFilter(xlFilterInPlace, , , True)

    FiltercBox.Clear
    SelectbDataToCompute.Clear
    
    Dim firstRow As Range
    
    Set firstRow = ws.Range(ws.Cells(1, 1), ws.Cells(1, ws.UsedRange.Columns.Count))

    For i = 1 To firstRow.Count
     FiltercBox.AddItem (firstRow.Cells(1, i).Value)
     SelectbDataToCompute.AddItem (firstRow.Cells(1, i).Value)
    Next i
    
    FiltercBox.ListIndex = 0
    SelectbDataToCompute.ListIndex = 0
    
End Sub


    ' REF EDIT CODE
    'Get the address, or reference, from the RefEdit control.
    ' Addr = RefEdit1.value

     'Set the SelRange Range object to the range specified in the
     'RefEdit control.
     'Set SelRange = Range(Addr)

Private Sub CommandButton1_Click()

     Dim SelRange As Range
     Dim Addr As String

     'Initialise la feuille résultat
     Dim cleanSheet As Worksheet
     Dim myClasses As Collection
     Dim dataSheet As Worksheet
     
     Set dataSheet = ActiveWorkbook.Sheets(ComboBox1.Value)
     
     Set cleanSheet = init_sheet(ComboBox1.Value, StepBox.Value)
     Set myClasses = init_class(cleanSheet, StepBox.Value)
     
     criteriaValue = FiltercCritere.Value
     criteriaIndexColonne = findNum(dataSheet, FiltercBox.Value)
     dataIndexColonne = findNum(dataSheet, SelectbDataToCompute.Value)
    
     Call work(dataSheet, cleanSheet, dataIndexColonne, myClasses, criteriaIndexColonne, criteriaValue)
     
     Unload UserForm1
End Sub


Private Function findNum(sheet, strSearch) As Integer
  
    Dim aCell As Range

    Set aCell = sheet.Rows(1).find(What:=strSearch, LookIn:=xlValues, _
    LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
    MatchCase:=False, SearchFormat:=False)

    If Not aCell Is Nothing Then
       findNum = aCell.column
    End If
    
End Function

Private Sub FiltercBox_Change()

   Dim rangeSelected As Range
   Dim numOfSelectedCriteria As Integer
   
   FiltercCritere.Clear
   
   Set objDic = CreateObject("Scripting.Dictionary")
   
   numOfSelectedCriteria = findNum(ActiveSheet, FiltercBox.Value)
   
   Set rangeSelected = ActiveSheet.Range(ActiveSheet.Range(Cells(1, numOfSelectedCriteria), Cells(1, numOfSelectedCriteria)), _
   ActiveSheet.Range(Cells(1, numOfSelectedCriteria), Cells(1, numOfSelectedCriteria)).End(xlDown))

    For Each c In rangeSelected
        If Not objDic.Exists(c.Value) Then
            If Not IsEmpty(c.Value) Then
                objDic.Add c.Value, c.Value
                End If
            End If
    Next
   
   Key = objDic.keys
   For i = 0 To objDic.Count - 1
     FiltercCritere.AddItem (Key(i))
   Next
   
End Sub

Private Sub UserForm1_Initialize()

'Initialise les feuilles
For i = 1 To Sheets.Count
 ComboBox1.AddItem (Sheets(i).Name)
Next i
ComboBox1.ListIndex = 0
    
'Initialise les classes
StepBox.AddItem (20)
StepBox.AddItem (30)
StepBox.AddItem (40)
StepBox.ListIndex = 0


End Sub

Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1

Public Sub init()

    UserForm1.Show

End Sub

Public Function init_sheet(wSheet, stepClass) As Worksheet
    Dim FCesaire As Worksheet
      
    Set FCesaire = Sheets.Add(After:=Worksheets(1))
    FCesaire.Name = "stat_" & stepClass

    FCesaire.Cells.ClearContents

    With FCesaire.Range("A3:K3")
        .Columns(1).Value = "Orientation Class"
        .Columns(1).Name = "OClass"
        
        .Columns(2).Value = "expected"
        .Columns(2).Name = "expected"
        
        
        .Columns(3).Value = "observed"
        .Columns(3).Name = "observed"
        
        .Columns(4).Name = "obsExp"
        .Columns(4).Value = "diff obs exp"
        
        .Columns(5).Name = "sqrExp"
        .Columns(5).Value = "racine exp"
        
        .Columns(6).Name = "cEij"
        .Columns(6).Value = "Eij"
        
        .Columns(7).Name = "cVij"
        .Columns(7).Value = "Vij"
        
        .Columns(8).Name = "cSqrVij"
        .Columns(8).Value = "racine Vij"
         
        .Columns(9).Name = "cDij"
        .Columns(9).Value = "dij"
        
        .Columns(10).Name = "cPcalcule"
        .Columns(10).Value = "P calculée"
        
        .Columns(11).Name = "cPcalculeNeg"
        .Columns(11).Value = "P calculée"
        
    End With
    
    'Compute title row
    
    For i = 1 To FCesaire.UsedRange.Columns.Count
        FCesaire.Columns(i).NumberFormat = "0.000"
    Next
        
    
    Set init_sheet = FCesaire
End Function

Public Function init_class(wSheet, stepClass) As Object
    
    Dim NbreClasse As Integer
   
    Dim ClassRange As Collection
    Set ClassRange = New Collection
    
    NbreClasse = Application.Round(360 / stepClass, 1)
    
    Dim index As Integer
    Dim min As Integer
    Dim max As Integer
    
    index = 0
    
    For i = 0 To (360 - stepClass) Step stepClass
        
        If i = 0 Then
           min = 0
           max = i + stepClass
        Else
            min = i + 1
            max = i + stepClass
        End If
        
        ClassRange.Add (Array(min, max))
        index = index + 1
    Next
    
    Set init_class = ClassRange
    
End Function

Public Sub work(dSheet As Worksheet, wSheet As Worksheet, dataIndexColonne, myClasses, criteriaIndexColonne, criteriaValue)
  
    Dim NbreLigne As Integer
    Dim NbObjets As Integer
    
    Dim mystep As Integer
    Dim countStep As Integer
    Dim nbObjectInClass As Integer
    Dim NbreClasse As Integer
    
    Dim columnWithHeader As Range
    Dim Filteredarea As Range
    Dim SelectedArea As Range
    
    Set columnWithHeader = dSheet.Columns(dataIndexColonne)
    Set rangeA = dSheet.Range(dSheet.Cells(2, dataIndexColonne), dSheet.Cells(2, dataIndexColonne))
    Set rangeB = dSheet.Range(dSheet.Cells(2, dataIndexColonne), dSheet.Cells(2, dataIndexColonne)).End(xlDown)
    Set SelectedArea = dSheet.Range(rangeA, rangeB)
     
    'Set SelectedArea = columnWithHeader.Offset(1, 0).Resize(columnWithHeader.Rows.Count - 1)
    
    SelectedArea.AutoFilter Field:=criteriaIndexColonne, Criteria1:="=" & criteriaValue
    
    'Compute total object
    NbObjets = SelectedArea.Cells.SpecialCells(xlCellTypeVisible).SpecialCells(xlCellTypeConstants).Count
    
    'Number of classes
    NbreClasse = myClasses.Count
    
    Dim expected As Double
    Dim diff As Double
    Dim sqrtExpected As Double
    Dim eij As Double
    Dim vij As Double
    Dim sqrVij As Double
    Dim dij As Double
    Dim normDist As Double
    Dim normDistReduced As Double
    Dim rowToWrite As Integer
    
    'Compute title row
    Dim firstRow As Range
    Set firstRow = wSheet.Range(wSheet.Cells(3, 1), wSheet.Cells(3, wSheet.UsedRange.Columns.Count))
    
    For i = 1 To NbreClasse
        
        rowToWrite = 1 + i
        Range("OClass").Cells(rowToWrite, 1).Value = myClasses.Item(i)(0) & " - " & myClasses.Item(i)(1)
        SelectedArea.AutoFilter Field:=dataIndexColonne, Criteria1:=">=" & myClasses.Item(i)(0), Operator:=xlAnd, Criteria2:="<=" & myClasses.Item(i)(1)
        
        expected = NbObjets / NbreClasse
        wSheet.Range("OClass").Cells(rowToWrite, 2).Value = expected
        
        On Error Resume Next
          nbObjectInClass = SelectedArea.Cells.SpecialCells(xlCellTypeVisible).Count
          
          If Err Then
            wSheet.Range("OClass").Cells(rowToWrite, 3).Value = 0
          Else
            wSheet.Range("OClass").Cells(rowToWrite, 3).Value = nbObjectInClass
          End If
          
    Next
    
    'Re-init row index
    rowToWrite = 0
    
    For i = 1 To NbreClasse
          rowToWrite = 1 + i
          objectExpected = wSheet.Range("OClass").Cells(rowToWrite, 2).Value
          objectReal = wSheet.Range("OClass").Cells(rowToWrite, 3).Value
          
          'Observed - expected
          diff = objectReal - objectExpected
          wSheet.Range("OClass").Cells(rowToWrite, 4).Value = diff
          
          'SQRT
          sqrtExpected = Sqr(expected)
          wSheet.Range("OClass").Cells(rowToWrite, 5).Value = sqrtExpected
          
          'Eij
           eij = diff / sqrtExpected
           wSheet.Range("OClass").Cells(rowToWrite, 6).Value = eij
        
           'Vij
           vij = 1 - (objectReal / NbObjets)
           wSheet.Range("OClass").Cells(rowToWrite, 7).Value = vij
           
           'sqrVij
           sqrVij = Sqr(vij)
           wSheet.Range("OClass").Cells(rowToWrite, 8).Value = sqrVij
           
           'dij
           dij = eij / sqrVij
           wSheet.Range("OClass").Cells(rowToWrite, 9).Value = dij
           
           'NORMDIST
           normDist = WorksheetFunction.normDist(dij, 0, 1, 1)
           wSheet.Range("OClass").Cells(rowToWrite, 10).Value = normDist
           
           'Compute color by line
           Dim rowToColor As Range
           Dim myColor As String
           Dim valueToTest As Double
            
           If normDist < 0 Then
                wSheet.Range("OClass").Cells(rowToWrite, 11).Value = normDist
           Else
                normDistReduced = 1 - normDist
                wSheet.Range("OClass").Cells(rowToWrite, 11).Value = normDistReduced
           End If
           
           valueToTest = wSheet.Range("OClass").Cells(rowToWrite, 11).Value
           
           If valueToTest <= 0.001 Then
                myColor = "green"
           ElseIf valueToTest <= 0.01 Then
                myColor = "darkblue"
           ElseIf valueToTest <= 0.05 Then
                myColor = "softblue"
           Else
            If valueToTest > 0.1 Then
                myColor = "red"
            ElseIf valueToTest > 0.05 Then
                myColor = "orange"
            End If
           End If
           
           Call colorLine(wSheet, (2 + rowToWrite), myColor)
           
    Next
    
    'vert: inf. ou = à 0,001 ;  bleu foncé :inf ou = à 0,01 ; bleu clair: inf. ou = à 0,05 : orange = sup à 0,05 ; rouge : sup à 0,1

    'Clean Display
    
    For Each fitColumn In firstRow
        fitColumn.Font.Bold = True
        fitColumn.Font.ThemeColor = xlThemeColorLight1
        
        fitColumn.EntireColumn.AutoFit
    Next
    
    
    dSheet.ShowAllData
    
    
End Sub

Public Sub colorLine(wSheet As Worksheet, index As Integer, color As String)
    
    Dim selected As Integer
    Dim shade As Double
    shade = 0
    
    If color = "red" Then
        selected = xlThemeColorAccent2
    ElseIf color = "green" Then
        selected = xlThemeColorAccent3
    ElseIf color = "softblue" Then
        selected = xlThemeColorAccent1
        shade = 0.399975585192419
    ElseIf color = "darkblue" Then
        selected = xlThemeColorAccent1
    Else
        selected = xlThemeColorAccent6
    End If
    
    wSheet.Range(wSheet.Cells(index, 1), wSheet.Cells(index, wSheet.UsedRange.Columns.Count)).Select
    
    'Set = wSheet.Range(index, wSheet.Cells(index, wSheet.Columns.Count).End(xlToLeft))
    'MsgBox wSheet.Columns.Count
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = selected
        .TintAndShade = shade
        .PatternTintAndShade = 0
    End With
    
    
    
End Sub


