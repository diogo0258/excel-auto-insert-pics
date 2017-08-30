Attribute VB_Name = "Módulo2"
Sub AutoAddPicsToCellsBasedOnFileNames()
    ' following tuto from here: http://www.robvanderwoude.com/vbstech_regexp.php
    Dim dirPath As String
    Dim fileName As String

    Dim sheetNum As Integer
    Dim cellId As String
    
    Dim validFileNameRe As Object
    Dim ReMatch As Object
    
    dirPath = Application.ThisWorkbook.Path & "\renamed-pics"
    ' thisworkbook refers to the file containing the macro.
    
    Set validFileNameRe = CreateObject("vbscript.regexp")
    With validFileNameRe
        .Pattern = "^(\d+) ([A-Z]+\d+)(\..+)$"
        .MultiLine = False
        .IgnoreCase = True
        .Global = False
    End With
    
    fileName = Dir(dirPath & "\*", vbNormal) ' if there's more than one match, returns only one. Subsequent Dir() calls will return the others
    While fileName <> ""
        Set ReMatch = validFileNameRe.Execute(fileName)
        
        ' We should get only 1 match since the Global property is FALSE
        If ReMatch.Count = 1 Then
            sheetNum = ReMatch.Item(0).Submatches(0)
            cellId = ReMatch.Item(0).Submatches(1)
            
            Sheets(sheetNum).Activate  ' AddPicByPathAndFitToCell uses activesheet
            Call AddPicByPathAndFitToCell(dirPath & "\" & fileName, Range(cellId))
        End If
        
        Set ReMatch = Nothing
        fileName = Dir()
    Wend
    
    Set validFileNameRe = Nothing
End Sub
