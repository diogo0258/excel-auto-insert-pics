Attribute VB_Name = "Módulo1"

Sub FitShapeToRange(MyShape As Shape, MyRange As Range)
' some ideas from code by JoeMo: https://www.mrexcel.com/forum/excel-questions/802311-resize-picture-macro-excel-2007-a.html
    Dim RangeSizeRatio As Single
    Dim ShapeSizeRatio As Single
    Dim ShapeIsRotated As Boolean
    
    RangeSizeRatio = MyRange.Width / MyRange.Height
    
    With MyShape
        .LockAspectRatio = msoTrue
        ShapeIsRotated = .Rotation = 90 Or .Rotation = 270
        
        If ShapeIsRotated Then
            ShapeSizeRatio = .Height / .Width
            If ShapeSizeRatio < RangeSizeRatio Then
                .Width = MyRange.Height
            Else
                .Height = MyRange.Width
            End If
        Else
            ShapeSizeRatio = .Width / .Height
            If ShapeSizeRatio > RangeSizeRatio Then
                .Width = MyRange.Width
            Else
                .Height = MyRange.Height
            End If
        End If
        
        .Top = MyRange.Top + MyRange.Height / 2 - .Height / 2
        .Left = MyRange.Left + MyRange.Width / 2 - .Width / 2
    End With
End Sub


Sub FitSelectedPicToTopLeftCell()  ' topleftcell can be a merged cell
    Dim MyRange As Range
    Dim MyShape As Shape
    
    If TypeName(Selection) = "Picture" Then
        Set MyRange = Selection.TopLeftCell.MergeArea
        Set MyShape = Selection.ShapeRange(1)
        
        Call FitShapeToRange(MyShape, MyRange)
    Else
        MsgBox "Select one picture before running this macro"
    End If
End Sub


Function AddPicByPathAndFitToCell(FilePath As String, Cell As Range) ' adds to activesheet; cell can be a merged cell, so use mergearea
' comment from https://answers.microsoft.com/en-us/office/forum/office_2010-customize/vba-changes-to-picturesinsert-shapeaddpicture/a14e60bc-a777-41f9-a4b3-d18a7a33beae
' In Excel 2010-VBA script. Object Pictures.Insert method now just inserts a path instead of a picture (2007 inserted picture).
' The only workaround appears to be to use the shape objects. Shape.AddPicture that has a parameter to "Save with Document"
    Dim MyRange As Range
    Dim MyShape As Shape
    
    Set MyRange = Cell.MergeArea
    Set MyShape = ActiveSheet.Shapes.AddPicture(fileName:=FilePath, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
            Left:=0, Top:=0, Width:=-1, Height:=-1)
    
    Call FitShapeToRange(MyShape, MyRange)
End Function


Function SelectFile() As String  ' returns only 1st selected
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .AllowMultiSelect = False
        If .Show <> -1 Then
            SelectFile = ""
        Else
            SelectFile = fd.SelectedItems(1)
        End If
    End With
End Function


Sub ExploreForPicAndFitToActiveCell()
Attribute ExploreForPicAndFitToActiveCell.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim FilePath As String
    
    FilePath = SelectFile()
    If FilePath <> "" Then
        Call AddPicByPathAndFitToCell(FilePath, ActiveCell)
    End If
End Sub


Sub AddOrFitPicToCell()
Attribute AddOrFitPicToCell.VB_ProcData.VB_Invoke_Func = "e\n14"
    If TypeOf Selection Is Range Then
        Call ExploreForPicAndFitToActiveCell
    ElseIf TypeOf Selection Is Picture Then
        Call FitSelectedPicToTopLeftCell
    End If
End Sub
