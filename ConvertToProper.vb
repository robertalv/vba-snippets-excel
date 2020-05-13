Sub ConvertToProper()
    Dim ws As Object
    Dim LCell As Range
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationSemiautomatic
    
    For Each ws In ActiveWorkbook.Sheets
        On Error Resume Next
        ws.Activate
        
        For Each LCell In Cells.SpecialCells(xlConstants, xlTextValues)
            LCell.Formula = StrConv(LCell.Formula, vbProperCase)
        Next
    Next ws
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub