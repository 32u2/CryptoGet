Imports System
Imports System.Collections.Generic
Imports ExcelDna.Integration

' includes examples 
Namespace AsyncFunctions
    Module ResizeTestFunctions
        Function MakeArray(ByVal rows As Integer, ByVal columns As Integer) As Object(,)
            Dim result As Object(,) = New Object(rows - 1, columns - 1) {}

            For i As Integer = 0 To rows - 1

                For j As Integer = 0 To columns - 1
                    result(i, j) = i + j
                Next
            Next

            Return result
        End Function

        Function MakeArrayDoubles(ByVal rows As Integer, ByVal columns As Integer) As Double(,)
            Dim result As Double(,) = New Double(rows - 1, columns - 1) {}

            For i As Integer = 0 To rows - 1

                For j As Integer = 0 To columns - 1
                    result(i, j) = i + (j / 1000.0)
                Next
            Next

            Return result
        End Function

        Function MakeMixedArrayAndResize(ByVal rows As Integer, ByVal columns As Integer) As Object
            Dim result As Object(,) = New Object(rows - 1, columns - 1) {}

            For j As Integer = 0 To columns - 1
                result(0, j) = "Col " & j
            Next

            For i As Integer = 1 To rows - 1

                For j As Integer = 0 To columns - 1
                    result(i, j) = i + (j * 0.1)
                Next
            Next

            Return ArrayResizer.Resize(result)
        End Function

        Function MakeArrayAndResize(ByVal rows As Integer, ByVal columns As Integer, ByVal unused As String, ByVal unusedtoo As String) As Object
            Dim result As Object(,) = MakeArray(rows, columns)
            Return ArrayResizer.Resize(result)
        End Function

        Function MakeArrayAndResizeDoubles(ByVal rows As Integer, ByVal columns As Integer) As Double(,)
            Dim result As Double(,) = MakeArrayDoubles(rows, columns)
            Return ArrayResizer.ResizeDoubles(result)
        End Function
    End Module

    Public Class ArrayResizer
        Inherits XlCall

        Public Shared Function Resize(ByVal array As Object(,)) As Object
            Dim caller = TryCast(Excel(xlfCaller), ExcelReference)
            If caller Is Nothing Then Return array
            Dim rows As Integer = array.GetLength(0)
            Dim columns As Integer = array.GetLength(1)
            If rows = 0 OrElse columns = 0 Then Return array

            If (caller.RowLast - caller.RowFirst + 1 = rows) AndAlso (caller.ColumnLast - caller.ColumnFirst + 1 = columns) Then
                Return array
            End If

            Dim rowLast = caller.RowFirst + rows - 1
            Dim columnLast = caller.ColumnFirst + columns - 1

            If rowLast > ExcelDnaUtil.ExcelLimits.MaxRows - 1 OrElse columnLast > ExcelDnaUtil.ExcelLimits.MaxColumns - 1 Then
                Return ExcelError.ExcelErrorValue
            End If

            ExcelAsyncUtil.QueueAsMacro(Function()
                                            Dim target = New ExcelReference(caller.RowFirst, rowLast, caller.ColumnFirst, columnLast, caller.SheetId)
                                            DoResize(target)
                                        End Function)
            Return array
        End Function

        Public Shared Function ResizeDoubles(ByVal array As Double(,)) As Double(,)
            Dim caller = TryCast(Excel(xlfCaller), ExcelReference)
            If caller Is Nothing Then Return array
            Dim rows As Integer = array.GetLength(0)
            Dim columns As Integer = array.GetLength(1)
            If rows = 0 OrElse columns = 0 Then Return array

            If (caller.RowLast - caller.RowFirst + 1 = rows) AndAlso (caller.ColumnLast - caller.ColumnFirst + 1 = columns) Then
                Return array
            End If

            Dim rowLast = caller.RowFirst + rows - 1
            Dim columnLast = caller.ColumnFirst + columns - 1

            If rowLast > ExcelDnaUtil.ExcelLimits.MaxRows - 1 OrElse columnLast > ExcelDnaUtil.ExcelLimits.MaxColumns - 1 Then
                Return Nothing
            End If

            ExcelAsyncUtil.QueueAsMacro(Function()
                                            Dim target = New ExcelReference(caller.RowFirst, rowLast, caller.ColumnFirst, columnLast, caller.SheetId)
                                            DoResize(target)
                                        End Function)
            Return array
        End Function

        Private Shared Sub DoResize(ByVal target As ExcelReference)
            Using New ExcelEchoOffHelper()

                Using New ExcelCalculationManualHelper()
                    Dim firstCell As ExcelReference = New ExcelReference(target.RowFirst, target.RowFirst, target.ColumnFirst, target.ColumnFirst, target.SheetId)
                    Dim formula As String = CStr(Excel(xlfGetCell, 41, firstCell))
                    Dim isFormulaArray As Boolean = CBool(Excel(xlfGetCell, 49, firstCell))

                    If isFormulaArray Then

                        Using New ExcelSelectionHelper(firstCell)
                            Excel(xlcSelectSpecial, 6)
                            Dim oldArray As ExcelReference = CType(Excel(xlfSelection), ExcelReference)
                            oldArray.SetValue(ExcelEmpty.Value)
                        End Using
                    End If

                    Dim isR1C1Mode As Boolean = CBool(Excel(xlfGetWorkspace, 4))
                    Dim formulaR1C1 As String = formula

                    If Not isR1C1Mode Then
                        Dim formulaR1C1Obj As Object
                        Dim formulaR1C1Return As XlReturn = TryExcel(xlfFormulaConvert, formulaR1C1Obj, formula, True, False, ExcelMissing.Value, firstCell)

                        If formulaR1C1Return <> XlReturn.XlReturnSuccess OrElse TypeOf formulaR1C1Obj Is ExcelError Then
                            Dim firstCellAddress As String = CStr(Excel(xlfReftext, firstCell, True))
                            Excel(xlcAlert, "Cannot resize array formula at " & firstCellAddress & " - formula might be too long when converted to R1C1 format.")
                            firstCell.SetValue("'" & formula)
                            Return
                        End If

                        formulaR1C1 = CStr(formulaR1C1Obj)
                    End If

                    Dim ignoredResult As Object
                    Dim formulaArrayReturn As XlReturn = TryExcel(xlcFormulaArray, ignoredResult, formulaR1C1, target)

                    If formulaArrayReturn <> XlReturn.XlReturnSuccess Then
                        Dim firstCellAddress As String = CStr(Excel(xlfReftext, firstCell, True))
                        Excel(xlcAlert, "Cannot resize array formula at " & firstCellAddress & " - result might overlap another array.")
                        firstCell.SetValue("'" & formula)
                    End If
                End Using
            End Using
        End Sub
    End Class

    Public Class ExcelEchoOffHelper
        Inherits XlCall
        Implements IDisposable

        Private oldEcho As Object

        Public Sub New()
            oldEcho = Excel(xlfGetWorkspace, 40)
            Excel(xlcEcho, False)
        End Sub

        Public Sub Dispose()
            Excel(xlcEcho, oldEcho)
        End Sub

        Private Sub IDisposable_Dispose() Implements IDisposable.Dispose
            Dispose()
        End Sub
    End Class

    Public Class ExcelCalculationManualHelper
        Inherits XlCall
        Implements IDisposable

        Private oldCalculationMode As Object

        Public Sub New()
            oldCalculationMode = Excel(xlfGetDocument, 14)
            Excel(xlcOptionsCalculation, 3)
        End Sub

        Public Sub Dispose()
            Excel(xlcOptionsCalculation, oldCalculationMode)
        End Sub

        Private Sub IDisposable_Dispose() Implements IDisposable.Dispose
            Dispose()
        End Sub
    End Class

    Public Class ExcelSelectionHelper
        Inherits XlCall
        Implements IDisposable

        Private oldSelectionOnActiveSheet As Object
        Private oldActiveCellOnActiveSheet As Object
        Private oldSelectionOnRefSheet As Object
        Private oldActiveCellOnRefSheet As Object

        Public Sub New(ByVal refToSelect As ExcelReference)
            oldSelectionOnActiveSheet = Excel(xlfSelection)
            oldActiveCellOnActiveSheet = Excel(xlfActiveCell)
            Dim refSheet As String = CStr(Excel(xlSheetNm, refToSelect))
            Excel(xlcWorkbookSelect, New Object() {refSheet})
            oldSelectionOnRefSheet = Excel(xlfSelection)
            oldActiveCellOnRefSheet = Excel(xlfActiveCell)
            Excel(xlcFormulaGoto, refToSelect)
        End Sub

        Public Sub Dispose()
            Excel(xlcSelect, oldSelectionOnRefSheet, oldActiveCellOnRefSheet)
            Dim oldActiveSheet As String = CStr(Excel(xlSheetNm, oldSelectionOnActiveSheet))
            Excel(xlcWorkbookSelect, New Object() {oldActiveSheet})
            Excel(xlcSelect, oldSelectionOnActiveSheet, oldActiveCellOnActiveSheet)
        End Sub

        Private Sub IDisposable_Dispose() Implements IDisposable.Dispose
            Dispose()
        End Sub
    End Class
End Namespace
