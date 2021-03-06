﻿Imports System.Collections.Generic
Imports ExcelDna.Integration
Imports ExcelDna.Integration.XlCall


' This class defines a few test functions that can be used to explore the automatic array resizing.
Public Module ResizeTestFunctions

        ' Just returns an array of the given size
        Public Function MakeArray(rows As Integer, columns As Integer) As Object
            Dim result As Object(,) = New Object(rows - 1, columns - 1) {}
            For i As Integer = 0 To rows - 1
                For j As Integer = 0 To columns - 1
                    result(i, j) = i + j
                Next
            Next

            Return result
        End Function

        Public Function MakeArrayDoubles(rows As Integer, columns As Integer) As Double(,)
            Dim result As Double(,) = New Double(rows - 1, columns - 1) {}
            For i As Integer = 0 To rows - 1
                For j As Integer = 0 To columns - 1
                    result(i, j) = i + (j / 1000.0)
                Next
            Next

            Return result
        End Function

        ' Makes an array, but automatically resizes the result
        Public Function MakeArrayAndResize(rows As Integer, columns As Integer) As Object
            Dim result As Object = MakeArray(rows, columns)

            ' Can also call Resize via Excel - so if the Resize add-in is not part of this code, it should still work
            ' (though calling direct is better for large arrays - it prevents extra marshaling).
            ' Return XlCall.Excel(XlCall.xlUDF, "Resize", result)

            Return ArrayResizer.Resize(result)
        End Function

        Public Function MakeArrayAndResizeDoubles(rows As Integer, columns As Integer) As Double(,)
            Dim result As Double(,) = MakeArrayDoubles(rows, columns)
            ' Call Resize via Excel - so if the Resize add-in is not part of this code, it should still work.
            Return ArrayResizer.ResizeDoubles(result)
        End Function

    End Module

    Public Module ArrayResizer

        ' This function will run in the UDF context.
        ' Needs extra protection to allow multithreaded use.
        Public Function Resize(array As Object(,)) As Object
            Dim caller As ExcelReference = TryCast(Excel(xlfCaller), ExcelReference)
            If caller Is Nothing Then
                Return array
            End If

            Dim rows As Integer = array.GetLength(0)
            Dim columns As Integer = array.GetLength(1)

            If rows = 0 OrElse columns = 0 Then
                Return array
            End If

            If (caller.RowLast - caller.RowFirst + 1 = rows) AndAlso (caller.ColumnLast - caller.ColumnFirst + 1 = columns) Then
                ' Size is already OK - just return result
                Return array
            End If

            Dim rowLast = caller.RowFirst + rows - 1
            Dim columnLast = caller.ColumnFirst + columns - 1

            If rowLast > ExcelDnaUtil.ExcelLimits.MaxRows - 1 OrElse columnLast > ExcelDnaUtil.ExcelLimits.MaxColumns - 1 Then
                ' Can't resize - goes beyond the end of the sheet - just return #VALUE
                ' (Can't give message here, or change cells)
                Return ExcelError.ExcelErrorValue
            End If

            ' TODO: Add guard for ever-changing result?
            ExcelAsyncUtil.QueueAsMacro(
          Sub()
                ' Create a reference of the right size
                Dim target = New ExcelReference(caller.RowFirst, rowLast, caller.ColumnFirst, columnLast, caller.SheetId)
                ' Will trigger a recalc by writing formula
                DoResize(target)
          End Sub)

            ' Return what we have - to prevent flashing #N/A
            Return array
        End Function

        Public Function ResizeDoubles(array As Double(,)) As Double(,)
            Dim caller As ExcelReference = TryCast(Excel(xlfCaller), ExcelReference)
            If caller Is Nothing Then
                Return array
            End If

            Dim rows As Integer = array.GetLength(0)
            Dim columns As Integer = array.GetLength(1)

            If rows = 0 OrElse columns = 0 Then
                Return array
            End If

            If (caller.RowLast - caller.RowFirst + 1 = rows) AndAlso (caller.ColumnLast - caller.ColumnFirst + 1 = columns) Then
                ' Size is already OK - just return result
                Return array
            End If

            Dim rowLast = caller.RowFirst + rows - 1
            Dim columnLast = caller.ColumnFirst + columns - 1

            If rowLast > ExcelDnaUtil.ExcelLimits.MaxRows - 1 OrElse columnLast > ExcelDnaUtil.ExcelLimits.MaxColumns - 1 Then
                ' Can't resize - goes beyond the end of the sheet - just return null (for #NUM!)
                ' (Can't give message here, or change cells)
                Return Nothing
            End If

            ' TODO: Add guard for ever-changing result?
            ExcelAsyncUtil.QueueAsMacro(
          Sub()
                ' Create a reference of the right size
                Dim target = New ExcelReference(caller.RowFirst, rowLast, caller.ColumnFirst, columnLast, caller.SheetId)
                ' Will trigger a recalc by writing formula
                DoResize(target)
          End Sub)

            ' Return what we have - to prevent flashing #N/A
            Return array
        End Function

        Private Sub DoResize(target As ExcelReference)
            ' Get the current state for reset later
            Dim oldEcho As Object = Excel(xlfGetWorkspace, 40)
            Dim oldCalculationMode As Object = Excel(xlfGetDocument, 14)
            Try
                Excel(xlcEcho, False)
                Excel(xlcOptionsCalculation, 3)

                Dim firstCell As New ExcelReference(target.RowFirst, target.RowFirst, target.ColumnFirst, target.ColumnFirst, target.SheetId)

                ' Get the formula in the first cell of the target
                Dim formula As String = DirectCast(Excel(xlfGetCell, 41, firstCell), String)
                Dim isFormulaArray As Boolean = CBool(Excel(xlfGetCell, 49, firstCell))
                If isFormulaArray Then
                    Dim oldSelectionOnActiveSheet As Object = Excel(xlfSelection)
                    Dim oldActiveCell As Object = Excel(xlfActiveCell)

                    ' Remember old selection and select the first cell of the target
                    Dim firstCellSheet As String = DirectCast(Excel(xlSheetNm, firstCell), String)
                    Excel(xlcWorkbookSelect, New Object() {firstCellSheet})
                    Dim oldSelectionOnArraySheet As Object = Excel(xlfSelection)
                    Excel(xlcFormulaGoto, firstCell)

                    ' Extend the selection to the whole array and clear
                    Excel(xlcSelectSpecial, 6)
                    Dim oldArray As ExcelReference = DirectCast(Excel(xlfSelection), ExcelReference)

                    oldArray.SetValue(ExcelEmpty.Value)
                    Excel(xlcSelect, oldSelectionOnArraySheet)
                    Excel(xlcFormulaGoto, oldSelectionOnActiveSheet)
                End If
                ' Get the formula and convert to R1C1 mode
                Dim isR1C1Mode As Boolean = CBool(Excel(xlfGetWorkspace, 4))
                Dim formulaR1C1 As String = formula
                If Not isR1C1Mode Then
                ' Set the formula into the whole target
                Dim formulaR1C1Obj As Object = {}
                Dim formulaR1C1Return As XlReturn
                    formulaR1C1Return = TryExcel(xlfFormulaConvert, formulaR1C1Obj, formula, True, False, ExcelMissing.Value, firstCell)
                    If formulaR1C1Return <> XlReturn.XlReturnSuccess OrElse TypeOf formulaR1C1Obj Is ExcelError Then
                        Dim firstCellAddress As String
                        firstCellAddress = CStr(Excel(xlfReftext, firstCell, True))
                        Excel(xlcAlert, "Cannot resize array formula at " & firstCellAddress & " - formula might be too long when converted to R1C1 format.")
                        firstCell.SetValue("'" & formula)
                        Return
                    End If
                    formulaR1C1 = CStr(formulaR1C1Obj)
                End If
                ' Must be R1C1-style references
                Dim ignoredResult As Object = Nothing
                'Debug.Print("Resizing START: " + target.RowLast);
                Dim retval As XlReturn = TryExcel(xlcFormulaArray, ignoredResult, formulaR1C1, target)
                'Debug.Print("Resizing FINISH");

                ' TODO: Find some dummy macro to clear the undo stack

                If retval <> XlReturn.XlReturnSuccess Then
                    Dim firstCellAddress As String = DirectCast(Excel(xlfReftext, firstCell, True), String)
                    Excel(xlcAlert, (Convert.ToString("Cannot resize array formula at ") & firstCellAddress) + " - result might overlap another array.")
                    ' Might have failed due to array in the way.
                    firstCell.SetValue(Convert.ToString("'") & formula)
                End If
            Finally
                Excel(xlcEcho, oldEcho)
                Excel(xlcOptionsCalculation, oldCalculationMode)
            End Try
        End Sub
    End Module
