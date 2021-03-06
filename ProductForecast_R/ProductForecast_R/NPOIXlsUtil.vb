﻿Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports System.Linq
Imports System.Web
Imports NPOI
Imports NPOI.HPSF
Imports NPOI.HSSF
Imports NPOI.HSSF.UserModel
Imports NPOI.POIFS
Imports NPOI.Util
Public Class NPOIXlsUtil

    'Public Shared Sub DataTableToExcelDownload(dt As DataTable, FileName As String)
    '    Dim ms1 As IO.MemoryStream = RenderDataTableToExcel(dt)
    '    With HttpContext.Current.Response
    '        .Clear()
    '        '.ContentType = "application/vnd.ms-excel"
    '        .AddHeader("Content-Disposition", String.Format("attachment; filename={0};", FileName + ".xls"))
    '        .ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    '        .BinaryWrite(ms1.ToArray())
    '        .End()
    '    End With
    'End Sub

    Public Shared Function RenderDataSetToExcel(ds As DataSet) As MemoryStream
        Dim workbook As New HSSFWorkbook(), ms As New MemoryStream()
        For idx As Integer = 0 To ds.Tables.Count - 1
            Dim SourceTable As DataTable = ds.Tables(idx), sheetName As String = "sheet" + (idx + 1).ToString()
            If Not String.IsNullOrEmpty(SourceTable.TableName) Then sheetName = SourceTable.TableName
            Dim sheet As HSSFSheet = workbook.CreateSheet(sheetName), headerRow As HSSFRow = sheet.CreateRow(0)

            ' handling header. 
            For Each column As DataColumn In SourceTable.Columns
                headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName)
            Next
            ' handling value. 
            Dim rowIndex As Integer = 1

            For Each row As DataRow In SourceTable.Rows
                Dim dataRow As HSSFRow = sheet.CreateRow(rowIndex)
                For Each column As DataColumn In SourceTable.Columns
                    dataRow.CreateCell(column.Ordinal).SetCellValue(row(column).ToString())
                Next
                rowIndex += 1
            Next
        Next
       

        workbook.Write(ms) : ms.Flush() : ms.Position = 0
        'sheet = Nothing : headerRow = Nothing : workbook = Nothing
        Return ms
    End Function

    Public Shared Function RenderDataTableToExcel(SourceTable As DataTable) As MemoryStream
        Dim workbook As New HSSFWorkbook(), ms As New MemoryStream(), sheet As HSSFSheet = workbook.CreateSheet(), headerRow As HSSFRow = sheet.CreateRow(0)
        ' handling header. 
        For Each column As DataColumn In SourceTable.Columns
            headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName)
        Next
        ' handling value. 
        Dim rowIndex As Integer = 1

        For Each row As DataRow In SourceTable.Rows
            Dim dataRow As HSSFRow = sheet.CreateRow(rowIndex)
            For Each column As DataColumn In SourceTable.Columns
                dataRow.CreateCell(column.Ordinal).SetCellValue(row(column).ToString())
            Next
            rowIndex += 1
        Next

        workbook.Write(ms) : ms.Flush() : ms.Position = 0
        sheet = Nothing : headerRow = Nothing : workbook = Nothing
        Return ms
    End Function

    Public Shared Sub RenderDataSetToExcel(ds As DataSet, FileName As String)
        Dim ms As MemoryStream = TryCast(RenderDataSetToExcel(ds), MemoryStream)
        Dim fs As New FileStream(FileName, FileMode.Create, FileAccess.Write)
        Dim data As Byte() = ms.ToArray()
        fs.Write(data, 0, data.Length) : fs.Flush() : fs.Close()
    End Sub

    Public Shared Sub RenderDataTableToExcel(SourceTable As DataTable, FileName As String)
        Dim ms As MemoryStream = TryCast(RenderDataTableToExcel(SourceTable), MemoryStream)
        Dim fs As New FileStream(FileName, FileMode.Create, FileAccess.Write)
        Dim data As Byte() = ms.ToArray()

        fs.Write(data, 0, data.Length) : fs.Flush() : fs.Close()
        data = Nothing : ms = Nothing : fs = Nothing
    End Sub

    Public Shared Function RenderDataTableFromExcel(ExcelFileStream As Stream, SheetName As String, HeaderRowIndex As Integer) As DataTable
        Dim workbook As New HSSFWorkbook(ExcelFileStream)
        Dim sheet As HSSFSheet = workbook.GetSheet(SheetName)

        Dim table As New DataTable()

        Dim headerRow As HSSFRow = sheet.GetRow(HeaderRowIndex)
        Dim cellCount As Integer = headerRow.LastCellNum

        For i As Integer = headerRow.FirstCellNum To cellCount - 1
            Dim column As New DataColumn(headerRow.GetCell(i).StringCellValue)
            table.Columns.Add(column)
        Next

        Dim rowCount As Integer = sheet.LastRowNum

        For i As Integer = (sheet.FirstRowNum + 1) To sheet.LastRowNum - 1
            Dim row As HSSFRow = sheet.GetRow(i)
            Dim dataRow As DataRow = table.NewRow()

            For j As Integer = row.FirstCellNum To cellCount - 1
                dataRow(j) = row.GetCell(j).ToString()
            Next
        Next

        ExcelFileStream.Close()
        workbook = Nothing
        sheet = Nothing
        Return table
    End Function

    Public Shared Function ExcelFilePathToDataTable(XlsPath As String) As DataTable
        Dim ms As New IO.MemoryStream()
        Using file As New IO.FileStream(XlsPath, IO.FileMode.Open, IO.FileAccess.Read)
            Dim bytes As Byte() = New Byte(file.Length - 1) {}
            file.Read(bytes, 0, CInt(file.Length))
            ms.Write(bytes, 0, CInt(file.Length))
        End Using

        Dim dt As DataTable = NPOIXlsUtil.RenderDataTableFromExcel(ms, 0, 0)
        Return dt
    End Function

    Public Shared Function ExcelToDataSet(XlsPath As String) As DataSet
        Dim ms As New IO.MemoryStream()
        Using file As New IO.FileStream(XlsPath, IO.FileMode.Open, IO.FileAccess.Read)
            Dim bytes As Byte() = New Byte(file.Length - 1) {}
            file.Read(bytes, 0, CInt(file.Length))
            ms.Write(bytes, 0, CInt(file.Length))
        End Using
        Dim workbook As New HSSFWorkbook(ms)
        Dim ds As New DataSet
        For curSheetIdx As Integer = 0 To workbook.Count - 1
            Dim sheet As HSSFSheet = workbook.GetSheetAt(curSheetIdx)

            Dim table As New DataTable()

            Dim headerRow As HSSFRow = sheet.GetRow(0)
            Dim cellCount As Integer = headerRow.LastCellNum

            For i As Integer = headerRow.FirstCellNum To cellCount - 1
                Dim column As New DataColumn(headerRow.GetCell(i).StringCellValue)
                table.Columns.Add(column)
            Next

            Dim rowCount As Integer = sheet.LastRowNum

            For i As Integer = (sheet.FirstRowNum + 1) To sheet.LastRowNum
                Dim row As HSSFRow = sheet.GetRow(i)
                Dim dataRow As DataRow = table.NewRow()

                For j As Integer = row.FirstCellNum To cellCount - 1
                    If row.GetCell(j) IsNot Nothing Then
                        dataRow(j) = row.GetCell(j).ToString()
                    End If
                Next

                table.Rows.Add(dataRow)
            Next
            table.TableName = sheet.SheetName
            ds.Tables.Add(table)
        Next
        ms.Close()
        Return ds
    End Function

    Public Shared Function RenderDataTableFromExcel(ExcelFileStream As Stream, SheetIndex As Integer, HeaderRowIndex As Integer) As DataTable
        Dim workbook As New HSSFWorkbook(ExcelFileStream)
        Dim sheet As HSSFSheet = workbook.GetSheetAt(SheetIndex)

        Dim table As New DataTable()

        Dim headerRow As HSSFRow = sheet.GetRow(HeaderRowIndex)
        Dim cellCount As Integer = headerRow.LastCellNum

        For i As Integer = headerRow.FirstCellNum To cellCount - 1
            Dim column As New DataColumn(headerRow.GetCell(i).StringCellValue)
            table.Columns.Add(column)
        Next

        Dim rowCount As Integer = sheet.LastRowNum

        For i As Integer = (sheet.FirstRowNum + 1) To sheet.LastRowNum
            Dim row As HSSFRow = sheet.GetRow(i)
            Dim dataRow As DataRow = table.NewRow()

            For j As Integer = row.FirstCellNum To cellCount - 1
                If row.GetCell(j) IsNot Nothing Then
                    dataRow(j) = row.GetCell(j).ToString()
                End If
            Next

            table.Rows.Add(dataRow)
        Next

        ExcelFileStream.Close()
        workbook = Nothing
        sheet = Nothing
        Return table
    End Function

    'Public Shared Sub DataTableToExcel2007Download(ByRef dt As DataTable, ByRef FileName As String)
    '    Dim wb As New XSSF.UserModel.XSSFWorkbook()

    '    'Dim dt As DataTable = ds.Tables(tbIdx)
    '    Dim ws As SS.UserModel.ISheet
    '    If dt.TableName <> String.Empty Then
    '        ws = wb.CreateSheet(dt.TableName)
    '    Else
    '        ws = wb.CreateSheet("Sheet1")
    '    End If

    '    ws.CreateRow(0)
    '    '第一行為欄位名稱
    '    For i As Integer = 0 To dt.Columns.Count - 1
    '        ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns(i).ColumnName)
    '    Next

    '    For i As Integer = 0 To dt.Rows.Count - 1
    '        ws.CreateRow(i + 1)
    '        For j As Integer = 0 To dt.Columns.Count - 1
    '            ws.GetRow(i + 1).CreateCell(j).SetCellValue(dt.Rows(i)(j).ToString())
    '        Next
    '    Next


    '    Dim ms As New IO.MemoryStream()
    '    '產生檔案
    '    wb.Write(ms) : ms.Flush()
    '    'ms.Position = 0

    '    With HttpContext.Current.Response
    '        .Clear()
    '        '.ContentType = "application/vnd.ms-excel"
    '        .AddHeader("Content-Disposition", String.Format("attachment; filename={0};", FileName + ".xlsx"))
    '        .ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    '        .BinaryWrite(ms.ToArray())
    '        .End()
    '    End With

    'End Sub
    'Public Shared Sub DataSetToExcel2007Download(ByRef ds As DataSet, ByRef FileName As String)
    '    Dim wb As New XSSF.UserModel.XSSFWorkbook()
    '    For tbIdx As Integer = 0 To ds.Tables.Count - 1
    '        Dim dt As DataTable = ds.Tables(tbIdx)
    '        Dim ws As SS.UserModel.ISheet
    '        If dt.TableName <> String.Empty Then
    '            ws = wb.CreateSheet(dt.TableName)
    '        Else
    '            ws = wb.CreateSheet("Sheet" + (tbIdx + 1).ToString())
    '        End If

    '        ws.CreateRow(0)
    '        '第一行為欄位名稱
    '        For i As Integer = 0 To dt.Columns.Count - 1
    '            ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns(i).ColumnName)
    '        Next

    '        For i As Integer = 0 To dt.Rows.Count - 1
    '            ws.CreateRow(i + 1)
    '            For j As Integer = 0 To dt.Columns.Count - 1
    '                ws.GetRow(i + 1).CreateCell(j).SetCellValue(dt.Rows(i)(j).ToString())
    '            Next
    '        Next
    '    Next

    '    Dim ms As New IO.MemoryStream()
    '    '產生檔案
    '    wb.Write(ms) : ms.Flush()
    '    'ms.Position = 0

    '    With HttpContext.Current.Response
    '        .Clear()
    '        '.ContentType = "application/vnd.ms-excel"
    '        .AddHeader("Content-Disposition", String.Format("attachment; filename={0};", FileName + ".xlsx"))
    '        .ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    '        .BinaryWrite(ms.ToArray())
    '        .End()
    '    End With

    'End Sub

End Class
