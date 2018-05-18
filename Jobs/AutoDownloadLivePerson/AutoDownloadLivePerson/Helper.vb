Imports System.Reflection

Public Class Helper
    Public Shared Function ListToDataTable(Of T)(data As IList(Of T)) As DataTable
        Dim properties As ComponentModel.PropertyDescriptorCollection = ComponentModel.TypeDescriptor.GetProperties(GetType(T))
        Dim table As New DataTable()
        For Each prop As ComponentModel.PropertyDescriptor In properties
            table.Columns.Add(prop.Name, If(Nullable.GetUnderlyingType(prop.PropertyType), prop.PropertyType))
        Next
        For Each item As T In data
            Dim row As DataRow = table.NewRow()
            For Each prop As ComponentModel.PropertyDescriptor In properties
                row(prop.Name) = If(prop.GetValue(item), DBNull.Value)
            Next
            table.Rows.Add(row)
        Next
        Return table

    End Function

    Public Shared Function DataTableToList(Of T As New)(table As DataTable) As IList(Of T)
        Dim properties As IList(Of PropertyInfo) = GetType(T).GetProperties().ToList()
        Dim result As IList(Of T) = New List(Of T)()

        '取得DataTable所有的row data
        For Each row In table.Rows
            Dim item = MappingItem(Of T)(DirectCast(row, DataRow), properties)
            result.Add(item)
        Next

        Return result
    End Function

    Private Shared Function MappingItem(Of T As New)(row As DataRow, properties As IList(Of PropertyInfo)) As T
        Dim item As New T()
        For Each [property] In properties
            If row.Table.Columns.Contains([property].Name) Then
                '針對欄位的型態去轉換
                If [property].PropertyType = GetType(DateTime) Then
                    Dim dt As New DateTime()
                    If DateTime.TryParse(row([property].Name).ToString(), dt) Then
                        [property].SetValue(item, dt, Nothing)
                    Else
                        [property].SetValue(item, Nothing, Nothing)
                    End If
                ElseIf [property].PropertyType = GetType(Decimal) Then
                    Dim val As New Decimal()
                    Decimal.TryParse(row([property].Name).ToString(), val)
                    [property].SetValue(item, val, Nothing)
                ElseIf [property].PropertyType = GetType(Double) Then
                    Dim val As New Double()
                    Double.TryParse(row([property].Name).ToString(), val)
                    [property].SetValue(item, val, Nothing)
                ElseIf [property].PropertyType = GetType(Integer) Then
                    Dim val As New Integer()
                    Integer.TryParse(row([property].Name).ToString(), val)
                    [property].SetValue(item, val, Nothing)
                Else
                    If row([property].Name) IsNot DBNull.Value Then
                        [property].SetValue(item, row([property].Name), Nothing)
                    End If
                End If
            End If
        Next
        Return item
    End Function

End Class
