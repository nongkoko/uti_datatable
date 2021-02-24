Option Strict On
Option Compare Text

Imports System.Web.UI.WebControls

Public Class datatableUtility

    Public Event onBeforeDrawRow(ByVal iRowIndex As Integer)

    Private _rowStyle As String
    Private _colStyle As New Dictionary(Of Integer, String)

    Public ReadOnly Property colStyle As Dictionary(Of Integer, String)
        Get
            Return _colStyle
        End Get
    End Property

    Public Property rowStyle As String
        Get
            Return _rowStyle
        End Get
        Set(ByVal value As String)
            _rowStyle = value
        End Set
    End Property

    Public Shared Sub drawASPtable(ByVal iTable As Table, ByVal iDataTable As DataTable, ByVal iDataColumnNameToRender() As String)

        Dim aRow As TableRow
        Dim aCell As TableCell

        Dim aDict As New Dictionary(Of Integer, Integer) 'index kolom vs start number
        Dim aRunNumb As Integer
        Dim aAngka As Integer
        aRunNumb = 0
        For Each aDCN As String In iDataColumnNameToRender
            If aDCN Like "autoNumber*" Then
                aAngka = CInt(Strings.Split(aDCN, ":")(1))
                aDict.Add(aRunNumb, aAngka)
            End If
            aRunNumb += 1
        Next

        Dim aRowNumber As Integer
        aRowNumber = -1
        For Each aDR As DataRow In iDataTable.Rows
            aRow = New TableRow


            aRunNumb = 0
            aRowNumber += 1

            iTable.Rows.Add(aRow)

            aRow.ID = "row" & aRowNumber
            For Each aDCN As String In iDataColumnNameToRender
                aCell = New TableCell
                aRow.Cells.Add(aCell)

                If aDict.TryGetValue(aRunNumb, aAngka) Then
                    aCell.Text = myCstr(aAngka)
                    aDict.Item(aRunNumb) = aAngka + 1
                Else
                    If Strings.InStr(aDCN, ":") <> 0 Then
                        Dim aInfo() As String
                        aInfo = Strings.Split(aDCN, ":", 2)

                        If Strings.InStr(aInfo(1), ":") <> 0 Then
                            Dim aInfo2() As String
                            Dim aObject As Object
                            aInfo2 = Strings.Split(aInfo(1), ":", 2)
                            Select Case aInfo2(0)
                                Case "date"
                                    If DBNull.Value.Equals(aDR.Item(aInfo(0))) Then
                                        aObject = Nothing
                                    Else
                                        aObject = CDate(aDR.Item(aInfo(0)))
                                    End If

                                Case "int"
                                    aObject = CInt(aDR.Item(aInfo(0)))
                                Case "double"
                                    aObject = CDbl(aDR.Item(aInfo(0)))
                            End Select

                            aCell.Text = Strings.Format(aObject, aInfo2(1))
                        Else
                            aCell.Text = Strings.Format(aDR.Item(aInfo(0)), aInfo(1))
                        End If
                    Else
                        aCell.Text = myCstr(aDR.Item(aDCN))
                    End If
                End If

                aRunNumb += 1
            Next
        Next

    End Sub

    Private Shared Function myCstr(ByVal iObject As Object) As String
        If DBNull.Value.Equals(iObject) Then
            Return ""
        End If
        Return CStr(iObject)
    End Function

    Public Sub toASPtable(ByVal iASPtable As Table, ByVal iDataTable As DataTable, ByVal iDataColumnNameToRender() As String)
        Dim aRow As TableRow
        Dim aCell As TableCell

        Dim aDict As New Dictionary(Of Integer, Integer) 'index kolom vs start number
        Dim aColIndex As Integer
        Dim aAngka As Integer
        aColIndex = 0
        For Each aDCN As String In iDataColumnNameToRender
            If aDCN Like "autoNumber*" Then
                aAngka = CInt(Strings.Split(aDCN, ":")(1))
                aDict.Add(aColIndex, aAngka)
            End If
            aColIndex += 1
        Next

        Dim aRowNumber As Integer
        aRowNumber = -1
        For Each aDR As DataRow In iDataTable.Rows
            aRow = New TableRow


            aColIndex = 0
            aRowNumber += 1

            iASPtable.Rows.Add(aRow)
            aRow.ID = "row" & aRowNumber
            RaiseEvent onBeforeDrawRow(aRowNumber)

            If _rowStyle <> "" Then
                aRow.Style.Value = myCstr(_rowStyle)
            End If
            For Each aDCN As String In iDataColumnNameToRender
                aCell = New TableCell
                aRow.Cells.Add(aCell)

                If aDict.TryGetValue(aColIndex, aAngka) Then
                    aCell.Text = myCstr(aAngka)
                    aDict.Item(aColIndex) = aAngka + 1
                Else
                    If Strings.InStr(aDCN, ":") <> 0 Then
                        Dim aInfo() As String
                        aInfo = Strings.Split(aDCN, ":", 2)

                        If Strings.InStr(aInfo(1), ":") <> 0 Then
                            Dim aInfo2() As String
                            Dim aObject As Object
                            aInfo2 = Strings.Split(aInfo(1), ":", 2)
                            Select Case aInfo2(0)
                                Case "date"
                                    If DBNull.Value.Equals(aDR.Item(aInfo(0))) Then
                                        aObject = Nothing
                                    Else
                                        aObject = CDate(aDR.Item(aInfo(0)))
                                    End If

                                Case "int"
                                    aObject = CInt(aDR.Item(aInfo(0)))
                                Case "double"
                                    aObject = CDbl(aDR.Item(aInfo(0)))
                            End Select

                            aCell.Text = Strings.Format(aObject, aInfo2(1))
                        Else
                            aCell.Text = Strings.Format(aDR.Item(aInfo(0)), aInfo(1))
                        End If
                    Else
                        aCell.Text = myCstr(aDR.Item(aDCN))
                    End If
                End If

                '==> adding style to cell by its column index
                Dim aStyle As String

                If _colStyle.TryGetValue(aColIndex, aStyle) Then
                    aCell.Style.Value = aStyle
                End If

                aColIndex += 1
            Next
        Next
    End Sub

End Class
