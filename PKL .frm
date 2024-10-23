Sub AutoFillPackingList()

    ' Khai báo bi?n cho Workbook và Worksheets
    Dim wsData As Worksheet
    Dim wsPacking As Worksheet
    Dim lastRowData As Long
    Dim lastRowPacking As Long
    Dim i As Long
    Dim poNumber As String
    Dim itemNumber As String
    Dim foundPO As Range
    Dim foundItem As Range
    Dim isRowValid As Boolean

    ' Thi?t l?p tham chi?u cho các b?ng tính
    Set wsData = ThisWorkbook.Sheets("Data") ' Ð?m b?o "Data" kh?p v?i tên b?ng tính th?c t? c?a b?n
    Set wsPacking = ThisWorkbook.Sheets("Packing List") ' Ð?m b?o "Packing List" kh?p v?i tên b?ng tính th?c t? c?a b?n

    ' Ki?m tra xem b?ng tính Data có tr?ng không
    If Application.WorksheetFunction.CountA(wsData.Cells) = 0 Then
        MsgBox "B?ng tính 'Data' dang tr?ng!"
        Exit Sub
    End If

    ' Tìm dòng cu?i cùng c?a d? li?u trong b?ng "Data"
    lastRowData = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    ' Tìm dòng cu?i cùng c?a b?ng "Packing List"
    lastRowPacking = wsPacking.Cells(wsPacking.Rows.Count, "A").End(xlUp).Row

    ' T?t c?p nh?t màn hình d? tang t?c d? th?c thi macro
    Application.ScreenUpdating = False

    ' L?p qua t?ng hàng trong b?ng "Packing List"
    For i = 2 To lastRowPacking
        ' L?y s? PO và item number t? Packing List
        poNumber = wsPacking.Cells(i, 1).Value ' S? PO ? c?t A
        itemNumber = wsPacking.Cells(i, 2).Value ' Item number ? c?t B

        ' Bi?n ki?m tra h?p l?
        isRowValid = True ' Gi? d?nh r?ng hàng là h?p l?

        ' Ki?m tra tính h?p l? cho s? PO
        Set foundPO = wsData.Columns("A").Find(poNumber, LookIn:=xlValues, LookAt:=xlWhole)
        If foundPO Is Nothing Then
            isRowValid = False ' Ðánh d?u hàng không h?p l? n?u không tìm th?y PO
        End If

        ' Ki?m tra tính h?p l? cho item number
        Set foundItem = wsData.Columns("B").Find(itemNumber, LookIn:=xlValues, LookAt:=xlWhole)
        If foundItem Is Nothing Then
            isRowValid = False ' Ðánh d?u hàng không h?p l? n?u không tìm th?y item number
        End If

        ' Tô màu d? n?u hàng không h?p l?
        If Not isRowValid Then
            wsPacking.Rows(i).Interior.Color = RGB(255, 0, 0) ' Tô màu d?
        Else
            wsPacking.Rows(i).Interior.ColorIndex = xlNone ' Xóa màu tô n?u hàng h?p l?
        End If
    Next i

    ' B?t l?i c?p nh?t màn hình
    Application.ScreenUpdating = True

    MsgBox "Ki?m tra hoàn thành!"

End Sub




