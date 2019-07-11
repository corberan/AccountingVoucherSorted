Option Explicit

Sub 分录号排序()
    Application.ScreenUpdating = False
    
    ' 获得当前工作表
    Dim wks As Worksheet
    Set wks = ActiveSheet
    
    ' 检查
    If wks.UsedRange.Rows.Count < 2 Then
        MsgBox "要进行排序，要求此工作簿不能少于两行"
        Exit Sub
    End If
    
    ' 定义 凭证字、凭证号、分录号 表头字符串，如有变化，请自行适配修改
    Dim pzzHeader As String, pzhHeader As String, flhHeader As String
    pzzHeader = "PZZ"
    pzhHeader = "PZH"
    flhHeader = "FLH"
    
    ' 需要处理的行数，仅用作调试，默认为工作簿使用到的最大行
    Dim targetRowCount As Long
    targetRowCount = wks.UsedRange.Rows.Count
    
    ' 找出 凭证字、凭证号、分录号 表头所在列的列序
    Dim pzzColIndex As Long, pzhColIndex As Long, flhColIndex As Long
    Dim cellValue As Variant
    
    Dim i As Long
    For i = 1 To wks.UsedRange.Columns.Count
        cellValue = wks.Cells(1, i).Value
        If IsEmpty(cellValue) = False And StrComp(TypeName(cellValue), "String") = 0 Then
            If StrComp(cellValue, pzzHeader) = 0 Then
                pzzColIndex = i
            ElseIf StrComp(cellValue, pzhHeader) = 0 Then
                pzhColIndex = i
            ElseIf StrComp(cellValue, flhHeader) = 0 Then
                flhColIndex = i
            End If
        End If
    Next
    
    If pzzColIndex = 0 Then
        MsgBox "缺少凭证字表头【" & pzzHeader & "】"
        Exit Sub
    End If
    
    If pzhColIndex = 0 Then
        MsgBox "缺少凭证号表头【" & pzhHeader & "】"
        Exit Sub
    End If
    
    If flhColIndex = 0 Then
        MsgBox "缺少分录号表头【" & flhHeader & "】"
        Exit Sub
    End If
    
    
    If MsgBox("将重置分录号表头【" & flhHeader & "】所在列【" & Split(wks.Columns(flhColIndex).Address(, False), ":")(0) & "】的数据，是否继续？", vbYesNo) <> vbYes Then
        Exit Sub
    End If
    
    '
    Dim rangeExp As String
    Dim rng As Range
    
    ' 清除分录号所在列的数据
    rangeExp = GetColRangeExpFromColIndex(wks, flhColIndex, targetRowCount)
    wks.Range(rangeExp).ClearContents
    
    ' 将凭证字列和凭证号列数据保存到数组，用作循环对比，而不是每次使用 cells 获取，加快程序运行时间
    ' 注意，单列 Range 赋值到数组，结果也是二维数组
    Dim pzzDataArr() As Variant, pzhDataArr() As Variant
    
    rangeExp = GetColRangeExpFromColIndex(wks, pzzColIndex, targetRowCount)
    Set rng = wks.Range(rangeExp)
    pzzDataArr = rng
    
    rangeExp = GetColRangeExpFromColIndex(wks, pzhColIndex, targetRowCount)
    Set rng = wks.Range(rangeExp)
    pzhDataArr = rng
        
    Debug.Assert LBound(pzzDataArr) = LBound(pzhDataArr)
    Debug.Assert UBound(pzzDataArr) = UBound(pzhDataArr)
    
    'Debug.Print LBound(pzzRowsDataArr), UBound(pzzRowsDataArr), LBound(pzhRowsDataArr), UBound(pzhRowsDataArr)
        
    ' 即用于保存分录号，也用于判断对应序号的行是否已经处理过
    ' 注意，要将数组赋值给 Range，需要定义类型为 Variant，同时上下标要正确
    Dim flhDataArr() As Variant
    ReDim flhDataArr(1 To targetRowCount, 1 To 1)
    
    '
    Dim pzzCellValue As Variant, pzhCellValue As Variant
    Dim anotherPzzCellValue As Variant, anotherPzhCellValue As Variant
    Dim flhIndex As Long
    
    ' 遍历
    Dim j As Long
    Dim indexMin As Long, indexMax As Long
    indexMin = LBound(pzzDataArr)
    indexMax = UBound(pzzDataArr)
    
    For i = indexMin To indexMax
        ' 没有分录号的行才需要处理
        If flhDataArr(i, 1) = 0 Then
            ' 初始分录号
            flhIndex = 1
            flhDataArr(i, 1) = flhIndex
            flhIndex = flhIndex + 1
            '
            pzzCellValue = pzzDataArr(i, 1)
            pzhCellValue = pzhDataArr(i, 1)
            ' 直接从 i + 1 开始即可
            For j = i + 1 To indexMax
                If flhDataArr(j, 1) = 0 Then
                    anotherPzzCellValue = pzzDataArr(j, 1)
                    anotherPzhCellValue = pzhDataArr(j, 1)
                    ' 凭证字和凭证号均相同的两行
                    If pzzCellValue = anotherPzzCellValue And pzhCellValue = anotherPzhCellValue Then
                        ' 累加分录号
                        flhDataArr(j, 1) = flhIndex
                        flhIndex = flhIndex + 1
                    End If
                End If
            Next j
        End If
    Next i
    
    rangeExp = GetColRangeExpFromColIndex(wks, flhColIndex, targetRowCount)
    Set rng = wks.Range(rangeExp)
    rng = flhDataArr
    
    Application.ScreenUpdating = True
    
    MsgBox "完成"
End Sub

Function GetColRangeExpFromColIndex(wks As Worksheet, colIndex As Long, rangeEnd As Long) As String
    Dim colLetter As String
    colLetter = Split(wks.Columns(colIndex).Address(, False), ":")(0)
    GetColRangeExpFromColIndex = colLetter & "2:" & colLetter & rangeEnd
End Function
