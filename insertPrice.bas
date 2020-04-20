Attribute VB_Name = "Module1"
Sub �Һ��ڰ����Է�()
'
' �Һ��ڰ����Է� ��ũ��
' �Һ��ڰ� ������ �ִ� �������� �Է��Ѵ�.
'
' �ٷ� ���� Ű: Ctrl+w


' SDN ǥ�� �̹��� �ֱ� VBA
' ����) thxwelchs@gmail.com
    Dim productColumn As String
    Dim priceColumn As String
    productColumn = "C"
    priceColumn = "I"
    
    Dim FDG As FileDialog
    Dim Selected As Integer: Dim i As Integer
    Dim SelectedFilePath As String
    
    Dim selectedFileName As String
    Dim bookFileName As String
    bookFileName = Application.ActiveWorkbook.Name
    
    Set FDG = Application.FileDialog(msoFileDialogFilePicker)
    
    With FDG
        .Title = "�Һ��ڰ��� ���� ������ �����ϼ���. (å�� ���ϰ� ������ ������ ��ġ�ؾ� �մϴ�.)"
        .Filters.Add "Only Excel File", "*.xls; *.xlsx; *.xlsm"
        .InitialFileName = Application.ActiveWorkbook.Path
        .AllowMultiSelect = False
        Selected = .Show
    End With
    
    If Selected = -1 Then
        Dim inputData As String
        MsgBox (ThisWorkbook.Path)
        selectedFileName = Right$(FDG.SelectedItems(1), Len(FDG.SelectedItems(1)) - InStrRev(FDG.SelectedItems(1), "\"))
        Dim selectedFileOnlyPath As String
        selectedFileOnlyPath = Split(FDG.SelectedItems(1), selectedFileName)(0)
        selectedFileOnlyPath = Left(selectedFileOnlyPath, Len(selectedFileOnlyPath) - 1)
        
        If ThisWorkbook.Path <> selectedFileOnlyPath Then
            MsgBox ("å�ڿ� ���� ������ ���� ������ ��ġ�ؾ� �մϴ�.")
            Exit Sub
        End If
        
        
        inputData = InputBox("ǰ��, ���� �� ��ȣ�� ���Ŀ� �°� �Է��ϼ���" & vbCrLf & "����: C,I" & vbCrLf & "�⺻��: " & productColumn & "," & priceColumn & "" & vbCrLf & "�Է����� ������ �⺻������ �����˴ϴ�.")
                
        If inputData <> "" Then
            If InStr(inputData, ",") = 0 Or Len(inputData) <> 3 Then
                MsgBox ("ǰ��, ���� �� ��ȣ�� ���Ŀ� ���� �ʽ��ϴ�.")
                Exit Sub
            End If
            
            Dim inputDatas() As String
            
            inputDatas = Split(inputData, ",")
            
            productColumn = inputDatas(0)
            priceColumn = inputDatas(1)
        End If
        
        
        
        Dim ws1 As Worksheet
        Dim ws2 As Worksheet
        Set ws1 = Workbooks(ThisWorkbook.Name).Worksheets(1)
        Set ws2 = Workbooks(selectedFileName).Worksheets(1)
        Dim rg1 As Range
        Dim rg2 As Range
        Set rg1 = ws1.Range("A1", "R" & ws1.Cells.SpecialCells(xlCellTypeLastCell).Row)
        Set rg2 = ws2.Range("A1", "I" & ws2.Cells.SpecialCells(xlCellTypeLastCell).Row)
        
        For Each cel In rg1.Cells
            With cel
                If cel.Value = "ǰ ��" Then
                    Dim check As Boolean
                    check = IsNumeric(cel.Offset(-6, 0).Value) And cel.Offset(1, 0).Value = "ǰ ��" And cel.Offset(2, 0).Value = "�� ��" And cel.Offset(3, 0).Value = "�� ��"
                    If check = True Then
                        Dim productName As String
                        productName = cel.Offset(0, 1).Value
                        Dim price As Long
                        price = findPrice(rg2, productName, productColumn, priceColumn)
                        
                        If price <> -1 Then
                            cel.Offset(3, 1).Value = price
                        End If
                        
                    End If
                End If
            End With
        Next cel
        
        ' MsgBox (ws1.Cells.SpecialCells(xlCellTypeLastCell).Row)
        ' MsgBox (ws2.Cells.SpecialCells(xlCellTypeLastCell).Row)
        
        ' SelectedFilePath = SelectedFilePath & FDG.SelectedItems(1)
        
        ' Set SelectedFile = GetObject(SelectedFilePath)
        
        ' MsgBox (ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row)
        
        Dim findRange As Range
        Set findRange = Range("A1", "I" & ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row)
        
        ' MsgBox (ActiveSheet.Name)
        ' MsgBox (currentFileActiveSheet)
        
        ' MsgBox (findPrice(findRange, "B-10326", productColumn, priceColumn))
        
    Else
        MsgBox ("������ �������� �ʾ� ��ũ�θ� �����մϴ�.")
    End If
    
    
    
    ' Debug.Print SelectedFilePath
End Sub


Function findPrice(sr As Range, productName As String, productColumn As String, priceColumn As String) As Long
    Dim price As Long
    price = -1
    For e = 2 To sr.Rows.Count Step 1
        If productName = sr.Cells(e, productColumn).Value Then
            price = sr.Cells(e, priceColumn).Value
            Exit For
        End If
    Next
    
    findPrice = price
End Function
