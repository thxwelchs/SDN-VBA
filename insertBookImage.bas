Sub imageInsertToProductTable()
    ' SDN 표지 이미지 넣기 VBA
    ' 문의) thxwelchs@gmail.com
    
    Dim folderPath As String
    folderPath = ""
    Dim RootFolder As String
    Dim scriptstr As String
    Dim ImageDirPath As String

    If IsMac() Then
        On Error Resume Next
        RootFolder = MacScript("return (path to desktop foldera) as String")

        If Val(Application.Version) < 15 Then
            scriptstr = "(choose folder with prompt ""이미지 경로 폴더 선택""" & _
                " default location alias """ & RootFolder & """) as string"
        Else
            scriptstr = "return posix path of (choose folder with prompt ""이미지 경로 폴더 선택""" & _
                " default location alias """ & RootFolder & """) as string"
        End If

        folderPath = MacScript(scriptstr)
        On Error GoTo 0
    Else
        With Application.FileDialog(msoFileDialogFolderPicker)
            .InitialFileName = Application.DefaultFilePath & "\"
            .Title = "이미지 경로 폴더 선택"
            .Show
            If .SelectedItems.Count > 0 Then
                folderPath = .SelectedItems(1)
            End If
        End With
    End If

    If folderPath <> "" Then
        Debug.Print folderPath
        ImageDirPath = folderPath
        lastRowCount = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
        Set Rng = ActiveSheet.Range("A1", "R" & lastRowCount)

        For Each cel In Rng.Cells
            With cel
                If cel.Value = "품 번" Then
                    Dim check As Boolean
                    check = IsNumeric(cel.Offset(-6, 0).Value) And cel.Offset(1, 0).Value = "품 명" And cel.Offset(2, 0).Value = "설 명" And cel.Offset(3, 0).Value = "가 격"
                    If check = True Then
                        Dim productName As String
                        productName = cel.Offset(0, 1).Value
                        
                        Dim startColumn As Integer
                        Dim startRow As Integer

                        startColumn = cel.Offset(-6, 0).Column + 1
                        startRow = cel.Offset(-6, 0).Row

                        Range(Cells(startRow, startColumn), Cells(startRow + 5, startColumn + 3)).Select



                        Dim IsFileProcessing As Boolean
                        IsFileProcessing = False

                        ' 이미지 확장자 검사 [나중에 추가 가능]
                        Dim FileExts(3) As String
                        FileExts(0) = ".jpg"
                        FileExts(1) = ".png"
                        FileExts(2) = ".gif"
                        Dim FileExtsLen As Integer
                        FileExtsLen = UBound(FileExts) - LBound(FileExts)

                        Dim FileExt As String
                        Dim FilePath As String
                        Dim FileOriginPath As String
                        FileOriginPath = ImageDirPath & "\" & productName
                        FilePath = ""

                        For i = 0 To (FileExtsLen - 1)
                            FileExt = FileExts(i)
                            On Error Resume Next
                                FilePath = Dir(FileOriginPath & FileExt)
                            On Error GoTo 0
                            If FilePath <> "" Then
                                IsFileProcessing = True
                                Exit For
                            End If
                        Next i
                        
                        If IsFileProcessing Then
                            With ActiveSheet.Pictures.Insert(FileOriginPath & FileExt).ShapeRange
                                .LockAspectRatio = msoFalse
                                .Height = Selection.Height - 1
                                .Width = Selection.Width - 1
                                .Left = Selection.Left
                                .Top = Selection.Top + 1
                                .Rotation = 0
                            End With
                        End If
                        
                    End If
                End If
            End With
        Next cel
    Else
        MsgBox ("폴더 경로를 확인하여주세요.")
    End If

    
End Sub


Function IsMac()
    IsMac = InStr(Application.OperatingSystem, "Mac")
End Function

