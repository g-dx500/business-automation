Attribute VB_Name = "Module1"
Sub open_Read_Only()

    Dim x As String
    x = Range("B2").Value

    '簡易チェック:空欄、ファイル存在確認
    If x = "" Then
        MsgBox "B2にファイルパスを入力してください。"
        Exit Sub
    End If

    If Dir(x) = "" Then
        MsgBox "指定されたファイルが見つかりません: " & x
        Exit Sub
    End If

    Workbooks.Open Filename:=x, ReadOnly:=True '読み取り専用で開く

End Sub


Sub copy_Sales()

    On Error GoTo ErrHandler

    'コピー元
    Dim shName As String: shName = Range("B3").Value
    Dim fltName As String: fltName = Range("B4").Value
    
    'コピー先
    Dim FoldarName As String: FolderName = Range("B5").Value
    Dim FilenameExtension As String: FilenameExtension = ".xlsx" '保存ファイルの拡張子
    Dim SaveName As String
    
    '■コピー処理-----------------------------------------
    'コピー元Excelファイルを開く
    Call open_Read_Only
    
    Dim wb1 As Workbook: Set wb1 = ActiveWorkbook

    'シート存在チェック
    Dim shExists As Boolean
    Dim ws As Worksheet
    
    For Each ws In wb1.Worksheets
        If ws.Name = shName Then shExists = True
    Next ws
    
    If Not shExists Then
        MsgBox "シート「" & shName & "」が見つかりません。"
        GoTo CleanExit
    End If

    wb1.Sheets(shName).Activate
    Range("A4").AutoFilter 1, fltName
    
    'コピー範囲選択
    Dim r As Long, c As Long, x As Long, h As Long
    Dim startCell As Range

    Set startCell = Range("A3")

    x = startCell.Offset(-1, 0).Row 'コピーしない行数
    h = startCell.Offset(1, 0).Row  'ヘッダー行

    r = Cells(Rows.Count, 1).End(xlUp).Row - x
    c = Cells(h, Columns.Count).End(xlToLeft).Column

    startCell.Resize(r, c).Select
    
    Selection.Copy
    '-----------------------------------------------------
    
    '■貼り付け処理---------------------------------------
    '新規ブックを開く
    Workbooks.Add
    Dim wb2 As Workbook
    Set wb2 = ActiveWorkbook

    '貼り付け
    ActiveWorkbook.Sheets(1).Activate
    Range("A1").Select
    ActiveCell.PasteSpecial Paste:=xlPasteValues '(BOOK間のコピペでは必ずPasteSpecial)
    ActiveCell.PasteSpecial Paste:=xlPasteFormats

    '作業日の日付で保存
    Dim Filename As String
    Dim n_WK As Integer
    n_WK = DatePart("ww", Date) - 1
    
        If n_WK < 10 Then
            Filename = "WK" & "0" & n_WK      '数字が1桁の場合は「0」を足す
        Else
            Filename = "WK" & n_WK
        End If
    
    Dim i As Long
    Dim saved As Boolean
    For i = 1 To 100
        SaveName = FolderName & "\" & Filename & "_" & i & FilenameExtension
        'SaveNameが見つかればその文字列を返すので、見つからなければ空白
        If Dir(SaveName) = "" Then
            wb2.SaveAs Filename:=SaveName '(Activeworkbook.SaveAsでも可)
            saved = True
            Exit For
        End If
    Next i

    If Not saved Then
        MsgBox "保存先に空き番号が見つかりませんでした(101個以上存在)。"
    End If

    wb2.Close False  '(Activeworkbook.Close Falseでも可)
    '-----------------------------------------------------
    
    wb1.Close False 'コピー元を閉じる
    
    'オブジェクト解放
    Set ws = Nothing
    Set wb1 = Nothing
    Set startCell = Nothing
    Set wb2 = Nothing

CleanExit:
    If Not wb1 Is Nothing Then wb1.Close False
    Set wb1 = Nothing
    Set wb2 = Nothing
    Exit Sub
    
ErrHandler:
    MsgBox "エラーが発生しました: " & Err.Description
    Resume CleanExit

End Sub
