Attribute VB_Name = "LogicHierarchy"
Option Explicit

' ============================================================
'  COBOL ロジック階層 (hand-simple / vba0526版アルゴリズム準拠)
'   - シート: コントロール / COBOLソース / ロジック階層 の 3 枚
'   - PROCEDURE DIVISION 以降だけ解析
'   - PM 流: 行末ピリオドを先に剥がしてから keyword 判定
'   - IF→[THEN]/[ELSE]、EVALUATE→WHEN、SEARCH→[AT END]/WHEN を分岐ノード化
'   - 罫線ツリー ├── / └── / │  (ChrW で生成)
'   - 開始行は COBOLソース シートへハイパーリンク
'   - スタック検査で END-xx/ELSE/WHEN の対応ミスを赤表示
' ============================================================

Private Const SH_CTRL As String = "コントロール"
Private Const SH_SRC  As String = "COBOLソース"
Private Const SH_TREE As String = "ロジック階層"
Private Const CODE_START As Long = 8
Private Const CODE_END   As Long = 72
Private Const C_HEADER  As Long = 14474460   ' RGB(220,220,220)
Private Const C_SECTION As Long = 13168895   ' RGB(255,240,200)
Private Const C_BRANCH  As Long = 16314338   ' RGB(226,239,248)

' ノード木 (フラット + 親インデックス)
Private nText() As String, nKind() As String, nLine() As String
Private nParent() As Long, nDepth() As Long, nCount As Long

' ============================================================
'  (1) コントロールシート (Meiryo UI + teal 角丸ボタン)
' ============================================================
Public Sub SetupControlSheet()
    Dim cHead As Long, cAcc As Long, cMute As Long, cCard As Long, cW As Long
    cHead = RGB(38, 70, 83): cAcc = RGB(42, 157, 143)
    cMute = RGB(120, 120, 120): cCard = RGB(244, 246, 248): cW = RGB(255, 255, 255)

    Dim ws As Worksheet: Set ws = EnsureSheet(SH_CTRL)
    ws.Move Before:=ThisWorkbook.Worksheets(1)
    ws.Cells.Clear
    Do While ws.Shapes.Count > 0: ws.Shapes(1).Delete: Loop
    ws.Activate
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False: ActiveWindow.DisplayHeadings = False
    On Error GoTo 0
    ws.Columns("A").ColumnWidth = 2.5
    ws.Columns("B:H").ColumnWidth = 13.5

    With ws.Range("B2:H3")
        .Merge: .Interior.Color = cHead: .IndentLevel = 1
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter
    End With
    ws.Range("B2").Value = "COBOL ロジック階層 ツール"
    With ws.Range("B2").Font: .Name = "Meiryo UI": .Size = 20: .Bold = True: .Color = cW: End With
    ws.Rows("2:3").RowHeight = 24

    With ws.Range("B4:H4"): .Merge: .HorizontalAlignment = xlLeft: .IndentLevel = 1: End With
    ws.Range("B4").Value = "PROCEDURE DIVISION の IF/EVALUATE/PERFORM/SEARCH 階層を可視化"
    With ws.Range("B4").Font: .Name = "Meiryo UI": .Size = 10: .Color = cMute: End With
    ws.Rows("4").RowHeight = 22

    ws.Range("B6").Value = "■ 使い方"
    With ws.Range("B6").Font: .Name = "Meiryo UI": .Size = 12: .Bold = True: .Color = cAcc: End With
    With ws.Range("B7:H8")
        .Merge: .Interior.Color = cCard: .WrapText = True
        .VerticalAlignment = xlCenter: .IndentLevel = 1
        .Font.Name = "Meiryo UI": .Font.Size = 11
    End With
    ws.Range("B7").Value = "1. 下のボタンを押し、COBOL ソース (.cbl) を選択" & Chr(10) & _
                           "2. COBOLソース と ロジック階層 シートが自動生成 (再実行で上書き)"
    ws.Rows("7:8").RowHeight = 20

    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(5, ws.Range("B10").Left, ws.Range("B10").Top, 320, 46)
    btn.Name = "btnAnalyze": btn.Fill.ForeColor.RGB = cAcc: btn.Line.Visible = msoFalse
    With btn.TextFrame2
        .VerticalAnchor = msoAnchorMiddle
        With .TextRange
            .Text = "COBOL ソースを選択して解析"
            .ParagraphFormat.Alignment = msoAlignCenter
            .Font.Size = 13: .Font.Bold = msoTrue
            .Font.Name = "Meiryo UI": .Font.Fill.ForeColor.RGB = cW
        End With
    End With
    btn.OnAction = "PickCobolAndAnalyze"
    ws.Rows("10:11").RowHeight = 26
    ws.Range("A1").Select
End Sub

' ============================================================
'  (2) ボタン OnAction
' ============================================================
Public Sub PickCobolAndAnalyze()
    Dim cblPath As String: cblPath = PickCobolFile()
    If Len(cblPath) = 0 Then Exit Sub
    Dim lines() As String: lines = ReadSourceLines(cblPath)
    BuildSourceSheet lines, cblPath
    BuildHierarchySheet lines, cblPath
End Sub

Private Function PickCobolFile() As String
    Dim fd As FileDialog: Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "COBOL ソースを選択": .AllowMultiSelect = False: .Filters.Clear
        .Filters.Add "COBOL/テキスト", "*.cbl;*.cob;*.txt;*.src;*.cpy"
        .Filters.Add "すべて", "*.*"
        If .Show = -1 Then PickCobolFile = .SelectedItems(1)
    End With
End Function

' ============================================================
'  (3) ファイル読込 + 文字コード自動判定 (BOM + UTF-8 検証)
' ============================================================
Private Function ReadSourceLines(ByVal path As String) As String()
    Dim cs As String: cs = DetectCharset(path)
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = cs: stm.Open: stm.LoadFromFile path
    Dim whole As String: whole = stm.ReadText(-1): stm.Close
    whole = Replace(whole, vbCrLf, vbLf): whole = Replace(whole, vbCr, vbLf)
    ReadSourceLines = Split(whole, vbLf)
End Function

Private Function DetectCharset(ByVal path As String) As String
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1: stm.Open: stm.LoadFromFile path
    Dim bin() As Byte: bin = stm.Read(-1): stm.Close
    Dim n As Long: n = UBound(bin) + 1
    If n >= 3 Then
        If bin(0) = &HEF And bin(1) = &HBB And bin(2) = &HBF Then DetectCharset = "UTF-8": Exit Function
    End If
    If n >= 2 Then
        If bin(0) = &HFF And bin(1) = &HFE Then DetectCharset = "Unicode": Exit Function
        If bin(0) = &HFE And bin(1) = &HFF Then DetectCharset = "unicodeFFFE": Exit Function
    End If
    If IsValidUtf8(bin, n) Then DetectCharset = "UTF-8" Else DetectCharset = "Shift_JIS"
End Function

Private Function IsValidUtf8(ByRef bin() As Byte, ByVal n As Long) As Boolean
    Dim i As Long, cont As Long, b As Long, k As Long
    Do While i < n
        b = bin(i)
        If b < &H80 Then
            cont = 0
        ElseIf b >= &HC2 And b <= &HDF Then
            cont = 1
        ElseIf b >= &HE0 And b <= &HEF Then
            cont = 2
        ElseIf b >= &HF0 And b <= &HF4 Then
            cont = 3
        Else
            IsValidUtf8 = False: Exit Function
        End If
        For k = 1 To cont
            i = i + 1
            If i >= n Then IsValidUtf8 = False: Exit Function
            If (bin(i) And &HC0) <> &H80 Then IsValidUtf8 = False: Exit Function
        Next k
        i = i + 1
    Loop
    IsValidUtf8 = True
End Function

' ============================================================
'  (4) COBOLソース シート (全行 + 物理行番号、ハイパーリンクの飛び先)
' ============================================================
Private Sub BuildSourceSheet(ByRef lines() As String, ByVal cblPath As String)
    Dim ws As Worksheet: Set ws = EnsureSheet(SH_SRC)
    ws.Cells.Clear
    ws.Range("A1").Value = SH_SRC
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14
    ws.Range("A2").Value = "(行番号 = ファイル内の物理行)  対象: " & cblPath
    ws.Cells(3, 1).Value = "行番号": ws.Cells(3, 2).Value = "ソース"
    ws.Range("A3:B3").Font.Bold = True
    ws.Range("A3:B3").Interior.Color = C_HEADER
    Dim n As Long: n = UBound(lines) - LBound(lines) + 1
    If n <= 0 Then Exit Sub
    Dim arr() As Variant: ReDim arr(1 To n, 1 To 2)
    Dim i As Long
    For i = 1 To n
        arr(i, 1) = i
        arr(i, 2) = CStr(lines(LBound(lines) + i - 1))
    Next i
    Application.ScreenUpdating = False
    ws.Range(ws.Cells(4, 1), ws.Cells(3 + n, 2)).Value = arr
    ws.Range(ws.Cells(4, 1), ws.Cells(3 + n, 2)).Font.Name = "MS Gothic"
    ws.Columns("A").ColumnWidth = 8: ws.Columns("B").ColumnWidth = 100
    ws.Activate: ws.Range("A4").Select
    On Error Resume Next
    ActiveWindow.FreezePanes = False: ActiveWindow.FreezePanes = True
    On Error GoTo 0
    Application.ScreenUpdating = True
End Sub

' ============================================================
'  (5) ロジック階層 シート
' ============================================================
Private Sub BuildHierarchySheet(ByRef lines() As String, ByVal cblPath As String)
    Dim errCount As Long: BuildTree lines, errCount
    Dim ws As Worksheet: Set ws = EnsureSheet(SH_TREE)
    ws.Cells.Clear
    ws.Range("A1").Value = "COBOL ロジック階層"
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14
    ws.Range("A3").Value = "対象ファイル": ws.Range("B3").Value = cblPath
    ws.Range("A4").Value = "総行数":       ws.Range("B4").Value = UBound(lines) - LBound(lines) + 1
    ws.Range("A5").Value = "検出した問題": ws.Range("B5").Value = errCount & " 件"
    If errCount > 0 Then ws.Range("B5").Font.Color = RGB(192, 0, 0)
    Dim hdr As Long: hdr = 7
    ws.Cells(hdr, 1).Value = "階層ツリー": ws.Cells(hdr, 2).Value = "開始行": ws.Cells(hdr, 3).Value = "種別"
    ws.Range(ws.Cells(hdr, 1), ws.Cells(hdr, 3)).Font.Bold = True
    ws.Range(ws.Cells(hdr, 1), ws.Cells(hdr, 3)).Interior.Color = C_HEADER
    Application.ScreenUpdating = False
    RenderTree ws, hdr + 1
    ws.Columns("A").ColumnWidth = 80: ws.Columns("B").ColumnWidth = 8: ws.Columns("C").ColumnWidth = 14
    Application.ScreenUpdating = True
    ws.Activate: ws.Range("A1").Select
    MsgBox "解析完了。 検出した問題: " & errCount & " 件", _
        IIf(errCount > 0, vbExclamation, vbInformation)
End Sub

' --- ソース → ノード木 ---
Private Sub BuildTree(ByRef lines() As String, ByRef errCount As Long)
    nCount = 0
    ReDim nText(1 To 256): ReDim nKind(1 To 256): ReDim nLine(1 To 256)
    ReDim nParent(1 To 256): ReDim nDepth(1 To 256)
    errCount = 0
    Dim ctrl As New Collection
    Dim curC As Long: curC = -1
    Dim inProc As Boolean
    Dim i As Long, raw As String, ind As String, code As String, cu As String, kind As String
    Dim srcLineNo As Long, newIdx As Long, brIdx As Long, isHdr As Boolean
    Dim pendingIfNode As Long, wasPending As Long

    For i = LBound(lines) To UBound(lines)
        srcLineNo = i - LBound(lines) + 1
        raw = lines(i)
        If Len(Trim(raw)) = 0 Then GoTo NextLine
        If Len(raw) >= 7 Then
            ind = Mid(raw, 7, 1)
            If ind = "*" Or ind = "/" Then GoTo NextLine
        End If
        code = Trim(MidSafe(raw, CODE_START, CODE_END - CODE_START + 1))
        If Len(code) = 0 Then GoTo NextLine

        ' 多行 IF 条件 トラッキング: 毎反復で前回値を保存しクリア
        wasPending = pendingIfNode
        pendingIfNode = 0

        ' PROCEDURE DIVISION より前はスキップ (語間の空白数不問)
        If Not inProc Then
            cu = UCase(code)
            If InStr(cu, "PROCEDURE") > 0 And InStr(cu, "DIVISION") > 0 Then
                inProc = True
                curC = AddNode("PROCEDURE DIVISION", "SECTION", CStr(srcLineNo), -1)
            End If
            GoTo NextLine
        End If

        ' 段落/SECTION ヘッダ判定は ピリオド付きの原文で
        isHdr = IsParagraphHeader(raw, code)
        ' PM 流: keyword 判定前にピリオド剥がし
        If Right(code, 1) = "." Then code = Trim(Left(code, Len(code) - 1))

        If isHdr Then
            If ctrl.Count > 0 Then
                AddNode "!! 前段落の階層が " & ctrl.Count & " 残存 (END-xxx 不足?)", "ERROR", CStr(srcLineNo), curC
                errCount = errCount + 1
                Do While ctrl.Count > 0: ctrl.Remove ctrl.Count: Loop
            End If
            curC = AddNode(Application.WorksheetFunction.Trim(code), "SECTION", CStr(srcLineNo), -1)
            GoTo NextLine
        End If

        ClassifyLine code, kind
        Select Case kind
            Case "IF"
                newIdx = AddNode(code, "IF", CStr(srcLineNo), curC)
                brIdx = AddNode("[THEN]", "BRANCH", "", newIdx)
                ctrl.Add Array("IF", newIdx, curC)
                curC = brIdx
                pendingIfNode = newIdx
            Case "ELSE"
                If ctrl.Count = 0 Or ctrl(ctrl.Count)(0) <> "IF" Then
                    AddNode "!! ELSE に対応する IF がありません", "ERROR", CStr(srcLineNo), curC
                    errCount = errCount + 1
                Else
                    brIdx = AddNode("[ELSE]", "BRANCH", "", ctrl(ctrl.Count)(1))
                    curC = brIdx
                End If
            Case "END-IF": CloseCtrl ctrl, "IF", CStr(srcLineNo), curC, errCount
            Case "EVALUATE"
                newIdx = AddNode(code, "EVALUATE", CStr(srcLineNo), curC)
                ctrl.Add Array("EVALUATE", newIdx, curC): curC = newIdx
            Case "WHEN"
                If ctrl.Count = 0 Or (ctrl(ctrl.Count)(0) <> "EVALUATE" And ctrl(ctrl.Count)(0) <> "SEARCH") Then
                    AddNode "!! WHEN に対応する EVALUATE/SEARCH がありません", "ERROR", CStr(srcLineNo), curC
                    errCount = errCount + 1
                Else
                    brIdx = AddNode(code, "WHEN", CStr(srcLineNo), ctrl(ctrl.Count)(1))
                    curC = brIdx
                End If
            Case "END-EVALUATE": CloseCtrl ctrl, "EVALUATE", CStr(srcLineNo), curC, errCount
            Case "SEARCH"
                newIdx = AddNode(code, "SEARCH", CStr(srcLineNo), curC)
                ctrl.Add Array("SEARCH", newIdx, curC): curC = newIdx
            Case "AT-END"
                If ctrl.Count > 0 And ctrl(ctrl.Count)(0) = "SEARCH" Then
                    brIdx = AddNode("[AT END]", "BRANCH", CStr(srcLineNo), ctrl(ctrl.Count)(1))
                    curC = brIdx
                Else
                    AddNode code, "ACTION", CStr(srcLineNo), curC
                End If
            Case "END-SEARCH": CloseCtrl ctrl, "SEARCH", CStr(srcLineNo), curC, errCount
            Case "PERFORM(inline)"
                newIdx = AddNode(code, "PERFORM", CStr(srcLineNo), curC)
                ctrl.Add Array("PERFORM", newIdx, curC): curC = newIdx
            Case "END-PERFORM": CloseCtrl ctrl, "PERFORM", CStr(srcLineNo), curC, errCount
            Case "THEN"
                ' [THEN] ブランチで表現済 → 無視
            Case Else
                If wasPending > 0 And Not IsActionVerb(code) Then
                    ' 多行 IF 条件の継続行: IF ノードのテキストに連結
                    nText(wasPending) = nText(wasPending) & " " & code
                    pendingIfNode = wasPending
                ElseIf IsActionVerb(code) Then
                    AddNode code, "ACTION", CStr(srcLineNo), curC
                End If
        End Select
NextLine:
    Next i
    If ctrl.Count > 0 Then
        AddNode "!! 解析終了時に階層が " & ctrl.Count & " 残存 (END-xxx 不足)", "ERROR", "", -1
        errCount = errCount + 1
    End If
End Sub

Private Sub CloseCtrl(ByRef ctrl As Collection, ByVal want As String, ByVal seqNo As String, _
                      ByRef curC As Long, ByRef errCount As Long)
    If ctrl.Count = 0 Then
        AddNode "!! END-" & want & " に対応する開始がありません", "ERROR", seqNo, curC
        errCount = errCount + 1: Exit Sub
    End If
    Dim top As Variant: top = ctrl(ctrl.Count)
    If top(0) <> want Then
        AddNode "!! 直近の開始(" & top(0) & ") と END-" & want & " が不一致", "ERROR", seqNo, curC
        errCount = errCount + 1
    End If
    curC = top(2): ctrl.Remove ctrl.Count
End Sub

Private Function AddNode(ByVal text As String, ByVal kind As String, _
                         ByVal lineNo As String, ByVal parentIdx As Long) As Long
    nCount = nCount + 1
    If nCount > UBound(nText) Then
        Dim ns As Long: ns = UBound(nText) * 2
        ReDim Preserve nText(1 To ns): ReDim Preserve nKind(1 To ns)
        ReDim Preserve nLine(1 To ns): ReDim Preserve nParent(1 To ns): ReDim Preserve nDepth(1 To ns)
    End If
    nText(nCount) = text: nKind(nCount) = kind: nLine(nCount) = lineNo: nParent(nCount) = parentIdx
    If parentIdx = -1 Then nDepth(nCount) = 0 Else nDepth(nCount) = nDepth(parentIdx) + 1
    AddNode = nCount
End Function

' --- ツリー描画 + 開始行ハイパーリンク ---
Private Sub RenderTree(ByVal ws As Worksheet, ByVal startRow As Long)
    If nCount = 0 Then Exit Sub
    Dim lastChild() As Long: ReDim lastChild(-1 To nCount)
    Dim i As Long
    For i = 1 To nCount: lastChild(nParent(i)) = i: Next i
    Dim r As Long: r = startRow
    Dim disp As String, sl As Long
    For i = 1 To nCount
        If nKind(i) = "SECTION" Then
            disp = ChrW(&H25A0) & " " & nText(i)
        Else
            disp = BuildPrefix(i, lastChild) & nText(i)
        End If
        ws.Cells(r, 1).Value = disp
        ws.Cells(r, 1).Font.Name = "MS Gothic"
        If Len(nLine(i)) > 0 Then
            sl = CLng(nLine(i))
            On Error Resume Next
            ws.Hyperlinks.Add Anchor:=ws.Cells(r, 2), Address:="", _
                SubAddress:="'" & SH_SRC & "'!A" & (sl + 3), _
                TextToDisplay:=CStr(sl)
            If Err.Number <> 0 Then ws.Cells(r, 2).Value = sl: Err.Clear
            On Error GoTo 0
        End If
        ws.Cells(r, 3).Value = KindLabel(nKind(i))
        ApplyRowColor ws, r, nKind(i)
        r = r + 1
    Next i
End Sub

Private Function BuildPrefix(ByVal idx As Long, ByRef lastChild() As Long) As String
    Dim chain() As Long: ReDim chain(1 To 128)
    Dim cnt As Long, a As Long: a = nParent(idx)
    Do While a <> -1
        If nParent(a) <> -1 Then cnt = cnt + 1: chain(cnt) = a
        a = nParent(a)
    Loop
    Dim s As String, j As Long, node As Long
    For j = cnt To 1 Step -1
        node = chain(j)
        If lastChild(nParent(node)) = node Then s = s & "    " Else s = s & ChrW(&H2502) & "   "
    Next j
    If lastChild(nParent(idx)) = idx Then
        s = s & ChrW(&H2514) & ChrW(&H2500) & " "
    Else
        s = s & ChrW(&H251C) & ChrW(&H2500) & " "
    End If
    BuildPrefix = s
End Function

Private Function KindLabel(ByVal kind As String) As String
    Select Case kind
        Case "PERFORM": KindLabel = "PERFORM(in)"
        Case "ERROR":   KindLabel = ChrW(&H2605) & "ERROR"
        Case Else:      KindLabel = kind
    End Select
End Function

Private Sub ApplyRowColor(ByVal ws As Worksheet, ByVal r As Long, ByVal kind As String)
    Dim rng As Range: Set rng = ws.Range(ws.Cells(r, 1), ws.Cells(r, 3))
    Select Case kind
        Case "SECTION": rng.Interior.Color = C_SECTION: ws.Cells(r, 1).Font.Bold = True
        Case "IF", "EVALUATE", "SEARCH", "WHEN", "BRANCH", "PERFORM"
            rng.Interior.Color = C_BRANCH
        Case "ERROR": rng.Interior.Color = RGB(255, 199, 206): ws.Cells(r, 1).Font.Bold = True
    End Select
End Sub

' --- 行種別判定 (ピリオドは呼び出し元で既に剥がし済) ---
Private Sub ClassifyLine(ByVal code As String, ByRef kind As String)
    If StartsWithKw(code, "END-IF") Then kind = "END-IF": Exit Sub
    If StartsWithKw(code, "END-EVALUATE") Then kind = "END-EVALUATE": Exit Sub
    If StartsWithKw(code, "END-PERFORM") Then kind = "END-PERFORM": Exit Sub
    If StartsWithKw(code, "END-SEARCH") Then kind = "END-SEARCH": Exit Sub
    If StartsWithKw(code, "ELSE") Then kind = "ELSE": Exit Sub
    If StartsWithKw(code, "AT END") Then kind = "AT-END": Exit Sub
    If StartsWithKw(code, "WHEN") Then kind = "WHEN": Exit Sub
    If StartsWithKw(code, "THEN") Then kind = "THEN": Exit Sub
    If StartsWithKw(code, "EVALUATE") Then kind = "EVALUATE": Exit Sub
    If StartsWithKw(code, "SEARCH") Then kind = "SEARCH": Exit Sub
    If StartsWithKw(code, "IF") Then kind = "IF": Exit Sub
    If StartsWithKw(code, "PERFORM") Then
        If IsInlinePerform(code) Then kind = "PERFORM(inline)" Else kind = "PERFORM-OUT"
        Exit Sub
    End If
    kind = "STMT"
End Sub

' --- 動詞ホワイトリスト (これ以外は条件継続行等として無視) ---
Private Function IsActionVerb(ByVal code As String) As Boolean
    Dim t As String: t = FirstToken(UCase(Trim(code)))
    Select Case t
        Case "PERFORM", "CALL", "MOVE", "COMPUTE", _
             "ADD", "SUBTRACT", "MULTIPLY", "DIVIDE", _
             "STRING", "UNSTRING", "INSPECT", "SET", "INITIALIZE", _
             "READ", "WRITE", "REWRITE", "DELETE", _
             "OPEN", "CLOSE", "ACCEPT", "DISPLAY", _
             "EXIT", "CONTINUE", "GOBACK", "STOP", "GO"
            IsActionVerb = True
    End Select
End Function

Private Function IsInlinePerform(ByVal code As String) As Boolean
    Dim rest As String: rest = Trim(Mid(UCase(Trim(code)), Len("PERFORM") + 1))
    If Len(rest) = 0 Then IsInlinePerform = True: Exit Function
    Dim t As String: t = FirstToken(rest)
    Select Case t
        Case "UNTIL", "VARYING", "WITH", "FOREVER", "TEST": IsInlinePerform = True
        Case Else: IsInlinePerform = IsNumeric(t)
    End Select
End Function

' --- 段落/SECTION ヘッダ判定 (ピリオド付き原文で呼ぶ) ---
Private Function IsParagraphHeader(ByVal raw As String, ByVal code As String) As Boolean
    Dim u As String: u = UCase(Trim(code))
    If Right(u, 1) <> "." Then Exit Function
    Dim body As String: body = Trim(Left(u, Len(u) - 1))   ' ピリオド除去
    ' SECTION ヘッダは列位置に依存せず内容で判定 (環境/桁ズレに強い)
    If body = "SECTION" Or Right(body, 8) = " SECTION" Then
        IsParagraphHeader = True: Exit Function
    End If
    ' 段落 (bare name) は Area A (8～11桁) 始まりのみ
    Dim col As Long: col = FirstNonSpaceCol(raw, CODE_START)
    If col < CODE_START Or col > 11 Then Exit Function
    Dim t As String: t = FirstToken(body)
    Select Case t
        Case "IF", "ELSE", "END-IF", "EVALUATE", "WHEN", "END-EVALUATE", _
             "PERFORM", "END-PERFORM", "SEARCH", "END-SEARCH", "MOVE", "CALL", _
             "COMPUTE", "ADD", "SUBTRACT", "MULTIPLY", "DIVIDE", "STRING", "UNSTRING", _
             "INSPECT", "SET", "GO", "GOBACK", "EXIT", "CONTINUE", "INITIALIZE", _
             "READ", "WRITE", "OPEN", "CLOSE", "ACCEPT", "DISPLAY", "THEN", "STOP"
            IsParagraphHeader = False
        Case Else: IsParagraphHeader = True
    End Select
End Function

Private Function StartsWithKw(ByVal code As String, ByVal kw As String) As Boolean
    Dim u As String: u = UCase(Trim(code))
    If u = kw Then StartsWithKw = True: Exit Function
    If Left(u, Len(kw) + 1) = kw & " " Then StartsWithKw = True
End Function

Private Function FirstToken(ByVal s As String) As String
    s = Trim(s)
    Dim p As Long: p = InStr(s, " ")
    If p = 0 Then FirstToken = s Else FirstToken = Left(s, p - 1)
End Function

Private Function MidSafe(ByVal s As String, ByVal startPos As Long, ByVal length As Long) As String
    If startPos > Len(s) Then Exit Function
    MidSafe = Mid(s, startPos, length)
End Function

Private Function FirstNonSpaceCol(ByVal s As String, ByVal fromCol As Long) As Long
    Dim i As Long
    For i = fromCol To Len(s)
        If Mid(s, i, 1) <> " " Then FirstNonSpaceCol = i: Exit Function
    Next i
End Function

Private Function EnsureSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nm
    End If
    Set EnsureSheet = ws
End Function
