Option Explicit

' ========= 列索引常量 =========
' A 列：盲孔(黑色梯形堆叠) + 埋孔(圆柱)；B 列：通孔(细金圆柱) + 背钻(绿色方块)
Private Const COL_VIA As Long = 1         ' 盲孔/埋孔图形列 (A)
Private Const COL_VIA_THROUGH As Long = 2 ' 通孔/背钻图形列 (B)

Private Const COL_ID As Long = 3          ' 标识 (C)
Private Const COL_TYPE As Long = 4        ' 层类别
Private Const COL_COPPER As Long = 5      ' 铜箔
Private Const COL_MAT_SHEET As Long = 6   ' 材料表
Private Const COL_CODE As Long = 7        ' PP Glass style
Private Const COL_PARAM As Long = 8       ' Param(含胶)/Resin(%)
Private Const COL_PLY As Long = 9         ' ply(梳理层数，厚度乘数)
Private Const COL_RESIDUAL As Long = 10   ' 残铜率(%)
Private Const COL_VENDOR As Long = 11     ' Vendor(mm)
Private Const COL_DK As Long = 12         ' DK
Private Const COL_DF As Long = 13         ' DF ★ 新增
Private Const COL_ACTUAL As Long = 14     ' Actual(mm)
Private Const COL_CUST_SPEC As Long = 15  ' 客规(um)
Private Const COL_PLANT_SPEC As Long = 16 ' 厂规(um) (P)

' ========= 图形前缀 & 孔型绘制常量 =========
Private Const G_PREFIX As String = "StackupGraphics_"

' 楔形（盲孔锥形）几何常量
Private Const WEDGE_LINE_VISIBLE As Boolean = True
Private Const MAX_PAIR_HEIGHT_RATIO As Double = 0.6        ' 楔形占可用高度比例上限
Private Const MIN_WEDGE_HALF_H As Double = 8               ' 楔形最小高（pt）
Private Const GAP_RATIO As Double = 0.15                   ' 上下楔形之间的间隙占比
Private Const COL_WIDTH_USAGE_RATIO As Double = 0.75       ' 楔形底宽不超过列宽的比例
Private Const TOP_TO_BOTTOM_RATIO As Double = 1.6          ' 上底宽 : 下底宽 比
Private Const MIN_TOP_WIDTH As Double = 6                  ' 楔形上底最小宽（pt）

' ============================================================
' Section 0: 通用清理工具（新增）
' ============================================================
' —— 删除指定工作表上的所有形状（含线/连接线/矩形/文本框）
' 用逐步删除避免 For Each 边遍历边删除导致的“遗漏”
Public Sub CleanAllShapes(ws As Worksheet)
    On Error Resume Next
    Do While ws.Shapes.Count > 0
        ws.Shapes(1).Delete
    Loop
    On Error GoTo 0
End Sub

' ============================================================
' Section 1: 通用工具 & 初始化
' ============================================================
Public Function EnsureSheet(wb As Workbook, sheetName As String) As Worksheet
    Set EnsureSheet = Nothing
    On Error Resume Next
    Set EnsureSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        EnsureSheet.Name = sheetName
    End If
End Function

Private Function TryGetSheet(wb As Workbook, sheetName As String) As Worksheet
    On Error Resume Next
    Set TryGetSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function WorksheetIsEmpty(ws As Worksheet) As Boolean
    WorksheetIsEmpty = (ws.Cells.Find(What:="*", LookIn:=xlFormulas, _
                     SearchOrder:=xlByRows, SearchDirection:=xlPrevious) Is Nothing)
End Function

' 清除窗口拆分/冻结状态，避免旧文件遗留的分割线被物化成 Line 图形
Private Sub ResetWindowPanes(ws As Worksheet)
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.Split = False
    ActiveWindow.SplitRow = 0
    ActiveWindow.SplitColumn = 0
    On Error GoTo 0
End Sub

Private Function ResetStackupSheetKeepSizing(wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ResetStackupSheetKeepSizing = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        ResetStackupSheetKeepSizing.Name = sheetName
    Else
        ' 先清除形状再清内容（防止遗留箭头线）
        CleanAllShapes ws
        ws.Cells.Clear
        Set ResetStackupSheetKeepSizing = ws
    End If
End Function

Private Function SafeParseLong(ByVal s As String, ByVal defaultVal As Long) As Long
    On Error GoTo Fail
    If Len(Trim$(s)) = 0 Then SafeParseLong = defaultVal: Exit Function
    If Not IsNumeric(s) Then SafeParseLong = defaultVal Else SafeParseLong = CLng(val(s))
    Exit Function
Fail:
    SafeParseLong = defaultVal
End Function

Private Function GetConfigValue(ws As Worksheet, key As String) As String
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If Trim$(CStr(ws.Cells(i, 1).Value)) = key Then
            GetConfigValue = Trim$(CStr(ws.Cells(i, 2).Value))
            Exit Function
        End If
    Next i
    GetConfigValue = ""
End Function

Private Sub InitializeConfigSheetIfMissing(ws As Worksheet)
    Dim doInit As Boolean: doInit = (Trim$(CStr(ws.Cells(1, 1).Value)) <> "Config Item")
    If WorksheetIsEmpty(ws) Then doInit = True
    If Not doInit Then Exit Sub

    ws.Cells.Clear
    Dim configData As Variant
    configData = Array( _
        Array("Config Item", "Value", "Description"), _
        Array("Execute", "TRUE", "TRUE=运行生成，FALSE=不运行"), _
        Array("总层数", "8", "PCB总层数 (1-50)"), _
        Array("CORE数量", "1", "CORE层数量"), _
        Array("是否对称", "TRUE", "TRUE=对称结构，FALSE=不对称"), _
        Array("埋孔", "NA", "埋孔层信息，如 L2-L6"), _
        Array("通孔", "是", "通孔：是=绘制 L1~LBottom 金色圆柱"), _
        Array("背钻", "NA", "背钻层信息，如 L1-L3;L8-L6"), _
        Array("量测方式", "SM-SM", "SM-SM / Au-Au / Cu-Cu"), _
        Array("残铜→厚度效应", "2", "残铜→厚度效应 (% per %)"), _
        Array("目标总厚", "1.2", "目标总厚度 (mm)"), _
        Array("使用油墨型号", "4000 MEH", "顶/底S/M显示的油墨名"), _
        Array("是否混压", "否", "是=问不同材料，填写相关层别"), _
        Array("材料1", "S1170G", "主要材料"), _
        Array("材料2", "NA", "（混压时用）"), _
        Array("启用冻结保护", "TRUE", "TRUE=冻结窗口并保护，FALSE=不上锁"), _
        Array("Laser阶数", "NA", "例如8L 若是2阶Laser，L1-L3和L8-L6"), _
        Array("是否手动增选CORE层别", "否", "例如8L 若是Core 数量 1 ，自动最中心"), _
        Array("CORE的层别", "NA", "例如8L 若是Core 数量 2 ，L3-L4层,L5-L6"), _
        Array("Core 厚度", "50.8", "以50.8为例，每次增加25.4um"), _
        Array("纯压PP Code", "1056", "例如1027"), _
        Array("混压时 PPCode", "NA", "那NA") _
    )
    Dim i As Long, j As Long
    For i = LBound(configData) To UBound(configData)
        For j = LBound(configData(i)) To UBound(configData(i))
            ws.Cells(i + 1, j + 1).Value = configData(i)(j)
        Next j
    Next i
    With ws
        .Rows(1).Font.Bold = True
        .Range("A1:C" & UBound(configData) + 1).Borders.LineStyle = xlContinuous
    End With
End Sub

' 确保孔型相关配置项存在（兼容旧 Config 表，缺失则追加）
Private Sub EnsureConfigExtraKeys(ws As Worksheet)
    EnsureConfigKey ws, "通孔", "是", "通孔：是=绘制 L1~LBottom 金色圆柱"
    EnsureConfigKey ws, "背钻", "NA", "背钻层信息，如 L1-L3;L8-L6"
End Sub

Private Sub EnsureConfigKey(ws As Worksheet, key As String, val As String, desc As String)
    Dim c As Range: Set c = GetCfgValueCell(ws, key)
    If Not c Is Nothing Then Exit Sub
    Dim r As Long: r = ws.Cells(ws.Rows.count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = key
    ws.Cells(r, 2).Value = val
    ws.Cells(r, 3).Value = desc
End Sub

Private Sub InitializeMaterialsIndex(ws As Worksheet)
    If WorksheetIsEmpty(ws) Then
        ws.Cells.Clear
        ws.Cells(1, 1).Value = "材料表名称"
        ws.Cells(2, 1).Value = "EMC-390"
        ws.Cells(3, 1).Value = "S-1150G"
        ws.Cells(4, 1).Value = "S-1150"
        ws.Cells(5, 1).Value = "S1150GH"
        ws.Cells(6, 1).Value = "S1170G"
        ws.Cells(7, 1).Value = "SDI06K"
        ws.Cells(8, 1).Value = "TU-862HF"
        ws.Cells(9, 1).Value = "SDI03K"
    End If
    With ws
        .Rows(1).Font.Bold = True
        .Range("A1:A" & Application.WorksheetFunction.Max(9, .Cells(.Rows.count, 1).End(xlUp).Row)).Borders.LineStyle = xlContinuous
    End With
End Sub

Private Sub ClearColorsExceptStackup()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Name <> "Stack-up(Optimized)" Then
            ws.Cells.Interior.Pattern = xlNone
            ws.Cells.Font.ColorIndex = xlColorIndexAutomatic
            On Error Resume Next: ws.Cells.FormatConditions.Delete: On Error GoTo 0
        End If
    Next ws
End Sub

' ============================================================
' Section 2: UDF（联动 Vendor(mm) & DK & DF）
' ============================================================
Public Function MatVendor(materialSheetName As String, t As String, code As String, rc As Variant, ply As Variant) As Variant
    Dim ws As Worksheet, rIdx As Long, v As Variant
    Set ws = TryGetSheet(ThisWorkbook, materialSheetName)
    If Not ws Is Nothing Then
        rIdx = FindMaterialRow(ws, t, code, rc, ply)
        If rIdx > 0 Then
            v = ws.Cells(rIdx, 5).Value ' Nominal(mm) = 1 ply 基准厚度
            If IsNumeric(v) Then
                Dim mult As Double: mult = 1
                If IsNumeric(ply) Then mult = CDbl(ply)
                MatVendor = CDbl(v) * mult   ' 厚度 × 梳理(ply)层数
            Else
                MatVendor = v
            End If
            Exit Function
        End If
    End If
    MatVendor = FindAcrossMaterialsIndex(5, t, code, rc, ply)
End Function

Public Function MatDK(materialSheetName As String, t As String, code As String, rc As Variant, ply As Variant) As Variant
    Dim ws As Worksheet, rIdx As Long, v As Variant
    Set ws = TryGetSheet(ThisWorkbook, materialSheetName)
    If Not ws Is Nothing Then
        rIdx = FindMaterialRow(ws, t, code, rc, ply)
        If rIdx > 0 Then
            v = ws.Cells(rIdx, 6).Value ' DK
            If IsNumeric(v) Then MatDK = CDbl(v) Else MatDK = v
            Exit Function
        End If
    End If
    MatDK = FindAcrossMaterialsIndex(6, t, code, rc, ply)
End Function

' ★ 新增：DF（耗散因子）
Public Function MatDF(materialSheetName As String, t As String, code As String, rc As Variant, ply As Variant) As Variant
    Dim ws As Worksheet, rIdx As Long, v As Variant
    Set ws = TryGetSheet(ThisWorkbook, materialSheetName)
    If Not ws Is Nothing Then
        rIdx = FindMaterialRow(ws, t, code, rc, ply)
        If rIdx > 0 Then
            v = ws.Cells(rIdx, 7).Value ' DF 在材料表第7列
            If IsNumeric(v) Then
                MatDF = CDbl(v)
                Exit Function
            ElseIf Len(Trim$(CStr(v))) > 0 Then
                MatDF = v
                Exit Function
            End If
            ' 该行 DF 为空：取同材料表内第一个非空 DF（材料级默认 DF）
            MatDF = FirstNonEmptyDF(ws)
            Exit Function
        End If
    End If
    MatDF = FindAcrossMaterialsIndex(7, t, code, rc, ply)
End Function

' 材料级 DF 兜底：取材料表内第一个非空 DF
Private Function FirstNonEmptyDF(ws As Worksheet) As Variant
    On Error GoTo Fail
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    Dim r As Long, v As Variant
    For r = 2 To lastRow
        v = ws.Cells(r, 7).Value
        If IsNumeric(v) Then FirstNonEmptyDF = CDbl(v): Exit Function
    Next r
Fail:
    FirstNonEmptyDF = ""
End Function

Private Function FindMaterialRow(ws As Worksheet, t As String, code As String, rc As Variant, ply As Variant) As Long
    On Error GoTo ExitHere
    FindMaterialRow = 0
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    Dim r As Long, t0 As String, c0 As String, rc0 As Variant, ply0 As Variant
    Dim resinPly1 As Long: resinPly1 = 0
    Dim resinAny As Long: resinAny = 0
    Dim codePly1 As Long: codePly1 = 0
    Dim codeAny As Long: codeAny = 0

    ' ply 作为“梳理层数”，是厚度乘数不再参与匹配；这里返回 1 ply 的基准行
    For r = 2 To lastRow
        t0 = UCase$(Trim$(CStr(ws.Cells(r, 1).Value))) ' Type
        c0 = Trim$(CStr(ws.Cells(r, 2).Value))         ' Code
        If t0 = UCase$(Trim$(t)) And c0 = Trim$(code) Then
            rc0 = ws.Cells(r, 3).Value                  ' Resin(%)
            ply0 = ws.Cells(r, 4).Value                 ' ply
            Dim resinMatch As Boolean: resinMatch = EqualNum(rc0, rc)
            If resinMatch And EqualNum(ply0, 1) And resinPly1 = 0 Then resinPly1 = r
            If resinMatch And resinAny = 0 Then resinAny = r
            If EqualNum(ply0, 1) And codePly1 = 0 Then codePly1 = r
            If codeAny = 0 Then codeAny = r
        End If
    Next r

    ' 优先级：Type+Code+含胶量(ply=1) → Type+Code+含胶量 → Type+Code(ply=1) → Type+Code
    If resinPly1 > 0 Then
        FindMaterialRow = resinPly1
    ElseIf resinAny > 0 Then
        FindMaterialRow = resinAny
    ElseIf codePly1 > 0 Then
        FindMaterialRow = codePly1
    ElseIf codeAny > 0 Then
        FindMaterialRow = codeAny
    End If
ExitHere:
End Function

Private Function FindAcrossMaterialsIndex(targetCol As Long, t As String, code As String, rc As Variant, ply As Variant) As Variant
    On Error GoTo Fail
    Dim wsIdx As Worksheet: Set wsIdx = TryGetSheet(ThisWorkbook, "MaterialsIndex")
    If wsIdx Is Nothing Then GoTo Fail
    Dim lastIdx As Long: lastIdx = wsIdx.Cells(wsIdx.Rows.count, 1).End(xlUp).Row
    Dim i As Long, nm As String, ws As Worksheet, rIdx As Long, v As Variant
    For i = 2 To lastIdx
        nm = Trim$(CStr(wsIdx.Cells(i, 1).Value))
        If Len(nm) > 0 Then
            Set ws = TryGetSheet(ThisWorkbook, nm)
            If Not ws Is Nothing Then
                rIdx = FindMaterialRow(ws, t, code, rc, ply)
                If rIdx > 0 Then
                    v = ws.Cells(rIdx, targetCol).Value
                    If IsNumeric(v) Then FindAcrossMaterialsIndex = CDbl(v) Else FindAcrossMaterialsIndex = v
                    Exit Function
                End If
            End If
        End If
    Next i
Fail:
    FindAcrossMaterialsIndex = ""
End Function

Private Function EqualNum(a As Variant, b As Variant) As Boolean
    On Error GoTo TextCmp
    If IsNumeric(a) And IsNumeric(b) Then
        EqualNum = (Abs(CDbl(a) - CDbl(b)) < 0.000001)
        Exit Function
    End If
TextCmp:
    EqualNum = (Trim$(CStr(a)) = Trim$(CStr(b)))
End Function

' ============================================================
' Section 3: Config UI（下拉与层对表）
' ============================================================
Public Function GetCfgValueCell(ws As Worksheet, itemName As String) As Range
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    Dim r As Long
    For r = 1 To lastRow
        If Trim$(CStr(ws.Cells(r, 1).Value)) = itemName Then
            Set GetCfgValueCell = ws.Cells(r, 2)
            Exit Function
        End If
    Next r
End Function

Public Function GetCfgValueMulti(ws As Worksheet, ParamArray keys() As Variant) As String
    Dim i As Long, v As String
    For i = LBound(keys) To UBound(keys)
        v = GetCfgValue(ws, CStr(keys(i)), "")
        If Len(Trim$(v)) > 0 Then
            GetCfgValueMulti = Trim$(v)
            Exit Function
        End If
    Next i
    GetCfgValueMulti = ""
End Function

Public Function GetCfgValueCellMulti(ws As Worksheet, ParamArray keys() As Variant) As Range
    Dim i As Long, c As Range
    For i = LBound(keys) To UBound(keys)
        Set c = GetCfgValueCell(ws, CStr(keys(i)))
        If Not c Is Nothing Then
            Set GetCfgValueCellMulti = c
            Exit Function
        End If
    Next i
End Function

Public Function GetCfgValue(ws As Worksheet, itemName As String, Optional defaultValue As String = "") As String
    Dim c As Range: Set c = GetCfgValueCell(ws, itemName)
    If c Is Nothing Then
        GetCfgValue = defaultValue
    Else
        GetCfgValue = Trim$(CStr(c.Value))
    End If
End Function

Public Function CollectPPList(materialSheetName As String) As Collection
    Dim col As New Collection
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(materialSheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set CollectPPList = col: Exit Function
    End If
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    Dim r As Long, t As String, code As String
    For r = 2 To lastRow
        t = Trim$(CStr(ws.Cells(r, 1).Value))
        code = Trim$(CStr(ws.Cells(r, 2).Value))
        If UCase$(t) = "PP" And Len(code) > 0 Then
            On Error Resume Next: col.Add code, code: On Error GoTo 0
        End If
    Next r
    Set CollectPPList = col
End Function

Public Function CollectPP_RC(materialSheetName As String, ppCode As String) As Collection
    Dim col As New Collection
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(materialSheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set CollectPP_RC = col: Exit Function
    End If
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    Dim r As Long, t As String, code As String
    For r = 2 To lastRow
        t = Trim$(CStr(ws.Cells(r, 1).Value))
        code = Trim$(CStr(ws.Cells(r, 2).Value))
        If UCase$(t) = "PP" And code = ppCode Then
            On Error Resume Next
            col.Add Format$(val(Replace(CStr(ws.Cells(r, 3).Value), "%", "")), "0.0"), _
                   CStr(val(Replace(CStr(ws.Cells(r, 3).Value), "%", "")))
            On Error GoTo 0
        End If
    Next r
    Set CollectPP_RC = col
End Function

Public Function WriteList(ws As Worksheet, anchor As Range, items As Collection, Optional clearRows As Long = 300) As Range
    Dim rng As Range: Set rng = ws.Range(anchor, anchor.Offset(clearRows, 0))
    rng.ClearContents
    Dim i As Long
    For i = 1 To items.count
        anchor.Offset(i - 1, 0).Value = items(i)
    Next i
    Set WriteList = ws.Range(anchor, anchor.Offset(items.count - 1, 0))
End Function

' —— 新增 showDropdown 参数：控制是否显示单元格下拉箭头
Public Sub SetDV_List(target As Range, listRange As Range, Optional allowBlank As Boolean = True, Optional showDropdown As Boolean = True)
    With target.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="=" & listRange.Address(True, True, xlA1, True)
        .IgnoreBlank = allowBlank
        .InCellDropdown = showDropdown   ' ← 可控下拉箭头
        .ShowError = True
        .ErrorTitle = "选择错误"
        .ErrorMessage = "请从下拉列表选择有效项。"
    End With
End Sub

Public Sub SetupConfigUI()
    Dim wsC As Worksheet: Set wsC = EnsureSheet(ThisWorkbook, "Config")

    ' Core 厚度步进下拉
    Dim baseUm As Double: baseUm = val(GetCfgValue(wsC, "Core 厚度", "50.8"))
    Dim coreList As Collection: Set coreList = GenCoreSteps(baseUm, 25.4, 50)
    Dim rngCore As Range: Set rngCore = WriteList(wsC, wsC.Range("Z2"), coreList, 500)
    Dim cCore As Range: Set cCore = GetCfgValueCell(wsC, "Core 厚度")
    If Not cCore Is Nothing Then SetDV_List cCore, rngCore, True, True

    ' PP 型号下拉（材料1+材料2 并集）
    Dim mat1 As String: mat1 = GetCfgValue(wsC, "材料1", "")
    Dim mat2 As String: mat2 = GetCfgValue(wsC, "材料2", "")
    Dim ppUnion As New Collection
    Dim l As Collection
    Set l = CollectPPList(mat1): AddAllUnique ppUnion, l
    Set l = CollectPPList(mat2): AddAllUnique ppUnion, l
    Dim rngPP As Range: Set rngPP = WriteList(wsC, wsC.Range("AA2"), ppUnion, 500)
    Dim cPurePP As Range: Set cPurePP = GetCfgValueCell(wsC, "纯压PP Code")
    Dim cMixPP As Range: Set cMixPP = GetCfgValueCell(wsC, "混压时 PPCode")
    If Not cPurePP Is Nothing Then SetDV_List cPurePP, rngPP, True, True
    If Not cMixPP Is Nothing Then SetDV_List cMixPP, rngPP, True, True

    ' 层对表（AD:层间；AE:PP；AF:RC(%)；AG:ply）
    Dim totalLayers As Long: totalLayers = val(GetCfgValue(wsC, "总层数", "8"))
    BuildInterlayerPPTable wsC, totalLayers, rngPP

    ' —— 额外：清除 Config 上可能遗留的形状（避免竖线箭头）
    CleanAllShapes wsC

    MsgBox "Config UI 初始化完成。"
End Sub

Private Sub AddAllUnique(ByRef target As Collection, ByVal src As Collection)
    Dim i As Long, k As String
    For i = 1 To src.count
        k = CStr(src(i))
        On Error Resume Next: target.Add src(i), k: On Error GoTo 0
    Next i
End Sub

Public Sub BuildInterlayerPPTable(wsC As Worksheet, totalLayers As Long, rngPP As Range)
    Dim anchor As Range: Set anchor = wsC.Range("AD1")
    anchor.Offset(0, 0).Value = "层间"
    anchor.Offset(0, 1).Value = "PP型号"
    anchor.Offset(0, 2).Value = "RC(%)"
    anchor.Offset(0, 3).Value = "ply"
    wsC.Range(anchor.Offset(1, 0), anchor.Offset(400, 3)).ClearContents

    Dim pairCount As Long: pairCount = Application.WorksheetFunction.Max(0, totalLayers - 1)
    Dim r As Long
    For r = 1 To pairCount
        anchor.Offset(r, 0).Value = "L" & r & "-L" & (r + 1)
        ' 这里可以选择隐藏下拉箭头：showDropdown:=False
        SetDV_List anchor.Offset(r, 1), rngPP, True, False
    Next r
End Sub

Public Sub SetRCForRow(wsC As Worksheet, targetPPCell As Range)
    Dim ppCode As String: ppCode = Trim$(CStr(targetPPCell.Value))
    If Len(ppCode) = 0 Then Exit Sub
    Dim mat1 As String: mat1 = GetCfgValue(wsC, "材料1", "")
    Dim mat2 As String: mat2 = GetCfgValue(wsC, "材料2", "")

    Dim rcUnion As New Collection
    Dim c As Collection
    Set c = CollectPP_RC(mat1, ppCode): AddAllUnique rcUnion, c
    Set c = CollectPP_RC(mat2, ppCode): AddAllUnique rcUnion, c
    If rcUnion.count = 0 Then Exit Sub

    Dim rngRC As Range: Set rngRC = WriteList(wsC, wsC.Range("AB2"), rcUnion, 500)
    Dim rcCell As Range: Set rcCell = targetPPCell.Offset(0, 1)
    SetDV_List rcCell, rngRC, True, False
End Sub

Public Function GenCoreSteps(baseUm As Double, stepUm As Double, Optional count As Long = 40) As Collection
    Dim col As New Collection, i As Long
    For i = 0 To count - 1
        col.Add Round(baseUm + stepUm * i, 1)
    Next i
    Set GenCoreSteps = col
End Function

Public Function IsReady(ws As Worksheet) As Boolean
    Dim execFlag As String: execFlag = UCase$(GetCfgValue(ws, "Execute", "FALSE"))
    If execFlag = "TRUE" Then IsReady = True: Exit Function
    Dim n As String: n = GetCfgValue(ws, "总层数", "")
    If Len(n) = 0 Then IsReady = False: Exit Function
    Dim isMix As String: isMix = GetCfgValue(ws, "是否混压", "否")
    If UCase$(isMix) = "是" Or UCase$(isMix) = "TRUE" Then
        IsReady = (Len(GetCfgValue(ws, "混压时 PPCode", "")) > 0)
    Else
        IsReady = (Len(GetCfgValue(ws, "纯压PP Code", "")) > 0)
    End If
End Function

' ============================================================
' Section 4: Runner（执行入口）
' ============================================================
Public Sub RunIfReady()
    Dim wsC As Worksheet: Set wsC = EnsureSheet(ThisWorkbook, "Config")
    If Not IsReady(wsC) Then
        MsgBox "Config 关键项未完成，请先选择。"
        Exit Sub
    End If
    Call BuildStackup
End Sub

Public Sub OneClick_Setup_And_Run()
    Call SetupConfigUI
    Call RunIfReady
End Sub

' ============================================================
' Section 5: 堆层生成（核心）
' ============================================================
Public Sub BuildStackup()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    On Error GoTo EH

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsCfg As Worksheet, wsIdx As Worksheet, wsOpt As Worksheet
    Set wsCfg = EnsureSheet(wb, "Config")
    Set wsIdx = EnsureSheet(wb, "MaterialsIndex")
    Set wsOpt = ResetStackupSheetKeepSizing(wb, "Stack-up(Optimized)")

    ' 清除旧窗口拆分/冻结状态（避免遗留分割线）
    ResetWindowPanes wsOpt

    InitializeConfigSheetIfMissing wsCfg
    EnsureConfigExtraKeys wsCfg
    InitializeMaterialsIndex wsIdx

    Dim exec As String: exec = GetConfigValue(wsCfg, "Execute")
    If UCase$(Trim$(exec)) <> "TRUE" Then
        MsgBox "Config 中 Execute=FALSE，未执行生成。", vbInformation
        GoTo ExitHere
    End If

    Dim n As Long: n = CLng(val(GetConfigValue(wsCfg, "总层数")))
    Dim measure As String: measure = GetConfigValue(wsCfg, "量测方式")
    Dim tNom As Double: tNom = CDbl(val(GetConfigValue(wsCfg, "目标总厚")))
    Dim inkName As String: inkName = GetConfigValue(wsCfg, "使用油墨型号")
    Dim effect As Double: effect = CDbl(val(GetConfigValue(wsCfg, "残铜→厚度效应")))
    Dim enableFreeze As Boolean: enableFreeze = (UCase$(GetConfigValue(wsCfg, "启用冻结保护")) = "TRUE")
    Dim isMixed As Boolean: isMixed = (UCase$(GetConfigValue(wsCfg, "是否混压")) = "是" _
                                    Or UCase$(GetConfigValue(wsCfg, "是否混压")) = "TRUE" _
                                    Or UCase$(GetConfigValue(wsCfg, "是否混压")) = "YES")
    Dim coreCount As Long: coreCount = CLng(val(GetConfigValue(wsCfg, "CORE数量")))
    Dim enforceSymmetry As Boolean: enforceSymmetry = (UCase$(GetConfigValue(wsCfg, "是否对称")) = "TRUE")
    Dim buriedSpec As String: buriedSpec = GetConfigValue(wsCfg, "埋孔")
    Dim material1 As String: material1 = GetConfigValue(wsCfg, "材料1")
    Dim material2 As String: material2 = GetConfigValue(wsCfg, "材料2")
    Dim laserRaw As String: laserRaw = GetConfigValue(wsCfg, "Laser阶数")
    Dim laserStage As Long: laserStage = SafeParseLong(laserRaw, 0)
    Dim throughRaw As String: throughRaw = GetConfigValue(wsCfg, "通孔")
    Dim throughOn As Boolean
    If UCase$(throughRaw) = "否" Or UCase$(throughRaw) = "FALSE" Or UCase$(throughRaw) = "NO" Then
        throughOn = False
    Else
        throughOn = True   ' 默认绘制通孔
    End If
    Dim backDrillSpec As String: backDrillSpec = GetConfigValue(wsCfg, "背钻")
    Dim manualCore As Boolean: manualCore = (UCase$(GetConfigValue(wsCfg, "是否手动增选CORE层别")) = "是" _
                                           Or UCase$(GetConfigValue(wsCfg, "是否手动增选CORE层别")) = "TRUE")
    Dim coreSpec As String: coreSpec = GetConfigValue(wsCfg, "CORE的层别")
    Dim coreThkUm As Double: coreThkUm = CDbl(val(GetConfigValue(wsCfg, "Core 厚度")))

    If n <= 0 Or n > 50 Then
        MsgBox "总层数必须在 1–50 之间", vbExclamation
        GoTo ExitHere
    End If

    EnsureMaterialSheetsSafe wb

    Dim mixRanges As Collection: Set mixRanges = New Collection
    If isMixed Then
        CollectMixedRanges n, mixRanges
        If mixRanges.count = 0 Then
            Dim midPos As Long: midPos = Application.WorksheetFunction.RoundDown(n / 2, 0)
            AddRange mixRanges, 1, Application.WorksheetFunction.Max(1, midPos), material1
            If midPos < n - 1 Then AddRange mixRanges, midPos + 1, n - 1, material2
        End If
    End If

    GenerateStackupLayout wsOpt, n, effect, measure, inkName, tNom, _
                          enableFreeze, isMixed, coreCount, enforceSymmetry, _
                          buriedSpec, material1, material2, mixRanges, laserStage, _
                          manualCore, coreSpec, coreThkUm, throughOn, backDrillSpec

    ' 映射 Config 层对选择到堆叠：覆盖 PP 行的 Code/Param/ply
    ApplyInterlayerSelections wsCfg, wsOpt, n

    SetupDataValidation wsOpt, wsIdx
    ApplyStyling wsOpt, enableFreeze
    EnsureMinimumColumnWidths wsOpt
    EnsureTotalsTitlesFit wsOpt
    ClearColorsExceptStackup
    OptimizeLayout

    ' 清理非本程序创建的残留形状（如遗留直线/箭头）
    RemoveStrayShapes wsOpt

    MsgBox "PCB堆层已生成并应用层对选择。", vbInformation

ExitHere:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

EH:
    MsgBox "错误: " & Err.Description, vbCritical
    Resume ExitHere
End Sub

Private Sub GenerateStackupLayout(ws As Worksheet, n As Long, effect As Double, _
    measure As String, inkName As String, _
    tNom As Double, enableFreeze As Boolean, _
    isMixed As Boolean, coreCount As Long, _
    enforceSymmetry As Boolean, buriedSpec As String, _
    material1 As String, material2 As String, _
    mixRanges As Collection, laserStage As Long, _
    manualCore As Boolean, coreSpec As String, _
    coreThkUm As Double, throughOn As Boolean, backDrillSpec As String)

    ' —— 开始就清除形状（避免遗留箭头线）
    CleanAllShapes ws

    ws.Cells.Clear
    WriteTableHeader ws

    Dim r As Long: r = 2
    WriteSMRow ws, r, "Top", inkName, measure
    r = r + 1

    Dim corePositions As Collection
    If manualCore Then
        Set corePositions = ParseManualCorePositions(n, coreSpec, coreCount)
        If corePositions Is Nothing Or corePositions.count = 0 Then
            Set corePositions = CalculateCorePositions(n, coreCount, enforceSymmetry)
        End If
    Else
        Set corePositions = CalculateCorePositions(n, coreCount, enforceSymmetry)
    End If

    Dim i As Long
    For i = 1 To n
        WriteCopperLayerRow ws, r, i, n
        r = r + 1
        If i < n Then
            Dim isCoreLayer As Boolean: isCoreLayer = IsInCollection(corePositions, i)
            Dim chosenMaterial As String
            If isMixed Then
                chosenMaterial = GetMaterialByRanges(i, mixRanges, material1)
            Else
                chosenMaterial = material1
            End If
            WriteDielectricRow ws, r, i, isCoreLayer, chosenMaterial, enforceSymmetry, n, coreThkUm, effect
            r = r + 1
        End If
    Next i

    WriteSMRow ws, r, "Bottom", inkName, measure
    r = r + 1

    ' —— 孔型绘制（A/B 图形列）——
    ' 盲孔：黑色梯形，上下对称
    If laserStage > 0 Then
        DrawBlindVias ws, n, laserStage
    End If

    ' 埋孔：圆柱（等径钻孔）
    If Len(Trim$(buriedSpec)) > 0 And UCase$(Trim$(buriedSpec)) <> "NA" Then
        DrawBuriedVias ws, n, buriedSpec
    End If

    ' 通孔：金色圆柱，L1 ~ LBottom
    If throughOn Then
        DrawThroughVias ws, n
    End If

    ' 背钻：灰色宽圆柱覆盖通孔两端残桩
    If Len(Trim$(backDrillSpec)) > 0 And UCase$(Trim$(backDrillSpec)) <> "NA" Then
        DrawBackDrillVias ws, n, backDrillSpec
    End If

    r = r + 1
    WriteTotalsRow ws, r, tNom, measure
End Sub

Private Sub WriteTableHeader(ws As Worksheet)
    ' 两列孔型图形列表头：A=盲孔/埋孔，B=通孔/背钻
    ws.Cells(1, COL_VIA).Value = "盲孔/埋孔"
    ws.Cells(1, COL_VIA_THROUGH).Value = "通孔/背钻"

    Dim headers As Variant
    headers = Array("标识", "层类别", "铜箔", "材料表", "Glass_Style", "Param(含胶量)", "ply", _
                    "残铜率(%)", "Vendor(mm)", "DK", "DF", "Actual(mm)", "客规(um)", "厂规(um)")
    Dim i As Long
    For i = 0 To UBound(headers)
        ws.Cells(1, COL_ID + i).Value = headers(i)
    Next i
End Sub

Private Sub WriteSMRow(ws As Worksheet, r As Long, position As String, inkName As String, measure As String)
    ws.Cells(r, COL_ID).Value = "S/M (" & position & ")"
    ws.Cells(r, COL_TYPE).Value = inkName
    ws.Cells(r, COL_CUST_SPEC).Formula = "=IF(""" & measure & """=""SM-SM"",18,"""")"
    ws.Cells(r, COL_PLANT_SPEC).Formula = "=IF(""" & measure & """=""SM-SM"",18,"""")"
    ws.Cells(r, COL_ACTUAL).FormulaR1C1 = _
        "=IF(ISNUMBER(RC" & COL_CUST_SPEC & "),ROUND(RC" & COL_CUST_SPEC & "/1000,4),"""")"
End Sub

Private Sub WriteCopperLayerRow(ws As Worksheet, r As Long, layerNum As Long, n As Long)
    ws.Cells(r, COL_ID).Value = "L" & layerNum
    ws.Cells(r, COL_TYPE).Value = IIf(layerNum Mod 2 = 0, "Plane", "Signal")
    ws.Cells(r, COL_COPPER).Value = IIf(layerNum <= 2 Or layerNum >= n - 1, "MLS-G", "MLS-G3")
    If layerNum <= 2 Or layerNum >= n - 1 Then
        ws.Cells(r, COL_RESIDUAL).Value = 65
        ws.Cells(r, COL_CUST_SPEC).Value = 23
    Else
        ws.Cells(r, COL_RESIDUAL).Value = IIf(layerNum Mod 2 = 0, 78, 85)
        ws.Cells(r, COL_CUST_SPEC).Value = IIf(ws.Cells(r, COL_RESIDUAL).Value = 85, 26, 23)
    End If
    ws.Cells(r, COL_ACTUAL).FormulaR1C1 = _
        "=IF(ISNUMBER(RC" & COL_CUST_SPEC & "),ROUND(RC" & COL_CUST_SPEC & "/1000,4),"""")"
End Sub

' === 从材料表抓默认 Code/Resin（支持纯压默认） ===
Private Function GetFirstCode(materialSheetName As String, t As String) As String
    Dim ws As Worksheet, lastRow As Long, r As Long
    Set ws = TryGetSheet(ThisWorkbook, materialSheetName)
    GetFirstCode = IIf(UCase$(t) = "CORE", "2116", "1080") ' 兜底默认
    If ws Is Nothing Then Exit Function
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = UCase$(t) Then
            If Len(Trim$(CStr(ws.Cells(r, 2).Value))) > 0 Then
                GetFirstCode = Trim$(CStr(ws.Cells(r, 2).Value))
                Exit Function
            End If
        End If
    Next r
End Function

Private Function GetResinValue(materialSheetName As String, t As String, code As String) As Variant
    Dim ws As Worksheet, lastRow As Long, r As Long
    Set ws = TryGetSheet(ThisWorkbook, materialSheetName)
    GetResinValue = IIf(UCase$(t) = "CORE", 68, 72) ' 兜底默认
    If ws Is Nothing Then Exit Function
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    For r = 2 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(r, 1).Value))) = UCase$(t) _
        And Trim$(CStr(ws.Cells(r, 2).Value)) = Trim$(code) Then
            GetResinValue = ws.Cells(r, 3).Value ' Resin(%)
            Exit Function
        End If
    Next r
End Function

Private Function IsPurePress() As Boolean
    Dim wsCfg As Worksheet: Set wsCfg = TryGetSheet(ThisWorkbook, "Config")
    If wsCfg Is Nothing Then IsPurePress = True: Exit Function
    IsPurePress = Not (UCase$(GetConfigValue(wsCfg, "是否混压")) = "是" _
                    Or UCase$(GetConfigValue(wsCfg, "是否混压")) = "TRUE" _
                    Or UCase$(GetConfigValue(wsCfg, "是否混压")) = "YES")
End Function

Private Sub WriteDielectricRow(ws As Worksheet, r As Long, _
    layerNum As Long, isCoreLayer As Boolean, _
    chosenMaterial As String, enforceSymmetry As Boolean, _
    n As Long, coreThkUm As Double, effect As Double)

    ws.Cells(r, COL_ID).Value = IIf(isCoreLayer, "CORE", "PP")
    ws.Cells(r, COL_MAT_SHEET).Value = chosenMaterial

    ' 默认 Code/Param（纯压下按材料表抓取）
    Dim defCode As String, defRC As Variant
    If isCoreLayer Then
        defCode = GetFirstCode(chosenMaterial, "CORE")
        defRC = GetResinValue(chosenMaterial, "CORE", defCode)
    Else
        defCode = GetFirstCode(chosenMaterial, "PP")
        defRC = GetResinValue(chosenMaterial, "PP", defCode)
    End If
    ws.Cells(r, COL_CODE).Value = defCode
    ws.Cells(r, COL_PARAM).Value = defRC
    ws.Cells(r, COL_PLY).Value = 1
    ws.Cells(r, COL_RESIDUAL).ClearContents

    ' UDF 联动（Vendor/DK/DF）
    ws.Cells(r, COL_VENDOR).FormulaR1C1 = "=MatVendor(RC" & COL_MAT_SHEET & ",RC" & COL_ID & ",RC" & COL_CODE & ",RC" & COL_PARAM & ",RC" & COL_PLY & ")"
    ws.Cells(r, COL_DK).FormulaR1C1 = "=MatDK(RC" & COL_MAT_SHEET & ",RC" & COL_ID & ",RC" & COL_CODE & ",RC" & COL_PARAM & ",RC" & COL_PLY & ")"
    ws.Cells(r, COL_DF).FormulaR1C1 = "=MatDF(RC" & COL_MAT_SHEET & ",RC" & COL_ID & ",RC" & COL_CODE & ",RC" & COL_PARAM & ",RC" & COL_PLY & ")"

    ' CORE 厚度覆盖 Vendor(mm)
    If UCase$(CStr(ws.Cells(r, COL_ID).Value)) = "CORE" Then
        If coreThkUm > 0 Then
            ws.Cells(r, COL_VENDOR).Formula = ""
            ws.Cells(r, COL_VENDOR).Value = Round(coreThkUm / 1000, 6)
        End If
    End If

    ' 客规(um)
    ws.Cells(r, COL_CUST_SPEC).FormulaR1C1 = _
        "=IF(RC" & COL_ID & "=""PP"",IF(ISNUMBER(RC" & COL_VENDOR & "),ROUND(RC" & COL_VENDOR & "*1000,0),""""),"""")"

    ' Actual(mm)：厚度 = Vendor（MatVendor 已含 ×ply 梳理层数），PP/CORE 一致
    ws.Cells(r, COL_ACTUAL).FormulaR1C1 = _
        "=IF(ISNUMBER(RC" & COL_VENDOR & "),RC" & COL_VENDOR & ","""")"
End Sub

Private Sub WriteTotalsRow(ws As Worksheet, r As Long, tNom As Double, measure As String)
    ws.Cells(r, COL_ID).Value = "Totals"
    ws.Cells(r, COL_TYPE).Value = "Total nominal (mm)"
    ws.Cells(r, COL_COPPER).Value = "Total actual (mm)"
    ws.Cells(r, COL_MAT_SHEET).Value = "Total delta (mm)"
    ws.Cells(r, COL_CODE).Value = "量测方式"
    ws.Cells(r, COL_CUST_SPEC).Value = "总板厚(um)"

    ' 在 Totals 的下一行写数值/公式
    ws.Cells(r + 1, COL_TYPE).Value = tNom

    ' 动态汇总 Actual(mm)
    Dim lastActRow As Long: lastActRow = ws.Cells(ws.Rows.count, COL_ACTUAL).End(xlUp).Row
    If lastActRow < 2 Then lastActRow = r - 1
    ws.Cells(r + 1, COL_COPPER).FormulaR1C1 = "=ROUND(SUM(R2C" & COL_ACTUAL & ":R" & lastActRow & "C" & COL_ACTUAL & "),4)"

    ' 总差值 = Total actual - Total nominal
    ws.Cells(r + 1, COL_MAT_SHEET).FormulaR1C1 = _
        "=ROUND(R" & (r + 1) & "C" & COL_COPPER & "-R" & (r + 1) & "C" & COL_TYPE & ",4)"

    ws.Cells(r + 1, COL_CODE).Value = measure

    ' 总板厚(um) = 客规(um)列求和（动态从第2行到 Totals 前一行）
    Dim lastSpecRow As Long: lastSpecRow = r - 1
    ws.Cells(r + 1, COL_CUST_SPEC).FormulaR1C1 = _
        "=SUM(R2C" & COL_CUST_SPEC & ":R" & lastSpecRow & "C" & COL_CUST_SPEC & ")"
End Sub

' =================== 孔型绘制（盲孔/埋孔/通孔/背钻，同列 A） ===================
' 埋孔：圆柱（等径钻孔），橙色，宽度中等
Private Sub DrawBuriedVias(ws As Worksheet, n As Long, buriedSpec As String)
    Dim pairs As Collection: Set pairs = ParseBuriedViasToPairs(buriedSpec, n)
    Dim item As Variant
    For Each item In pairs
        Dim startL As Long: startL = CLng(item(0))
        Dim endL As Long: endL = CLng(item(1))
        DrawCylinderVia ws, COL_VIA, RowIndexForLayer(startL), RowIndexForLayer(endL), _
                        RGB(255, 165, 0), RGB(180, 90, 0), _
                        "埋 L" & startL & "-L" & endL, RGB(80, 40, 0), 0.45, 0.1
    Next item
End Sub

' 通孔：金色细圆柱（缩小钻孔），L1 ~ LBottom（整板贯穿），B 列
Private Sub DrawThroughVias(ws As Worksheet, n As Long)
    DrawCylinderVia ws, COL_VIA_THROUGH, RowIndexForLayer(1), RowIndexForLayer(n), _
                    RGB(255, 215, 0), RGB(184, 134, 11), _
                    "通 L1-L" & n, RGB(90, 60, 0), 0.28, 0.1
End Sub

' 背钻：绿色方块标记（B 列靠右放置，不与通孔重叠）
Private Sub DrawBackDrillVias(ws As Worksheet, n As Long, spec As String)
    Dim pairs As Collection: Set pairs = ParseBuriedViasToPairs(spec, n)
    Dim item As Variant
    For Each item In pairs
        Dim sL As Long: sL = CLng(item(0))
        Dim eL As Long: eL = CLng(item(1))
        DrawBackDrillSquares ws, COL_VIA_THROUGH, sL, eL, "背钻 L" & sL & "-L" & eL
    Next item
End Sub

' 在背钻起止层各画一个绿色方块（靠右偏移，与居中的通孔不重合）
Private Sub DrawBackDrillSquares(ws As Worksheet, col As Long, sL As Long, eL As Long, labelText As String)
    Dim lo As Long: lo = Application.WorksheetFunction.Min(sL, eL)
    Dim hi As Long: hi = Application.WorksheetFunction.Max(sL, eL)
    Dim rTop As Long: rTop = RowIndexForLayer(lo)
    Dim rBot As Long: rBot = RowIndexForLayer(hi)

    Dim leftPos As Double: leftPos = ws.Columns(col).Left
    Dim widthPos As Double: widthPos = ws.Columns(col).Width
    Dim sq As Double: sq = Application.WorksheetFunction.Min(widthPos * 0.3, 10)
    If sq < 6 Then sq = 6
    Dim xLeft As Double: xLeft = leftPos + widthPos * 0.6
    Dim yTop As Double: yTop = ws.Cells(rTop, col).Top + 1
    Dim yBot As Double: yBot = ws.Cells(rBot, col).Top + ws.Cells(rBot, col).Height - 1

    Dim shp As Shape, shpName As String, j As Long
    For j = 1 To 2
        If j = 1 Then
            Set shp = ws.Shapes.AddShape(msoShapeRectangle, xLeft, yTop, sq, sq)
            shpName = G_PREFIX & "BackDrill_" & lo
        Else
            Set shp = ws.Shapes.AddShape(msoShapeRectangle, xLeft, yBot - sq, sq, sq)
            shpName = G_PREFIX & "BackDrill_" & hi
        End If
        shp.Name = shpName
        shp.Fill.ForeColor.RGB = RGB(0, 176, 80)      ' 绿色
        shp.Fill.Transparency = 0.1
        shp.Line.ForeColor.RGB = RGB(0, 100, 40)
        shp.Line.Weight = 1#
        shp.ZOrder msoBringToFront
    Next j

    AddViaLabel ws, xLeft, (yTop + yBot) / 2, labelText, RGB(0, 100, 40)
End Sub

Private Function ParseBuriedViasToPairs(spec As String, n As Long) As Collection
    Set ParseBuriedViasToPairs = New Collection
    If Len(Trim$(spec)) = 0 Then Exit Function
    Dim normalized As String
    normalized = UCase$(spec)
    normalized = Replace(normalized, "L", "")
    normalized = Replace(normalized, ";", " ")
    normalized = Replace(normalized, ",", " ")
    normalized = Replace(normalized, "，", " ")
    normalized = Replace(normalized, "/", "-")
    Dim parts() As String: parts = Split(normalized, " ")
    Dim p As Variant
    For Each p In parts
        If InStr(1, p, "-", vbTextCompare) > 0 Then
            Dim ab() As String: ab = Split(p, "-")
            If UBound(ab) = 1 Then
                Dim sL As Long, eL As Long
                sL = CLng(val(ab(0)))
                eL = CLng(val(ab(1)))
                sL = Application.WorksheetFunction.Max(1, Application.WorksheetFunction.Min(n, sL))
                eL = Application.WorksheetFunction.Max(1, Application.WorksheetFunction.Min(n, eL))
                If sL <= eL Then ParseBuriedViasToPairs.Add Array(sL, eL)
            End If
        End If
    Next p
End Function

' ========= 盲孔（激光孔）绘制：黑色梯形按“阶”堆叠，上下对称 =========
' 每个“阶”代表一次激光钻，画一个梯形；多阶即多个梯形上下堆叠（如 3 阶 = 3 个梯形）
Private Sub DrawBlindVias(ws As Worksheet, n As Long, laserStage As Long)
    If laserStage < 1 Then Exit Sub
    Dim topStart As Long: topStart = 1
    Dim topEnd As Long: topEnd = Application.WorksheetFunction.Min(n, 1 + laserStage)
    Dim botStart As Long: botStart = n
    Dim botEnd As Long: botEnd = Application.WorksheetFunction.Max(1, n - laserStage)

    Dim k As Long
    ' 顶侧：每个阶一个梯形（锥尖朝下），仅在首个梯形加整段标签
    For k = topStart To topEnd - 1
        DrawBlindTrapezoid ws, COL_VIA, k, k + 1, False, _
                           IIf(k = topStart, "盲 L" & topStart & "-L" & topEnd, "")
    Next k

    ' 底侧：镜像（锥尖朝上）
    For k = botStart To botEnd + 1 Step -1
        DrawBlindTrapezoid ws, COL_VIA, k, k - 1, True, _
                           IIf(k = botStart, "盲 L" & botEnd & "-L" & botStart, "")
    Next k
End Sub

' 画单个盲孔梯形。bottomSurface=False 表示表层在上（上宽下窄）；True 表示表层在下（上窄下宽）
Private Sub DrawBlindTrapezoid(ws As Worksheet, col As Long, fromLayer As Long, toLayer As Long, _
                               ByVal bottomSurface As Boolean, labelText As String)
    Dim lo As Long: lo = Application.WorksheetFunction.Min(fromLayer, toLayer)
    Dim hi As Long: hi = Application.WorksheetFunction.Max(fromLayer, toLayer)
    If lo = hi Then Exit Sub

    Dim rSurf As Long: rSurf = RowIndexForLayer(lo)
    Dim rTip As Long: rTip = RowIndexForLayer(hi)

    Dim leftPos As Double: leftPos = ws.Columns(col).Left
    Dim widthPos As Double: widthPos = ws.Columns(col).Width
    Dim xCenter As Double: xCenter = leftPos + widthPos / 2

    Dim topY As Double: topY = ws.Cells(rSurf, col).Top
    Dim botY As Double: botY = ws.Cells(rTip, col).Top + ws.Cells(rTip, col).Height
    If botY <= topY Then Exit Sub

    Dim bodyW As Double: bodyW = Application.WorksheetFunction.Min(widthPos * COL_WIDTH_USAGE_RATIO, 18)
    bodyW = Application.WorksheetFunction.Max(bodyW, MIN_TOP_WIDTH * TOP_TO_BOTTOM_RATIO)
    Dim tipW As Double: tipW = Application.WorksheetFunction.Max(bodyW / TOP_TO_BOTTOM_RATIO, MIN_TOP_WIDTH)

    Dim shp As Shape
    If Not bottomSurface Then
        Set shp = AddTrapezoidWedge(ws, xCenter, topY, botY, bodyW, tipW)          ' 上宽下窄
    Else
        Set shp = AddTrapezoidWedge(ws, xCenter, topY, botY, bodyW, tipW, True)    ' 上窄下宽
    End If
    shp.Fill.ForeColor.RGB = RGB(0, 0, 0)        ' 盲孔黑色
    shp.Line.ForeColor.RGB = RGB(0, 0, 0)
    shp.Name = G_PREFIX & "Blind_" & lo & "_" & hi

    ' 独立文本框标签（自由形状上直接加文字会产生多余的 Line 图形）
    If Len(labelText) > 0 Then
        AddViaLabel ws, xCenter, (topY + botY) / 2, labelText, RGB(255, 255, 255)
    End If
End Sub

' 构建梯形楔形（可选镜像）。默认上宽下窄；flip=True 时上窄下宽
Private Function AddTrapezoidWedge(ws As Worksheet, _
    xCenter As Double, yTop As Double, yBottom As Double, _
    topW As Double, bottomW As Double, _
    Optional ByVal flip As Boolean = False) As Shape

    Dim upW As Double: upW = IIf(Not flip, topW, bottomW)
    Dim dnW As Double: dnW = IIf(Not flip, bottomW, topW)

    Dim ff As FreeformBuilder
    Set ff = ws.Shapes.BuildFreeform(msoEditingCorner, xCenter - upW / 2, yTop)
    ff.AddNodes msoSegmentLine, msoEditingAuto, xCenter + upW / 2, yTop
    ff.AddNodes msoSegmentLine, msoEditingAuto, xCenter + dnW / 2, yBottom
    ff.AddNodes msoSegmentLine, msoEditingAuto, xCenter - dnW / 2, yBottom
    ff.AddNodes msoSegmentLine, msoEditingAuto, xCenter - upW / 2, yTop

    Dim shp As Shape
    Set shp = ff.ConvertToShape
    With shp
        .Line.Visible = WEDGE_LINE_VISIBLE
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(0, 0, 0)   ' 默认黑色，调用方可覆盖
        .Fill.Transparency = 0#
        .ZOrder msoBringToFront
    End With
    Set AddTrapezoidWedge = shp
End Function

' 独立文字标签（用于自由形状等不适合直接加文字的形状）
Private Sub AddViaLabel(ws As Worksheet, xCenter As Double, yMid As Double, txt As String, fontRGB As Long)
    Dim tb As Shape
    Set tb = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, xCenter - 30, yMid - 8, 60, 16)
    tb.Name = G_PREFIX & "Label"
    tb.TextFrame.Characters.Text = txt
    tb.TextFrame.Characters.Font.Size = 8
    tb.TextFrame.Characters.Font.Color = fontRGB
    tb.TextFrame.HorizontalAlignment = xlHAlignCenter
    tb.TextFrame.VerticalAlignment = xlVAlignCenter
    tb.Line.Visible = msoFalse
    tb.Fill.Visible = msoFalse
End Sub

' 删除非本程序创建（名称不以 G_PREFIX 开头）的残留形状，如遗留的直线/箭头
Private Sub RemoveStrayShapes(ws As Worksheet)
    On Error Resume Next
    Dim shp As Shape
    Dim namesToDel As New Collection
    For Each shp In ws.Shapes
        If Left$(shp.Name, Len(G_PREFIX)) <> G_PREFIX Then
            namesToDel.Add shp.Name
        End If
    Next shp
    Dim nm As Variant
    For Each nm In namesToDel
        ws.Shapes(CStr(nm)).Delete
    Next nm
    On Error GoTo 0
End Sub

' 安全标签（TextFrame2 优先，失败回退 TextFrame）
Private Sub LabelShape(shp As Shape, txt As String, fontRGB As Long)
    On Error GoTo OldTextFrame
    With shp.TextFrame2
        .AutoSize = msoAutoSizeNone
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = txt
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextRange.Font.Size = 10
        .TextRange.Font.Fill.ForeColor.RGB = fontRGB
    End With
    Exit Sub
OldTextFrame:
    On Error Resume Next
    shp.TextFrame.Characters.Text = txt
    shp.TextFrame.Characters.Font.Size = 10
    shp.TextFrame.Characters.Font.Color = fontRGB
    shp.TextFrame.HorizontalAlignment = xlHAlignCenter
    shp.TextFrame.VerticalAlignment = xlVAlignCenter
End Sub

' 圆柱孔（等径钻孔）：跨 rowTop~rowBottom，居中于列 col
' widthRatio=占列宽比例（0.28 通孔偏细 / 0.45 埋孔 / 0.7 背钻偏宽）；transparency=填充透明度
Private Sub DrawCylinderVia(ws As Worksheet, col As Long, rowTop As Long, rowBottom As Long, _
                            fillRGB As Long, lineRGB As Long, labelText As String, fontRGB As Long, _
                            Optional ByVal widthRatio As Double = 0.45, Optional ByVal transparency As Double = 0.1)
    If rowTop <= 0 Or rowBottom <= 0 Then Exit Sub
    Dim lo As Long: lo = Application.WorksheetFunction.Min(rowTop, rowBottom)
    Dim hi As Long: hi = Application.WorksheetFunction.Max(rowTop, rowBottom)

    Dim leftPos As Double: leftPos = ws.Columns(col).Left
    Dim widthPos As Double: widthPos = ws.Columns(col).Width
    Dim xCenter As Double: xCenter = leftPos + widthPos / 2
    Dim bodyW As Double: bodyW = widthPos * widthRatio
    If bodyW < 4 Then bodyW = 4

    Dim topY As Double: topY = ws.Rows(lo).Top
    Dim botY As Double: botY = ws.Rows(hi).Top + ws.Rows(hi).Height

    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, xCenter - bodyW / 2, topY, bodyW, botY - topY)
    With shp
        .Fill.ForeColor.RGB = fillRGB
        .Fill.Transparency = transparency
        .Line.ForeColor.RGB = lineRGB
        .Line.Weight = 1#
        .ZOrder msoBringToFront
    End With
    shp.Name = G_PREFIX & "ViaC" & col & "_" & lo & "_" & hi
    LabelShape shp, labelText, fontRGB
End Sub

Private Function RowIndexForLayer(layerNum As Long) As Long
    RowIndexForLayer = 3 + (layerNum - 1) * 2
End Function

Private Function CalculateCorePositions(n As Long, coreCount As Long, enforceSymmetry As Boolean) As Collection
    Set CalculateCorePositions = New Collection
    If coreCount <= 0 Then Exit Function
    If enforceSymmetry And coreCount = 1 Then
        Dim midPos As Long
        midPos = Application.WorksheetFunction.RoundUp(n / 2, 0)
        CalculateCorePositions.Add midPos
    Else
        Dim stepSize As Double, pos As Double, i As Long
        stepSize = n / (coreCount + 1)
        For i = 1 To coreCount
            pos = Application.WorksheetFunction.Round(stepSize * i, 0)
            If pos > 0 And pos < n Then CalculateCorePositions.Add CLng(pos)
        Next i
    End If
End Function

Private Function ParseManualCorePositions(n As Long, coreSpec As String, coreCount As Long) As Collection
    Dim result As New Collection
    Dim s As String: s = Trim$(coreSpec)
    If Len(s) = 0 Then Set ParseManualCorePositions = result: Exit Function
    s = Replace(s, "，", ",")
    s = Replace(UCase$(s), ";", ",")
    s = Replace(UCase$(s), " ", "")
    Dim parts() As String: parts = Split(s, ",")
    Dim i As Long
    For i = LBound(parts) To UBound(parts)
        If InStr(parts(i), "-") > 0 Then
            Dim ab() As String: ab = Split(parts(i), "-")
            If UBound(ab) = 1 Then
                Dim aL As Long, bL As Long
                aL = CLng(val(Replace(ab(0), "L", "")))
                bL = CLng(val(Replace(ab(1), "L", "")))
                aL = Application.WorksheetFunction.Max(1, Application.WorksheetFunction.Min(n, aL))
                bL = Application.WorksheetFunction.Max(1, Application.WorksheetFunction.Min(n, bL))
                Dim pos As Long: pos = Application.WorksheetFunction.Min(aL, bL)
                If pos > 0 And pos < n Then result.Add pos
            End If
        End If
    Next i

    If result.count = 0 And UBound(parts) = 0 And InStr(parts(0), "-") > 0 Then
        Dim ab2() As String: ab2 = Split(parts(0), "-")
        If UBound(ab2) = 1 Then
            Dim startL As Long, endL As Long
            startL = CLng(val(Replace(ab2(0), "L", "")))
            endL = CLng(val(Replace(ab2(1), "L", "")))
            startL = Application.WorksheetFunction.Max(1, Application.WorksheetFunction.Min(n, startL))
            endL = Application.WorksheetFunction.Max(1, Application.WorksheetFunction.Min(n, endL))
            Dim p As Long
            For p = startL To endL - 1
                If p > 0 And p < n Then result.Add p
            Next p
            Do While result.count > coreCount
                result.Remove result.count
            Loop
        End If
    End If
    Set ParseManualCorePositions = result
End Function

Private Function IsInCollection(col As Collection, val As Long) As Boolean
    Dim item As Variant
    For Each item In col
        If CLng(item) = val Then IsInCollection = True: Exit Function
    Next item
    IsInCollection = False
End Function

' ========= 混压范围 =========
Private Sub CollectMixedRanges(n As Long, ByRef mixRanges As Collection)
    Dim prompt As String
    prompt = "请输入材料范围（铜层范围），格式：起层-止层=材料表" & vbCrLf & _
             "示例：1-4=S-1150G；支持多段，完成后点取消。" & vbCrLf & _
             "提示：介质层位于 Li 与 L(i+1) 之间，最大止层为 " & (n - 1)
    Do
        Dim s As String
        s = InputBox(prompt, "混压材料范围设置")
        If Len(s) = 0 Then Exit Do
        Dim startL As Long, endL As Long, mat As String
        If ParseRangeSpec(s, startL, endL, mat) Then
            startL = Application.WorksheetFunction.Max(1, startL)
            endL = Application.WorksheetFunction.Min(n - 1, endL)
            If startL <= endL Then
                AddRange mixRanges, startL, endL, mat
            Else
                MsgBox "非法范围：起层大于止层。", vbExclamation
            End If
        Else
            MsgBox "格式错误。请按“起层-止层=材料表”输入，例如：1-4=S-1150G", vbExclamation
        End If
    Loop
End Sub

Private Function ParseRangeSpec(ByVal s As String, ByRef startL As Long, ByRef endL As Long, ByRef mat As String) As Boolean
    On Error GoTo Fail
    Dim parts() As String: parts = Split(Replace(UCase$(s), "/", "-"), "=")
    If UBound(parts) <> 1 Then GoTo Fail
    Dim rng As String: rng = Trim$(Replace(parts(0), "L", ""))
    mat = Trim$(parts(1))
    Dim ab() As String: ab = Split(rng, "-")
    If UBound(ab) <> 1 Then GoTo Fail
    startL = CLng(val(Trim$(ab(0))))
    endL = CLng(val(Trim$(ab(1))))
    If Len(mat) = 0 Then GoTo Fail
    ParseRangeSpec = True
    Exit Function
Fail:
    ParseRangeSpec = False
End Function

Private Sub AddRange(ByRef mixRanges As Collection, ByVal startL As Long, ByVal endL As Long, ByVal mat As String)
    mixRanges.Add Array(startL, endL, mat)
End Sub

Private Function GetMaterialByRanges(i As Long, mixRanges As Collection, defaultMat As String) As String
    Dim item As Variant
    For Each item In mixRanges
        If i >= CLng(item(0)) And i <= CLng(item(1)) Then
            GetMaterialByRanges = CStr(item(2))
            Exit Function
        End If
    Next item
    GetMaterialByRanges = defaultMat
End Function

' ========= 材料表保障/表头/样例/Key =========
Public Sub EnsureMaterialSheetsSafe(wb As Workbook)
    On Error GoTo EH
    Dim materialSheets As Variant
    materialSheets = Array("EMC-390", "S-1150G", "S-1150", "S1170G", "S1150GH", "SDI06K", "TU-862HF", "SDI03K")

    Dim nm As Variant, ws As Worksheet
    For Each nm In materialSheets
        Set ws = TryGetSheet(wb, CStr(nm))
        If ws Is Nothing Then
            Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
            ws.Name = CStr(nm)
            InitMaterialHeader ws
            WriteSampleRows ws, CStr(nm)
            NormalizeMaterialHeader ws
            FillKeyIfMissing ws
            ' 清除材料表形状（保险）
            CleanAllShapes ws
        Else
            If WorksheetIsEmpty(ws) Then
                InitMaterialHeader ws
                WriteSampleRows ws, CStr(nm)
            End If
            NormalizeMaterialHeader ws
            FillKeyIfMissing ws
            CleanAllShapes ws
        End If
    Next nm
    Exit Sub
EH:
    MsgBox "材料表保障失败: " & Err.Description, vbCritical
End Sub

Private Sub InitMaterialHeader(ws As Worksheet)
    Dim headers As Variant
    headers = Array("Type", "Code", "Resin(% )", "ply", "Nominal(mm)", "DK", "DF", "Key") ' DF新增，Key右移
    Dim i As Long
    For i = 0 To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i
    With ws
        .Rows(1).Font.Bold = True
        .Range("A1:H1").Borders.LineStyle = xlContinuous
    End With
End Sub

Private Sub WriteSampleRows(ws As Worksheet, sheetName As String)
    Select Case sheetName
        Case "EMC-390"
            AppendMaterialRow ws, "PP", "1080", 72, 1, 0.062, 4.2, 0.018
            AppendMaterialRow ws, "PP", "2116", 68, 1, 0.112, 4#, 0.02
            AppendMaterialRow ws, "CORE", "2116", 68, 2, 0.224, 4#, 0.02
        Case "S-1150G"
            AppendMaterialRow ws, "PP", "1067", 75, 1, 0.051, 4.3, 0.016
            AppendMaterialRow ws, "PP", "1080", 72, 1, 0.062, 4.2, 0.018
            AppendMaterialRow ws, "CORE", "2116", 68, 2, 0.224, 4#, 0.02
        Case "S-1150"
            AppendMaterialRow ws, "PP", "106", 80, 1, 0.038, 4.5, 0.02
            AppendMaterialRow ws, "PP", "1080", 72, 1, 0.062, 4.2, 0.018
            AppendMaterialRow ws, "CORE", "2116", 68, 1, 0.112, 4#, 0.02
        Case "S1170G", "S1150GH", "SDI06K", "TU-862HF", "SDI03K"
            ' 你的工作簿中已有完整数据，这里仅规范表头与Key（DF列请补齐）
    End Select
End Sub

Private Sub AppendMaterialRow(ws As Worksheet, mType As String, mCode As String, _
    resin As Double, ply As Long, nominalMm As Double, dk As Double, df As Double) ' ★新增 df
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lastRow < 2 Then lastRow = 1
    Dim dataRow As Long: dataRow = lastRow + 1
    ws.Cells(dataRow, 1).Value = mType
    ws.Cells(dataRow, 2).Value = mCode
    ws.Cells(dataRow, 3).Value = resin
    ws.Cells(dataRow, 4).Value = ply
    ws.Cells(dataRow, 5).Value = nominalMm
    ws.Cells(dataRow, 6).Value = dk
    ws.Cells(dataRow, 7).Value = df ' DF
End Sub

Private Sub NormalizeMaterialHeader(ws As Worksheet)
    Dim hdr As String
    hdr = CStr(ws.Cells(1, 5).Value)
    If Len(hdr) = 0 Then
        ws.Cells(1, 5).Value = "Nominal(mm)"
    ElseIf InStr(1, hdr, "(um)", vbTextCompare) > 0 Then
        ws.Cells(1, 5).Value = Replace(hdr, "(um)", "(mm)")
    ElseIf InStr(1, hdr, "(mm)", vbTextCompare) = 0 Then
        ws.Cells(1, 5).Value = "Nominal(mm)"
    End If
End Sub

Private Sub ConvertNominalUmToMmIfNeeded(ws As Worksheet)
    ' 若需从um转mm可在此实现；当前材料表多为mm
End Sub

Private Sub FillKeyIfMissing(ws As Worksheet)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    Dim i As Long, keyCol As Long: keyCol = 8 ' Key列
    For i = 2 To lastRow
        If Len(Trim$(CStr(ws.Cells(i, keyCol).Value))) = 0 Then
            ws.Cells(i, keyCol).Formula = "=A" & i & "&B" & i & "&C" & i & "&D" & i
        End If
    Next i
End Sub

Private Sub SetupDataValidation(ws As Worksheet, wsIdx As Worksheet)
    Dim lastIdx As Long: lastIdx = wsIdx.Cells(wsIdx.Rows.count, 1).End(xlUp).Row
    Dim lastOptRow As Long: lastOptRow = ws.Cells(ws.Rows.count, COL_ID).End(xlUp).Row
    If lastIdx > 1 Then
        Dim dvRange As Range
        Set dvRange = ws.Range(ws.Cells(3, COL_MAT_SHEET), ws.Cells(lastOptRow, COL_MAT_SHEET))
        With dvRange.Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:="=MaterialsIndex!$A$2:$A$" & lastIdx
            .IgnoreBlank = True
            .InCellDropdown = False      ' ← 关闭下拉箭头（仅保留验证）
            .InputTitle = "选择材料表"
            .InputMessage = "请从下拉列表中选择材料表"
            .ErrorMessage = "只能选择列表中的材料表"
        End With
    End If
End Sub

Private Sub ApplyStyling(ws As Worksheet, enableFreeze As Boolean)
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, COL_ID).End(xlUp).Row
    ws.Cells.ClearFormats

    ' A、B 两列孔型图形列表头（灰色底）
    With ws.Range("A1:B1")
        .RowHeight = 22
        .Font.Bold = True
        .Interior.Color = RGB(89, 89, 89)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ' 叠构数据表头从 C 到 P（14 列）
    With ws.Range("C1:P1")
        .RowHeight = 22
        .Font.Bold = True
        .Interior.Color = RGB(0, 112, 192)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    ws.Range("A:P").HorizontalAlignment = xlCenter
    ws.Range("A:P").VerticalAlignment = xlCenter

    Dim r As Long
    For r = 2 To lastRow
        Select Case True
            Case InStr(CStr(ws.Cells(r, COL_ID).Value), "S/M") > 0
                ws.Range("A" & r & ":P" & r).Interior.Color = RGB(183, 225, 205)
            Case CStr(ws.Cells(r, COL_ID).Value) = "PP"
                ws.Range("A" & r & ":P" & r).Interior.Color = RGB(204, 229, 255)
            Case CStr(ws.Cells(r, COL_ID).Value) = "CORE"
                ws.Range("A" & r & ":P" & r).Interior.Color = RGB(255, 242, 204) ' CORE黄色
            Case Left$(CStr(ws.Cells(r, COL_ID).Value), 1) = "L"
                ws.Range("A" & r & ":P" & r).Interior.Color = RGB(225, 245, 254)
            Case CStr(ws.Cells(r, COL_ID).Value) = "Totals"
                ws.Range("A" & r & ":P" & r).Font.Bold = True
                ws.Range("A" & r & ":P" & r).Interior.Color = RGB(242, 242, 242)
        End Select
    Next r

    With ws.Range("A1:P" & lastRow).Borders
        .LineStyle = xlContinuous
        .Color = RGB(128, 128, 128)
        .Weight = xlThin
    End With
    With ws.Range("A1:P1").Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Color = RGB(0, 0, 0)
        .Weight = xlMedium
    End With

    If ws.AutoFilterMode Then ws.AutoFilterMode = False

    ' 先清除旧的冻结/拆分，再按配置重新冻结（C2 冻结表头与 A/B 孔型列）
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ActiveWindow.Split = False
    ActiveWindow.SplitRow = 0
    ActiveWindow.SplitColumn = 0
    On Error GoTo 0

    If enableFreeze Then
        ws.Activate
        ws.Range("C2").Select
        ActiveWindow.FreezePanes = True
    Else
        ws.Range("A1").Select
    End If

    ' ---- 强制清除删除线（整表 & 关键列/行）----
    ws.Range("A:P").Font.Strikethrough = False          ' 整表
    ws.Columns(COL_CODE).Font.Strikethrough = False      ' Code列
    For r = 2 To lastRow
        If InStr(1, CStr(ws.Cells(r, COL_ID).Value), "S/M", vbTextCompare) > 0 Then
            ws.Range("A" & r & ":P" & r).Font.Strikethrough = False
        End If
    Next r
End Sub

' 调整兜底最小列宽，并提升 B/C 列下限，便于 Totals 标题完整显示
Private Sub EnsureMinimumColumnWidths(ws As Worksheet)
    If ws.Columns(COL_VIA).ColumnWidth < 16 Then ws.Columns(COL_VIA).ColumnWidth = 16
    If ws.Columns(COL_VIA_THROUGH).ColumnWidth < 16 Then ws.Columns(COL_VIA_THROUGH).ColumnWidth = 16
    If ws.Columns(COL_PARAM).ColumnWidth < 14 Then ws.Columns(COL_PARAM).ColumnWidth = 14
    If ws.Columns(COL_VENDOR).ColumnWidth < 12 Then ws.Columns(COL_VENDOR).ColumnWidth = 12
    If ws.Columns(COL_DK).ColumnWidth < 10 Then ws.Columns(COL_DK).ColumnWidth = 10
    If ws.Columns(COL_DF).ColumnWidth < 10 Then ws.Columns(COL_DF).ColumnWidth = 10
    If ws.Columns(COL_ACTUAL).ColumnWidth < 12 Then ws.Columns(COL_ACTUAL).ColumnWidth = 12
    If ws.Columns(COL_MAT_SHEET).ColumnWidth < 18 Then ws.Columns(COL_MAT_SHEET).ColumnWidth = 18
    If ws.Columns(COL_CODE).ColumnWidth < 12 Then ws.Columns(COL_CODE).ColumnWidth = 12
    If ws.Columns(COL_TYPE).ColumnWidth < 20 Then ws.Columns(COL_TYPE).ColumnWidth = 20 ' B: Total nominal
    If ws.Columns(COL_COPPER).ColumnWidth < 20 Then ws.Columns(COL_COPPER).ColumnWidth = 20 ' C: Total actual
    If ws.Columns(COL_CUST_SPEC).ColumnWidth < 12 Then ws.Columns(COL_CUST_SPEC).ColumnWidth = 12
    If ws.Columns(COL_PLANT_SPEC).ColumnWidth < 12 Then ws.Columns(COL_PLANT_SPEC).ColumnWidth = 12
End Sub

Private Sub OptimizeLayout()
    Dim wb As Workbook: Set wb = ThisWorkbook
    On Error Resume Next
    wb.Worksheets("Config").Move Before:=wb.Worksheets(1)
    wb.Worksheets("MaterialsIndex").Move Before:=wb.Worksheets(2)
    wb.Worksheets("Stack-up(Optimized)").Move Before:=wb.Worksheets(3)
    On Error GoTo 0
End Sub

' ============================================================
' Section 6: 层对选择 → 堆叠（自动映射 & 对称镜像）
' ============================================================
Private Sub ApplyInterlayerSelections(wsCfg As Worksheet, wsOpt As Worksheet, totalLayers As Long)
    Dim anchor As Range: Set anchor = wsCfg.Range("AD1")
    Dim pairCount As Long: pairCount = Application.WorksheetFunction.Max(0, totalLayers - 1)
    Dim i As Long, optRow As Long
    Dim codeVal As String, rcVal As Variant, plyVal As Variant
    Dim hasPly As Boolean
    hasPly = HasHeaderAt(wsCfg, "AG1", "ply")

    Dim enforceSymmetry As Boolean
    enforceSymmetry = (UCase$(GetCfgValue(wsCfg, "是否对称", "FALSE")) = "TRUE")

    ' 对称时只处理上半部分（源），下半部分由镜像公式回填，避免源/镜像互相覆盖产生循环引用
    Dim iMax As Long
    If enforceSymmetry Then
        iMax = Application.WorksheetFunction.RoundUp(pairCount / 2, 0)
    Else
        iMax = pairCount
    End If

    For i = 1 To iMax
        codeVal = Trim$(CStr(anchor.Offset(i, 1).Value))          ' AE: PP型号
        rcVal = CleanRCValue(anchor.Offset(i, 2).Value)            ' AF: RC(%)
        If hasPly Then plyVal = CleanPlyValue(anchor.Offset(i, 3).Value)

        If Len(codeVal) > 0 And Not IsEmpty(rcVal) And Len(Trim$(CStr(rcVal))) > 0 Then
            optRow = RowIndexForLayer(i) + 1
            If UCase$(CStr(wsOpt.Cells(optRow, COL_ID).Value)) = "PP" Then
                wsOpt.Cells(optRow, COL_CODE).Value = codeVal
                wsOpt.Cells(optRow, COL_PARAM).Value = rcVal
                If hasPly And Not IsEmpty(plyVal) Then
                    wsOpt.Cells(optRow, COL_PLY).Value = plyVal
                End If
                ' 重挂 UDF（含 DF）
                wsOpt.Cells(optRow, COL_VENDOR).FormulaR1C1 = "=MatVendor(RC" & COL_MAT_SHEET & ",RC" & COL_ID & ",RC" & COL_CODE & ",RC" & COL_PARAM & ",RC" & COL_PLY & ")"
                wsOpt.Cells(optRow, COL_DK).FormulaR1C1 = "=MatDK(RC" & COL_MAT_SHEET & ",RC" & COL_ID & ",RC" & COL_CODE & ",RC" & COL_PARAM & ",RC" & COL_PLY & ")"
                wsOpt.Cells(optRow, COL_DF).FormulaR1C1 = "=MatDF(RC" & COL_MAT_SHEET & ",RC" & COL_ID & ",RC" & COL_CODE & ",RC" & COL_PARAM & ",RC" & COL_PLY & ")"
                wsOpt.Cells(optRow, COL_CUST_SPEC).FormulaR1C1 = _
                    "=IF(RC" & COL_ID & "=""PP"",IF(ISNUMBER(RC" & COL_VENDOR & "),ROUND(RC" & COL_VENDOR & "*1000,0),""""),"""")"
            End If
        ElseIf IsPurePress Then
            ' 纯压 & 层对表未填：按材料1（材料表名）默认抓取
            optRow = RowIndexForLayer(i) + 1
            If UCase$(CStr(wsOpt.Cells(optRow, COL_ID).Value)) = "PP" Then
                Dim matNm As String: matNm = Trim$(CStr(wsOpt.Cells(optRow, COL_MAT_SHEET).Value))
                Dim dCode As String, dRC As Variant
                dCode = GetFirstCode(matNm, "PP")
                dRC = GetResinValue(matNm, "PP", dCode)
                wsOpt.Cells(optRow, COL_CODE).Value = dCode
                wsOpt.Cells(optRow, COL_PARAM).Value = dRC
                wsOpt.Cells(optRow, COL_VENDOR).FormulaR1C1 = "=MatVendor(RC" & COL_MAT_SHEET & ",RC" & COL_ID & ",RC" & COL_CODE & ",RC" & COL_PARAM & ",RC" & COL_PLY & ")"
                wsOpt.Cells(optRow, COL_DK).FormulaR1C1 = "=MatDK(RC" & COL_MAT_SHEET & ",RC" & COL_ID & ",RC" & COL_CODE & ",RC" & COL_PARAM & ",RC" & COL_PLY & ")"
                wsOpt.Cells(optRow, COL_DF).FormulaR1C1 = "=MatDF(RC" & COL_MAT_SHEET & ",RC" & COL_ID & ",RC" & COL_CODE & ",RC" & COL_PARAM & ",RC" & COL_PLY & ")"
                wsOpt.Cells(optRow, COL_CUST_SPEC).FormulaR1C1 = _
                    "=IF(RC" & COL_ID & "=""PP"",IF(ISNUMBER(RC" & COL_VENDOR & "),ROUND(RC" & COL_VENDOR & "*1000,0),""""),"""")"
            End If
        End If

        ' 对称镜像：把上半部 optRow 的 PP 信息镜像到下半部（公式链接）
        If enforceSymmetry Then
            Dim mirror As Long: mirror = pairCount - i + 1
            Dim mRow As Long: mRow = RowIndexForLayer(mirror) + 1
            If mRow <> optRow Then
                If UCase$(CStr(wsOpt.Cells(mRow, COL_ID).Value)) = "PP" Then
                    wsOpt.Cells(mRow, COL_MAT_SHEET).FormulaR1C1 = "=R" & optRow & "C" & COL_MAT_SHEET
                    wsOpt.Cells(mRow, COL_CODE).FormulaR1C1 = "=R" & optRow & "C" & COL_CODE
                    wsOpt.Cells(mRow, COL_PARAM).FormulaR1C1 = "=R" & optRow & "C" & COL_PARAM
                    wsOpt.Cells(mRow, COL_PLY).FormulaR1C1 = "=R" & optRow & "C" & COL_PLY
                    wsOpt.Cells(mRow, COL_VENDOR).FormulaR1C1 = "=MatVendor(RC" & COL_MAT_SHEET & ",RC" & COL_ID & ",RC" & COL_CODE & ",RC" & COL_PARAM & ",RC" & COL_PLY & ")"
                    wsOpt.Cells(mRow, COL_DK).FormulaR1C1 = "=MatDK(RC" & COL_MAT_SHEET & ",RC" & COL_ID & ",RC" & COL_CODE & ",RC" & COL_PARAM & ",RC" & COL_PLY & ")"
                    wsOpt.Cells(mRow, COL_DF).FormulaR1C1 = "=MatDF(RC" & COL_MAT_SHEET & ",RC" & COL_ID & ",RC" & COL_CODE & ",RC" & COL_PARAM & ",RC" & COL_PLY & ")"
                    wsOpt.Cells(mRow, COL_CUST_SPEC).FormulaR1C1 = _
                        "=IF(RC" & COL_ID & "=""PP"",IF(ISNUMBER(RC" & COL_VENDOR & "),ROUND(RC" & COL_VENDOR & "*1000,0),""""),"""")"
                End If
            End If
        End If
    Next i

    ' 统一重挂 UDF 与 Actual（含PP残铜影响）并完全重算
    ReapplyUDFForAllDielectrics wsOpt
    Application.CalculateFullRebuild
End Sub

Private Function HasHeaderAt(ws As Worksheet, addr As String, expectText As String) As Boolean
    On Error GoTo Fail
    HasHeaderAt = (UCase$(Trim$(CStr(ws.Range(addr).Value))) = UCase$(Trim$(expectText)))
    Exit Function
Fail:
    HasHeaderAt = False
End Function

Private Function CleanRCValue(v As Variant) As Variant
    Dim s As String: s = Trim$(CStr(v))
    If Len(s) = 0 Then
        CleanRCValue = ""
        Exit Function
    End If
    s = Replace(s, "%", "")
    s = Replace(s, "％", "")
    s = Trim$(s)
    If IsNumeric(s) Then
        CleanRCValue = CDbl(s)
    Else
        CleanRCValue = v
    End If
End Function

Private Function CleanPlyValue(v As Variant) As Variant
    Dim s As String: s = Trim$(CStr(v))
    If Len(s) = 0 Then
        CleanPlyValue = Empty
        Exit Function
    End If
    If IsNumeric(s) Then
        CleanPlyValue = CLng(val(s))
    Else
        CleanPlyValue = v
    End If
End Function

Private Sub ReapplyUDFForAllDielectrics(wsOpt As Worksheet)
    Dim lastRow As Long
    lastRow = wsOpt.Cells(wsOpt.Rows.count, COL_ID).End(xlUp).Row

    Dim r As Long, idText As String
    For r = 2 To lastRow
        idText = UCase$(Trim$(CStr(wsOpt.Cells(r, COL_ID).Value)))
        If idText = "PP" Or idText = "CORE" Then
            wsOpt.Cells(r, COL_VENDOR).FormulaR1C1 = _
                "=MatVendor(RC" & COL_MAT_SHEET & ",RC" & COL_ID & ",RC" & COL_CODE & ",RC" & COL_PARAM & ",RC" & COL_PLY & ")"
            wsOpt.Cells(r, COL_DK).FormulaR1C1 = _
                "=MatDK(RC" & COL_MAT_SHEET & ",RC" & COL_ID & ",RC" & COL_CODE & ",RC" & COL_PARAM & ",RC" & COL_PLY & ")"
            wsOpt.Cells(r, COL_DF).FormulaR1C1 = _
                "=MatDF(RC" & COL_MAT_SHEET & ",RC" & COL_ID & ",RC" & COL_CODE & ",RC" & COL_PARAM & ",RC" & COL_PLY & ")"
            wsOpt.Cells(r, COL_CUST_SPEC).FormulaR1C1 = _
                "=IF(RC" & COL_ID & "=""PP"",IF(ISNUMBER(RC" & COL_VENDOR & "),ROUND(RC" & COL_VENDOR & "*1000,0),""""),"""")"
            ' Actual(mm)：厚度 = Vendor（MatVendor 已含 ×ply），PP/CORE 一致
            wsOpt.Cells(r, COL_ACTUAL).FormulaR1C1 = _
                "=IF(ISNUMBER(RC" & COL_VENDOR & "),RC" & COL_VENDOR & ","""")"
        End If
    Next r
End Sub

' ========= Totals 标题适配（新增） =========
Private Function FindTotalsRow(ws As Worksheet) As Long
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, COL_ID).End(xlUp).Row
    Dim rr As Long
    For rr = 2 To lastRow
        If Trim$(CStr(ws.Cells(rr, COL_ID).Value)) = "Totals" Then
            FindTotalsRow = rr
            Exit Function
        End If
    Next rr
End Function

Private Sub EnsureTotalsTitlesFit(ws As Worksheet)
    Dim tr As Long: tr = FindTotalsRow(ws)
    If tr > 0 Then
        ws.Cells(tr, COL_TYPE).WrapText = True
        ws.Cells(tr, COL_COPPER).WrapText = True
        ws.Columns(COL_TYPE).AutoFit
        ws.Columns(COL_COPPER).AutoFit
        If ws.Columns(COL_TYPE).ColumnWidth < 20 Then ws.Columns(COL_TYPE).ColumnWidth = 20
        If ws.Columns(COL_COPPER).ColumnWidth < 20 Then ws.Columns(COL_COPPER).ColumnWidth = 20
        ws.Rows(tr).AutoFit
    End If
End Sub

' ============================================================
' Section 7: 一键清理（可选）
' ============================================================
Public Sub QuickClean_Shapes_And_Strikethrough()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet

    ' 1) Stack-up(Optimized)：清形状 + 清删除线
    On Error Resume Next
    Set ws = wb.Worksheets("Stack-up(Optimized)")
    On Error GoTo 0
    If Not ws Is Nothing Then
        CleanAllShapes ws
        Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, COL_ID).End(xlUp).Row
        ws.Range("A:P").Font.Strikethrough = False
        ws.Columns(COL_CODE).Font.Strikethrough = False  ' Code 列
        Dim r As Long
        For r = 2 To lastRow
            If InStr(1, CStr(ws.Cells(r, COL_ID).Value), "S/M", vbTextCompare) > 0 Then
                ws.Range("A" & r & ":P" & r).Font.Strikethrough = False
         End If
        Next r
    End If

    ' 2) Config：清形状 + 清删除线
    On Error Resume Next
    Set ws = wb.Worksheets("Config")
    On Error GoTo 0
    If Not ws Is Nothing Then
        CleanAllShapes ws
        ws.Range("A:Z").Font.Strikethrough = False
    End If

    ' 3) 材料表：可选清理
    Dim mats As Variant: mats = Array("EMC-390", "S-1150G", "S-1150", "S1170G", "S1150GH", "SDI06K", "TU-862HF", "SDI03K")
    Dim nm As Variant
    For Each nm In mats
        On Error Resume Next
        Set ws = wb.Worksheets(CStr(nm))
        On Error GoTo 0
        If Not ws Is Nothing Then
            CleanAllShapes ws
            ws.Cells.Font.Strikethrough = False
        End If
    Next nm
End Sub


