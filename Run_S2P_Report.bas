Option Explicit

Private TargetWB As Workbook
Private RunMode As String
Private PlatformKind As String
Private RefDict As Object
Private CoreNames() As String
Private CoreRowStart() As Long

Private Const N_POINTS As Long = 801

' ========================= 진입점 =========================
Public Sub Run_S2P_Report()
    Dim baseFolder As String
    Dim idx As Long

    Application.ScreenUpdating = False              ' 화면갱신 정지
    Application.Calculation = xlCalculationManual   ' 수동= 자동 계산 중지
    Application.EnableEvents = False                ' 이벤트 비활성화

    Set TargetWB = ActiveWorkbook                   ' Targetwb=현재 활성 워크북

    PlatformKind = SelectPlatformKind()          ' 1) TOWER / BLUESTAR
    If PlatformKind = "" Then GoTo CLEANUP       ' 공백일 시 종료

    RunMode = SelectRunMode()                    ' 2) SRU / AUX
    If RunMode = "" Then GoTo CLEANUP            ' 공백일 시 종료

    InitCoreAndRows                              ' 코어/행(셀) 매핑 초기화
    Set RefDict = CreateObject("Scripting.Dictionary")

    baseFolder = PickFolder()
    If Len(baseFolder) = 0 Then GoTo CLEANUP
    If Right$(baseFolder, 1) <> "\" And Right$(baseFolder, 1) <> "/" Then baseFolder = baseFolder & "\"

    ' 1) 참조 세트: SRU01 / AUX01
    If Not LoadReferenceSet(baseFolder, 1) Then
        MsgBox IIf(RunMode = "SRU", "SRU01", "AUX01") & " 레퍼런스 로딩 중 치명적 오류가 발생했습니다. LOG 시트를 확인하세요.", vbExclamation
        GoTo CLEANUP
    End If

    ' 2) 타깃 세트: 02 ~ 16
    For idx = 2 To 16
        If Not ProcessTargetSet(baseFolder, idx) Then
            ' 오류가 있어도 진행
        End If
    Next idx

    MsgBox "완료! " & IIf(RunMode = "SRU", "SRU02~SRU16", "AUX02~AUX16") & "에 결과를 기입했습니다.", vbInformation ' 메시지박스

CLEANUP:            ' 종료
    Application.EnableEvents = True     ' 이벤트 활성화
    Application.Calculation = xlCalculationAutomatic    ' 자동계산 활성화
    Application.ScreenUpdating = True   ' 화면갱신 활성화
End Sub

' ========================= 선택 UI =========================
Private Function SelectPlatformKind() As String
    Dim ans As String
    ans = Application.InputBox(Prompt:="플랫폼을 입력하세요: TOWER 또는 BLUESTAR", Title:="SELECT (TOWER, BLUESTAR)", Type:=2)      ' 메시지창 띄우고 ans에 값 저장
    If ans = "False" Then Exit Function         ' ans에 값이 없을 시 종료
    ans = UCase$(Trim$(ans))    'Trim$ : 공백 제거 , UCase$ : 대문자화
    If ans <> "TOWER" And ans <> "BLUESTAR" Then    ' ans가 Tower또는 Bluestar가 아닐 경우 실행
        MsgBox "TOWER 또는 BLUESTAR만 입력 가능합니다.", vbExclamation  ' 메시지 박스 이후 종료
        Exit Function
    End If
    SelectPlatformKind = ans
End Function

Private Function SelectRunMode() As String
    Dim ans As String
    ans = Application.InputBox(Prompt:="모드를 입력하세요: SRU 또는 AUX", Title:="SELECT (SRU, AUX)", Type:=2)
    If ans = "False" Then Exit Function
    ans = UCase$(Trim$(ans))
    If ans <> "SRU" And ans <> "AUX" Then
        MsgBox "SRU 또는 AUX만 입력 가능합니다.", vbExclamation
        Exit Function
    End If
    SelectRunMode = ans
End Function

' ========================= 참조/타깃 처리 =========================
Private Function LoadReferenceSet(ByVal baseFolder As String, ByVal idx As Long) As Boolean
    Dim k As Long, filePath As String
    Dim rawVals() As Double, unwrapped() As Double
    Dim ok As Boolean, key As String
    Dim modes As Variant, m As Variant

    On Error GoTo ERRH      ' Runtime 오류 시 ERRH로 이동
    LoadReferenceSet = True

    For k = LBound(CoreNames) To UBound(CoreNames)      ' k변수로 반복문 실행.
        modes = ModesForCore(CoreNames(k))
        For Each m In modes
            filePath = BuildFilePath(baseFolder, idx, CoreNames(k), CStr(m))
            rawVals = ReadS2PPhaseColumn(filePath, ok)
            If Not ok Then
                LogIssue "REF MISSING/FORMAT", filePath
                LoadReferenceSet = False
            Else
                unwrapped = UnwrapPhase(rawVals)
                key = ModeKey(CStr(m), CoreNames(k))
                RefDict(key) = unwrapped
            End If
        Next m
    Next k
    Exit Function
ERRH:
    LoadReferenceSet = False
    LogIssue "REF ERROR", (IIf(RunMode = "SRU", "SRU01", "AUX01")) & " loading failed: " & Err.Description
End Function

Private Function ProcessTargetSet(ByVal baseFolder As String, ByVal idx As Long) As Boolean
    Dim ws As Worksheet, shName As String
    Dim k As Long, colIdx As Long, r0 As Long
    Dim rawVals() As Double, unwrapped() As Double
    Dim ok As Boolean, key As String, refVals() As Double
    Dim pct4 As Double, pct3 As Double, pctL As Double, maxStep50 As Double
    Dim modes As Variant, m As Variant

    On Error GoTo ERRH
    ProcessTargetSet = True

    shName = IIf(RunMode = "SRU", "SRU", "AUX") & Format$(idx, "00")
    Set ws = EnsureSheet(shName)

    ' SRU 모드일 때 이전 시트명이 SBFxx라면 자동 개명
    If RunMode = "SRU" Then TryRenameSBFtoSRU idx

    For k = LBound(CoreNames) To UBound(CoreNames)
        modes = ModesForCore(CoreNames(k))
        For Each m In modes
            colIdx = ColumnForMode(CStr(m))                 ' NORMAL=F(6), LOW=G(7), BYPASS=G(7)
            rawVals = ReadS2PPhaseColumn(BuildFilePath(baseFolder, idx, CoreNames(k), CStr(m)), ok)
            If Not ok Then
                LogIssue "TARGET MISSING/FORMAT", BuildFilePath(baseFolder, idx, CoreNames(k), CStr(m))
                GoTo NextMode
            End If

            unwrapped = UnwrapPhase(rawVals)
            key = ModeKey(CStr(m), CoreNames(k))
            If Not RefDict.Exists(key) Then
                LogIssue "NO REF", "Ref missing for key: " & key
                GoTo NextMode
            End If
            refVals = RefDict(key)

            ComputeMetrics CoreNames(k), refVals, unwrapped, pct4, pct3, pctL, maxStep50

            ' ▼ ABS + Round(2) 적용
            pct4 = Abs2(pct4)
            pct3 = Abs2(pct3)
            pctL = Abs2(pctL)
            maxStep50 = Abs2(maxStep50)

            r0 = CoreRowStart(k)
            ws.Cells(r0 - 1, colIdx).Value = maxStep50
            ws.Cells(r0 - 1, colIdx).NumberFormat = "0.00"

            ws.Cells(r0 + 0, colIdx).Value = pct4
            ws.Cells(r0 + 1, colIdx).Value = pct3
            ws.Cells(r0 + 2, colIdx).Value = pctL
            ws.Range(ws.Cells(r0, colIdx), ws.Cells(r0 + 2, colIdx)).NumberFormat = "0.00"

NextMode:
        Next m
    Next k

    Exit Function
ERRH:
    ProcessTargetSet = False
    LogIssue "SET ERROR", (IIf(RunMode = "SRU", "SRU", "AUX")) & Format$(idx, "00") & ": " & Err.Description
End Function

' ========================= 핵심 계산 =========================
' ref/tgt(801) + coreName ⇒ 퍼센트 3개 + 50MHz Step Max
Private Sub ComputeMetrics(ByVal coreName As String, _
                           ByRef refVals() As Double, ByRef tgtVals() As Double, _
                           ByRef pct_gt4 As Double, ByRef pct_3to4 As Double, ByRef pct_lt3 As Double, _
                           ByRef max_step50 As Double)
    Dim i As Long
    Dim PM() As Double
    Dim APM_non() As Double, mask_non() As Boolean
    Dim APM_sl() As Double, mask_sl() As Boolean
    Dim PT_non As Double, PT_sl As Double, have_non As Boolean, have_sl As Boolean
    Dim MT As Double, maxAbs As Double
    Dim cnt4 As Long, cnt3 As Long, cntL As Long
    
    ReDim PM(1 To N_POINTS)
    For i = 1 To N_POINTS
        PM(i) = tgtVals(i) - refVals(i)   ' Phase Matching
    Next i
    
    ' PATH별 파라미터 취득
    Dim nonStart As Long, nonSize As Long, nonMode As String
    Dim slStart As Long, slSize As Long, slMode As String
    Dim flipSign As Boolean
    GetPathParams coreName, nonStart, nonSize, nonMode, slStart, slSize, slMode, flipSign
    
    ' 평균 배열 생성(인덱스별로 해당 평균을 넣고, 유효여부는 mask로 표시)
    BuildAPMArrays PM, nonStart, nonSize, nonMode, slStart, slSize, slMode, APM_non, mask_non, APM_sl, mask_sl
    
    ' 트래킹 및 Max 계산
    maxAbs = 0#
    For i = 1 To N_POINTS
        have_non = mask_non(i)
        have_sl = mask_sl(i)
        
        If have_non Then
            If flipSign Then
                PT_non = APM_non(i) - PM(i)
            Else
                PT_non = PM(i) - APM_non(i)
            End If
        End If
        
        If have_sl Then
            If flipSign Then
                PT_sl = APM_sl(i) - PM(i)
            Else
                PT_sl = PM(i) - APM_sl(i)
            End If
        End If
        
        ' Max Tracking 결합 규칙
        If have_non And have_sl Then
            MT = MaxD(AbsD(PT_non), AbsD(PT_sl))
        ElseIf have_non Then
            MT = AbsD(PT_non)
        ElseIf have_sl Then     ' ㅇㅇ
            MT = AbsD(PT_sl)
        Else
            ' 윈도우에 속하지 않으면 PM 절대값으로 대체(801분모 유지)
            MT = AbsD(PM(i))
        End If
        
        If MT > maxAbs Then maxAbs = MT
        
        If MT > 4# Then
            cnt4 = cnt4 + 1
        ElseIf MT >= 3# And MT <= 4# Then
            cnt3 = cnt3 + 1
        End If
    Next i
    
    cntL = N_POINTS - cnt3 - cnt4
    pct_gt4 = 100# * cnt4 / N_POINTS
    pct_3to4 = 100# * cnt3 / N_POINTS
    pct_lt3 = 100# * cntL / N_POINTS
    max_step50 = maxAbs
End Sub

' 위상 언랩
Private Function UnwrapPhase(ByRef rawVals() As Double) As Double()
    Dim i As Long, delta As Double, cur As Double
    Dim outArr() As Double
    ReDim outArr(1 To N_POINTS)
    outArr(1) = rawVals(1)
    For i = 2 To N_POINTS
        cur = rawVals(i)
        delta = cur - outArr(i - 1)
        Do While delta > 180#: cur = cur - 360#: delta = cur - outArr(i - 1): Loop
        Do While delta < -180#: cur = cur + 360#: delta = cur - outArr(i - 1): Loop
        outArr(i) = cur
    Next i
    UnwrapPhase = outArr
End Function

' ========================= 파일 I/O / 유틸 =========================
Private Function ReadS2PPhaseColumn(ByVal filePath As String, ByRef ok As Boolean) As Double()
    Dim ff As Integer, ln As String, tokens() As String
    Dim arr() As Double, count As Long
    Dim fmt As String, reV As Double, imV As Double, ang As Double, tok As String

    ok = False
    If Dir$(filePath, vbNormal) = "" Then Exit Function     ' filePath가 정상적이지 않으면 종료

    ReDim arr(1 To N_POINTS)
    ff = FreeFile   ' Open에 쓸 고유 파일번호 확보(중복 방지). 하드코딩(#1 등) 대신 항상 FreeFile 사용
    On Error GoTo ERRH
    Open filePath For Input As #ff      ' filepath열고  읽기전용모드 후 #번호 매기기
    Do While Not EOF(ff)                ' End of file에 도달할때까지 반복
        Line Input #ff, ln              ' 파일 안의 첫 줄을 읽어 ln에 저장
        ln = Trim$(ln)                  ' 양 끝 공백 제거
        If ln = "" Then GoTo ContinueLine   ' ln에 내용이 없을 시 ContinueLine으로 이동

        If Left$(ln, 1) = "#" Then          ' ln 왼쪽 첫번째가 #일 경우
            If InStr(1, " " & UCase$(ln) & " ", " RI ") > 0 Then fmt = "RI"     'ln(대문자+양끝 공백 보정) 안에 RI라는 완전한 단어 토큰이 들어있으면, 데이터 포맷을 RI로 설정
            If InStr(1, " " & UCase$(ln) & " ", " DB ") > 0 Then fmt = "DB"     'ln(대문자+양끝 공백 보정) 안에 DB라는 완전한 단어 토큰이 들어있으면, 데이터 포맷을 DB로 설정
            If InStr(1, " " & UCase$(ln) & " ", " MA ") > 0 Then fmt = "MA"     'ln(대문자+양끝 공백 보정) 안에 MA라는 완전한 단어 토큰이 들어있으면, 데이터 포맷을 MA로 설정
            GoTo ContinueLine
        End If
        If Left$(ln, 1) = "!" Then GoTo ContinueLine

        tokens = SplitOnWhitespace(ln)
        If UBound(tokens) < 4 Then GoTo ContinueLine

        If fmt = "RI" Then
            tok = Replace$(tokens(3), "D", "E"): reV = Val(tok)
            tok = Replace$(tokens(4), "D", "E"): imV = Val(tok)
            ang = Application.WorksheetFunction.Atan2(reV, imV) * 180# / Application.WorksheetFunction.Pi()
        Else
            tok = Replace$(tokens(4), "D", "E")         ' 5번째 배열 값 저장(D,E는 정규화)
            ang = Val(tok)                              ' 값으로 변환
        End If

        count = count + 1
        If count <= N_POINTS Then arr(count) = ang
        If count >= N_POINTS Then Exit Do
ContinueLine:
    Loop

    Close #ff
    If count = N_POINTS Then ok = True: ReadS2PPhaseColumn = arr
    Exit Function
ERRH:
    On Error Resume Next
    Close #ff
    LogIssue "READ ERROR", filePath & " | " & Err.Description
End Function

Private Function SplitOnWhitespace(ByVal s As String) As String()
    s = Replace$(s, vbTab, " ")                 ' s안의 문자를 공백 하나로 변경
    s = Application.WorksheetFunction.Trim(s)   ' s안의 다중 공백을 한 칸으로 지정
    SplitOnWhitespace = Split(s, " ")           ' 공백 기준 분할
End Function

' 파일 경로 조립
Private Function BuildFilePath(ByVal baseFolder As String, ByVal idx As Long, _
                               ByVal coreName As String, ByVal modeName As String) As String
    If RunMode = "SRU" Then
        BuildFilePath = baseFolder & "(" & "SRU" & Format$(idx, "00") & ") " _
                        & coreName & "_" & UCase$(modeName) & "_PHASE.s2p"
    Else
        ' AUX
        If UCase$(modeName) = "BYPASS" Then
            BuildFilePath = baseFolder & "(" & "AUX" & Format$(idx, "00") & ") " _
                            & "J37-J9_NORMAL_BYPASS_PHASE.s2p"
        Else
            BuildFilePath = baseFolder & "(" & "AUX" & Format$(idx, "00") & ") " _
                            & coreName & "_" & UCase$(modeName) & "_PHASE.s2p"
        End If
    End If
End Function

Private Function ColumnForMode(ByVal modeName As String) As Long
    Select Case UCase$(modeName)
        Case "NORMAL": ColumnForMode = 6   ' F
        Case "LOW":    ColumnForMode = 7   ' G
        Case "BYPASS": ColumnForMode = 7   ' G
        Case Else:     ColumnForMode = 6
    End Select
End Function

Private Function ModeKey(ByVal modeName As String, ByVal coreName As String) As String
    ModeKey = UCase$(modeName) & "|" & coreName
End Function

Private Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a >= b Then MaxD = a Else MaxD = b
End Function

Private Function AbsD(ByVal a As Double) As Double
    If a >= 0# Then AbsD = a Else AbsD = -a
End Function

Private Function PickFolder() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "S2P 파일 폴더를 선택하세요 (" & IIf(RunMode = "SRU", "SRU01~SRU16", "AUX01~AUX16") & ")"
        .AllowMultiSelect = False
        If .Show = -1 Then PickFolder = .SelectedItems(1)
    End With
End Function

' 시트 보장 + SRU 모드에서 SBF→SRU 자동 개명 처리
Private Function EnsureSheet(ByVal shName As String) As Worksheet
    Dim ws As Worksheet, s As Worksheet, legacy As String
    For Each ws In TargetWB.Worksheets
        If StrComp(Trim$(ws.Name), Trim$(shName), vbTextCompare) = 0 Then Set EnsureSheet = ws: Exit Function
    Next ws
    If RunMode = "SRU" Then
        legacy = "SBF" & Right$(shName, 2)
        For Each s In TargetWB.Worksheets
            If StrComp(Trim$(s.Name), legacy, vbTextCompare) = 0 Then
                On Error Resume Next
                s.Name = shName
                On Error GoTo 0
                Set EnsureSheet = s
                Exit Function
            End If
        Next s
    End If
    Set EnsureSheet = TargetWB.Worksheets.Add(After:=TargetWB.Sheets(TargetWB.Sheets.count))
    EnsureSheet.Name = shName
End Function

' LOG: 대상 통합문서에 기록
Private Sub LogIssue(ByVal kind As String, ByVal msg As String)
    Dim ws As Worksheet, r As Long, s As Worksheet
    For Each s In TargetWB.Worksheets
        If StrComp(Trim$(s.Name), "LOG", vbTextCompare) = 0 Then Set ws = s: Exit For
    Next s
    If ws Is Nothing Then
        Set ws = TargetWB.Worksheets.Add(Before:=TargetWB.Sheets(1))
        ws.Name = "LOG"
        ws.Range("A1:C1").Value = Array("When", "Type", "Message")
        ws.Columns("A:C").ColumnWidth = 48
    End If
    r = ws.Cells(ws.Rows.count, "A").End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = kind
    ws.Cells(r, 3).Value = msg
End Sub

' ========================= 초기화(코어/행) =========================
Private Sub InitCoreAndRows()
    Dim i As Long
    Dim tmpNames As Variant, tmpRows As Variant

    If RunMode = "SRU" Then
        If PlatformKind = "BLUESTAR" Then
            ' --- SRU / BLUESTAR: 20코어 (기존) ---
            tmpNames = Array( _
                "J1-J9_HF1", "J1-J9_HF2", "J1-J9_HF3", "J1-J9_HF4", "J1-J9_HF5", _
                "J1-J9_HF6", "J1-J9_HF7", "J1-J9_HF8", "J1-J9_HF9", "J2-J9_HF1", _
                "J3-J9_LF2", "J3-J9_LF6", "J3-J9_LF7", "J3-J9_LF8", "J3-J9_LF9", _
                "J4-J9_LF2", "J5-J9_LF1", "J5-J9_LF3", "J5-J9_LF4", "J5-J9_LF5" _
            )
            tmpRows = Array( _
                206, 245, 290, 336, 380, _
                425, 470, 515, 560, 608, _
                651, 695, 740, 785, 830, _
                876, 920, 965, 1010, 1055 _
            )
        Else
            ' --- SRU / TOWER: 24코어 (LF21/LF22 추가, 순서/셀 변경) ---
            ' J1 HF1~J3 LF2 까지는 동일, 이후 순서:
            ' J3: LF21, LF22, LF6, LF7, LF8, LF9
            ' J4: LF2, LF21, LF22
            ' J5: LF1, LF3, LF4, LF5
            tmpNames = Array( _
                "J1-J9_HF1", "J1-J9_HF2", "J1-J9_HF3", "J1-J9_HF4", "J1-J9_HF5", _
                "J1-J9_HF6", "J1-J9_HF7", "J1-J9_HF8", "J1-J9_HF9", "J2-J9_HF1", _
                "J3-J9_LF2", "J3-J9_LF21", "J3-J9_LF22", "J3-J9_LF6", "J3-J9_LF7", _
                "J3-J9_LF8", "J3-J9_LF9", "J4-J9_LF2", "J4-J9_LF21", "J4-J9_LF22", _
                "J5-J9_LF1", "J5-J9_LF3", "J5-J9_LF4", "J5-J9_LF5" _
            )
            ' >4.0° 행(r0): 앞 11개는 기존, J3 LF21부터는 제공 범위(690~693 등)에서 +1
            tmpRows = Array( _
                206, 245, 290, 336, 380, _
                425, 470, 515, 560, 608, _
                651, 691, 736, 785, 830, _
                875, 920, 966, 1006, 1051, _
                1100, 1145, 1190, 1235 _
            )
        End If

    Else
        ' ---------- AUX ----------
        If PlatformKind = "BLUESTAR" Then
            ' 기존 순서: 1.HF1 2.LF2 3.LF1 4.BYPASS
            tmpNames = Array("J37-J9_HF1", "J37-J9_LF2", "J37-J9_LF1", "J37-J9_NORMAL_BYPASS")
            ' 제공 CEL: 160~163 / 199~202 / 243~246 / 285~288 → >4.0° 행은 +1
            tmpRows = Array(161, 200, 244, 286)
        Else
            ' TOWER 순서: 1.HF1 2.LF2 3.LF21 4.LF22 5.LF1 6.BYPASS
            tmpNames = Array("J37-J9_HF1", "J37-J9_LF2", "J37-J9_LF21", "J37-J9_LF22", "J37-J9_LF1", "J37-J9_NORMAL_BYPASS")
            ' 제공 CEL: 160~163 / 199~202 / 241~244 / 286~289 / 333~336 / 375~378 → >4.0° 행은 +1
            tmpRows = Array(161, 200, 242, 287, 334, 376)
        End If
    End If

    ReDim CoreNames(0 To UBound(tmpNames))          ' 변수 재정의 및 형식 지정
    ReDim CoreRowStart(0 To UBound(tmpRows))
    For i = LBound(tmpNames) To UBound(tmpNames)
        CoreNames(i) = CStr(tmpNames(i))
        CoreRowStart(i) = CLng(tmpRows(i))
    Next i
End Sub

' 해당 코어에 적용될 모드 집합
Private Function ModesForCore(ByVal coreName As String) As Variant
    If RunMode = "SRU" Then
        ModesForCore = Array("NORMAL", "LOW")
    Else
        If coreName = "J37-J9_NORMAL_BYPASS" Then
            ModesForCore = Array("BYPASS")                ' BYPASS는 단독, G열
        Else
            ModesForCore = Array("NORMAL", "LOW")         ' HF1/LF1/LF2/LF21/LF22
        End If
    End If
End Function

' SRU 모드에서 SBFxx → SRUxx 개명 시도
Private Sub TryRenameSBFtoSRU(ByVal idx As Long)
    Dim legacy As String, target As String
    Dim s As Worksheet
    legacy = "SBF" & Format$(idx, "00")
    target = "SRU" & Format$(idx, "00")
    For Each s In TargetWB.Worksheets
        If StrComp(Trim$(s.Name), legacy, vbTextCompare) = 0 Then
            On Error Resume Next
            s.Name = target
            On Error GoTo 0
            Exit For
        End If
    Next s
End Sub

' ---- PATH별 윈도/시작/잔여처리/부호 규칙 ----
' nonMode/slMode: "DROP" | "INCLUDE" | "ANCHOR_END_FULL"
' ---- PATH별 윈도/시작/잔여처리/부호 규칙 ----
' nonMode/slMode: "DROP" | "INCLUDE" | "ANCHOR_END_FULL"
Private Sub GetPathParams(ByVal coreName As String, _
                          ByRef nonStart As Long, ByRef nonSize As Long, ByRef nonMode As String, _
                          ByRef slStart As Long, ByRef slSize As Long, ByRef slMode As String, _
                          ByRef flipSign As Boolean)
    Dim nm As String
    nm = UCase$(coreName)
    
    ' 기본 규칙(기존 PATH)
    nonStart = 1: nonSize = 27: nonMode = "INCLUDE"   ' 마지막 18 포함
    slStart = 14: slSize = 27: slMode = "DROP"        ' 14~796만
    flipSign = False                                   ' PM - AVG
    
    ' ---- 구체 항목을 먼저 검사 (충돌 방지: LF21/LF22가 LF2로 오인되지 않도록) ----
    If InStr(nm, "_HF1") > 0 Then
        ' 1) HF1
        nonStart = 1: nonSize = 4:  nonMode = "DROP"            ' 200묶음(800) 나머지1 버림
        slStart = 3: slSize = 4:   slMode = "INCLUDE"           ' E9부터 4씩, 마지막 3 포함
        flipSign = True
        Exit Sub
    End If
    
    If InStr(nm, "_LF21") > 0 Then
        ' 3) LF21
        nonStart = 1: nonSize = 41: nonMode = "ANCHOR_END_FULL" ' 마지막 E767~E807(41)로 앵커
        slStart = 21: slSize = 41:  slMode = "DROP"             ' E27부터 41씩, 나머지2 버림
        flipSign = True
        Exit Sub
    End If
    
    If InStr(nm, "_LF22") > 0 Then
        ' LF22는 기본(27 규칙) 유지 → 위의 기본값 그대로 사용
        ' flipSign=False (PM-AVG)
        Exit Sub
    End If
    
    ' 여기서부터는 더 일반적인 것들
    If MatchPathExact(nm, "_LF2") Then
        ' 2) LF2 (정확히 _LF2만 매칭; _LF21/_LF22는 위에서 이미 처리)
        nonStart = 1: nonSize = 6:  nonMode = "DROP"            ' 나머지5 버림
        slStart = 4: slSize = 6:    slMode = "DROP"             ' E10부터 6씩
        flipSign = True
        Exit Sub
    End If
    
    If InStr(nm, "_LF6") > 0 Then
        ' 4) LF6
        nonStart = 1: nonSize = 33: nonMode = "DROP"            ' 나머지9 버림
        slStart = 17: slSize = 33:  slMode = "INCLUDE"          ' E23부터 33씩, 마지막 26 포함
        flipSign = True
        Exit Sub
    End If
    
    If InStr(nm, "_LF3") > 0 Then
        ' 5) LF3
        nonStart = 1: nonSize = 88: nonMode = "DROP"            ' 나머지9 버림
        slStart = 45: slSize = 88:  slMode = "INCLUDE"          ' E51부터 88씩, 마지막 53 포함
        flipSign = True
        Exit Sub
    End If
    
    If InStr(nm, "_LF4") > 0 Then
        ' 6) LF4
        nonStart = 1: nonSize = 75: nonMode = "INCLUDE"         ' 마지막 51 포함
        slStart = 38: slSize = 75:  slMode = "DROP"             ' E44부터 75씩, 나머지14 버림
        flipSign = True
        Exit Sub
    End If
    
    If InStr(nm, "_LF5") > 0 Then
        ' 7) LF5
        nonStart = 1: nonSize = 81: nonMode = "INCLUDE"         ' 마지막 72 포함
        slStart = 41: slSize = 81:  slMode = "DROP"             ' E47부터 81씩, 나머지32 버림
        flipSign = True
        Exit Sub
    End If
    
    ' 그 외 PATH는 기본 규칙 유지 (27/14, PM-AVG)
End Sub

' "_LF2"가 들어있되, 뒤에 숫자가 이어지지 않는 경우만 True (즉 _LF21/_LF22는 제외)
Private Function MatchPathExact(ByVal nm As String, ByVal tag As String) As Boolean
    Dim p As Long, nextChar As String
    p = InStr(nm, tag)
    If p = 0 Then Exit Function
    If p + Len(tag) - 1 = Len(nm) Then
        MatchPathExact = True
    Else
        nextChar = Mid$(nm, p + Len(tag), 1)
        MatchPathExact = Not (nextChar >= "0" And nextChar <= "9")
    End If
End Function


' ---- 평균 배열 구성: 각 인덱스가 속한 구간 평균값을 직접 매핑 ----
Private Sub BuildAPMArrays(ByRef PM() As Double, _
                           ByVal nonStart As Long, ByVal nonSize As Long, ByVal nonMode As String, _
                           ByVal slStart As Long, ByVal slSize As Long, ByVal slMode As String, _
                           ByRef APM_non() As Double, ByRef mask_non() As Boolean, _
                           ByRef APM_sl() As Double, ByRef mask_sl() As Boolean)
    Dim i As Long, k As Long
    Dim n As Long: n = N_POINTS
    Dim startIdx As Long, endIdx As Long
    Dim full As Long, remN As Long
    Dim sum As Double, cnt As Long
    
    ReDim APM_non(1 To n): ReDim mask_non(1 To n)
    ReDim APM_sl(1 To n):  ReDim mask_sl(1 To n)
    
    ' ----- Non-overlap -----
    If nonSize > 0 Then
        ' 1) 기본 full 윈도
        If nonStart < 1 Then nonStart = 1
        If nonStart > n Then GoTo NON_DONE
        full = (n - nonStart + 1) \ nonSize
        For k = 0 To full - 1
            startIdx = nonStart + k * nonSize
            endIdx = startIdx + nonSize - 1
            If endIdx > n Then Exit For
            sum = 0#: cnt = 0
            For i = startIdx To endIdx: sum = sum + PM(i): cnt = cnt + 1: Next i
            If cnt > 0 Then
                sum = sum / cnt
                For i = startIdx To endIdx
                    APM_non(i) = sum: mask_non(i) = True
                Next i
            End If
        Next k
        
        ' 2) 잔여 처리
        remN = (n - nonStart + 1) - full * nonSize
        Select Case UCase$(nonMode)
            Case "INCLUDE"
                If remN > 0 Then
                    startIdx = nonStart + full * nonSize
                    endIdx = n
                    sum = 0#: cnt = 0
                    For i = startIdx To endIdx: sum = sum + PM(i): cnt = cnt + 1: Next i
                    sum = sum / cnt
                    For i = startIdx To endIdx
                        APM_non(i) = sum: mask_non(i) = True
                    Next i
                End If
            Case "ANCHOR_END_FULL"
                ' 마지막 구간을 끝쪽에 고정(길이=nonSize), 앞과 겹칠 수 있음
                If nonSize <= n Then
                    startIdx = n - nonSize + 1
                    endIdx = n
                    sum = 0#: cnt = 0
                    For i = startIdx To endIdx: sum = sum + PM(i): cnt = cnt + 1: Next i
                    sum = sum / cnt
                    For i = startIdx To endIdx
                        APM_non(i) = sum: mask_non(i) = True
                    Next i
                End If
            Case Else
                ' DROP: 아무것도 안 함
        End Select
    End If
NON_DONE:
    
    ' ----- Sliding(스텝=윈도 크기) -----
    If slSize > 0 Then
        If slStart < 1 Then slStart = 1
        If slStart <= n Then
            full = (n - slStart + 1) \ slSize
            ' full 윈도
            For k = 0 To full - 1
                startIdx = slStart + k * slSize
                endIdx = startIdx + slSize - 1
                If endIdx > n Then Exit For
                sum = 0#: cnt = 0
                For i = startIdx To endIdx: sum = sum + PM(i): cnt = cnt + 1: Next i
                sum = sum / cnt
                For i = startIdx To endIdx
                    APM_sl(i) = sum: mask_sl(i) = True
                Next i
            Next k
            ' 잔여 처리
            remN = (n - slStart + 1) - full * slSize
            If UCase$(slMode) = "INCLUDE" And remN > 0 Then
                startIdx = n - remN + 1
                endIdx = n
                sum = 0#: cnt = 0
                For i = startIdx To endIdx: sum = sum + PM(i): cnt = cnt + 1: Next i
                sum = sum / cnt
                For i = startIdx To endIdx
                    APM_sl(i) = sum: mask_sl(i) = True
                Next i
            End If
        End If
    End If
End Sub

' Excel ROUND 기반: 절대값 후 소수 둘째자리
Private Function Abs2(ByVal v As Double) As Double
    Abs2 = Application.WorksheetFunction.Round(Abs(v), 2)
End Function
