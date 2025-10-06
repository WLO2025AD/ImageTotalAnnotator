#Requires AutoHotkey v2.0
#SingleInstance Force
#Warn
SetWorkingDir A_ScriptDir

; ===== 사용자 경로 =====
ImageAnnotatorHTML := A_ScriptDir "\ImageAnnotator_v3.0.html"
ZoomItPath         := A_ScriptDir "\tools\ZoomIt64.exe"
ZoomItGuidePath    := A_ScriptDir "\docs\Zoomit 사용방법 251006.pdf"
;
; ======================

global gVisible := false

; ===== 패널 GUI =====
panel := Gui("+AlwaysOnTop +ToolWindow -Caption", "Image Annotator v2.2")
panel.MarginX := 10, panel.MarginY := 8
panel.BackColor := "0x202225"

title := panel.Add("Text", "cWhite BackgroundTrans w280", "🖊  Image Annotator v2.2")
title.SetFont("s10 bold")
panel.Add("Text", "cSilver BackgroundTrans", "어디서든 빠르게 주석/표시 (좌/우 도킹 + 온/오프)")

panel.Add("Text", "xm cAqua BackgroundTrans", "──────── Tools ────────")
btnEdge   := panel.Add("Button", "w280", "🌐 Edge 웹캡처 (Ctrl+Shift+S)")
btnSnip   := panel.Add("Button", "w280", "📸 Windows 스니핑 (Win+Shift+S)")
btnZDraw  := panel.Add("Button", "w280", "🖍 ZoomIt 드로잉 시작 (Ctrl+2)")
btnZOpts  := panel.Add("Button", "w280", "⚙ ZoomIt 실행/옵션 열기")
btnWAnn   := panel.Add("Button", "w280", "📄 Word: 이미지 주석 매크로 실행")
btnHTML   := panel.Add("Button", "w280", "🗂 Image Annotator (HTML) 열기")

panel.Add("Text", "xm cAqua BackgroundTrans", "──── 캡처 → Annotator ────")
btnShotWin  := panel.Add("Button", "w280", "🖼 현재 창 캡처 → Annotator (Ctrl+Alt+X)")
btnShotFull := panel.Add("Button", "w280", "🖼 전체 화면 캡처 → Annotator (Ctrl+Alt+Shift+X)")

panel.Add("Text", "xm cAqua BackgroundTrans", "──── ZoomIt 색상 퀵 버튼 ────")
row := panel.Add("Text", "w280 h0")
btnR := panel.Add("Button", "w40",  "R")
btnG := panel.Add("Button", "w40 xp+48 yp", "G")
btnB := panel.Add("Button", "w40 xp+48 yp", "B")
btnY := panel.Add("Button", "w40 xp+48 yp", "Y")
btnP := panel.Add("Button", "w40 xp+48 yp", "P")
for b in [btnR, btnG, btnB, btnY, btnP]
    b.ToolTip := "ZoomIt 드로잉 중에 누르면 색상 변경"

panel.Add("Text", "xm cAqua BackgroundTrans", "──── Help & Guide ────")
btnHotkeys := panel.Add("Button", "w280", "🧭 통합 핫키 모음 (팝업)")
btnZGuide  := panel.Add("Button", "w280", "📕 ZoomIt 가이드 PDF 열기")

panel.Add("Text", "xm cAqua BackgroundTrans", "──────── View ────────")
btnLeft  := panel.Add("Button", "w135", "⬅ 왼쪽 도킹")
btnRight := panel.Add("Button", "w135 xp+145 yp", "오른쪽 도킹 ➡")
btnHide  := panel.Add("Button", "xm w280", "🧹 패널 숨기기 (Ctrl+Alt+A)")

grip := panel.Add("Text", "xm w280 h8 BackgroundTrans")
OnMessage(0x201, (*) => PostMessage(0xA1, 2, , grip.Hwnd))  ; 드래그 이동

; ----------------------------------------------------------------------
; ===== 유틸 함수 (OnEvent 콜백 함수) =====
; ----------------------------------------------------------------------

OpenAnnotatorAndPaste() {
    if !FileExist(ImageAnnotatorHTML) {
        MsgBox("ImageAnnotator_v3.0.html 이 없습니다:`n" ImageAnnotatorHTML, "Error")
        return
    }
    Run(ImageAnnotatorHTML)
    Sleep 900
    Send("^v")  ; HTML이 paste 이벤트로 이미지 로드
}

EdgeWebCaptureCallback(*) {
    if WinActive("ahk_exe msedge.exe")
        Send("^+s")
    else
        Send("#+s")
}

StartZoomItDrawing(*) {
    if ProcessExist("ZoomIt64.exe") || ProcessExist("ZoomIt.exe") {
        Send("^2")
    } else if FileExist(ZoomItPath) {
        Run(ZoomItPath)
        Sleep 600
        Send("^2")
    } else {
        MsgBox("ZoomIt not found:`n" ZoomItPath, "Error")
    }
}

OpenZoomItOptions(*) {
    if FileExist(ZoomItPath)
        Run(ZoomItPath)
    else
        MsgBox("ZoomIt not found:`n" ZoomItPath, "Error")
}

RunWordAnnotateMacro(*) {
    if WinActive("ahk_exe WINWORD.EXE") {
        Send("!{F8}")
        Sleep 200
        Send("AnnotateSelectedImage{Enter}")
    } else {
        MsgBox("Word가 활성 창이 아닙니다. Word로 전환 후 이미지를 선택하세요.", "Info")
    }
}

OpenAnnotatorHTML(*) {
    if FileExist(ImageAnnotatorHTML)
        Run(ImageAnnotatorHTML)
    else
        MsgBox("ImageAnnotator_v3.0.html 이 없습니다:`n" ImageAnnotatorHTML, "Error")
}

CaptureWindowAndAnnotate(*) {
    Send("!{PrintScreen}")   ; 활성 창 캡처 → 클립보드
    OpenAnnotatorAndPaste()
}

CaptureFullAndAnnotate(*) {
    Send("{PrintScreen}")    ; 전체 화면 캡처 → 클립보드
    OpenAnnotatorAndPaste()
}

ShowHotkeysPopup(*) {
    hk := Gui("+AlwaysOnTop +ToolWindow", "통합 핫키 모음")
    hk.MarginX := 10, hk.MarginY := 10
    txt :=
    (
    "● 패널`n"
    "  Ctrl+Alt+A : 패널 표시/숨김`n"
    "`n"
    "● 캡처→Annotator`n"
    "  Ctrl+Alt+X : 현재 창 캡처→Annotator`n"
    "  Ctrl+Alt+Shift+X : 전체 화면 캡처→Annotator`n"
    "`n"
    "● ZoomIt (드로잉 시작 Ctrl+2)`n"
    "  색: R/G/B/Y/O/P  |  지우기: E  |  텍스트: T  |  종료: Esc`n"
    "`n"
    "● 웹`n"
    "  Ctrl+Alt+E : Edge 웹캡처(Edge 활성 시) / 스니핑`n"
    "  Ctrl+Alt+S : Windows 스니핑(Win+Shift+S)`n"
    "`n"
    "● Word`n"
    "  Ctrl+Alt+W : AnnotateSelectedImage 매크로`n"
    )
    hk.Add("Edit", "w440 h240 ReadOnly", txt)
    hk.Add("Button", "w440", "닫기").OnEvent("Click", (*) => hk.Destroy())
    hk.Show()
}

OpenZoomItGuide(*) {
    if FileExist(ZoomItGuidePath)
        Run(ZoomItGuidePath)
    else
        MsgBox("ZoomIt 가이드 PDF 경로를 확인하세요:`n" ZoomItGuidePath, "Info")
}

; ----------------------------------------------------------------------
; ===== 버튼 이벤트 (함수 연결) =====
; ----------------------------------------------------------------------
btnEdge.OnEvent("Click", EdgeWebCaptureCallback)

btnSnip.OnEvent("Click", (*) => Send("#+s"))

btnZDraw.OnEvent("Click", StartZoomItDrawing)

btnZOpts.OnEvent("Click", OpenZoomItOptions)

btnWAnn.OnEvent("Click", RunWordAnnotateMacro)

btnHTML.OnEvent("Click", OpenAnnotatorHTML)

btnShotWin.OnEvent("Click", CaptureWindowAndAnnotate)

btnShotFull.OnEvent("Click", CaptureFullAndAnnotate)

; Zoomit 색상 퀵 버튼 이벤트
btnR.OnEvent("Click", (*) => Send("r"))
btnG.OnEvent("Click", (*) => Send("g"))
btnB.OnEvent("Click", (*) => Send("b"))
btnY.OnEvent("Click", (*) => Send("y"))
btnP.OnEvent("Click", (*) => Send("p"))

btnHotkeys.OnEvent("Click", ShowHotkeysPopup)

btnZGuide.OnEvent("Click", OpenZoomItGuide)

; ===== 도킹/토글 함수 (간단한 로직은 화살표 함수 유지) =====
btnLeft.OnEvent("Click", (*) => Dock("L"))
btnRight.OnEvent("Click", (*) => Dock("R"))
btnHide.OnEvent("Click", (*) => TogglePanel())

TogglePanel() {
    global gVisible
    if gVisible {
        panel.Hide()
        gVisible := false
    } else {
        Dock("R")
        panel.Show()
        gVisible := true
    }
}

Dock(side:="R") {
    ; 변수 이름 L, T, R, B를 _L, _T, _R, _B로 변경하여 경고 해결
    MonitorGetWorkArea(1, &_L, &_T, &_R, &_B)
    w := 300, h := 640
    x := (side = "L") ? (_L + 8) : (_R - w - 8)
    y := _T + 80
    panel.Show("x" x " y" y " w" w " h" h)
}

; ===== 글로벌 단축키 =====
^!a::TogglePanel()
^!i::btnHTML.Click()
^!e::btnEdge.Click()
^!s::btnSnip.Click()
^!z::btnZDraw.Click()
^!w::btnWAnn.Click()
^!h::btnHotkeys.Click()
^!x::btnShotWin.Click()
^!+x::btnShotFull.Click()

; 스크립트 시작 시 패널 표시
TogglePanel()
return