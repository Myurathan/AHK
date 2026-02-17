#Requires AutoHotkey v2.0
#SingleInstance Force

; ==============================================================================
; SETTINGS
; ==============================================================================
global bgPath := "" ; Set to A_ScriptDir "\launcher-bg.jpg" if you have a background image

global cols   := 4
global bw     := 260
global bh     := 36
global gap    := 10
global HdrH   := 36
global MARG   := 16

; Visible rows before scroll
global MaxVisibleRows := 8
global ScrollBarW := 12

; Theme
global Theme_HeaderColor := "E11D2B"
global Theme_WindowBg    := "0F172A"
global Theme_TextColor   := "E5E7EB"

; Styling
global SearchH      := 34
global SearchRadius := 10
global Btn_Radius   := 10
global BtnTextPad   := 16
global Btn_FontName := "Segoe UI"
global Btn_FontSize := 10
global Btn_FontBold := true

; Effects
global UseAcrylic     := true
global Acrylic_Tint   := 0xE61A1F2A
global UseDwmShadow   := true
global CloseOnDefocus := true
global FadeInEnabled  := true

; INI Path
global iniPath := A_ScriptDir "\launcher.usage.ini"

; ==============================================================================
; BUTTONS
; ==============================================================================
global items := []

items.Push({ name: "Excel", cmd: "excel.exe" })
items.Push({ name: "Google Antigravity", cmd: "C:\Users\myura\AppData\Local\Programs\Antigravity\Antigravity.exe" })
items.Push({ name: "Claude", cmd: "msedge.exe --app=https://claude.ai" })
items.Push({ name: "Gemini", cmd: "msedge.exe --app=https://gemini.google.com" })
items.Push({ name: "Notepad++", cmd: "notepad++.exe" })
; items.Push({ name: "Calculator", cmd: "calc.exe" })

; ==============================================================================
; GLOBALS & INITIALIZATION
; ==============================================================================
global LauncherGui := ""
global hLauncher   := 0
global hQuery      := 0
global SearchBoxObj := ""
global SliderObj    := ""

global contentW := 0, searchW := 0
global currentH := 200
global hasBg := false

; Maps and State
global btnMap       := Map() 
global hwndToIdx    := Map()
global visibleOrder := []
global selPos       := 0
global pressedIdx   := 0
global usage        := Map()
global lastQueryLower := ""

; GDI+ Handles
global gToken := 0
global gHbmN := 0, gHbmH := 0, gHbmD := 0, gHbmSearch := 0
global hBrushEdit := 0
global Edit_BkColor := 0x0028160E ; BGR
global Edit_TxtColor := 0x00EBE7E5 ; BGR

; Scroll State
global scrollRow    := 0
global rowsTotal    := 0
global rowsVisible  := 0
global maxScrollRow := 0
global rowStep      := 0
global lastKeyNavTick := 0

; Cursor
global hCurHand := DllCall("LoadCursor", "Ptr", 0, "Ptr", 32649, "Ptr")

; Start GDI+
Launcher_GdipStart()
OnExit(Launcher_OnExit)

; Register Messages
OnMessage(0x0200, Launcher_WM_MOUSEMOVE)      ; WM_MOUSEMOVE
OnMessage(0x0020, Launcher_WM_SETCURSOR)      ; WM_SETCURSOR
OnMessage(0x0006, Launcher_WM_ACTIVATE)       ; WM_ACTIVATE
OnMessage(0x0201, Launcher_WM_LBUTTONDOWN)    ; WM_LBUTTONDOWN
OnMessage(0x0202, Launcher_WM_LBUTTONUP)      ; WM_LBUTTONUP
OnMessage(0x0133, Launcher_WM_CTLCOLOREDIT)   ; WM_CTLCOLOREDIT
OnMessage(0x020A, Launcher_WM_MOUSEWHEEL)     ; WM_MOUSEWHEEL
OnMessage(0x0100, Launcher_WM_KEYDOWN)        ; WM_KEYDOWN

; Build UI ONCE at startup
BuildUI()

; ==============================================================================
; HOTKEY
; ==============================================================================
^+Space::
{
    if WinExist("ahk_id " hLauncher) {
        if (DllCall("IsWindowVisible", "Ptr", hLauncher)) {
            LauncherGui.Hide()
        } else {
            ResetAndShow()
        }
    } else {
        ; Should rarely happen since we build at start, but safeguard:
        BuildUI()
        ResetAndShow()
    }
}

; ==============================================================================
; UI BUILDER
; ==============================================================================
BuildUI() {
    global

    ; Only build if it doesn't exist
    if (LauncherGui != "")
        return

    searchW  := cols * bw + (cols - 1) * gap
    contentW := searchW + (2 * MARG)

    Launcher_LoadUsage()
    Launcher_EnsureButtonBitmaps()
    Launcher_EnsureSearchBitmap(searchW, SearchH, SearchRadius)

    btnMap       := Map()
    hwndToIdx    := Map()
    visibleOrder := []
    selPos       := 0
    pressedIdx   := 0
    scrollRow    := 0

    ; Create GUI
    LauncherGui := Gui("+AlwaysOnTop -Caption +ToolWindow +LastFound +Owner", "Launcher")
    LauncherGui.BackColor := Theme_WindowBg
    LauncherGui.MarginX := MARG
    LauncherGui.MarginY := MARG
    LauncherGui.SetFont("s10", "Segoe UI")
    hLauncher := LauncherGui.Hwnd

    ; Background Image
    hasBg := (bgPath != "" && FileExist(bgPath))
    if (hasBg)
        LauncherGui.Add("Picture", "x0 y" HdrH " w" contentW " h1 vBgPic", bgPath)

    ; Header
    LauncherGui.Add("Progress", "x0 y0 w" contentW " h" HdrH " c" Theme_HeaderColor " vHdrBar", 100)
    LauncherGui.SetFont("s10 Bold", "Segoe UI")
    LauncherGui.Add("Text", "x20 y9 BackgroundTrans cFFFFFF", "Launcher")
    
    ; Close Button
    LauncherGui.SetFont("s12 Bold", "Segoe UI")
    HdrClose := LauncherGui.Add("Text", "x" (contentW - 36) " y7 w26 h22 Center BackgroundTrans cFFFFFF", Chr(215))
    HdrClose.OnEvent("Click", (*) => Launcher_HideWindow())
    
    ; Drag Handler
    HdrDrag := LauncherGui.Add("Text", "x0 y0 w" contentW " h" HdrH " BackgroundTrans")
    HdrDrag.OnEvent("Click", (*) => PostMessage(0xA1, 2, 0, , "ahk_id " hLauncher))
    
    LauncherGui.SetFont("s10", "Segoe UI")

    ; Search Background
    searchY := 52
    LauncherGui.Add("Picture", "x" MARG " y" searchY " w" searchW " h" SearchH " vSearchBg", "HBITMAP:" gHbmSearch)

    ; Search Edit
    editX := MARG + 12
    editY := searchY + 7
    editW := searchW - 24
    editH := SearchH - 14
    SearchBoxObj := LauncherGui.Add("Edit", "x" editX " y" editY " w" editW " h" editH " vQuery WantTab -E0x200")
    SearchBoxObj.OnEvent("Change", Launcher_Filter)
    hQuery := SearchBoxObj.Hwnd
    CueBanner(hQuery, "Type to filter...")

    ; Buttons
    fontOpts := "s" Btn_FontSize (Btn_FontBold ? " Bold" : "")
    LauncherGui.SetFont(fontOpts, Btn_FontName)
    
    ; Pre-calculate GUI height needed for all buttons
    totalRows := Ceil(items.Length / cols)
    neededHeight := 100 + (totalRows * bh) + ((totalRows - 1) * gap) + MARG + 50
    LauncherGui.Show("Hide w" contentW " h" neededHeight)

    yButtons := 100
    Loop items.Length {
        i := A_Index
        
        ; Button Image - create empty, Launcher_Render will set bitmap
        PicObj := LauncherGui.Add("Picture", "x" MARG " y" yButtons " w" bw " h" bh)
        PicObj.OnEvent("Click", Launcher_Click_Bound.Bind(i))
        
        ; Button Text
        txX := MARG + BtnTextPad
        txW := bw - (BtnTextPad * 2)
        TxtObj := LauncherGui.Add("Text", "x" txX " y" yButtons " w" txW " h" bh " BackgroundTrans c" Theme_TextColor " +0x200", items[i].name)
        TxtObj.OnEvent("Click", Launcher_Click_Bound.Bind(i))

        ; Map metadata - store bitmap handles per button
        btnMap[i] := { Pic: PicObj, Txt: TxtObj, state: "", visible: false }
        hwndToIdx[PicObj.Hwnd] := i
        hwndToIdx[TxtObj.Hwnd] := i
        
        ; Increment Y position for next button
        yButtons += (bh + gap)
    }

    ; Scrollbar
    scrollX := contentW - MARG + Floor((MARG - ScrollBarW) / 2)
    if (scrollX < contentW - ScrollBarW - 2)
        scrollX := contentW - ScrollBarW - 2

    scrollY := 100
    scrollH := (MaxVisibleRows * bh) + ((MaxVisibleRows - 1) * gap)
    if (scrollH < bh)
        scrollH := bh

    SliderObj := LauncherGui.Add("Slider", "x" scrollX " y" scrollY " w" ScrollBarW " h" scrollH " Vertical AltSubmit NoTicks -TabStop Range0-0", 0)
    SliderObj.OnEvent("Change", Launcher_Scroll)
    SliderObj.Visible := false

    lastQueryLower := ""
    
    ; Initial Setup (Hide immediately)
    Launcher_LayoutButtons("")
    LauncherGui.Show("Hide w" contentW " h" currentH)
}

ResetAndShow() {
    global
    if !LauncherGui
        return

    SearchBoxObj.Value := ""
    lastQueryLower := ""
    scrollRow := 0
    Launcher_LayoutButtons("")

    LauncherGui.Show("w" contentW " h" currentH " Center NA")

    if (UseDwmShadow)
        EnableDwmShadow(hLauncher)
    if (UseAcrylic)
        EnableAcrylic(hLauncher, Acrylic_Tint)

    SearchBoxObj.Focus()
    if (FadeInEnabled)
        FadeIn(hLauncher, 210, 255, 12)
}

Launcher_HideWindow() {
    global
    if (LauncherGui)
        LauncherGui.Hide()
    ResetGlobals()
}

ResetGlobals() {
    global
    selPos := 0
    pressedIdx := 0
    visibleOrder := []
    scrollRow := 0
}

; ==============================================================================
; EVENTS & ACTIONS
; ==============================================================================
Launcher_Filter(*) {
    global lastQueryLower, scrollRow
    q := SearchBoxObj.Value
    lastQueryLower := StrLower(q)
    scrollRow := 0
    Launcher_LayoutButtons(lastQueryLower)
}

Launcher_Click_Bound(idx, *) {
    Launcher_RunIndex(idx)
}

Launcher_Scroll(*) {
    global scrollRow
    Launcher_SetScrollRow(SliderObj.Value, true)
}

; ==============================================================================
; LAYOUT & RENDER
; ==============================================================================
Launcher_LayoutButtons(queryLower) {
    global visibleOrder, selPos, SliderObj, cols, bw, gap
    
    if (!IsObject(SliderObj))
        return
    
    visibleOrder := Launcher_GetOrder(queryLower)

    if (visibleOrder.Length >= 1)
        selPos := 1
    else
        selPos := 0

    Launcher_UpdateScrollMetrics()
    Launcher_Render()
}

Launcher_UpdateScrollMetrics() {
    global

    cnt := visibleOrder.Length
    
    if (cnt < 1) {
        rowsTotal := 0
        rowsVisible := 0
        maxScrollRow := 0
        scrollRow := 0
        SliderObj.Visible := false
        Launcher_ResizeToViewport()
        return
    }

    rowsTotal := Ceil(cnt / cols)
    rowsVisible := rowsTotal
    if (rowsVisible > MaxVisibleRows)
        rowsVisible := MaxVisibleRows

    maxScrollRow := rowsTotal - rowsVisible
    if (maxScrollRow < 0)
        maxScrollRow := 0

    if (scrollRow > maxScrollRow)
        scrollRow := maxScrollRow
    if (scrollRow < 0)
        scrollRow := 0

    ; Resize slider
    btnAreaH := rowsVisible * bh + (rowsVisible - 1) * gap
    if (btnAreaH < bh)
        btnAreaH := bh
    
    if (maxScrollRow > 0) {
        SliderObj.Opt("+Range0-" maxScrollRow)
        SliderObj.Value := scrollRow
        SliderObj.Move(, , , btnAreaH)
        SliderObj.Visible := true
    } else {
        SliderObj.Value := 0
        SliderObj.Visible := false
    }

    Launcher_ResizeToViewport()
}

Launcher_ResizeToViewport() {
    global currentH

    y0 := 100
    if (rowsVisible <= 0) {
        lastBottom := y0
    } else {
        btnAreaH := rowsVisible * bh + (rowsVisible - 1) * gap
        lastBottom := y0 + btnAreaH
    }

    totalH := lastBottom + MARG + 18
    minH := 160
    if (totalH < minH)
        totalH := minH

    currentH := totalH

    if (hLauncher && DllCall("IsWindowVisible", "Ptr", hLauncher)) {
        LauncherGui.Show("w" contentW " h" currentH " NA")
    }
}

Launcher_Render() {
    global

    ; Hide all first
    Loop items.Length {
        i := A_Index
        btnMap[i].Pic.Visible := false
        btnMap[i].Txt.Visible := false
        btnMap[i].visible := false
    }

    cnt := visibleOrder.Length
    if (cnt < 1) {
        UpdateBackground()
        return
    }

    y0 := 100
    for pos, idx in visibleOrder {
        row := (pos - 1) // cols
        col := Mod(pos - 1, cols)

        if (row < scrollRow || row >= scrollRow + rowsVisible)
            continue

        visRow := row - scrollRow
        x := MARG + col * (bw + gap)
        y := y0 + visRow * (bh + gap)

        btnMap[idx].Pic.Move(x, y, bw, bh)
        
        txX := x + BtnTextPad
        txW := bw - (BtnTextPad * 2)
        btnMap[idx].Txt.Move(txX, y, txW, bh)

        btnMap[idx].Pic.Visible := true
        btnMap[idx].Txt.Visible := true
        btnMap[idx].visible := true
        
        ; Set state AFTER button is visible
        Launcher_SetState(idx, "N")
    }

    Launcher_ApplySelectionVisual()
    UpdateBackground()
}

UpdateBackground() {
    global hasBg
    if (!hasBg || !hLauncher)
        return
    
    WinGetPos(,,, &h, "ahk_id " hLauncher)
    if (h) {
        newH := h - HdrH
        try LauncherGui["BgPic"].Move(0, HdrH, contentW, newH)
    }
}

; ==============================================================================
; SELECTION & RUN
; ==============================================================================
Launcher_SetScrollRow(newRow, adjustSelection := true) {
    global scrollRow, selPos, lastKeyNavTick
    
    if (maxScrollRow <= 0) {
        scrollRow := 0
        SliderObj.Value := 0
        Launcher_Render()
        return
    }

    if (newRow < 0)
        newRow := 0
    if (newRow > maxScrollRow)
        newRow := maxScrollRow

    if (newRow = scrollRow)
        return

    scrollRow := newRow
    SliderObj.Value := scrollRow

    if (adjustSelection) {
        firstPos := scrollRow * cols + 1
        maxPos := visibleOrder.Length
        if (maxPos < 1) {
            selPos := 0
        } else {
            if (firstPos > maxPos)
                firstPos := maxPos
            selPos := firstPos
        }
    }

    lastKeyNavTick := A_TickCount
    Launcher_Render()
}

Launcher_ScrollBy(rowsDelta) {
    Launcher_SetScrollRow(scrollRow + rowsDelta, true)
}

Launcher_MoveSel(delta) {
    global selPos, scrollRow, lastKeyNavTick
    
    if (visibleOrder.Length < 1)
        return

    if (selPos < 1)
        selPos := 1

    newPos := selPos + delta
    if (newPos < 1)
        newPos := 1
    if (newPos > visibleOrder.Length)
        newPos := visibleOrder.Length

    if (newPos = selPos)
        return

    selPos := newPos
    lastKeyNavTick := A_TickCount

    rowSel := (selPos - 1) // cols
    oldScroll := scrollRow

    if (rowSel < scrollRow)
        scrollRow := rowSel
    else if (rowSel >= scrollRow + rowsVisible)
        scrollRow := rowSel - (rowsVisible - 1)

    if (scrollRow < 0)
        scrollRow := 0
    if (scrollRow > maxScrollRow)
        scrollRow := maxScrollRow

    if (scrollRow != oldScroll) {
        SliderObj.Value := scrollRow
        Launcher_Render()
    } else {
        Launcher_ApplySelectionVisual()
    }
}

Launcher_ApplySelectionVisual() {
    for pos, idx in visibleOrder
        Launcher_SetState(idx, "N")

    if (selPos >= 1 && selPos <= visibleOrder.Length) {
        idx := visibleOrder[selPos]
        Launcher_SetState(idx, "H")
    }
}

Launcher_RunSelected() {
    global visibleOrder, selPos
    if (visibleOrder.Length < 1)
        return
    if (selPos < 1)
        selPos := 1
    idx := visibleOrder[selPos]
    Launcher_RunIndex(idx)
}

Launcher_RunIndex(idx) {
    global usage
    if (idx <= 0 || idx > items.Length)
        return

    nm := items[idx].name
    cnt := usage.Has(nm) ? usage[nm] : 0
    cnt++
    usage[nm] := cnt
    try IniWrite(cnt, iniPath, "Usage", nm)

    o := items[idx]
    Launcher_HideWindow()

    if IsObject(o) {
        try {
            Run(o.cmd)
        } catch {
            MsgBox("Error running: " o.cmd)
        }
    }
}

; ==============================================================================
; STATE & BITMAPS
; ==============================================================================
Launcher_SetState(idx, state) {
    global bw, bh, Btn_Radius, gHbmN, gHbmH, gHbmD

    if (!btnMap[idx].visible)
        return
    if (btnMap[idx].state = state)
        return
    btnMap[idx].state := state

    hbm := (state = "H") ? gHbmH : (state = "D") ? gHbmD : gHbmN

    if (hbm) {
        try {
            ; Use HBITMAP:* to force a copy, preventing AHK taking ownership of our global handle
            btnMap[idx].Pic.Value := "HBITMAP:*" hbm
        }
    }
}

; ==============================================================================
; WINDOW MESSAGES (THE HARD PART)
; ==============================================================================
Launcher_WM_MOUSEMOVE(wParam, lParam, msg, hwnd) {
    global lastKeyNavTick, pressedIdx, selPos

    if (!hLauncher || hwnd != hLauncher)
        return

    if (A_TickCount - lastKeyNavTick < 250)
        return

    MouseGetPos(,, &win, &ctrl, 2)
    if (win != hLauncher)
        return

    if (pressedIdx) {
        if (hwndToIdx.Has(ctrl) && hwndToIdx[ctrl] = pressedIdx)
            Launcher_SetState(pressedIdx, "D")
        else
            Launcher_SetState(pressedIdx, "H")
        return
    }

    if (hwndToIdx.Has(ctrl)) {
        idx := hwndToIdx[ctrl]
        if (idx && btnMap[idx].visible) {
            pos := Launcher_PosInVisibleOrder(idx)
            if (pos) {
                selPos := pos
                Launcher_ApplySelectionVisual()
            }
        }
    }
}

Launcher_PosInVisibleOrder(idx) {
    for pos, v in visibleOrder
        if (v = idx)
            return pos
    return 0
}

Launcher_WM_LBUTTONDOWN(wParam, lParam, msg, hwnd) {
    global pressedIdx
    if (!hLauncher || hwnd != hLauncher)
        return
    MouseGetPos(,, &win, &ctrl, 2)
    if (win != hLauncher)
        return
    if (hwndToIdx.Has(ctrl)) {
        idx := hwndToIdx[ctrl]
        if (idx && btnMap[idx].visible) {
            pressedIdx := idx
            Launcher_SetState(idx, "D")
            DllCall("SetCapture", "Ptr", hLauncher)
        }
    }
}

Launcher_WM_LBUTTONUP(wParam, lParam, msg, hwnd) {
    global pressedIdx
    if (!hLauncher || hwnd != hLauncher)
        return
    if (pressedIdx) {
        DllCall("ReleaseCapture")
        Launcher_SetState(pressedIdx, "H")
        pressedIdx := 0
    }
}

Launcher_WM_SETCURSOR(wParam, lParam, msg, hwnd) {
    if (!hLauncher || hwnd != hLauncher)
        return
    MouseGetPos(,, &win, &ctrl, 2)
    if (win != hLauncher)
        return
    if (hwndToIdx.Has(ctrl)) {
        DllCall("SetCursor", "Ptr", hCurHand)
        return true
    }
}

Launcher_WM_ACTIVATE(wParam, lParam, msg, hwnd) {
    if (!CloseOnDefocus)
        return
    if (!hLauncher || hwnd != hLauncher)
        return
    if (wParam = 0)
        Launcher_HideWindow()
}

Launcher_WM_MOUSEWHEEL(wParam, lParam, msg, hwnd) {
    if (!hLauncher || !WinActive("ahk_id " hLauncher))
        return

    delta := (wParam >> 16) & 0xFFFF
    if (delta & 0x8000)
        delta := -(0x10000 - delta)

    if (delta > 0)
        Launcher_ScrollBy(-1)
    else if (delta < 0)
        Launcher_ScrollBy(1)

    return 0
}

Launcher_WM_KEYDOWN(wParam, lParam, msg, hwnd) {
    if (!hLauncher || !WinActive("ahk_id " hLauncher))
        return
    if (hwnd != hQuery)
        return

    vk := wParam

    if (vk = 0x1B) { ; Esc
        Launcher_HideWindow()
        return 0
    }
    if (vk = 0x0D) { ; Enter
        Launcher_RunSelected()
        return 0
    }
    if (vk = 0x26) { ; Up
        Launcher_MoveSel(-cols)
        return 0
    }
    if (vk = 0x28) { ; Down
        Launcher_MoveSel(cols)
        return 0
    }
    if (vk = 0x25) { ; Left
        Launcher_MoveSel(-1)
        return 0
    }
    if (vk = 0x27) { ; Right
        Launcher_MoveSel(1)
        return 0
    }
    if (vk = 0x09) { ; Tab
        if GetKeyState("Shift", "P")
            Launcher_MoveSel(-1)
        else
            Launcher_MoveSel(1)
        return 0
    }
    if (vk = 0x21) { ; PageUp
        Launcher_SetScrollRow(scrollRow - rowsVisible, true)
        return 0
    }
    if (vk = 0x22) { ; PageDown
        Launcher_SetScrollRow(scrollRow + rowsVisible, true)
        return 0
    }
    if (vk = 0x24) { ; Home
        global scrollRow := 0
        SliderObj.Value := 0
        global selPos := (visibleOrder.Length >= 1) ? 1 : 0
        Launcher_Render()
        return 0
    }
    if (vk = 0x23) { ; End
        Launcher_SetScrollRow(maxScrollRow, true)
        return 0
    }
}

Launcher_WM_CTLCOLOREDIT(wParam, lParam, msg, hwnd) {
    global hBrushEdit
    if (lParam != hQuery)
        return
    DllCall("SetTextColor", "Ptr", wParam, "UInt", Edit_TxtColor)
    DllCall("SetBkColor", "Ptr", wParam, "UInt", Edit_BkColor)
    if (!hBrushEdit)
        hBrushEdit := DllCall("CreateSolidBrush", "UInt", Edit_BkColor, "Ptr")
    return hBrushEdit
}

; ==============================================================================
; UTILS
; ==============================================================================
Launcher_LoadUsage() {
    global usage
    usage := Map()
    Loop items.Length {
        nm := items[A_Index].name
        try {
            cnt := IniRead(iniPath, "Usage", nm, 0)
            usage[nm] := Integer(cnt)
        } catch {
            usage[nm] := 0
        }
    }
}

Launcher_GetOrder(queryLower) {
    arr := []
    
    if (queryLower = "") {
        Loop items.Length {
            i := A_Index
            nm := items[i].name
            cnt := usage.Has(nm) ? usage[nm] : 0
            arr.Push({i: i, score: 0, use: cnt, name: StrLower(nm)})
        }
        Launcher_SortArr(arr, true)
        return Launcher_IdxList(arr)
    }

    Loop items.Length {
        i := A_Index
        nmL := StrLower(items[i].name)
        sc := Launcher_FuzzyScore(nmL, queryLower)
        if (sc > 0) {
            nm := items[i].name
            cnt := usage.Has(nm) ? usage[nm] : 0
            arr.Push({i: i, score: sc, use: cnt, name: nmL})
        }
    }

    Launcher_SortArr(arr, false)
    return Launcher_IdxList(arr)
}

Launcher_IdxList(arr) {
    out := []
    for k, o in arr
        out.Push(o.i)
    return out
}

Launcher_SortArr(arr, usageOnly := false) {
    ; Bubble sort sufficient for small lists
    n := arr.Length
    if (n <= 1)
        return
    Loop n - 1 {
        a := A_Index
        b := a + 1
        while (b <= n) {
            if (Launcher_IsBetter(arr[b], arr[a], usageOnly)) {
                tmp := arr[a], arr[a] := arr[b], arr[b] := tmp
            }
            b++
        }
    }
}

Launcher_IsBetter(o1, o2, usageOnly) {
    if (usageOnly) {
        if (o1.use != o2.use)
            return (o1.use > o2.use)
        return (StrCompare(o1.name, o2.name) < 0)
    }
    if (o1.score != o2.score)
        return (o1.score > o2.score)
    if (o1.use != o2.use)
        return (o1.use > o2.use)
    return (StrCompare(o1.name, o2.name) < 0)
}

Launcher_FuzzyScore(hay, needle) {
    if (needle = "")
        return 0

    pos := InStr(hay, needle)
    if (pos) {
        sc := 2000 - (pos * 10)
        if (pos = 1)
            sc += 400
        return sc
    }
    
    ; Simple subsequence matching
    i := 1
    skips := 0
    len := StrLen(needle)
    Loop len {
        ch := SubStr(needle, A_Index, 1)
        found := InStr(hay, ch, false, i)
        if (!found)
            return 0
        skips += (found - i)
        i := found + 1
    }
    sc := 900 - skips
    if (sc < 1)
        sc := 1
    return sc
}

FadeIn(hwnd, from := 0, to := 255, step := 15) {
    try WinSetTransparent(from, "ahk_id " hwnd)
    t := from
    while (t < to) {
        t += step
        if (t > to)
            t := to
        try WinSetTransparent(t, "ahk_id " hwnd)
        Sleep(8)
    }
}

CueBanner(hEdit, text) {
    buf := Buffer((StrLen(text) + 1) * 2, 0)
    StrPut(text, buf, "UTF-16")
    SendMessage(0x1501, 1, buf.Ptr, , "ahk_id " hEdit)
}

EnableDwmShadow(hwnd) {
    if !hwnd
        return
    try {
        if !DllCall("GetModuleHandle", "Str", "dwmapi.dll", "Ptr")
            DllCall("LoadLibrary", "Str", "dwmapi.dll", "Ptr")

        MARGINS := Buffer(16, 0)
        NumPut("Int", -1, MARGINS, 0)
        NumPut("Int", -1, MARGINS, 4)
        NumPut("Int", -1, MARGINS, 8)
        NumPut("Int", -1, MARGINS, 12)
        DllCall("dwmapi\DwmExtendFrameIntoClientArea", "Ptr", hwnd, "Ptr", MARGINS.Ptr)
        
        pref := 2
        DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", hwnd, "Int", 33, "Int*", &pref, "Int", 4)
    }
}

EnableAcrylic(hwnd, gradientColor := 0xCC202020) {
    if !hwnd
        return false
    try {
        hUser := DllCall("GetModuleHandle", "Str", "user32.dll", "Ptr")
        pSet  := DllCall("GetProcAddress", "Ptr", hUser, "AStr", "SetWindowCompositionAttribute", "Ptr")
        if (!pSet)
            return false

        ACCENT := Buffer(16, 0)
        NumPut("Int", 4, ACCENT, 0)
        NumPut("Int", 2, ACCENT, 4)
        NumPut("Int", gradientColor, ACCENT, 8)
        NumPut("Int", 0, ACCENT, 12)

        DATA := Buffer(A_PtrSize * 2 + 8, 0)
        NumPut("Int", 19, DATA, 0)
        NumPut("Ptr", ACCENT.Ptr, DATA, A_PtrSize)
        NumPut("UInt", 16, DATA, A_PtrSize * 2)

        DllCall(pSet, "Ptr", hwnd, "Ptr", DATA.Ptr)
    }
}

; ==============================================================================
; GDI+ LIBRARIES (EMBEDDED)
; ==============================================================================
Launcher_GdipStart() {
    global gToken
    if (gToken)
        return
    if !DllCall("GetModuleHandle", "Str", "gdiplus.dll", "Ptr")
        DllCall("LoadLibrary", "Str", "gdiplus.dll", "Ptr")

    si := Buffer(24, 0) ; size 24 for 64-bit usually safe, 16 for 32-bit. 
    NumPut("UInt", 1, si, 0)
    status := DllCall("gdiplus\GdiplusStartup", "Ptr*", &gToken, "Ptr", si.Ptr, "Ptr", 0)
    
    if (status != 0 || !gToken) {
        MsgBox("Error: GDI+ failed to initialize (Status: " status ").`n`nThe launcher cannot start.", "GDI+ Error", "Icon!")
        ExitApp
    }
}

Launcher_OnExit(*) {
    ; Clean up global bitmaps (if any)
    Launcher_DeleteHbm(gHbmN), Launcher_DeleteHbm(gHbmH), Launcher_DeleteHbm(gHbmD)
    Launcher_DeleteHbm(gHbmSearch)
    
    DllCall("DeleteObject", "Ptr", hBrushEdit)
    
    if (gToken)
        DllCall("gdiplus\GdiplusShutdown", "Ptr", gToken)
}

Launcher_EnsureButtonBitmaps() {
    global gHbmN, gHbmH, gHbmD
    if (gHbmN && gHbmH && gHbmD)
        return

    Launcher_DeleteHbm(gHbmN), Launcher_DeleteHbm(gHbmH), Launcher_DeleteHbm(gHbmD)
    gHbmN := Launcher_MakeShinyButtonBitmap(bw, bh, Btn_Radius, "N")
    gHbmH := Launcher_MakeShinyButtonBitmap(bw, bh, Btn_Radius, "H")
    gHbmD := Launcher_MakeShinyButtonBitmap(bw, bh, Btn_Radius, "D")
    
    if (!gHbmN || !gHbmH || !gHbmD) {
        MsgBox("Error: Failed to create button bitmaps (gHbmN=" gHbmN ").", "Bitmap Error", "Icon!")
        ExitApp
    }
}

Launcher_EnsureSearchBitmap(w, h, r) {
    global gHbmSearch
    if (gHbmSearch)
        return
    gHbmSearch := Launcher_MakeSearchBitmap(w, h, r)
}

Launcher_DeleteHbm(hbm) {
    if (hbm)
        DllCall("DeleteObject", "Ptr", hbm)
}

Launcher_CloneHBITMAP(srcHbm, w, h) {
    if (!srcHbm)
        return 0
    
    ; Create a compatible DC
    hdc := DllCall("GetDC", "Ptr", 0, "Ptr")
    hdcMem1 := DllCall("CreateCompatibleDC", "Ptr", hdc, "Ptr")
    hdcMem2 := DllCall("CreateCompatibleDC", "Ptr", hdc, "Ptr")
    
    ; Create new bitmap
    newHbm := DllCall("CreateCompatibleBitmap", "Ptr", hdc, "Int", w, "Int", h, "Ptr")
    
    ; Select bitmaps into DCs
    oldBmp1 := DllCall("SelectObject", "Ptr", hdcMem1, "Ptr", srcHbm, "Ptr")
    oldBmp2 := DllCall("SelectObject", "Ptr", hdcMem2, "Ptr", newHbm, "Ptr")
    
    ; Copy bitmap
    DllCall("BitBlt", "Ptr", hdcMem2, "Int", 0, "Int", 0, "Int", w, "Int", h
          , "Ptr", hdcMem1, "Int", 0, "Int", 0, "UInt", 0x00CC0020) ; SRCCOPY
    
    ; Cleanup
    DllCall("SelectObject", "Ptr", hdcMem1, "Ptr", oldBmp1)
    DllCall("SelectObject", "Ptr", hdcMem2, "Ptr", oldBmp2)
    DllCall("DeleteDC", "Ptr", hdcMem1)
    DllCall("DeleteDC", "Ptr", hdcMem2)
    DllCall("ReleaseDC", "Ptr", 0, "Ptr", hdc)
    
    return newHbm
}

Launcher_MakeSearchBitmap(w, h, r) {
    pBitmap := 0
    DllCall("gdiplus\GdipCreateBitmapFromScan0", "Int", w, "Int", h, "Int", 0, "Int", 0xE200B, "Ptr", 0, "Ptr*", &pBitmap)
    pG := 0
    DllCall("gdiplus\GdipGetImageGraphicsContext", "Ptr", pBitmap, "Ptr*", &pG)
    DllCall("gdiplus\GdipSetSmoothingMode", "Ptr", pG, "Int", 4)

    pPath := Gdip_CreateRoundRectPath(0, 0, w, h, r)

    Rect := Buffer(16, 0)
    NumPut("Int", 0, Rect, 0)
    NumPut("Int", 0, Rect, 4)
    NumPut("Int", w, Rect, 8)
    NumPut("Int", h, Rect, 12)

    pBrushGrad := 0
    DllCall("gdiplus\GdipCreateLineBrushFromRectI", "Ptr", Rect.Ptr, "UInt", 0xFF111B2F, "UInt", 0xFF0B1220, "Int", 1, "Int", 3, "Ptr*", &pBrushGrad)
    DllCall("gdiplus\GdipFillPath", "Ptr", pG, "Ptr", pBrushGrad, "Ptr", pPath)
    DllCall("gdiplus\GdipDeleteBrush", "Ptr", pBrushGrad)

    pPen := 0
    DllCall("gdiplus\GdipCreatePen1", "UInt", 0x33FFFFFF, "Float", 1.0, "Int", 2, "Ptr*", &pPen)
    DllCall("gdiplus\GdipDrawPath", "Ptr", pG, "Ptr", pPen, "Ptr", pPath)
    DllCall("gdiplus\GdipDeletePen", "Ptr", pPen)

    DllCall("gdiplus\GdipDeletePath", "Ptr", pPath)

    hbm := 0
    DllCall("gdiplus\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "Ptr*", &hbm, "UInt", 0x00000000)
    DllCall("gdiplus\GdipDeleteGraphics", "Ptr", pG)
    DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
    return hbm
}

Launcher_MakeShinyButtonBitmap(w, h, r, state) {
    if (state = "H") {
        cTop := 0xFF36527D, cBot := 0xFF1A2A46, cBorder := 0x66FFFFFF
        cGlow := 0x33E11D2B, glossA := 0x34FFFFFF, shadowA := 0x2A000000, accent := 0xFFE11D2B
    } else if (state = "D") {
        cTop := 0xFF1B2A46, cBot := 0xFF0E1628, cBorder := 0x44FFFFFF
        cGlow := 0x22000000, glossA := 0x18FFFFFF, shadowA := 0x26000000, accent := 0xFFB11621
    } else {
        cTop := 0xFF2A3A55, cBot := 0xFF111A2D, cBorder := 0x40FFFFFF
        cGlow := 0x00000000, glossA := 0x26FFFFFF, shadowA := 0x24000000, accent := 0x00000000
    }

    pBitmap := 0
    DllCall("gdiplus\GdipCreateBitmapFromScan0", "Int", w, "Int", h, "Int", 0, "Int", 0xE200B, "Ptr", 0, "Ptr*", &pBitmap)
    pG := 0
    DllCall("gdiplus\GdipGetImageGraphicsContext", "Ptr", pBitmap, "Ptr*", &pG)
    DllCall("gdiplus\GdipSetSmoothingMode", "Ptr", pG, "Int", 4)

    ; Shadow
    pPathS := Gdip_CreateRoundRectPath(2, 2, w-4, h-4, r)
    pBrushShadow := 0
    DllCall("gdiplus\GdipCreateSolidFill", "UInt", shadowA, "Ptr*", &pBrushShadow)
    DllCall("gdiplus\GdipFillPath", "Ptr", pG, "Ptr", pBrushShadow, "Ptr", pPathS)
    DllCall("gdiplus\GdipDeleteBrush", "Ptr", pBrushShadow)
    DllCall("gdiplus\GdipDeletePath", "Ptr", pPathS)

    ; Main
    pPath := Gdip_CreateRoundRectPath(1, 1, w-4, h-4, r)
    Rect := Buffer(16, 0)
    NumPut("Int", 1, Rect, 0), NumPut("Int", 1, Rect, 4)
    NumPut("Int", w-4, Rect, 8), NumPut("Int", h-4, Rect, 12)

    pBrushGrad := 0
    DllCall("gdiplus\GdipCreateLineBrushFromRectI", "Ptr", Rect.Ptr, "UInt", cTop, "UInt", cBot, "Int", 1, "Int", 3, "Ptr*", &pBrushGrad)
    DllCall("gdiplus\GdipFillPath", "Ptr", pG, "Ptr", pBrushGrad, "Ptr", pPath)
    DllCall("gdiplus\GdipDeleteBrush", "Ptr", pBrushGrad)

    ; Gloss
    hGloss := Floor((h-4) * 0.52)
    pPathGloss := Gdip_CreateRoundRectPath(1, 1, w-4, hGloss, r)
    Rect2 := Buffer(16, 0)
    NumPut("Int", 1, Rect2, 0), NumPut("Int", 1, Rect2, 4)
    NumPut("Int", w-4, Rect2, 8), NumPut("Int", hGloss, Rect2, 12)

    pBrushGloss := 0
    DllCall("gdiplus\GdipCreateLineBrushFromRectI", "Ptr", Rect2.Ptr, "UInt", glossA, "UInt", 0x00FFFFFF, "Int", 1, "Int", 3, "Ptr*", &pBrushGloss)
    DllCall("gdiplus\GdipFillPath", "Ptr", pG, "Ptr", pBrushGloss, "Ptr", pPathGloss)
    DllCall("gdiplus\GdipDeleteBrush", "Ptr", pBrushGloss)
    DllCall("gdiplus\GdipDeletePath", "Ptr", pPathGloss)

    ; Accent
    if (accent != 0) {
        pBrushA := 0
        DllCall("gdiplus\GdipCreateSolidFill", "UInt", accent, "Ptr*", &pBrushA)
        DllCall("gdiplus\GdipFillRectangleI", "Ptr", pG, "Ptr", pBrushA, "Int", 1+3, "Int", 1+4, "Int", 3, "Int", h-4-8)
        DllCall("gdiplus\GdipDeleteBrush", "Ptr", pBrushA)
    }

    ; Glow
    if (cGlow != 0) {
        pPenGlow := 0
        DllCall("gdiplus\GdipCreatePen1", "UInt", cGlow, "Float", 2.0, "Int", 2, "Ptr*", &pPenGlow)
        DllCall("gdiplus\GdipDrawPath", "Ptr", pG, "Ptr", pPenGlow, "Ptr", pPath)
        DllCall("gdiplus\GdipDeletePen", "Ptr", pPenGlow)
    }

    ; Border
    pPen := 0
    DllCall("gdiplus\GdipCreatePen1", "UInt", cBorder, "Float", 1.0, "Int", 2, "Ptr*", &pPen)
    DllCall("gdiplus\GdipDrawPath", "Ptr", pG, "Ptr", pPen, "Ptr", pPath)
    DllCall("gdiplus\GdipDeletePen", "Ptr", pPen)

    DllCall("gdiplus\GdipDeletePath", "Ptr", pPath)

    hbm := 0
    DllCall("gdiplus\GdipCreateHBITMAPFromBitmap", "Ptr", pBitmap, "Ptr*", &hbm, "UInt", 0x00000000)
    DllCall("gdiplus\GdipDeleteGraphics", "Ptr", pG)
    DllCall("gdiplus\GdipDisposeImage", "Ptr", pBitmap)
    return hbm
}

Gdip_CreateRoundRectPath(x, y, w, h, r) {
    pPath := 0
    DllCall("gdiplus\GdipCreatePath", "Int", 0, "Ptr*", &pPath)

    d := r * 2
    x2 := x + w
    y2 := y + h

    DllCall("gdiplus\GdipAddPathArc",  "Ptr", pPath, "Float", x,      "Float", y,      "Float", d, "Float", d, "Float", 180, "Float", 90)
    DllCall("gdiplus\GdipAddPathLine", "Ptr", pPath, "Float", x+r,    "Float", y,      "Float", x2-r, "Float", y)
    DllCall("gdiplus\GdipAddPathArc",  "Ptr", pPath, "Float", x2-d,   "Float", y,      "Float", d, "Float", d, "Float", 270, "Float", 90)
    DllCall("gdiplus\GdipAddPathLine", "Ptr", pPath, "Float", x2,     "Float", y+r,    "Float", x2, "Float", y2-r)
    DllCall("gdiplus\GdipAddPathArc",  "Ptr", pPath, "Float", x2-d,   "Float", y2-d,   "Float", d, "Float", d, "Float", 0,   "Float", 90)
    DllCall("gdiplus\GdipAddPathLine", "Ptr", pPath, "Float", x2-r,   "Float", y2,     "Float", x+r, "Float", y2)
    DllCall("gdiplus\GdipAddPathArc",  "Ptr", pPath, "Float", x,      "Float", y2-d,   "Float", d, "Float", d, "Float", 90,  "Float", 90)
    DllCall("gdiplus\GdipAddPathLine", "Ptr", pPath, "Float", x,      "Float", y2-r,   "Float", x, "Float", y+r)

    DllCall("gdiplus\GdipClosePathFigure", "Ptr", pPath)
    return pPath
}