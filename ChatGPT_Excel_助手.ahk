#Requires AutoHotkey v2.0
#SingleInstance Force
#Warn All, StdOut
#Include src\ChatGPTExcelCore.ahk
#Include src\MouseStatus.ahk

SendMode("Input")
SetWorkingDir(A_ScriptDir)

assistant := ChatGPTExcelAssistant()
if !assistant.Initialize()
    ExitApp

class ChatGPTExcelAssistant {
    __New() {
        this.enabled := true
        this.cleanedUp := false
        this.excel := 0
        this.workbook := 0
        this.platform := 0
        this.platformSelection := ""
        this.questionStart := 0
        this.answerStart := 0
        this.citationStart := 0
        this.rowOffset := 0
        this.citationColumnOffset := 0
        this.lastSentQuestion := ""
        this.lastSavedAnswer := ""
        this.closeContext := ""
        this.urlCopyVersion := 0
        this.status := MouseStatus(true)
        this.hotIfHandler := 0
        this.exitHandler := 0
    }

    Initialize() {
        platformKey := this.SelectPlatform()
        if platformKey = ""
            return false
        this.platform := GetPlatformConfig(platformKey)

        workbookPath := FileSelect(
            3,
            ,
            "選擇含有問題、回答與引用網址欄位的 Excel 活頁簿",
            "Excel 活頁簿 (*.xlsx; *.xlsm; *.xls)"
        )

        if workbookPath = ""
            return false

        try {
            this.excel := ComObject("Excel.Application")
            this.excel.Visible := true
            this.excel.DisplayAlerts := true
            this.workbook := this.excel.Workbooks.Open(workbookPath)
        } catch Error as err {
            MsgBox(
                "無法開啟 Excel 活頁簿。`n`n" . err.Message,
                "ChatGPT Excel AHK",
                "Iconx"
            )
            return false
        }

        this.questionStart := this.SelectSingleRange(
            "請在剛開啟的活頁簿中點選『問題起始格』，再按確定。",
            "選擇問題起始格"
        )
        if !IsObject(this.questionStart)
            return false

        this.answerStart := this.SelectSingleRange(
            "請在同一本活頁簿中點選『回答起始格』，再按確定。",
            "選擇回答起始格"
        )
        if !IsObject(this.answerStart)
            return false

        this.citationStart := this.SelectSingleRange(
            "請在同一本活頁簿中點選『引用網址起始格』，再按確定。",
            "選擇引用網址起始格"
        )
        if !IsObject(this.citationStart)
            return false

        this.RegisterHotkeys()
        this.exitHandler := ObjBindMethod(this, "HandleExit")
        OnExit(this.exitHandler)

        ranges := this.CurrentRanges()
        this.status.ShowSuccess(
            BuildStatus(
                "ready",
                this.CellLabel(ranges.question),
                this.CellLabel(ranges.answer),
                "",
                "",
                this.CellLabel(ranges.citation),
                "",
                this.platform.name
            ),
            2500
        )
        return true
    }

    SelectPlatform() {
        this.platformSelection := ""
        picker := Gui("+AlwaysOnTop -MinimizeBox", "選擇 AI 平台")
        picker.SetFont("s10", "Microsoft JhengHei UI")
        picker.AddText("w330", "請選擇本次要使用的網頁版 AI 平台：")
        gptButton := picker.AddButton("xm y+12 w160 h36 Default", "GPT（ChatGPT）")
        geminiButton := picker.AddButton("x+10 yp w160 h36", "Gemini")
        cancelButton := picker.AddButton("xm y+12 w330", "取消")

        gptButton.OnEvent(
            "Click",
            ObjBindMethod(this, "ChoosePlatform", picker, "gpt")
        )
        geminiButton.OnEvent(
            "Click",
            ObjBindMethod(this, "ChoosePlatform", picker, "gemini")
        )
        cancelButton.OnEvent("Click", ObjBindMethod(this, "CancelPlatform", picker))
        picker.OnEvent("Close", ObjBindMethod(this, "CancelPlatform", picker))

        picker.Show("AutoSize")
        WinWaitClose("ahk_id " . picker.Hwnd)
        return this.platformSelection
    }

    ChoosePlatform(picker, platformKey, *) {
        this.platformSelection := platformKey
        picker.Destroy()
    }

    CancelPlatform(picker, *) {
        this.platformSelection := ""
        picker.Destroy()
    }

    SelectSingleRange(prompt, title) {
        try selected := this.excel.InputBox(prompt, title, , , , , , 8)
        catch Error as err {
            MsgBox(
                "無法取得儲存格。`n`n" . err.Message,
                "ChatGPT Excel AHK",
                "Iconx"
            )
            return 0
        }

        if !IsObject(selected) {
            MsgBox("已取消選擇，腳本不會繼續。", "ChatGPT Excel AHK", "Iconi")
            return 0
        }

        try {
            if selected.Cells.CountLarge != 1 {
                MsgBox("請只選擇一個儲存格。", "ChatGPT Excel AHK", "Iconx")
                return 0
            }

            selectedWorkbookPath := selected.Parent.Parent.FullName
            if StrLower(selectedWorkbookPath) != StrLower(this.workbook.FullName) {
                MsgBox("所有起始格必須位於剛才選擇的活頁簿。", "ChatGPT Excel AHK", "Iconx")
                return 0
            }
        } catch Error as err {
            MsgBox(
                "選取範圍無法驗證。`n`n" . err.Message,
                "ChatGPT Excel AHK",
                "Iconx"
            )
            return 0
        }

        return selected
    }

    RegisterHotkeys() {
        Hotkey("F1", ObjBindMethod(this, "ToggleEnabled"))
        Hotkey("F2", ObjBindMethod(this, "ExitAssistant"))

        this.hotIfHandler := ObjBindMethod(this, "ActionsEnabled")
        HotIf(this.hotIfHandler)
        Hotkey("z", ObjBindMethod(this, "OpenSelectedPlatform"))
        Hotkey("x", ObjBindMethod(this, "SendCurrentQuestion"))
        Hotkey("c", ObjBindMethod(this, "CopyAndSaveAnswer"))
        Hotkey("s", ObjBindMethod(this, "SaveCurrentUrl"))
        Hotkey("v", ObjBindMethod(this, "DismissWarning"))
        Hotkey("a", ObjBindMethod(this, "CloseFocusedChromeTab"))
        HotIf()
    }

    ActionsEnabled(*) {
        return this.enabled
    }

    ToggleEnabled(*) {
        this.urlCopyVersion += 1
        this.enabled := !this.enabled
        if this.enabled {
            ranges := this.CurrentRanges()
            this.status.ShowInfo(
                BuildStatus("resumed", this.CellLabel(ranges.question)),
                false
            )
        } else {
            this.status.ShowInfo(BuildStatus("paused"), true)
        }
    }

    OpenSelectedPlatform(*) {
        this.urlCopyVersion += 1
        chromePath := this.FindChromePath()
        if chromePath = "" {
            this.status.ShowWarning(BuildStatus("chrome_missing"))
            return
        }

        quote := Chr(34)
        command := quote . chromePath . quote
            . " --incognito --new-window "
            . quote . this.platform.url . quote

        try Run(command)
        catch Error {
            this.status.ShowWarning(BuildStatus("chrome_launch_failed"))
            return
        }

        this.closeContext := ""
        this.ShowCurrentQuestionPreview()
    }

    SendCurrentQuestion(*) {
        this.urlCopyVersion += 1
        if !this.IsFocusedChrome() {
            this.status.ShowWarning(BuildStatus("wrong_window"))
            return
        }

        ranges := this.CurrentRanges()
        questionLabel := this.CellLabel(ranges.question)
        questionText := this.ReadCellText(ranges.question)
        if questionText = "" {
            this.status.ShowWarning(BuildStatus("empty_question", questionLabel))
            return
        }

        A_Clipboard := questionText
        if !ClipWait(1) {
            this.status.ShowWarning(BuildStatus("clipboard_failed"))
            return
        }

        ; Chrome 會把 Shift+Esc 攔截成「Chrome 工作管理員」，因此直接對目前焦點貼上。
        Sleep(Random(250, 500))
        Send("^v")
        Sleep(Random(180, 350))
        Send("{Enter}")
        this.lastSentQuestion := questionText
        this.closeContext := ""
        this.status.ShowSuccess(BuildStatus("sent", questionLabel), 2500)
    }

    SaveCurrentUrl(*) {
        this.urlCopyVersion += 1
        if !this.IsFocusedChrome() {
            this.status.ShowWarning(BuildStatus("wrong_window"))
            return
        }

        chromeHwnd := WinExist("A")
        copyVersion := this.urlCopyVersion
        this.status.ShowInfo(BuildStatus("url_copying"), true)
        copyResult := ReadSourceUrlWithRetry(
            ObjBindMethod(this, "CopyAddressBarAttempt", chromeHwnd, copyVersion)
        )
        if copyResult.issue = "cancelled"
            return
        if copyResult.issue != "" {
            this.status.ShowWarning(BuildStatus(copyResult.issue))
            return
        }
        sourceUrl := copyResult.url

        ranges := this.CurrentRanges()
        citationLabel := this.CellLabel(ranges.citation)
        if !this.ConfirmOverwrite(ranges.citation, citationLabel, "引用網址")
            return

        writeOk := false
        saveOk := false
        try {
            ranges.citation.NumberFormat := "@"
            ranges.citation.Value2 := sourceUrl
            writeOk := true
        } catch Error {
            this.status.ShowWarning(
                BuildStatus("url_write_failed", "", "", "", "", citationLabel)
            )
            return
        }

        try {
            this.workbook.Save()
            saveOk := true
        } catch Error {
            this.status.ShowWarning(BuildStatus("url_save_failed"))
            return
        }

        if !CanAdvance(true, writeOk, saveOk)
            return

        urlPreview := BuildPreview(sourceUrl, 40, 1)
        this.citationColumnOffset += 1
        this.closeContext := "source"
        nextRanges := this.CurrentRanges()
        this.status.ShowSuccess(
            BuildStatus(
                "url_saved",
                "",
                "",
                "",
                urlPreview,
                citationLabel,
                this.CellLabel(nextRanges.citation)
            ),
            2500
        )
    }

    CopyAddressBarAttempt(chromeHwnd, copyVersion, attempt) {
        this.EnsureUrlCopyTarget(chromeHwnd, copyVersion)
        SendEvent("^l")
        Sleep(Random(350, 500) + (attempt - 1) * 200)
        this.EnsureUrlCopyTarget(chromeHwnd, copyVersion)

        ; 每次重試都重新清空，避免把問題或上一次的剪貼簿誤當成這次結果。
        A_Clipboard := ""
        SendEvent("^c")
        hasText := ClipWait(2)
        this.EnsureUrlCopyTarget(chromeHwnd, copyVersion)
        return hasText ? A_Clipboard : ""
    }

    EnsureUrlCopyTarget(chromeHwnd, copyVersion) {
        ; 新動作取消舊的複製，不鎖住任何快捷鍵。
        if !this.enabled || this.urlCopyVersion != copyVersion
            throw TargetError("網址複製已由另一個動作取消")

        if !WinActive("ahk_id " . chromeHwnd) {
            this.status.ShowInfo(BuildStatus("url_copy_cancelled"))
            throw TargetError("網址複製的來源視窗已失去焦點")
        }
    }

    CopyAndSaveAnswer(*) {
        this.urlCopyVersion += 1
        if !this.IsFocusedChrome() {
            this.status.ShowWarning(BuildStatus("wrong_window"))
            return
        }

        ranges := this.CurrentRanges()
        answerText := A_Clipboard
        currentQuestion := this.ReadCellText(ranges.question)
        questionForCheck := this.lastSentQuestion != "" ? this.lastSentQuestion : currentQuestion
        clipboardIssue := AnswerClipboardIssue(
            answerText,
            questionForCheck,
            this.lastSavedAnswer
        )
        switch clipboardIssue {
            case "empty":
                this.status.ShowWarning(BuildStatus("copy_failed"))
                return
            case "url":
                this.status.ShowWarning(BuildStatus("answer_is_url"))
                return
            case "question":
                this.status.ShowWarning(BuildStatus("answer_is_question"))
                return
            case "duplicate":
                this.status.ShowWarning(BuildStatus("answer_duplicate"))
                return
        }

        answerLabel := this.CellLabel(ranges.answer)
        if !this.ConfirmOverwrite(ranges.answer, answerLabel, "回答")
            return

        copyOk := true
        writeOk := false
        saveOk := false

        try {
            ranges.answer.NumberFormat := "@"
            ranges.answer.Value2 := answerText
            writeOk := true
        } catch Error {
            this.status.ShowWarning(BuildStatus("write_failed", "", answerLabel))
            return
        }

        try {
            this.workbook.Save()
            saveOk := true
        } catch Error {
            this.status.ShowWarning(BuildStatus("save_failed"))
            return
        }

        if !CanAdvance(copyOk, writeOk, saveOk)
            return

        answerPreview := BuildPreview(answerText, 20, 2)
        this.lastSavedAnswer := answerText
        this.closeContext := "answer"
        this.lastSentQuestion := ""
        this.rowOffset += 1
        this.citationColumnOffset := 0
        nextRanges := this.CurrentRanges()
        this.status.ShowSuccess(
            BuildStatus(
                "saved",
                "",
                answerLabel,
                this.CellLabel(nextRanges.question),
                answerPreview,
                this.CellLabel(nextRanges.citation)
            ),
            2500
        )
    }

    DismissWarning(*) {
        if this.status.HasWarning()
            this.status.ClearWarning()
    }

    CloseFocusedChromeTab(*) {
        this.urlCopyVersion += 1
        if !this.IsFocusedChrome() {
            this.status.ShowWarning(BuildStatus("wrong_window"))
            return
        }

        Send("^w")
        statusKind := this.closeContext = "source" ? "closed_source"
            : this.closeContext = "answer" ? "closed_answer" : "closed"
        this.closeContext := ""
        this.status.ShowSuccess(BuildStatus(statusKind), 2500)
    }

    ShowCurrentQuestionPreview() {
        ranges := this.CurrentRanges()
        questionLabel := this.CellLabel(ranges.question)
        questionText := this.ReadCellText(ranges.question)

        if questionText = "" {
            this.status.ShowWarning(BuildStatus("empty_question", questionLabel))
            return
        }

        preview := BuildPreview(questionText, 20, 2)
        this.status.ShowInfo(
            BuildStatus(
                "preview",
                questionLabel,
                "",
                "",
                preview,
                "",
                "",
                this.platform.name
            ),
            true
        )
    }

    CurrentRanges() {
        return {
            question: this.questionStart.Offset(this.rowOffset, 0),
            answer: this.answerStart.Offset(this.rowOffset, 0),
            citation: this.citationStart.Offset(
                this.rowOffset,
                this.citationColumnOffset
            )
        }
    }

    RangeHasContent(range) {
        try {
            if range.HasFormula
                return true
            value := range.Value2
        } catch Error {
            return true
        }

        if IsObject(value)
            return true
        return value . "" != ""
    }

    ConfirmOverwrite(range, cellLabel, dataName) {
        if !this.RangeHasContent(range)
            return true

        choice := MsgBox(
            cellLabel . " 已有資料。`n`n是否以目前的" . dataName . "覆蓋？",
            "確認覆蓋",
            "YesNo Icon!"
        )
        if choice = "Yes"
            return true

        this.status.ShowInfo(
            BuildStatus("overwrite_cancelled", "", "", "", "", cellLabel),
            false
        )
        return false
    }

    CellLabel(range) {
        return range.Parent.Name . "!" . range.Address(false, false)
    }

    ReadCellText(range) {
        try value := range.Value2
        catch Error
            return ""

        if IsObject(value)
            return ""
        return Trim(value . "")
    }

    IsFocusedChrome() {
        try return IsChromeProcess(WinGetProcessName("A"))
        catch Error
            return false
    }

    FindChromePath() {
        registryKeys := [
            "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
            "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
            "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"
        ]

        for registryKey in registryKeys {
            try {
                path := RegRead(registryKey)
                if path != "" && FileExist(path)
                    return path
            }
        }

        candidates := [
            A_ProgramFiles . "\Google\Chrome\Application\chrome.exe",
            EnvGet("ProgramFiles(x86)") . "\Google\Chrome\Application\chrome.exe",
            EnvGet("LOCALAPPDATA") . "\Google\Chrome\Application\chrome.exe"
        ]

        for path in candidates {
            if path != "" && FileExist(path)
                return path
        }
        return ""
    }

    ExitAssistant(*) {
        this.status.ClearWarning()
        this.status.ShowSuccess(BuildStatus("ended"), 800)
        Sleep(800)
        this.Cleanup()
        ExitApp
    }

    HandleExit(exitReason, exitCode) {
        this.Cleanup()
    }

    Cleanup() {
        if this.cleanedUp
            return
        this.cleanedUp := true

        this.status.Shutdown()
        this.questionStart := 0
        this.answerStart := 0
        this.citationStart := 0
        this.platform := 0
        this.workbook := 0
        this.excel := 0
    }
}
