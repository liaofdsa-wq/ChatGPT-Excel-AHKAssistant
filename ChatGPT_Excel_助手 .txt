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
        this.questionStart := 0
        this.answerStart := 0
        this.rowOffset := 0
        this.status := MouseStatus(true)
        this.hotIfHandler := 0
        this.exitHandler := 0
    }

    Initialize() {
        workbookPath := FileSelect(
            3,
            ,
            "選擇含有問題與回答欄位的 Excel 活頁簿",
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

        this.RegisterHotkeys()
        this.exitHandler := ObjBindMethod(this, "HandleExit")
        OnExit(this.exitHandler)

        ranges := this.CurrentRanges()
        this.status.ShowSuccess(
            BuildStatus(
                "ready",
                this.CellLabel(ranges.question),
                this.CellLabel(ranges.answer)
            ),
            2500
        )
        return true
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
                MsgBox("問題格與回答格必須位於剛才選擇的活頁簿。", "ChatGPT Excel AHK", "Iconx")
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
        Hotkey("z", ObjBindMethod(this, "OpenIncognitoChatGPT"))
        Hotkey("x", ObjBindMethod(this, "SendCurrentQuestion"))
        Hotkey("c", ObjBindMethod(this, "CopyAndSaveAnswer"))
        Hotkey("v", ObjBindMethod(this, "DismissWarning"))
        Hotkey("a", ObjBindMethod(this, "CloseFocusedChromeWindow"))
        HotIf()
    }

    ActionsEnabled(*) {
        return this.enabled
    }

    ToggleEnabled(*) {
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

    OpenIncognitoChatGPT(*) {
        if !this.CanRunAction()
            return

        chromePath := this.FindChromePath()
        if chromePath = "" {
            this.status.ShowWarning(BuildStatus("chrome_missing"))
            return
        }

        quote := Chr(34)
        command := quote . chromePath . quote
            . " --incognito --new-window "
            . quote . "https://chatgpt.com/" . quote

        try Run(command)
        catch Error {
            this.status.ShowWarning(BuildStatus("chrome_launch_failed"))
            return
        }

        this.ShowCurrentQuestionPreview()
    }

    SendCurrentQuestion(*) {
        if !this.CanRunAction()
            return
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
        this.status.ShowSuccess(BuildStatus("sent", questionLabel), 2500)
    }

    CopyAndSaveAnswer(*) {
        if !this.CanRunAction()
            return
        if !this.IsFocusedChrome() {
            this.status.ShowWarning(BuildStatus("wrong_window"))
            return
        }

        answerText := A_Clipboard
        if Trim(answerText) = "" {
            this.status.ShowWarning(BuildStatus("copy_failed"))
            return
        }

        ranges := this.CurrentRanges()
        answerLabel := this.CellLabel(ranges.answer)
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
        this.rowOffset += 1
        nextRanges := this.CurrentRanges()
        this.status.ShowSuccess(
            BuildStatus(
                "saved",
                "",
                answerLabel,
                this.CellLabel(nextRanges.question),
                answerPreview
            ),
            2500
        )
    }

    DismissWarning(*) {
        if this.status.HasWarning()
            this.status.ClearWarning()
    }

    CloseFocusedChromeWindow(*) {
        if !this.CanRunAction()
            return
        if !this.IsFocusedChrome() {
            this.status.ShowWarning(BuildStatus("wrong_window"))
            return
        }

        Send("!{F4}")
        this.status.ShowSuccess(BuildStatus("closed"), 2500)
    }

    CanRunAction() {
        return !this.status.HasWarning()
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
            BuildStatus("preview", questionLabel, "", "", preview),
            true
        )
    }

    CurrentRanges() {
        return {
            question: this.questionStart.Offset(this.rowOffset, 0),
            answer: this.answerStart.Offset(this.rowOffset, 0)
        }
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
        this.workbook := 0
        this.excel := 0
    }
}
