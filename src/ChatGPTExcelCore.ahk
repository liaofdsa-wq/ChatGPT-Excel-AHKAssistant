#Requires AutoHotkey v2.0

BuildPreview(text, maxChars := 120, maxLines := 4) {
    if (maxChars < 1 || maxLines < 1)
        return ""

    normalized := StrReplace(text . "", "`r`n", "`n")
    normalized := StrReplace(normalized, "`r", "`n")
    sourceLines := StrSplit(normalized, "`n")
    visibleLines := []
    lineLimit := Min(sourceLines.Length, maxLines)

    Loop lineLimit
        visibleLines.Push(sourceLines[A_Index])

    preview := ""
    for index, line in visibleLines
        preview .= (index > 1 ? "`n" : "") . line

    wasTrimmed := sourceLines.Length > maxLines
    if (StrLen(preview) > maxChars) {
        preview := SubStr(preview, 1, maxChars)
        wasTrimmed := true
    }

    preview := RTrim(preview, " `t`r`n")
    return wasTrimmed ? preview . "…" : preview
}

BuildStatus(kind, questionLabel := "", answerLabel := "", nextQuestionLabel := "", preview := "") {
    switch kind {
        case "ready":
            return "準備完成｜問題 " . questionLabel . " → 回答 " . answerLabel
        case "preview":
            return "待貼入｜" . questionLabel . (preview = "" ? "" : "`n問題：" . preview)
        case "sent":
            return "已送出問題 " . questionLabel . "｜等待回答"
        case "empty_question":
            return "送出失敗｜" . questionLabel . " 沒有內容"
        case "wrong_window":
            return "操作失敗｜請先對焦 ChatGPT 頁面"
        case "saved":
            return "已儲存回答 " . answerLabel . "｜下一題 " . nextQuestionLabel
                . (preview = "" ? "" : "`n回答：" . preview)
        case "copy_failed":
            return "剪貼簿沒有回答｜先手動複製，按 V 清除後再按 C"
        case "write_failed":
            return "寫入失敗｜" . answerLabel . " 未寫入、位置沒有移動"
        case "save_failed":
            return "儲存失敗｜位置沒有移動，按 V 清除後再按 C 重試"
        case "closed":
            return "已關閉目前的 Chrome 視窗"
        case "paused":
            return "快捷鍵已暫停｜按 F1 恢復"
        case "resumed":
            return "快捷鍵已恢復｜目前問題 " . questionLabel
        case "ended":
            return "AHK 已結束"
        case "chrome_missing":
            return "開啟失敗｜找不到 Google Chrome"
        case "chrome_launch_failed":
            return "開啟失敗｜無法啟動 Google Chrome"
        case "clipboard_failed":
            return "貼上失敗｜無法寫入剪貼簿"
        default:
            throw ValueError("未知的狀態種類：" . kind)
    }
}

CanAdvance(copyOk, writeOk, saveOk) {
    return !!copyOk && !!writeOk && !!saveOk
}

IsChromeProcess(processName) {
    return StrLower(Trim(processName . "")) = "chrome.exe"
}
