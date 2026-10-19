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

GetPlatformConfig(platformKey) {
    switch StrLower(Trim(platformKey . "")) {
        case "gpt":
            return {
                key: "gpt",
                name: "GPT（ChatGPT）",
                url: "https://chatgpt.com/"
            }
        case "gemini":
            return {
                key: "gemini",
                name: "Gemini",
                url: "https://gemini.google.com/app"
            }
        default:
            throw ValueError("未知的 AI 平台：" . platformKey)
    }
}

IsValidSourceUrl(url) {
    value := Trim(url . "")
    if value = "" || RegExMatch(value, "\s")
        return false

    if !RegExMatch(
        value,
        "i)^https?://([^/?#:@\s]+)(?::\d+)?(?:[/?#]|$)",
        &match
    )
        return false

    host := StrLower(match[1])
    blockedDomains := ["chatgpt.com", "chat.openai.com", "gemini.google.com"]
    for domain in blockedDomains {
        if host = domain
            return false

        suffix := "." . domain
        if StrLen(host) > StrLen(suffix)
            && SubStr(host, StrLen(host) - StrLen(suffix) + 1) = suffix
            return false
    }
    return true
}

ReadSourceUrlWithRetry(readAttempt, maxAttempts := 3) {
    candidate := ""
    Loop maxAttempts {
        try candidate := Trim(readAttempt.Call(A_Index) . "")
        catch TargetError {
            return {url: "", issue: "cancelled"}
        } catch Error {
            candidate := ""
            continue
        }

        if IsValidSourceUrl(candidate)
            return {url: candidate, issue: ""}
    }

    return {
        url: "",
        issue: candidate = "" ? "url_copy_failed" : "invalid_url"
    }
}

AnswerClipboardIssue(text, currentQuestion := "", lastSavedAnswer := "") {
    answer := Trim(text . "")
    if answer = ""
        return "empty"
    if RegExMatch(answer, "i)^https?://\S+$")
        return "url"

    question := Trim(currentQuestion . "")
    if question != "" && StrCompare(answer, question, true) = 0
        return "question"

    previousAnswer := Trim(lastSavedAnswer . "")
    if previousAnswer != "" && StrCompare(answer, previousAnswer, true) = 0
        return "duplicate"
    return ""
}

BuildStatus(
    kind,
    questionLabel := "",
    answerLabel := "",
    nextQuestionLabel := "",
    preview := "",
    citationLabel := "",
    nextCitationLabel := "",
    platformName := ""
) {
    switch kind {
        case "ready":
            return "準備完成"
                . (platformName = "" ? "" : "｜" . platformName)
                . "｜問題 " . questionLabel . " → 回答 " . answerLabel
                . (citationLabel = "" ? "" : "｜網址 " . citationLabel)
                . "`nZ 開啟 AI"
        case "preview":
            return "待貼入"
                . (platformName = "" ? "" : " " . platformName)
                . "｜" . questionLabel
                . (preview = "" ? "" : "`n問題：" . preview)
                . "`nX 貼上並送出問題"
        case "sent":
            return "已送出問題 " . questionLabel . "｜等待回答"
                . "`nS 開啟來源後存網址／C 複製回答後儲存"
        case "empty_question":
            return "送出失敗｜" . questionLabel . " 沒有內容"
        case "wrong_window":
            return "操作失敗｜請先對焦 Chrome 的 AI／來源頁面"
        case "saved":
            return "已儲存回答 " . answerLabel . "｜下一題 " . nextQuestionLabel
                . (citationLabel = "" ? "" : "｜網址 " . citationLabel)
                . (preview = "" ? "" : "`n回答：" . preview)
                . "`nA 關閉 AI 分頁"
        case "copy_failed":
            return "剪貼簿沒有回答｜先手動複製，再按 C"
        case "answer_is_url":
            return "儲存失敗｜剪貼簿是網址，不是完整回答"
        case "answer_is_question":
            return "儲存失敗｜剪貼簿仍是目前問題，請重新複製回答"
        case "answer_duplicate":
            return "儲存失敗｜與上一筆回答相同，請確認已重新複製"
        case "write_failed":
            return "寫入失敗｜" . answerLabel . " 未寫入、位置沒有移動"
        case "save_failed":
            return "儲存失敗｜位置沒有移動，可直接按 C 重試"
        case "url_copying":
            return "正在複製網址｜等待網址列回應"
        case "url_copy_cancelled":
            return "網址複製已取消｜已離開原本視窗`nS 回到來源分頁後重試"
        case "url_copy_failed":
            return "網址取得失敗｜重試後仍無法複製，位置沒有移動"
                . "`nS 來源分頁可操作後重試"
        case "invalid_url":
            return "網址無效｜需為 HTTP／HTTPS，且不能是 AI 對話頁"
        case "url_write_failed":
            return "網址寫入失敗｜" . citationLabel . " 未寫入、位置沒有移動"
        case "url_save_failed":
            return "網址儲存失敗｜位置沒有移動，可直接按 S 重試"
        case "url_saved":
            return "已儲存網址 " . citationLabel . "｜下一格 " . nextCitationLabel
                . (preview = "" ? "" : "`n網址：" . preview)
                . "`nA 關閉來源分頁"
        case "overwrite_cancelled":
            return "已取消覆蓋｜" . citationLabel . " 保留原有資料"
        case "closed":
            return "已關閉目前分頁"
                . "`nZ 開啟 AI／S 存來源網址／C 複製回答後儲存"
        case "closed_source":
            return "已關閉目前分頁"
                . "`nS 開啟其他來源後存網址／C 複製回答後儲存"
        case "closed_answer":
            return "已關閉目前分頁"
                . "`nZ 開啟下一題"
        case "paused":
            return "快捷鍵已暫停`nF1 恢復快捷鍵"
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
