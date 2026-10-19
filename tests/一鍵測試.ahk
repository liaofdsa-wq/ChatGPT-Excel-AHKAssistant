#Requires AutoHotkey v2.0
#SingleInstance Force
#Warn All, StdOut
#Include ..\src\ChatGPTExcelCore.ahk
#Include ..\src\MouseStatus.ahk

failures := []
cliMode := A_Args.Length > 0 && A_Args[1] = "--cli"

AssertEqual(actual, expected, testName) {
    global failures
    if (actual !== expected) {
        failures.Push(testName . "`n  預期：" . expected . "`n  實際：" . actual)
    }
}

AssertTrue(actual, testName) {
    AssertEqual(actual ? true : false, true, testName)
}

AssertFalse(actual, testName) {
    AssertEqual(actual ? true : false, false, testName)
}

RepeatText(text, count) {
    output := ""
    Loop count
        output .= text
    return output
}

class UrlCopySequence {
    __New(results) {
        this.results := results
        this.calls := 0
    }

    Read(*) {
        this.calls += 1
        value := this.results[this.calls]
        if IsObject(value)
            throw value
        return value
    }
}

; 會抓到：預覽裁切被移除或邊界多切／少切一個字。
AssertEqual(BuildPreview("短問題"), "短問題", "短問題不加省略號")
AssertEqual(
    BuildPreview(RepeatText("甲", 121)),
    RepeatText("甲", 120) . "…",
    "超過 120 字時裁切並加省略號"
)
AssertEqual(
    BuildPreview("第一行`n第二行`n第三行`n第四行`n第五行"),
    "第一行`n第二行`n第三行`n第四行…",
    "預覽最多四行"
)
AssertEqual(
    BuildPreview("第一行`r`n第二行"),
    "第一行`n第二行",
    "Windows 換行正規化"
)
AssertEqual(
    BuildPreview("一二三四五六七八九十一二三四五六七八九十一", 20, 2),
    "一二三四五六七八九十一二三四五六七八九十…",
    "提示摘要最多二十字"
)

; 會抓到：平台代碼對到錯誤網址，導致 Z 開錯網站。
gptPlatform := GetPlatformConfig("gpt")
geminiPlatform := GetPlatformConfig("gemini")
AssertEqual(gptPlatform.name, "GPT（ChatGPT）", "GPT 平台名稱")
AssertEqual(gptPlatform.url, "https://chatgpt.com/", "GPT 平台網址")
AssertEqual(geminiPlatform.name, "Gemini", "Gemini 平台名稱")
AssertEqual(geminiPlatform.url, "https://gemini.google.com/app", "Gemini 平台網址")

; 會抓到：S 接受非網頁網址，或把 AI 對話頁誤當成引用來源。
AssertTrue(IsValidSourceUrl("https://example.com/article?id=1"), "接受 HTTPS 引用網址")
AssertTrue(IsValidSourceUrl("http://example.com/source"), "接受 HTTP 引用網址")
AssertTrue(IsValidSourceUrl("https://openai.com/policies/usage-policies/"), "允許一般 OpenAI 網頁")
AssertFalse(IsValidSourceUrl("chrome://settings"), "拒絕瀏覽器內部網址")
AssertFalse(IsValidSourceUrl("ftp://example.com/file"), "拒絕非 HTTP 網址")
AssertFalse(IsValidSourceUrl("https://chatgpt.com/c/123"), "拒絕 ChatGPT 對話頁")
AssertFalse(IsValidSourceUrl("https://gemini.google.com/app/123"), "拒絕 Gemini 對話頁")
AssertTrue(IsValidSourceUrl("https://evilchatgpt.com/article"), "不誤擋名稱相似網域")

; 會抓到：載入時第一次複製沒資料就放棄、重試超出上限，或切換視窗後仍儲存。
if IsSet(ReadSourceUrlWithRetry) {
    delayedUrl := UrlCopySequence(["", "about:blank", "https://support.microsoft.com/excel"])
    result := ReadSourceUrlWithRetry(ObjBindMethod(delayedUrl, "Read"))
    AssertEqual(result.url, "https://support.microsoft.com/excel", "等到可複製網址後才回傳成功")
    AssertEqual(result.issue, "", "載入後複製成功不留下舊錯誤")

    readyUrl := UrlCopySequence(["https://example.com/source"])
    result := ReadSourceUrlWithRetry(ObjBindMethod(readyUrl, "Read"))
    AssertEqual(result.url, "https://example.com/source", "首次複製成功即可使用")
    AssertEqual(readyUrl.calls, 1, "成功後不再複製，避免重複處理")

    unavailableUrl := UrlCopySequence(["", "", "", "https://example.com/too-late"])
    result := ReadSourceUrlWithRetry(ObjBindMethod(unavailableUrl, "Read"))
    AssertEqual(result.issue, "url_copy_failed", "多次取不到剪貼簿時回報複製失敗")
    AssertEqual(unavailableUrl.calls, 3, "複製失敗有重試上限")
    AssertEqual(result.url, "", "複製失敗不提供可寫入的網址")

    invalidUrl := UrlCopySequence(["about:blank", "about:blank", "about:blank"])
    result := ReadSourceUrlWithRetry(ObjBindMethod(invalidUrl, "Read"))
    AssertEqual(result.issue, "invalid_url", "持續為空白分頁時不誤存網址")
    AssertEqual(result.url, "", "無效網址不進入 Excel 寫入流程")

    busyClipboard := UrlCopySequence([Error("clipboard busy"), "https://example.com/recovered"])
    result := ReadSourceUrlWithRetry(ObjBindMethod(busyClipboard, "Read"))
    AssertEqual(result.url, "https://example.com/recovered", "剪貼簿短暫忙碌後可重試成功")

    cancelledCopy := UrlCopySequence([TargetError("target changed"), "https://example.com/other-tab"])
    result := ReadSourceUrlWithRetry(ObjBindMethod(cancelledCopy, "Read"))
    AssertEqual(result.issue, "cancelled", "切換目標時取消原本的複製")
    AssertEqual(result.url, "", "取消後不把其他分頁網址交給 Excel")
    AssertEqual(cancelledCopy.calls, 1, "取消後不繼續重試")
} else {
    failures.Push("載入中的網址複製缺少有限次重試功能")
}

; 會抓到：C 把問題、網址或未重新複製的上一筆回答寫入 Excel。
AssertEqual(AnswerClipboardIssue("", "問題", "上一筆"), "empty", "拒絕空回答")
AssertEqual(AnswerClipboardIssue("問題", "問題", "上一筆"), "question", "拒絕把目前問題存成回答")
AssertEqual(AnswerClipboardIssue("上一筆", "問題", "上一筆"), "duplicate", "拒絕重複保存上一筆回答")
AssertEqual(AnswerClipboardIssue("https://example.com", "問題", "上一筆"), "url", "拒絕把網址存成回答")
AssertEqual(AnswerClipboardIssue("完整回答`n第二行", "問題", "上一筆"), "", "接受一般多行回答")

; 會抓到：提示文案接錯儲存格，讓使用者誤判目前列。
AssertEqual(
    BuildStatus("ready", "測試!A2", "測試!B2", , , "測試!C2", , "GPT（ChatGPT）"),
    "準備完成｜GPT（ChatGPT）｜問題 測試!A2 → 回答 測試!B2｜網址 測試!C2`nZ 開啟 AI",
    "準備完成提示"
)
AssertEqual(
    BuildStatus("preview", "測試!A2", , , "請列出三種顏色"),
    "待貼入｜測試!A2`n問題：請列出三種顏色`nX 貼上並送出問題",
    "待貼問題預覽"
)
AssertEqual(
    BuildStatus("saved", , "測試!B2", "測試!A3"),
    "已儲存回答 測試!B2｜下一題 測試!A3`nA 關閉 AI 分頁",
    "成功寫入後提示下一題"
)
AssertEqual(
    BuildStatus("saved", , "測試!B2", "測試!A3", "回答摘要", "測試!C3"),
    "已儲存回答 測試!B2｜下一題 測試!A3｜網址 測試!C3`n回答：回答摘要`nA 關閉 AI 分頁",
    "成功寫入後顯示回答摘要與新列網址起點"
)
AssertEqual(
    BuildStatus("copy_failed"),
    "剪貼簿沒有回答｜先手動複製，再按 C",
    "剪貼簿沒有回答時提示手動複製"
)
AssertEqual(
    BuildStatus("url_saved", , , , "https://example.com/a", "測試!C2", "測試!D2"),
    "已儲存網址 測試!C2｜下一格 測試!D2`n網址：https://example.com/a`nA 關閉來源分頁",
    "網址成功後提示下一格"
)

AssertEqual(
    BuildStatus("closed_source"),
    "已關閉目前分頁`nS 開啟其他來源後存網址／C 複製回答後儲存",
    "存網址後關分頁提示繼續收來源或存回答"
)
AssertEqual(
    BuildStatus("closed_answer"),
    "已關閉目前分頁`nZ 開啟下一題",
    "存回答後關分頁提示開啟下一題"
)

; 會抓到：複製、寫入或儲存任一步驟失敗仍錯誤跳列。
AssertTrue(CanAdvance(true, true, true), "三步成功才可前進")
AssertFalse(CanAdvance(false, true, true), "複製失敗不可前進")
AssertFalse(CanAdvance(true, false, true), "寫入失敗不可前進")
AssertFalse(CanAdvance(true, true, false), "儲存失敗不可前進")

; 會抓到：對非 Chrome 視窗誤送快捷鍵或誤關視窗。
AssertTrue(IsChromeProcess("chrome.exe"), "辨識 Chrome")
AssertTrue(IsChromeProcess("CHROME.EXE"), "Chrome 名稱不分大小寫")
AssertFalse(IsChromeProcess("msedge.exe"), "拒絕非 Chrome 程序")

; 會抓到：舊警告阻擋新提示，或 V 無法清除目前警告。
status := MouseStatus(false)
AssertTrue(status.ShowWarning("錯誤"), "警告可以顯示")
AssertTrue(status.HasWarning(), "記錄目前顯示的警告")
AssertTrue(status.ShowSuccess("已儲存網址", 2500), "成功提示可直接取代舊警告")
AssertFalse(status.HasWarning(), "新動作成功後不保留舊警告")
AssertEqual(status.lastText, "已儲存網址", "顯示最新動作結果")
status.ShowWarning("另一個錯誤")
AssertTrue(status.ShowInfo("快捷鍵已暫停", true), "暫停提示可直接取代舊警告")
AssertFalse(status.HasWarning(), "暫停提示不被當成警告")
status.ShowWarning("可清除的錯誤")
status.ClearWarning()
AssertFalse(status.HasWarning(), "V 可清除目前警告")
AssertFalse(status.visible, "清除後隱藏提示")
AssertTrue(status.ShowSuccess("完成", 2500), "清除後可以顯示成功訊息")
status.Shutdown()

if failures.Length {
    message := "測試失敗：" . failures.Length . " 項`n`n"
    for failure in failures
        message .= failure . "`n`n"
    if cliMode
        FileAppend("CORE_TESTS_FAIL count=" . failures.Length . "`n", "*")
    else
        MsgBox message, "ChatGPT Excel AHK 測試", "Iconx"
    ExitApp 1
}

if cliMode
    FileAppend("CORE_TESTS_PASS`n", "*")
else
    MsgBox "全部核心測試通過。", "ChatGPT Excel AHK 測試", "Iconi"
ExitApp 0
