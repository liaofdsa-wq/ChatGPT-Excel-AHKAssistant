#Requires AutoHotkey v2.0
#SingleInstance Force
#Warn All, StdOut
#Include ..\src\ChatGPTExcelCore.ahk
#Include ..\src\MouseStatus.ahk

failures := []

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

; 會抓到：提示文案接錯儲存格，讓使用者誤判目前列。
AssertEqual(
    BuildStatus("ready", "測試!A2", "測試!B2"),
    "準備完成｜問題 測試!A2 → 回答 測試!B2",
    "準備完成提示"
)
AssertEqual(
    BuildStatus("preview", "測試!A2", , , "請列出三種顏色"),
    "待貼入｜測試!A2`n問題：請列出三種顏色",
    "待貼問題預覽"
)
AssertEqual(
    BuildStatus("saved", , "測試!B2", "測試!A3"),
    "已儲存回答 測試!B2｜下一題 測試!A3",
    "成功寫入後提示下一題"
)
AssertEqual(
    BuildStatus("saved", , "測試!B2", "測試!A3", "回答摘要"),
    "已儲存回答 測試!B2｜下一題 測試!A3`n回答：回答摘要",
    "成功寫入後顯示回答摘要"
)
AssertEqual(
    BuildStatus("copy_failed"),
    "剪貼簿沒有回答｜先手動複製，按 V 清除後再按 C",
    "剪貼簿沒有回答時提示手動複製"
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

; 會抓到：紅色警告被成功訊息蓋掉，或 V 無法解除警告。
status := MouseStatus(false)
AssertTrue(status.ShowWarning("錯誤"), "警告可以顯示")
AssertTrue(status.HasWarning(), "顯示警告後進入鎖定")
AssertFalse(status.ShowSuccess("不應顯示"), "成功訊息不能蓋掉警告")
status.ClearWarning()
AssertFalse(status.HasWarning(), "清除警告後解除鎖定")
AssertTrue(status.ShowSuccess("完成", 2500), "清除後可以顯示成功訊息")
status.Shutdown()

if failures.Length {
    message := "測試失敗：" . failures.Length . " 項`n`n"
    for failure in failures
        message .= failure . "`n`n"
    MsgBox message, "ChatGPT Excel AHK 測試", "Iconx"
    ExitApp 1
}

MsgBox "全部核心測試通過。", "ChatGPT Excel AHK 測試", "Iconi"
ExitApp 0
