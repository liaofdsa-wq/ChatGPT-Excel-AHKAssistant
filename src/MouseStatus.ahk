#Requires AutoHotkey v2.0

class MouseStatus {
    __New(renderTooltip := true) {
        this.tooltipEnabled := renderTooltip
        this.warningVisible := false
        this.visible := false
        this.lastText := ""
        this.lastKind := ""
        this.hasMousePosition := false
        this.lastMouseX := 0
        this.lastMouseY := 0
        this.followTimer := ObjBindMethod(this, "FollowMouse")
        this.hideTimer := ObjBindMethod(this, "HideExpired")

        if this.tooltipEnabled
            SetTimer(this.followTimer, 25)
    }

    ; 保留既有參數以相容呼叫；提示一律常駐，由下一個動作更新。
    ShowSuccess(text, ttlMs := 0) {
        return this.ShowMessage(text, "success", 0)
    }

    ShowInfo(text, sticky := true) {
        return this.ShowMessage(text, "info", 0)
    }

    ShowWarning(text) {
        return this.ShowMessage(text, "warning", 0)
    }

    ClearWarning() {
        this.warningVisible := false
        this.lastText := ""
        this.lastKind := ""
        this.visible := false

        if this.tooltipEnabled {
            SetTimer(this.hideTimer, 0)
            ToolTip(, , , 1)
        }
    }

    HasWarning() {
        return this.warningVisible
    }

    ShowMessage(text, kind, ttlMs := 0) {
        this.warningVisible := kind = "warning"
        this.lastText := text
        this.lastKind := kind
        this.visible := true

        if !this.tooltipEnabled
            return true

        this.RenderTooltip(true)

        SetTimer(this.hideTimer, 0)
        if ttlMs > 0
            SetTimer(this.hideTimer, -ttlMs)
        return true
    }

    FollowMouse() {
        this.RenderTooltip(false)
    }

    RenderTooltip(force) {
        if !this.tooltipEnabled || !this.visible
            return

        CoordMode("Mouse", "Screen")
        CoordMode("ToolTip", "Screen")
        MouseGetPos(&mouseX, &mouseY)
        if !force && this.hasMousePosition
            && mouseX = this.lastMouseX && mouseY = this.lastMouseY
            return

        this.hasMousePosition := true
        this.lastMouseX := mouseX
        this.lastMouseY := mouseY
        ToolTip(this.lastText, mouseX + 18, mouseY + 24, 1)
    }

    HideExpired() {
        if !this.tooltipEnabled || this.warningVisible
            return

        this.visible := false
        ToolTip(, , , 1)
    }

    Shutdown() {
        this.warningVisible := false
        this.visible := false

        if this.tooltipEnabled {
            SetTimer(this.followTimer, 0)
            SetTimer(this.hideTimer, 0)
            ToolTip(, , , 1)
        }
    }
}
