#Requires AutoHotkey v2.0

class MouseStatus {
    __New(renderTooltip := true) {
        this.tooltipEnabled := renderTooltip
        this.warningLocked := false
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

    ShowSuccess(text, ttlMs := 2500) {
        if this.warningLocked
            return false
        return this.ShowMessage(text, "success", ttlMs)
    }

    ShowInfo(text, sticky := false) {
        if this.warningLocked
            return false
        return this.ShowMessage(text, "info", sticky ? 0 : 2500)
    }

    ShowWarning(text) {
        this.warningLocked := true
        return this.ShowMessage(text, "warning", 0, true)
    }

    ClearWarning() {
        this.warningLocked := false
        this.lastText := ""
        this.lastKind := ""
        this.visible := false

        if this.tooltipEnabled {
            SetTimer(this.hideTimer, 0)
            ToolTip(, , , 1)
        }
    }

    HasWarning() {
        return this.warningLocked
    }

    ShowMessage(text, kind, ttlMs := 0, force := false) {
        if this.warningLocked && kind != "warning" && !force
            return false

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
        if !this.tooltipEnabled || this.warningLocked
            return

        this.visible := false
        ToolTip(, , , 1)
    }

    Shutdown() {
        this.warningLocked := false
        this.visible := false

        if this.tooltipEnabled {
            SetTimer(this.followTimer, 0)
            SetTimer(this.hideTimer, 0)
            ToolTip(, , , 1)
        }
    }
}
