// Ribbon callbacks. First entry point executed by WPS on add-in load.

var TIDE_PANE_KEY = "tide_taskpane_id"
var TIDE_TAB_KEY = "tide_open_tab"

function OnAddinLoad(ribbonUI) {
    if (typeof (window.Application.ribbonUI) != "object") {
        window.Application.ribbonUI = ribbonUI
    }
    if (typeof (window.Application.Enum) != "object") {
        window.Application.Enum = WPS_Enum
    }
    return true
}

function tidePaneUrl() {
    return GetUrlPath() + "/ui/tide.html"
}

// Open the task pane, optionally jumping to one tab.
function TideOpenPane(tab) {
    if (tab) {
        window.Application.PluginStorage.setItem(TIDE_TAB_KEY, tab)
    }
    var paneId = window.Application.PluginStorage.getItem(TIDE_PANE_KEY)
    if (!paneId) {
        var pane = window.Application.CreateTaskPane(tidePaneUrl())
        window.Application.PluginStorage.setItem(TIDE_PANE_KEY, pane.ID)
        pane.Visible = true
        return
    }
    var existing = window.Application.GetTaskPane(paneId)
    if (existing) {
        existing.Visible = true
    } else {
        var fresh = window.Application.CreateTaskPane(tidePaneUrl())
        window.Application.PluginStorage.setItem(TIDE_PANE_KEY, fresh.ID)
        fresh.Visible = true
    }
}

function OnAction(control) {
    var id = control.Id
    switch (id) {
        case "btnOpenPane":
            TideOpenPane("matches")
            break
        case "btnScanSheet":
            window.Application.PluginStorage.setItem("tide_scan_request", String(Date.now()))
            TideOpenPane("matches")
            break
        case "btnHistory":
            TideOpenPane("history")
            break
        case "btnSync":
            TideOpenPane("sync")
            break
        default:
            break
    }
}
