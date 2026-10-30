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
// Replaced by pack.sh on every release: a new stamp forces a fresh pane URL,
// so updates load without restarting WPS.
var TIDE_PANE_BUILD = "202610301515"

function TideOpenPane(tab) {
    if (tab) {
        window.Application.PluginStorage.setItem(TIDE_TAB_KEY, tab)
    }
    var storage = window.Application.PluginStorage
    var paneId = storage.getItem(TIDE_PANE_KEY)
    var storedBuild = storage.getItem("tide_pane_build")
    if (paneId && storedBuild === TIDE_PANE_BUILD) {
        var existing = window.Application.GetTaskPane(paneId)
        if (existing) {
            existing.Visible = true
            return
        }
    }
    if (paneId) {
        try {
            var stale = window.Application.GetTaskPane(paneId)
            if (stale) stale.Visible = false
        } catch (err) { /* pane already gone */ }
    }
    var pane = window.Application.CreateTaskPane(tidePaneUrl() + "?b=" + TIDE_PANE_BUILD)
    storage.setItem(TIDE_PANE_KEY, pane.ID)
    storage.setItem("tide_pane_build", TIDE_PANE_BUILD)
    pane.Visible = true
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
