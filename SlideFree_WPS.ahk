#Requires AutoHotkey v2.0

; 获取正在运行的 WPS 演示应用
GetWPS() {
    try {
        return ComObjActive("KWPP.Application")
    } catch {
        return
    }
}

; 获取当前放映窗口
GetWPSSlideShowWindow(wps) {
    try {
        if wps.SlideShowWindows.Count > 0
            return wps.SlideShowWindows(1)
    } catch {
        return
    }
}

; 判断是否正在放映
IsSlideShowRunning(wps) {
    try {
        return wps.SlideShowWindows.Count > 0
    } catch {
        return false
    }
}

; 切换到放映窗口
ActivateSlideShowWindow(wps) {
    slideShowWindow := GetWPSSlideShowWindow(wps)
    if slideShowWindow {
        ; 激活放映窗口
        slideShowWindow.Activate()
    }
}

; 下一页
WPS_Next() {
    wps := GetWPS()
    if !wps
        return
    if IsSlideShowRunning(wps) {
        ActivateSlideShowWindow(wps)  ; 确保窗口被激活
        wps.SlideShowWindows(1).View.Next()
    }
}

; 上一页
WPS_Prev() {
    wps := GetWPS()
    if !wps
        return
    if IsSlideShowRunning(wps) {
        ActivateSlideShowWindow(wps)  ; 确保窗口被激活
        wps.SlideShowWindows(1).View.Previous()
    }
}

; —— 绑定翻页键 —— 
$PgDn::WPS_Next()
$PgUp::WPS_Prev()