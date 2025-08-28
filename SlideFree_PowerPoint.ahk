#Requires AutoHotkey v2.0

; 获取正在运行的 PowerPoint 应用
GetPPT() {
    try {
        return ComObjActive("PowerPoint.Application")
    } catch {
        return
    }
}

; 判断是否正在放映
IsSlideShowRunning(ppt) {
    try {
        return ppt.SlideShowWindows.Count > 0
    } catch {
        return false
    }
}

; 下一页
PPT_Next() {
    ppt := GetPPT()
    if !ppt
        return
    if IsSlideShowRunning(ppt)
        ppt.SlideShowWindows(1).View.Next()
}

; 上一页
PPT_Prev() {
    ppt := GetPPT()
    if !ppt
        return
    if IsSlideShowRunning(ppt)
        ppt.SlideShowWindows(1).View.Previous()
}

; —— 绑定翻页键 —— 
$PgDn::PPT_Next()
$PgUp::PPT_Prev()