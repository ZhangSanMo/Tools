; AutoHotkey v2 脚本：Word & Excel 文本批量替换 GUI 版
; 日期：2024
; 版本：3.0 - 增加 Excel 支持，大幅优化批处理性能，增加临时文件过滤

#Requires AutoHotkey v2.0
#SingleInstance Force

; === GUI 定义与布局 ===
try {
    global MyGui := Gui(, "Office 文档批量替换工具 (Word/Excel)")
    MyGui.SetFont("s10")

    ; --- 文件夹选择区 ---
    ; 调整了 GroupBox 高度，内部控件使用统一的左边距 x26
    MyGui.Add("GroupBox", "x10 y10 w520 h115", "第一步：选择目标及文件类型")
    
    MyGui.Add("Text", "x26 y35", "文件夹路径:")
    global FolderPathEdit := MyGui.Add("Edit", "x+m yp-4 w310 r1 vFolderPath")
    MyGui.Add("Button", "x+m yp-1 w80", "浏览...").OnEvent("Click", SelectFolder)
    
    global RecursiveCheck := MyGui.Add("Checkbox", "x26 y+15 vIsRecursive Checked", "包含所有子文件夹")
    
    global ProcessWordCheck := MyGui.Add("Checkbox", "x26 y+10 vProcessWord Checked", "处理 Word 文件 (*.doc, *.docx)")
    ; 增加间距并确保与 Word 选项在同一行
    global ProcessExcelCheck := MyGui.Add("Checkbox", "x+30 yp vProcessExcel Checked", "处理 Excel 文件 (*.xls, *.xlsx)")

    ; --- 替换规则管理区 ---
    ; 修复了偏移问题，强制设为 x10 与第一步对齐
    MyGui.Add("GroupBox", "x10 y140 w520 h250", "第二步：管理替换规则")
    
    MyGui.Add("Text", "x26 y165", "规则列表 (双击可修改):")
    
    ; 调整列表宽度，刚好适应 GroupBox
    global ReplacementsLV := MyGui.Add("ListView", "x26 y+10 w488 h150 HScroll Grid vReplacementsLV", ["查找内容", "替换为"])
    ReplacementsLV.ModifyCol(1, "240")
    ReplacementsLV.ModifyCol(2, "240")
    ReplacementsLV.OnEvent("DoubleClick", EditReplacement)

    ; 精确计算按钮位置，使其右对齐
    MyGui.Add("Button", "x344 y+10 w80", "添加...").OnEvent("Click", AddReplacement)
    MyGui.Add("Button", "x+10 yp w80", "删除选中").OnEvent("Click", DeleteReplacement)

    ; --- 执行与状态区 ---
    ; 强制 x10 对齐
    MyGui.Add("GroupBox", "x10 y405 w520 h80", "第三步：开始执行")
    
    global StartButton := MyGui.Add("Button", "x26 y430 w120 h40", "开始替换").OnEvent("Click", StartProcessing)
    global StatusText := MyGui.Add("Text", "x+15 yp+10 w350", "状态：准备就绪。")

    ; --- 初始化默认规则 ---
    ReplacementsLV.Add(, "旧文本", "新文本")

    MyGui.OnEvent("Close", (*) => ExitApp())
    MyGui.Show("w540 h500") ; 调整了整体窗口高度以适配布局
} catch as e {
    MsgBox "创建 GUI 失败! `n错误: " e.Message
}

; === GUI 事件处理函数 ===
SelectFolder(*) {
    try {
        selectedFolder := DirSelect(, 3, "请选择包含文档的文件夹")
        if IsSet(selectedFolder) && selectedFolder != "" {
            FolderPathEdit.Value := selectedFolder
        }
    }
}
AddReplacement(*) {
    oldText := InputBox("请输入要查找的原始内容：", "添加规则 - 第1步/共2步").Value
    if !IsSet(oldText) || oldText = ""
        return
    newText := InputBox("请输入要替换成的新内容：", "添加规则 - 第2步/共2步").Value
    if !IsSet(newText)
        return
    ReplacementsLV.Add(, oldText, newText)
}
DeleteReplacement(*) {
    focusedRow := ReplacementsLV.GetNext(0, "F")
    if focusedRow > 0 {
        if MsgBox("确定要删除选中的规则吗？", "确认删除", "YesNo") = "Yes"
            ReplacementsLV.Delete(focusedRow)
    } else {
        MsgBox "请先在列表中选择一个要删除的规则。", "提示"
    }
}
EditReplacement(lv, row) {
    oldText := lv.GetText(row, 1)
    newText := lv.GetText(row, 2)
    newOldText := InputBox("请修改要查找的原始内容：", "修改规则",, oldText).Value
    if !IsSet(newOldText) || newOldText = ""
        return
    newNewText := InputBox("请修改要替换成的新内容：", "修改规则",, newText).Value
    if !IsSet(newNewText)
        return
    lv.Modify(row, , newOldText, newNewText)
}

; === 核心处理逻辑 ===
StartProcessing(*) {
    folderPath := FolderPathEdit.Value
    if !DirExist(folderPath) {
        MsgBox "错误：指定的文件夹不存在，请重新选择。", "错误"
        return
    }
    
    processWord := ProcessWordCheck.Value
    processExcel := ProcessExcelCheck.Value
    if (!processWord && !processExcel) {
        MsgBox "错误：请至少选择一种要处理的文件类型（Word 或 Excel）。", "错误"
        return
    }

    if ReplacementsLV.GetCount() = 0 {
        MsgBox "错误：替换规则列表为空，请至少添加一条规则。", "错误"
        return
    }

    MyGui.Opt("+Disabled")
    StatusText.Value := "正在扫描文件..."

    ; 提取替换规则
    replacements := Map()
    Loop ReplacementsLV.GetCount() {
        replacements[ReplacementsLV.GetText(A_Index, 1)] := ReplacementsLV.GetText(A_Index, 2)
    }

    isRecursive := RecursiveCheck.Value
    errorLogFile := A_ScriptDir "\error_log.txt"
    if FileExist(errorLogFile)
        FileDelete(errorLogFile)

    ; 收集符合条件的文件，并过滤掉 "~$" 开头的 Office 临时隐藏文件
    fileList := []
    Loop Files, folderPath "\*.*", (isRecursive ? "RF" : "F") {
        ext := StrLower(A_LoopFileExt)
        fileName := A_LoopFileName
        
        ; 忽略临时文件
        if InStr(fileName, "~$") = 1
            continue
            
        if (processWord && InStr(ext, "doc") = 1) || (processExcel && InStr(ext, "xls") = 1) {
            fileList.Push(A_LoopFileFullPath)
        }
    }

    totalFiles := fileList.Length
    if totalFiles = 0 {
        StatusText.Value := "提示：在指定位置未找到符合条件的文档。"
        MyGui.Opt("-Disabled")
        return
    }

    StatusText.Value := "正在启动 Office 进程..."
    
    wordApp := ""
    excelApp := ""
    
    ; 性能优化：在循环外部预先启动需要的 COM 进程
    try {
        if processWord {
            wordApp := ComObject("Word.Application")
            wordApp.Visible := false
            wordApp.DisplayAlerts := 0
        }
        if processExcel {
            excelApp := ComObject("Excel.Application")
            excelApp.Visible := false
            excelApp.DisplayAlerts := false
        }
    } catch as e {
        MsgBox "启动 Office 进程失败，请确保电脑已安装对应的软件。`n" e.Message, "环境错误"
        if wordApp
            wordApp.Quit()
        if excelApp
            excelApp.Quit()
        MyGui.Opt("-Disabled")
        return
    }

    processedCount := 0
    errorCount := 0

    ; 开始遍历处理文件
    for index, currentFile in fileList {
        fileName := StrSplit(currentFile, "\").Pop()
        StatusText.Value := "进度: " index "/" totalFiles "`n正在处理: " fileName
        ext := StrLower(RegExReplace(currentFile, ".*\.([^.]+)$", "$1"))

        try {
            if InStr(ext, "doc") = 1 {
                ; 处理 Word
                doc := wordApp.Documents.Open(currentFile)
                ProcessSingleDocument(doc, replacements)
                doc.Close(-1) ; -1 对应 wdSaveChanges
            } 
            else if InStr(ext, "xls") = 1 {
                ; 处理 Excel
                wb := excelApp.Workbooks.Open(currentFile)
                ProcessSingleExcel(wb, replacements)
                wb.Close(1) ; 1 对应 xlSaveChanges
            }
            processedCount++
        } catch as e {
            errorCount++
            logMessage := A_Now " | " currentFile " | 错误: " e.Message "`n"
            FileAppend(logMessage, errorLogFile)
        }
    }

    ; 清理并退出 Office 进程
    StatusText.Value := "正在清理并关闭后台进程..."
    if IsObject(wordApp)
        try wordApp.Quit(0) ; 0 对应 wdDoNotSaveChanges (文件已在此前保存)
    if IsObject(excelApp)
        try excelApp.Quit()

    finalMessage := "✅ 处理完成！`n`n成功处理 " processedCount " 个文件。`n失败 " errorCount " 个。"
    if (errorCount > 0) {
        finalMessage .= "`n`n失败详情请查看日志文件：`n" errorLogFile
    }
    StatusText.Value := "状态：处理完成。"
    MyGui.Opt("-Disabled")
    MsgBox finalMessage, "任务完成"
}

; === 函数：对单个 Word 文档执行所有替换操作 ===
ProcessSingleDocument(doc, replacements) {
    storyTypes := [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]
    for _, storyType in storyTypes {
        try {
            range := doc.StoryRanges(storyType)
            while (range) {
                findObj := range.Find
                findObj.ClearFormatting()
                findObj.MatchCase := false
                findObj.MatchWholeWord := false
                findObj.MatchWildcards := false
                findObj.Forward := true
                findObj.Wrap := 1
                
                repl := findObj.Replacement
                repl.ClearFormatting()

                for oldText, newText in replacements {
                    findObj.Text := oldText
                    repl.Text := newText
                    findObj.Execute(,,,,,,,,,, 2) ; 2 = wdReplaceAll
                }
                range := range.NextStoryRange
            }
        } catch {
            continue
        }
    }
}

; === 函数：对单个 Excel 文档执行所有替换操作 ===
ProcessSingleExcel(wb, replacements) {
    for ws in wb.Worksheets {
        for oldText, newText in replacements {
            try {
                ; 参数：What, Replacement, LookAt (2=xlPart 局部匹配)
                ws.Cells.Replace(oldText, newText, 2)
            } catch {
                continue
            }
        }
    }
}
