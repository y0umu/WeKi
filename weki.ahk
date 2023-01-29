/*
* WeKi
*
* 已实现的功能：
* - 微信自动群发消息
* - 半自动转发
*
* 群发用法：
* 按下 Win + Shift + z 复制并处理转发列表 > 按下 Win + Shift + a 输入要群发的文字消息并发送
*
*
* 半自动转发用法:
* 按下 Win + Shift + z 复制并处理转发列表 > 打开微信转发界面 > 按下 Win + Shift + x 执行转发
* Win + Shift + x 最多转发9人，转发名单超过9人时，请多次使用Win + Shift + X 直到遍历了转发名单
* Note: 半自动转发通常用来群发大量消息以及文件
*/

;****************************
; 参数配置
;****************************
; 高分辨率的不同预设
preset := "" ; 目前的可能取值有： "hidpi175", ""
if (preset = "hidpi175") {
; 175% 缩放下的预设
    imgDuoXuan := "wechat_multi-forward_assets\duoxuan_hidpi.png"
    rectDuoXuan := [280, 430, 120, 195]  ; 多选按钮的搜寻区域。坐标形式：[left, right, top, bottom]
    rectSend := [490, 630, 580, 610] ; 发送按钮区域。坐标形式：[left, right, top, bottom]
}
else {
    imgDuoXuan := "wechat_multi-forward_assets\duoxuan.png"
    rectDuoXuan := [270, 330, 100, 145]  ; 多选按钮的搜寻区域。坐标形式：[left, right, top, bottom]
    rectSend := [415, 523 ,489 ,510] ; 发送按钮区域。坐标形式：[left, right, top, bottom]
}

; Win + Shift + a 群发消息是否:
; 将 {name} 替换为全名
; 将 {lastname} 替换为全名的第一个字（一般为姓氏，但不能处理复姓）
templateSubstitute := true

; 全局变量
nameList := []
nameListInd := 0

; 明确 Click、ImageSearch 的坐标都是相对于窗口的（去除窗口装饰）
CoordMode "Mouse", "Client"
CoordMode "Pixel", "Client"

;****************************
; 函数
;****************************
printList(lst, title := "预览") {
    i := 1
    txt := "数组大小 = " . lst.Length . "`n"
    while (i <= lst.Length) {
        ; avoid too long output
        if (i > 15) {
            txt .= "...`n"
            break
        }
        txt .= "[" i "] " . lst[i] . " (length:" . StrLen(lst[i]) . ")`n"
        i += 1
    }
    MsgBox txt, title
}

; 检查 nameListInd 、 nameList 是否有效。
; 这两个变量在初始化或一次群发/完成之后会清空
validNameList() {
    if (nameListInd = 0 or nameList.Length = 0) {
        MsgBox "收件人列表已经清空了。请先按 Win + Shift + z 复制收件人列表"
        return false
    }
    return true
}

; 在 (x1, x2, y1, y2) 围成的矩形区域中进行一次点击
; 点击坐标是随机的，模拟人类的点击
clickBox(x1, x2, y1, y2) {
    Click Random(x1, x2), Random(y1, y2)
}

; 将鼠标移至(x1, x2, y1, y2) 围成的矩形区域中的某个随机位置
mouseMoveBox(x1, x2, y1, y2) {
    MouseMove Random(x1, x2), Random(y1, y2)
}


;****************************
; 按键绑定
;****************************
/*
#d::
{
    MsgBox rectDuoXuan[1] "," rectDuoXuan[3] "," rectDuoXuan[2] "," rectDuoXuan[4]
}
*/

; Win + Shift + z
; 处理来自剪贴板的输入收件人名单（换行符分隔），并将结果保存在 global nameList
#+z::
{
    global nameList := []
    global nameListInd := 1
    Send "^c"
    Sleep 50
    nameList := StrSplit(A_Clipboard, "`n", "`r")
    if (nameList.Length = 0) {
        nameListInd := 0
        return
    }
    ; 从 excel 复制出来的列在最后会有空值，删除之
    if (StrLen(nameList[-1]) = 0) {
        nameList.Pop()
    }
    printList(nameList, "剪切板预览")
}

; Win + Shift + x
; 执行转发
#+x::
{
    global nameList
    global nameListInd ; 当前指针
    if (not validNameList()) {
        return
    }
    ; 寻找“多选”按钮
    duoxuanFound := ImageSearch(&foundX, &foundY, rectDuoXuan[1], rectDuoXuan[3], rectDuoXuan[2], rectDuoXuan[4], "*25 " . imgDuoXuan)
    if (not duoxuanFound) {
        MsgBox "未找到多选按钮。请再试一次或调试本程序。"
        return
    }
    ;MouseMove foundX, foundY  ; MouseMove
    clickBox(foundX, foundX + 34, foundY, foundY + 17) ; duoxuan.png 是长35，宽18的图像
    Sleep 50
    ; 聚焦输入框
    Send "^f"
    Sleep 50
    ; 选中联系人
    i := 0
    Loop {
        SendText nameList[nameListInd + i] ; 输入联系人姓名
        Sleep 200 ; 微信此时会搜索联系人，这是一个耗时的工作
        Send "{Enter}"  ; 选中该联系人
        i += 1
        ; 退出条件1（超过名单长度）
        if (nameListInd + i > nameList.Length) {
            nameList := []
            nameListInd := 0
            mouseMoveBox(rectSend[1], rectSend[2], rectSend[3], rectSend[4]) ; 将鼠标移至“发送”按钮上，由用户最终确定是不是发送
            return
        }
        ; 退出条件2（已经选中9人（因为微信一次只让转9人））
        if (i = 9) {
            nameListInd += 9
            mouseMoveBox(rectSend[1], rectSend[2], rectSend[3], rectSend[4]) ; 将鼠标移至“发送”按钮上，由用户最终确定是不是发送
            return
        }
    }
}

; Win + Shift + a
; 弹出对话框请用户输入，并执行群发
#+a::
{
    global nameList
    global nameListInd ; 当前指针
    if (not validNameList()) {
        return
    }
    
    ; 输入消息
    ret := InputBox("请输入要群发的消息")
    if (ret.Result = "Cancel") {
        return
    }
    msg_ := ret.Value
    
    ; 寻找微信主窗口。这是因为输入消息后微信主窗口会失去焦点
    if (WinExist("ahk_class WeChatMainWndForPC") ) {
        WinActivate
    }
    else {
        ; 用 Ctrl + Alt + w 快捷键调出微信主窗口
        Send "^!w"
        Sleep 50
    }
    
    Loop {
        ; 聚焦联系人输入框
        Send "^f"
        Sleep 50
        ; 搜索联系人
        SendText nameList[nameListInd]
        Sleep 1000  ; 这里的搜索不仅仅搜联系人，比转发更加耗时
        Send "{Enter}"  ; 选中该联系人
        Sleep 500
        ; 发送消息
        msg := msg_
        if (templateSubstitute) {
            msg := StrReplace(msg, "{name}", nameList[nameListInd])
            msg := StrReplace(msg, "{lastname}", SubStr(nameList[nameListInd], 1, 1))
        }
        SendText msg
        Sleep 100
        Send "{Enter}" ; 此时触发发送动作
        
        nameListInd += 1
        ; 退出条件（超过名单长度）
        if (nameListInd > nameList.Length) {
            nameList := []
            nameListInd := 0
            return
        }
    }
}