'******************************************************************************
'* File       : PdPDM_AutoAttach.vbs
'* Purpose    : 从Sheet附加表到物理视图
'* Title      : 从Sheet附加表到物理视图
'* Category   : 导入附加
'* Version    : v1.0
'* Company    : www.duanzhihui.com
'* Author     : 段智慧
'* Description:
'*              CRUD
'*                  C 新增
'*                  R 只读  直接忽略。
'*                  U 更新
'*                  D 删除
'* History    :
'*              2016-03-09  v1.0    段智慧 增加
'******************************************************************************

Option Explicit

Dim mdl                                             ' the current model
Set mdl = ActiveModel
If (mdl Is Nothing) Then
   MsgBox "There is no Active Model"
End If

Dim HaveExcel
Dim RQ
RQ = vbYes 'MsgBox("Is Excel Installed on your machine ?", vbYesNo + vbInformation, "Confirmation")
If RQ = vbYes Then
    HaveExcel = True
    ' Open & Create Excel Document
    Dim exl  '
    Set exl = CreateObject("Excel.Application")

    Dim path, ws
    set ws=CreateObject("WScript.Shell")
    path = ws.CurrentDirectory + "\PdPDM_AttachTables.xlsx"

    path = InputBox ("请输入包含模型的Excel文件路径。", "文件路径", path)
    output "Excel文件路径为: " + path

    exl.Workbooks.Open path
    exl.Workbooks(1).Worksheets(1).Activate    '指定要打开的sheet名称
Else
    HaveExcel = False
End If

'on error Resume Next
Dim dgrm, tbl, sym, row
row = 2
With exl.Workbooks(1).Worksheets(1)
    do While .Cells(row, 1).Value <> ""                     '退出
        select case Ucase(.Cells(row, 1).Value)
        case "C"
            set dgrm = mdl.FindChildByCode(.Cells(row, 2).Value, PdPDM.cls_PhysicalDiagram, "", nothing, False)
            set tbl = mdl.FindChildByCode(.Cells(row, 3).Value, PdPDM.cls_Table, "", nothing, False)
            set sym = dgrm.FindSymbol(tbl)
            if sym is nothing then
                dgrm.AttachObject tbl
                output "第" + CStr(row) + "行，附加表：" + tbl + "。"
            end if
        case "R"
        case "U"
        case "D"
        case Else
        end select
       'exl.Range("A"+Cstr(row)).Value = "R"                '将 CRUD 设为默认值 R
        row = row + 1
    Loop
End With
exl.Workbooks(1).Close True
