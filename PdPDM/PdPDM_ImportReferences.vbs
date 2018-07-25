'******************************************************************************
'* File       : PdPDM_ImportReferences.vbs
'* Purpose    : 从Sheet中导入关系
'* Title      : 导入关系
'* Category   : 导入
'* Version    : v1.1
'* Company    : www.duanzhihui.com
'* Author     : 段智慧
'* Description: 从Sheet中导入关系
'*              CRUD
'*                  C 新增  如果关系存在则忽略，否则新增；如果连接存在则忽略，否则新增。
'*                  R 只读  直接忽略。
'*                  U 更新  如果关系存在则更新，否则新增；如果连接存在则删除后新增，否则新增。
'*                  D 删除  删除关系。
'* History    : 2016-04-07  v1.0    段智慧  新增
'*              2017-06-20  v1.1    段智慧  opensource计划
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
    exl.Visible = True

    Dim path, ws
    set ws=CreateObject("WScript.Shell")
    path = ws.CurrentDirectory + "\PdPDM_ImportReferences.xlsx"

    path = InputBox ("请输入包含模型的Excel文件路径。", "文件路径", path)
    output "Excel文件路径为: " + path

    exl.Workbooks.Open path
    exl.Workbooks(1).Worksheets(1).Activate    '指定要打开的sheet名称
Else
    HaveExcel = False
End If

dim par, obj1, obj2, ref, row, num
'on error Resume Next

row = 2
With exl.Workbooks(1).Worksheets(1)
    do While .Cells(row, 1).Value <> ""                                                                         '退出
        set obj1 = ActiveModel.FindChildByCode(.Cells(row, 6).Value, PdPDM.cls_Table, "", nothing, False)       '指定 主表
        set obj2 = ActiveModel.FindChildByCode(.Cells(row, 7).Value, PdPDM.cls_Table, "", nothing, False)       '指定 从表
        set par = obj2.Parent
        if par is nothing or obj1 is nothing or obj2 is nothing then
            output "第" + CStr(row) + "行，WARNING【无对象】：" + CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +")。"
        else

            if .Cells(row, 12).Value = 1 then
                num = ""
            else
                num = "-" + CStr(.Cells(row, 13).Value)
            end if
            exl.Range("B"+Cstr(row)).Value = par.Code
            exl.Range("C"+Cstr(row)).Value = "外键_" + obj2.Name + "-" + obj1.Name + num

            select case UCase(.Cells(row, 1).Value)
            case "C"
                output "第" + CStr(row) + "行，新增关系：" + CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +")。"
                CreateReference par, obj1, obj2, exl, row
            case "R"
                output "第" + CStr(row) + "行，只读关系：" + CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +")。"
            case "U"
                output "第" + CStr(row) + "行，更新关系：" + CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +")。"
                UpdateReference par, obj1, obj2, exl, row
            case "D"
                output "第" + CStr(row) + "行，删除关系：" + CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +")。"
                DeleteReference par, exl, row
            case Else
                output "第" + CStr(row) + "行，忽略关系：" + CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +")。"
            end select
           'exl.Range("A"+Cstr(row)).Value = "R"            '将 CRUD 设为默认值 R
        end if
        row = row + 1
    Loop
End With

exl.Workbooks(1).Close True

sub CreateReference(par, obj1, obj2, exl, row)
    dim ref
    With exl.Workbooks(1).Worksheets(1)
        set ref = mdl.FindChildByCode(.Cells(row, 4).Value, PdPDM.cls_Reference, "", nothing, False)
        if not ref is nothing then
            output "|__"+CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +") 关系存在，忽略关系。"
        Else
            set ref = par.References.CreateNew          '创建 关系
            ref.Name = .Cells(row, 3).Value             '指定 关系名称
            ref.Code = .Cells(row, 4).Value             '指定 关系编码
            ref.Comment = .Cells(row, 5).Value          '指定 关系注释
            set ref.ParentTable = obj1                  '指定 关系主表
            set ref.ChildTable = obj2                   '指定 关系从表
            CreateJoins par, ref, exl, CInt(.Cells(row, 11).Value)              '指定 关系连接
        end if
    End With
End sub

sub UpdateReference(par, obj1, obj2, exl, row)
    dim ref
    With exl.Workbooks(1).Worksheets(1)
        set ref = mdl.FindChildByCode(.Cells(row, 4).Value, PdPDM.cls_Reference, "", nothing, False)
        if not ref is nothing then
            ref.Name = .Cells(row, 3).Value             '指定 关系名称
           'ref.Code = .Cells(row, 4).Value             '指定 关系编码
            ref.Comment = .Cells(row, 5).Value          '指定 关系注释
            set ref.ParentTable = obj1                  '指定 关系主表
            set ref.ChildTable = obj2                   '指定 关系从表
            UpdateJoins par, ref, exl, CInt(.Cells(row, 11).Value)              '指定 关系连接
        Else
            output "|__"+CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +") 关系不存在，新增关系。"
            CreateReference par, obj1, obj2, exl, row
        end if
    End With
End sub

sub DeleteReference(par, exl, row)
    dim ref
    With exl.Workbooks(1).Worksheets(1)
        set ref = mdl.FindChildByCode(.Cells(row, 4).Value, PdPDM.cls_Reference, "", nothing, False)
        if not ref is nothing then
            ref.Delete
        end if
    End With
End sub

sub CreateJoins(par, ref, exl, row)
    dim jn
    With exl.Workbooks(1).Worksheets(2)
        if ref.Joins.Count = 0 then
            do
                CreateJoin exl, row, ref, -1
                row = row + 1
            loop while .Cells(row, 3).Value = .Cells(row - 1, 3).Value
        Else
            output "|__"+CStr(.Cells(row, 2).Value) + "(" + CStr(.Cells(row, 3).Value) +") 关系连接存在，修改连接。"
            UpdateJoins par, ref, exl, row
        end if
    End With
End sub

sub UpdateJoins(par, ref, exl, row)
    dim jn, idx
    With exl.Workbooks(1).Worksheets(2)
        if ref.Joins.Count = 0 then
            output "|__"+CStr(.Cells(row, 2).Value) + "(" + CStr(.Cells(row, 3).Value) +") 关系连接不存在，新增连接。"
            CreateJoins par, ref, exl, row
        Else
            do
                idx = FindJoin(exl, row, ref)               '   获取关联字段位置
                if idx = -1 then
                    CreateJoin exl, row, ref, idx
                Else
                    UpdateJoin exl, row, ref, idx
                End if
                row = row + 1
            loop while .Cells(row, 3).Value = .Cells(row - 1, 3).Value
        end if
    End With
End sub


sub CreateJoin(exl, row, ref, idx)
    dim jn
    With exl.Workbooks(1).Worksheets(2)
        set jn = ref.Joins.CreateNewAt(idx)                                                                                         '创建 连接
        set jn.ParentTableColumn = ref.ParentTable.FindChildByCode(.Cells(row, 4).Value, PdPDM.cls_Column, "", nothing, False)      '指定 连接父字段
        set jn.ChildTableColumn  = ref.ChildTable.FindChildByCode(.Cells(row, 5).Value, PdPDM.cls_Column, "", nothing, False)       '指定 连接子字段
    End With
End sub

sub UpdateJoin(exl, row, ref, idx)
    dim jn
    With exl.Workbooks(1).Worksheets(2)
        set jn = ref.Joins.Item(idx)                                                                                                '获取 连接
       'set jn.ParentTableColumn = ref.ParentTable.FindChildByCode(.Cells(row, 4).Value, PdPDM.cls_Column, "", nothing, False)      '指定 连接父字段
        set jn.ChildTableColumn  = ref.ChildTable.FindChildByCode(.Cells(row, 5).Value, PdPDM.cls_Column, "", nothing, False)       '指定 连接子字段
    End With
End sub

Function FindJoin(exl, row, ref)
    dim jn, ptc, idx
    With exl.Workbooks(1).Worksheets(2)
        set ptc = ref.ParentTable.FindChildByCode(.Cells(row, 4).Value, PdPDM.cls_Column, "", nothing, False)                       '获取 连接父字段
        idx = 0
        FindJoin = -1
        do while idx < ref.Joins.Count
            set jn = ref.Joins.Item(idx)
            if jn.ParentTableColumn.Code = ptc.Code then
                FindJoin = idx
            end if
            idx = idx + 1
        Loop
    End With
End Function
