'******************************************************************************
'* File       : PdPDM_ConfigTables.vbs
'* Purpose    : 从excel导入表配置
'* Title      : 配置表
'* Category   : 导入模型
'* Version    : v2.1
'* Company    : www.duanzhihui.com
'* Author     : 段智慧
'* Description: 配置表物理选项和默认字段
'*              CRUD
'*                  C 新增  如果存在则忽略，否则新增。
'*                  R 只读  直接忽略。
'*                  U 更新  如果存在则更新，否则新增。
'*                  D 删除  删除。
'* History    :
'*              2016-03-31  v1.0    段智慧  新增 默认字段配置。
'*              2016-07-08  v2.0    段智慧  增加 物理选项配置；支持通配符 * 。
'*              2017-08-18  v2.1    段智慧 修改 CBoolean ，解决TRUE、FALSE当成字符串处理，默认输出FALSE问题。
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
    Dim exl
    Set exl = CreateObject("Excel.Application")

    Dim path, ws
    set ws=CreateObject("WScript.Shell")
    path = ws.CurrentDirectory + "\PdPDM_ConfigTables.xlsx"

    path = InputBox ("请输入包含模型的Excel文件路径。", "文件路径", path)
    output "Excel文件路径为: " + path

    exl.Workbooks.Open path
    exl.Workbooks(1).Worksheets(1).Activate    '指定要打开的sheet名称
Else
    HaveExcel = False
End If

dim par, tbl, row, spar, stblname, stblcode, typ
'on error Resume Next

row = 2
With exl.Workbooks(1).Worksheets(1)
    do While .Cells(row, 1).Value <> ""                     '退出
        spar = CStr(.Cells(row, 2).Value)
        stblname = CStr(.Cells(row, 3).Value)
        stblcode = CStr(.Cells(row, 4).Value)
        typ = CStr(.Cells(row, 5).Value)

        set par = mdl.FindChildByCode(spar, PdPDM.cls_Package)
        if par is nothing then
            set par = mdl
        end if

        if spar = "*" Then
            For Each par in mdl.Packages
                For Each tbl in par.Tables
                    ConfigTable typ, par, exl, row, tbl
                Next
            Next
        ElseIf stblcode = "*" Then
            For Each tbl in par.Tables
                ConfigTable typ, par, exl, row, tbl
            Next
        Else
            set tbl = par.FindChildByCode(stblcode, PdPDM.cls_Table)
            if tbl is Nothing Then
                output "第" + CStr(row) + "行，表不存在：" + stblname + "(" + stblcode +")。"
            Else
                ConfigTable typ, par, exl, row, tbl
            end if
        end if
        row = row + 1
    Loop
End With

exl.Workbooks(1).Close False

sub ConfigTable(typ, par, exl, row, tbl)
    select case typ
    case "DefaultColumns"
        output "第" + CStr(row) + "行，默认字段：" + tbl + "。"
        DefaultColumns par, exl, row, tbl
    case "PhysicalOption"
        output "第" + CStr(row) + "行，物理选项：" + tbl + "。"
        PhysicalOption par, exl, row, tbl
    case Else
        output "第" + CStr(row) + "行，忽略配置：" + tbl + "。"
    end select
End sub

sub DefaultColumns(par, exl, row, tbl)
    dim col
    With exl.Workbooks(1).Worksheets(1)
        select case .Cells(row, 1).Value
        case "C"
            output "|__新增默认字段"
            CreateDefaultColumns par, exl, tbl, CLng(.Cells(row, 9).Value)
        case "R"
            output "|__只读默认字段"
        case "U"
            output "|__修改默认字段"
            UpdateDefaultColumns par, exl, tbl, CLng(.Cells(row, 9).Value)
        case "D"
            output "|__删除默认字段"
            DeleteDefaultColumns par, exl, tbl, CLng(.Cells(row, 9).Value)
        case Else
            output "|__忽略默认字段"
        end select
    End With
End sub

sub CreateDefaultColumns(par, exl, tbl, row)
    dim col, idx
    idx = tbl.Columns.Count
    With exl.Workbooks(1).Worksheets(2)
        do
            set col = tbl.FindChildByCode(.Cells(row, 5).Value, PdPDM.cls_Column)
            if col is Nothing Then
                set col = tbl.FindChildByName(.Cells(row, 4).Value, PdPDM.cls_Column)
            end if
            if col is nothing then
                CreateColumn exl, row, tbl, idx
            end if
            row = row + 1
            idx = idx + 1
        loop while .Cells(row, 3).Value = .Cells(row - 1, 3).Value
    End With
End sub

sub UpdateDefaultColumns(par, exl, tbl, row)
    dim col, idx
    idx = tbl.Columns.Count
    With exl.Workbooks(1).Worksheets(2)
        do
            set col = tbl.FindChildByCode(.Cells(row, 5).Value, PdPDM.cls_Column)
            if col is Nothing Then
                set col = tbl.FindChildByName(.Cells(row, 4).Value, PdPDM.cls_Column)
            end if
            if col is nothing then
                CreateColumn exl, row, tbl, idx
            else
                UpdateColumn exl, row, tbl, col, idx
            end if
            row = row + 1
            idx = idx + 1
        loop while .Cells(row, 3).Value = .Cells(row - 1, 3).Value
    End With
End sub

sub DeleteDefaultColumns(par, exl, tbl, row)
    dim col
    With exl.Workbooks(1).Worksheets(2)
        do
            set col = tbl.FindChildByCode(.Cells(row, 5).Value, PdPDM.cls_Column)
            if col is Nothing Then
                set col = tbl.FindChildByName(.Cells(row, 4).Value, PdPDM.cls_Column)
            end if
            if not col is nothing then
                tbl.Columns.Remove col, True
            end if
            row = row + 1
        loop while .Cells(row, 3).Value = .Cells(row - 1, 3).Value
    End With
End sub


'   CreateColumn    与 PdPDM_ImportTables.vbs 保持一致
sub CreateColumn(exl, row, tbl, idx)
    dim col
    With exl.Workbooks(1).Worksheets(2)
        set col = tbl.Columns.CreateNewAt(idx)                                  '创建 字段
        col.Name = .Cells(row, 4).Value                                         '指定 字段名称
        col.Code = .Cells(row, 5).Value                                         '指定 字段编码
        col.DataType = .Cells(row, 6).Value                                     '指定 数据类型
        col.Primary = CBoolean(.Cells(row, 7).Value)                            '指定 主键
        if not col.Primary Then
          col.Mandatory = CBoolean(.Cells(row, 8).Value)                        '指定 强制
        end if
        col.Comment = .Cells(row, 9).Value                                      '指定 注释
        if col.Comment = "" Then
            col.Comment = col.Name
        end if
        col.Description = .Cells(row, 10).Value                                 '指定 描述
        col.Annotation = .Cells(row, 11).Value                                  '指定 备注
    End With
End sub

'   UpdateColumn    与 PdPDM_ImportTables.vbs 保持一致
sub UpdateColumn(exl, row, tbl, col, idx)
    Dim i
    With exl.Workbooks(1).Worksheets(2)
        col.Name = .Cells(row, 4).Value                                         '指定 字段名称
        col.Code = .Cells(row, 5).Value                                         '指定 字段编码
        col.DataType = .Cells(row, 6).Value                                     '指定 数据类型
        col.Primary = CBoolean(.Cells(row, 7).Value)                            '指定 主键
        if not col.Primary Then
          col.Mandatory = CBoolean(.Cells(row, 8).Value)                        '指定 强制
        end if
        col.Comment = .Cells(row, 9).Value                                      '指定 注释
        if col.Comment = "" Then
            col.Comment = col.Name
        end if
        col.Description = .Cells(row, 10).Value                                 '指定 描述
        col.Annotation = .Cells(row, 11).Value                                  '指定 备注
    End With

    '移到指定位置
    for i = idx to tbl.Columns.Count - 1
        if col.Code = tbl.Columns.Item(i).Code Then
            tbl.Columns.Move idx, i
            exit for
        end if
    Next
End sub

'   CBoolean    与 PdPDM_ImportTables.vbs 保持一致
Function CBoolean(exp)
    select case exp
    case "TRUE"
        CBoolean = TRUE
    case "FALSE"
        CBoolean = FALSE
    case TRUE
        CBoolean = TRUE
    case FALSE
        CBoolean = FALSE        
    case "是"
        CBoolean = TRUE
    case "否"
        CBoolean = FALSE
    case "1"
        CBoolean = TRUE
    case "0"
        CBoolean = FALSE
    case "Y"
        CBoolean = TRUE
    case "N"
        CBoolean = FALSE
    case Else
        CBoolean = FALSE
    end select
End Function

sub PhysicalOption(par, exl, row, tbl)
    dim col, path, value
    With exl.Workbooks(1).Worksheets(1)
        path = CStr(.Cells(row, 7).Value)
        value = CStr(.Cells(row, 8).Value)
        select case .Cells(row, 1).Value
        case "C"
            output "|__新增物理选项"
            tbl.AddPhysicalOption(path)
            tbl.SetPhysicalOptionValue path, value
        case "R"
            output "|__只读物理选项"
        case "U"
            output "|__修改物理选项"
            tbl.SetPhysicalOptionValue path, value
        case "D"
            output "|__删除物理选项"
            tbl.DeletePhysicalOption path
        case Else
            output "|__忽略物理选项"
        end select
    End With
End sub
