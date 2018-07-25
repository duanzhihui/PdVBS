'******************************************************************************
'* File       : PdPDM_ImportTables.vbs
'* Purpose    : 从excel导入表
'* Title      : 导入表
'* Category   : 导入模型
'* Version    : v2.3
'* Company    : www.duanzhihui.com
'* Author     : 段智慧
'* Description: 从excel导入表
'*              CRUD
'*                  C 新增  如果表存在则忽略，否则新增；如果字段存在则忽略，否则新增。
'*                  R 只读  直接忽略。
'*                  U 更新  如果表存在则更新，否则新增；如果字段存在则更新，否则新增。
'*                  D 删除  删除表。
'* History    : 2016-03-09  v1.3    段智慧 增加字段默认“Comment”
'*              2016-04-09  v1.4    段智慧 增加表“Description”，修改表和字段“Comment”
'*              2017-04-10  v1.5    段智慧 增加过程“SetPrimaryKey”
'*              2017-05-10  v2.0    段智慧 增加表“Description、Annotation”
'*              2017-06-14  v2.1    段智慧 修改“Comment”取法是的import与export数据一致。
'*              2017-06-15  v2.2    段智慧 查找不区分大小写，增加 table.parent 修改。
'*              2018-01-05  v2.3    段智慧 解决 select case 大小写问题。
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
    path = ws.CurrentDirectory + "\PdPDM_ImportTables.xlsx"

    path = InputBox ("请输入包含模型的Excel文件路径。", "文件路径", path)
    output "Excel文件路径为: " + path

    exl.Workbooks.Open path
    exl.Workbooks(1).Worksheets(1).Activate
Else
    HaveExcel = False
End If

dim par, tbl, row
'on error Resume Next

row = 2
With exl.Workbooks(1).Worksheets(1)
    do While .Cells(row, 1).Value <> ""                     '退出

        set par = mdl.FindChildByCode(CStr(.Cells(row, 2).Value), PdPDM.cls_Package, "", nothing, False)
        if par is nothing then
            set par = mdl
        end if

        select case Ucase(.Cells(row, 1).Value)
        case "C"
            output "第" + CStr(row) + "行，新增表：" + CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +")。"
            CreateTable par, exl, row
        case "R"
            output "第" + CStr(row) + "行，只读表：" + CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +")。"
        case "U"
            output "第" + CStr(row) + "行，更新表：" + CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +")。"
            UpdateTable par, exl, row
        case "D"
            output "第" + CStr(row) + "行，删除表：" + CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +")。"
            DeleteTable par, exl, row
        case Else
            output "第" + CStr(row) + "行，忽略表：" + CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +")。"
        end select
       'exl.Range("A"+Cstr(row)).Value = "R"                '将 CRUD 设为默认值 R
        row = row + 1
    Loop
End With

exl.Workbooks(1).Close True

sub CreateTable(par, exl, row)
    dim tbl
    With exl.Workbooks(1).Worksheets(1)
        set tbl = mdl.FindChildByCode(.Cells(row, 4).Value, PdPDM.cls_Table, "", nothing, False)
        if not tbl is nothing then
            output "|__"+CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +") 表存在，忽略表。"
        Else
            set tbl = par.Tables.CreateNew                                      '创建 表
            tbl.Name = .Cells(row, 3).Value                                     '指定 表名称
            tbl.Code = .Cells(row, 4).Value                                     '指定 表编码
            if .Cells(row, 5).Value = "" and tbl.Name <> tbl.Code then          '指定 注释
                tbl.Comment = tbl.Name
            Else
                tbl.Comment = .Cells(row, 5).Value
            End if
            tbl.Description = .Cells(row, 6).Value                              '指定 描述
            tbl.Annotation = .Cells(row, 7).Value                               '指定 备注
            tbl.Owner = mdl.FindChildByCode(.Cells(row, 8).Value, PdPDM.cls_User, "", nothing, False)
        end if
        CreateColumns par, exl, CLng(.Cells(row, 9).Value)
        SetPrimaryKey tbl
    End With
End sub


sub UpdateTable(par, exl, row)
    dim tbl
    With exl.Workbooks(1).Worksheets(1)
        set tbl = mdl.FindChildByCode(.Cells(row, 4).Value, PdPDM.cls_Table, "", nothing, False)
        if not tbl is nothing then
            if not tbl.parent is par then
                Dim sel
                Set sel = ActiveModel.CreateSelection
                sel.Objects.Add(tbl)
                sel.MoveToPackage(par)
            end if
            tbl.Name = .Cells(row, 3).Value                                     '指定 表名称
           'tbl.Code = .Cells(row, 4).Value                                     '指定 表编码
            if .Cells(row, 5).Value = "" and tbl.Name <> tbl.Code then          '指定 注释
                tbl.Comment = tbl.Name
            Else
                tbl.Comment = .Cells(row, 5).Value
            End if
            tbl.Description = .Cells(row, 6).Value                              '指定 描述
            tbl.Annotation = .Cells(row, 7).Value                               '指定 备注
            tbl.Owner = mdl.FindChildByCode(.Cells(row, 8).Value, PdPDM.cls_User, "", nothing, False)
        Else
            output "|__"+CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +") 表不存在，新增表。"
            CreateTable par, exl, row
        end if
        UpdateColumns par, exl, CLng(.Cells(row, 9).Value)
       'SetPrimaryKey tbl
    End With
End sub

sub DeleteTable(par, exl, row)
    dim tbl
    With exl.Workbooks(1).Worksheets(1)
        set tbl = mdl.FindChildByCode(.Cells(row, 4).Value, PdPDM.cls_Table, "", nothing, False)
        if not tbl is nothing then
            tbl.Delete
        end if
    End With
End sub

sub CreateColumns(par, exl, row)
    dim tbl, col
    With exl.Workbooks(1).Worksheets(2)
        set tbl = par.FindChildByCode(.Cells(row, 3).Value, PdPDM.cls_Table, "", nothing, False)
        if tbl.Columns.Count = 0 then
            do
                CreateColumn exl, row, tbl, -1
                row = row + 1
            loop while .Cells(row, 3).Value = .Cells(row - 1, 3).Value
        Else
            output "|__"+CStr(.Cells(row, 2).Value) + "(" + CStr(.Cells(row, 3).Value) +") 表字段存在，忽略字段。"
        end if
    End With
End sub

sub UpdateColumns(par, exl, row)
    dim tbl, col, idx
    idx = 0
    With exl.Workbooks(1).Worksheets(2)
        set tbl = par.FindChildByCode(.Cells(row, 3).Value, PdPDM.cls_Table, "", nothing, False)
        if tbl.Columns.Count = 0 then
            output "|__"+CStr(.Cells(row, 2).Value) + "(" + CStr(.Cells(row, 3).Value) +") 表字段不存在，新增字段。"
            CreateColumns par, exl, row
        Else
            do
                set col = tbl.FindChildByCode(.Cells(row, 5).Value, PdPDM.cls_Column, "", nothing, False)
                if col is Nothing Then                      'V1.2   Code 找不到，找 Name。
                    set col = tbl.FindChildByName(.Cells(row, 4).Value, PdPDM.cls_Column, "", nothing, False)
                end if
                if not col is nothing then
                    UpdateColumn exl, row, tbl, col, idx
                else
                    CreateColumn exl, row, tbl, idx
                end if
                row = row + 1
                idx = idx + 1
            loop while .Cells(row, 3).Value = .Cells(row - 1, 3).Value

            '删除多余字段
            do while idx < tbl.Columns.Count
                tbl.Columns.RemoveAt(idx)
            Loop
        end if
    End With
End sub

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
        if .Cells(row, 9).Value = "" and col.name <> col.code Then              '指定 注释
            col.Comment = col.Name
        else
            col.Comment = .Cells(row, 9).Value
        end if
        col.Description = .Cells(row, 10).Value                                 '指定 描述
        col.Annotation = .Cells(row, 11).Value                                  '指定 备注
    End With
End sub

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
        if .Cells(row, 9).Value = "" and col.name <> col.code Then              '指定 注释
            col.Comment = col.Name
        else
            col.Comment = .Cells(row, 9).Value
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

sub SetPrimaryKey(tbl)
    dim pk
    if not tbl.PrimaryKey is nothing Then
        set pk = tbl.PrimaryKey
        pk.Name = "主键_" + tbl.Name
        pk.Code = "PK_" + tbl.Code
        pk.Comment = pk.Name
    end if
End sub

Function CBoolean(exp)
    select case Ucase(exp)
    case "TRUE", "是", "1", "Y", TRUE
        CBoolean = TRUE
    case "FALSE", "否", "0", "N", FALSE
        CBoolean = FALSE
    case Else
        CBoolean = FALSE
    end select
End Function
