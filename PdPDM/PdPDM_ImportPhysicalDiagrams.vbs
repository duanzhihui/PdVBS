'*******************************************************************************
'* File       : PdPDM_ImportPhysicalDiagrams.vbs
'* Purpose    : 从Sheet中导入PhysicalDiagrams
'* Title      : 从Sheet中导入PhysicalDiagrams
'* Category   : 导入模型
'* Version    : v1.0
'* Company    : www.duanzhihui.com
'* Author     : 段智慧
'* Description:
'*              CRUD
'*                  C 新增  如果PhysicalDiagram存在则忽略，否则新增。
'*                  R 只读  直接忽略。
'*                  U 更新  如果PhysicalDiagram存在则更新，否则新增。
'*                  D 删除  删除PhysicalDiagram。
'* History    :
'*              2016-03-31  v1.0    段智慧  新增
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
    path = ws.CurrentDirectory + "\PdPDM_ImportPhysicalDiagrams.xlsx"

    path = InputBox ("请输入包含模型的Excel文件路径。", "文件路径", path)
    output "Excel文件路径为: " + path

    exl.Workbooks.Open path
    exl.Workbooks(1).Worksheets(1).Activate    '指定要打开的sheet名称
Else
    HaveExcel = False
End If

dim par, dgrm, row
on error Resume Next

row = 2
With exl.Workbooks(1).Worksheets(1)
    do While .Cells(row, 1).Value <> ""                     '退出

        set par = mdl.FindChildByCode(CStr(.Cells(row, 2).Value), PdPDM.cls_Package, "", nothing, False)
        if par is nothing then
            set par = mdl
        end if

        select case Ucase(.Cells(row, 1).Value)
        case "C"
            output "第" + CStr(row) + "行，新增PhysicalDiagram：" + CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +")。"
            CreatePhysicalDiagram par, exl, row
        case "R"
            output "第" + CStr(row) + "行，只读PhysicalDiagram：" + CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +")。"
        case "U"
            output "第" + CStr(row) + "行，更新PhysicalDiagram：" + CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +")。"
            UpdatePhysicalDiagram par, exl, row
        case "D"
            output "第" + CStr(row) + "行，删除PhysicalDiagram：" + CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +")。"
            DeletePhysicalDiagram par, exl, row
        case Else
            output "第" + CStr(row) + "行，忽略PhysicalDiagram：" + CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +")。"
        end select
       'exl.Range("A"+Cstr(row)).Value = "R"                '将 CRUD 设为默认值 R
        row = row + 1
    Loop
End With

exl.Workbooks(1).Close True

sub CreatePhysicalDiagram(par, exl, row)
    dim dgrm
    With exl.Workbooks(1).Worksheets(1)
        set dgrm = par.FindChildByCode(.Cells(row, 4).Value, PdPDM.cls_PhysicalDiagram, "", nothing, False)
        if not dgrm is nothing then
            output "|__"+CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +") PhysicalDiagram存在，忽略PhysicalDiagram。"
        Else
            set dgrm = par.PhysicalDiagrams.CreateNew    '创建 PhysicalDiagram
            dgrm.Name = .Cells(row, 3).Value            '指定 PhysicalDiagram名称
            dgrm.Code = .Cells(row, 4).Value            '指定 PhysicalDiagram编码
            dgrm.Comment = .Cells(row, 5).Value         '指定 PhysicalDiagram注释
        end if
    End With
End sub

sub UpdatePhysicalDiagram(par, exl, row)
    dim dgrm
    With exl.Workbooks(1).Worksheets(1)
        set dgrm = par.FindChildByCode(.Cells(row, 4).Value, PdPDM.cls_PhysicalDiagram, "", nothing, False)
        if not dgrm is nothing then
            dgrm.Name = .Cells(row, 3).Value            '指定 PhysicalDiagram名称
           'dgrm.Code = .Cells(row, 4).Value            '指定 PhysicalDiagram编码
            dgrm.Comment = .Cells(row, 5).Value         '指定 PhysicalDiagram注释
        Else
            output "|__"+CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +") PhysicalDiagram不存在，新增PhysicalDiagram。"
            CreatePhysicalDiagram par, exl, row
        end if
    End With
End sub

sub DeletePhysicalDiagram(par, exl, row)
    dim dgrm
    With exl.Workbooks(1).Worksheets(1)
        set dgrm = par.FindChildByCode(.Cells(row, 4).Value, PdPDM.cls_PhysicalDiagram, "", nothing, False)
        if not dgrm is nothing then
            dgrm.Delete
        end if
    End With
End sub
