'******************************************************************************
'* File:        PdPDM_ImportPackages.vbs
'* Purpose:     从Sheet中导入Packages
'* Title:
'* Category:
'* Version:     1.1
'* Company:
'* Author:      段智慧
'* Description:
'*              CRUD
'*                  C 新增  如果Package存在则忽略，否则新增。
'*                  R 只读  直接忽略。
'*                  U 更新  如果Package存在则更新，否则新增。
'*                  D 删除  删除Package。
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
    path = ws.CurrentDirectory + "\PdPDM_ImportPackages.xlsx"

    path = InputBox ("请输入包含模型的Excel文件路径。", "文件路径", path)
    output "Excel文件路径为: " + path

    exl.Workbooks.Open path
    exl.Workbooks(1).Worksheets(1).Activate    '指定要打开的sheet名称
Else
    HaveExcel = False
End If

dim par, pkg, row
on error Resume Next

row = 2
With exl.Workbooks(1).Worksheets(1)
    do While .Cells(row, 1).Value <> ""                     '退出

        set par = mdl.FindChildByCode(CStr(.Cells(row, 2).Value), PdPDM.cls_Package)
        if par is nothing then
            set par = mdl
        end if

        select case .Cells(row, 1).Value
        case "C"
            output "第" + CStr(row) + "行，新增Package：" + CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +")。"
            CreatePackage par, exl, row
        case "R"
            output "第" + CStr(row) + "行，只读Package：" + CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +")。"
        case "U"
            output "第" + CStr(row) + "行，更新Package：" + CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +")。"
            UpdatePackage par, exl, row
        case "D"
            output "第" + CStr(row) + "行，删除Package：" + CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +")。"
            DeletePackage par, exl, row
        case Else
            output "第" + CStr(row) + "行，忽略Package：" + CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +")。"
        end select
        exl.Range("A"+Cstr(row)).Value = "R"
        row = row + 1
    Loop
End With

exl.Workbooks(1).Close True

sub CreatePackage(par, exl, row)
    dim pkg
    With exl.Workbooks(1).Worksheets(1)
        set pkg = par.FindChildByCode(.Cells(row, 4).Value, PdPDM.cls_Package)
        if not pkg is nothing then
            output "|__"+CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +") Package存在，忽略Package。"
        Else
            set pkg = par.Packages.CreateNew            '创建 Package
            pkg.Name = .Cells(row, 3).Value             '指定 Package名称
            pkg.Code = .Cells(row, 4).Value             '指定 Package编码
            pkg.Comment = .Cells(row, 5).Value          '指定 Package注释
            pkg.DefaultDiagram.Name =  pkg.Name         '指定 DefaultDiagram名称
            pkg.DefaultDiagram.Code =  pkg.Code         '指定 DefaultDiagram编码
        end if
    End With
End sub

sub UpdatePackage(par, exl, row)
    dim pkg
    With exl.Workbooks(1).Worksheets(1)
        set pkg = par.FindChildByCode(.Cells(row, 4).Value, PdPDM.cls_Package)
        if not pkg is nothing then
            pkg.Name = .Cells(row, 3).Value             '指定 Package名称
           'pkg.Code = .Cells(row, 4).Value             '指定 Package编码
            pkg.Comment = .Cells(row, 5).Value          '指定 Package注释
            pkg.DefaultDiagram.Name =  pkg.Name         '指定 DefaultDiagram名称
            pkg.DefaultDiagram.Code =  pkg.Code         '指定 DefaultDiagram编码
        Else
            output "|__"+CStr(.Cells(row, 3).Value) + "(" + CStr(.Cells(row, 4).Value) +") Package不存在，新增Package。"
            CreatePackage par, exl, row
        end if
    End With
End sub

sub DeletePackage(par, exl, row)
    dim pkg
    With exl.Workbooks(1).Worksheets(1)
        set pkg = par.FindChildByCode(.Cells(row, 4).Value, PdPDM.cls_Package)
        if not pkg is nothing then
            pkg.Delete
        end if
    End With
End sub
