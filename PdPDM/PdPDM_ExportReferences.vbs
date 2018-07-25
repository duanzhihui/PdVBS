'******************************************************************************
'* File       : PdPDM_ExportReferences.vbs
'* Purpose    : 导出引用到Excle
'* Title      : 导出引用
'* Category   : 导出模型
'* Version    : v1.0
'* Company    : www.duanzhihui.com
'* Author     : 段智慧
'* Description: 导出表到Excle
'* History    : 2018-04-04  v1.0    段智慧  新增脚本。
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

    Dim tplt, path, ws
    set ws=CreateObject("WScript.Shell")
    tplt = ws.CurrentDirectory + "\PdPDM_ExportReferences_Template.xlsx"
    path = ws.CurrentDirectory + "\PdPDM_ExportReferences_xxx.xlsx"

    path = InputBox ("请输入包含模型的Excel文件路径。", "文件路径", path)
    output "Excel文件路径为: " + path

    CreateObject("Scripting.FileSystemObject").CopyFile tplt, path, true

    exl.Workbooks.Open path
Else
    HaveExcel = False
End If

'on error Resume Next

Dim fldr, nb1, nb2, ar1(50000, 9), ar2(100000, 2)
Set fldr = ActiveDiagram.Parent
nb1 =0
nb2 =0

ListObjects(fldr)
exl.Workbooks(1).ForceFullCalculation = False
output CStr(now) + " Write list of references"
exl.Workbooks(1).Worksheets("References").Activate
exl.Range("A2").Resize(nb1, 10).Value = ar1
output CStr(now) + " Write list of Joins"
exl.Workbooks(1).Worksheets("Joins").Activate
exl.Range("C2").Resize(nb2, 3).Value = ar2
output CStr(now) + " Save excle"
exl.Workbooks(1).ForceFullCalculation = True
exl.Workbooks(1).Close True

Sub ListObjects(fldr)
    output CStr(now) + " Scan "+fldr.ClassName+" "+fldr.Code
    Dim ref, jn
    For Each ref In fldr.References
        if not ref.IsShortcut then
            ar1(nb1, 0) = "R"
            ar1(nb1, 1) = ref.Parent.Code
            ar1(nb1, 2) = ref.Name
            ar1(nb1, 3) = ref.Code
            ar1(nb1, 4) = ref.Comment
            ar1(nb1, 5) = ref.ParentTable.Code
            ar1(nb1, 6) = ref.ChildTable.Code
            ar1(nb1, 7) = ref.Cardinality
            ar1(nb1, 8) = ref.ParentRole
            ar1(nb1, 9) = ref.ChildRole
            nb1 = nb1 + 1
            
            output CStr(now) + " List "+ref.ClassName+" "+ref.Code + " Joins"
            for Each jn in ref.Joins
                ar2(nb2, 0) = jn.Parent.Code
                ar2(nb2, 1) = jn.ParentTableColumn.Code
                ar2(nb2, 2) = jn.ChildTableColumn.Code
                nb2 = nb2+1
            Next
        end if
    Next

    ' go into the sub-packages
    Dim f ' running folder
    For Each f In fldr.Packages
        ListObjects f
    Next
End Sub
