'*******************************************************************************
'* File       : PdPDM_AutoAttachTables.vbs
'* Purpose    : 自动附加表到默认视图
'* Title      : 自动附加表到默认视图
'* Category   : 自动附加
'* Version    : 1.0
'* Company    : www.duanzhihui.com
'* Author     : 段智慧
'* Description: 自动附加表到默认视图，如果是指定表的视图请用 PdPDM_AttachTables.vbs
'* History    :
'*              2016-03-31  v1.0    段智慧  新增
'******************************************************************************
Option Explicit

Dim mdl                                             ' the current model
Set mdl = ActiveModel
If (mdl Is Nothing) Then
   MsgBox "There is no Active Model"
End If

'on error Resume Next
Dim pkg, dgrm, tbl, sym
For Each pkg in mdl.Packages
    set dgrm = pkg.DefaultDiagram
    For Each tbl in pkg.Tables
        set sym = dgrm.FindSymbol(tbl)
        if sym is nothing Then
           dgrm.AttachObject tbl
        end if
    Next
    dgrm.AutoLayoutWithOptions 0, 1
Next
