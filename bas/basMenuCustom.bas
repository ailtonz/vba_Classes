Attribute VB_Name = "basMenuCustom"
Sub MenuCustom()
Dim cbWSMenuBar As CommandBar
Dim Ctrl As CommandBarControl, muCustom As CommandBarControl
Dim iHelpIndex As Integer

Set cbWSMenuBar = Application.CommandBars("Worksheet menu bar")
iHelpIndex = cbWSMenuBar.Controls("Ajuda").index

Set muCustom = cbWSMenuBar.Controls.add(Type:=msoControlPopup, before:=iHelpIndex, temporary:=True)
For Each Ctrl In cbWSMenuBar.Controls
    If Ctrl.Caption = "&Programas do MrExcel" Then
        cbWSMenuBar.Controls("Programas do MrExcel").Delete
    End If
Next Ctrl

With muCustom
    .Caption = "&Programas do MrExcel"
    With .Controls.add(Type:=msoControlButton)
        .Caption = "importar e formatar"
        .OnAction = "ImportFormat"
    End With
    
    With .Controls.add(Type:=msoControlButton)
        .Caption = "Calcular fim de ano"
        .OnAction = "CalcYearEnd"
    End With
End With


End Sub
