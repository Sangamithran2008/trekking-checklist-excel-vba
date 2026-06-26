Attribute VB_Name = "Module3"
Sub MarkItem()

    'This macro will mark an item as done by placing 1 in the next cell

    If ActiveCell.Value <> "" Then

        If ActiveCell.Offset(0, -1).Value <> 1 Then
            ActiveCell.Offset(0, -1).Value = 1
        Else
            ActiveCell.Offset(0, -1).Value = -1
        End If

    End If

End Sub
