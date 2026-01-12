Partial Public Class AdvancedTreeControl
    Public Event NodeMouseDown(pNode As TreeItem, e As MouseEventArgs)
    Public Event NodeMouseUp(pNode As TreeItem, e As MouseEventArgs)
    Public Event NodeDoubleClicked(pNode As TreeItem, e As MouseEventArgs)
    Public Event NodeChecked(pNode As TreeItem)
End Class
