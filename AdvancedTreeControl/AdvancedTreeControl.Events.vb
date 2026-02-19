Partial Public Class AdvancedTreeControl
    Public Event NodeMouseDown(pNode As TreeItem, e As MouseEventArgs)
    Public Event NodeMouseUp(pNode As TreeItem, e As MouseEventArgs)
    Public Event NodeDoubleClicked(pNode As TreeItem, e As MouseEventArgs)
    Public Event NodeChecked(pNode As TreeItem)
    Public Event NodeRadioSelected(nodeOn As TreeItem, nodeOff As TreeItem)
    Public Event RequestLazyLoad(sender As Object, item As TreeItem)
    Public Event RightIconClicked(pNode As TreeItem, e As MouseEventArgs)
End Class
