Imports System.Reflection

Partial Public Class AdvancedTreeControl
    ' Funcția pentru adăugarea unui element nou în arbore
    Public Function AddItem(pKey As String, pCaption As String,
                            Optional pParent As TreeItem = Nothing,
                            Optional pLeftIconClosed As Image = Nothing,
                            Optional pLeftIconOpen As Image = Nothing,
                            Optional pRightIcon As Image = Nothing,
                            Optional pTag As String = Nothing,
                            Optional pExpanded As Boolean = False,
                            Optional pLazyNode As Boolean = False) As TreeItem

        Dim it As New TreeItem With {
            .Key = pKey,
            .Tag = pTag,
            .Caption = pCaption,
            .Parent = pParent,
            .LeftIconClosed = pLeftIconClosed,
            .LeftIconOpen = pLeftIconOpen,
            .RightIcon = pRightIcon,
            .Expanded = pExpanded,
            .LazyNode = pLazyNode
        }

        If pParent Is Nothing Then
            it.Level = 0
            Items.Add(it)
        Else
            it.Level = pParent.Level + 1
            pParent.Children.Add(it)
        End If

        Me.Invalidate()
        Return it
    End Function

    ' Funcția care primește string-ul din VBA și returnează valoarea
    Public Function ProcessPropertyRequest(cmd As String) As String
        ' Format așteptat: "GET_PROPERTY||PropName||[OptionalNodeID]"
        Dim parts() As String = cmd.Split(separator, StringSplitOptions.None)

        If parts.Length < 2 Then Return "ERROR: Invalid Format"

        Dim propName As String = parts(1)
        Dim result As String = "NOT_FOUND"

        Try
            ' === CAZUL 1: PROPRIETATE A CONTROLULUI (GLOBAL) ===
            If parts.Length = 2 Then
                ' Căutăm proprietatea în clasa AdvancedTreeControl (Me)
                Dim propInfo As PropertyInfo = Me.GetType().GetProperty(propName, BindingFlags.Public Or BindingFlags.Instance Or BindingFlags.IgnoreCase)

                If propInfo IsNot Nothing Then
                    Dim val = propInfo.GetValue(Me, Nothing)
                    result = FormatValue(val)
                Else
                    result = "ERROR: Property '" & propName & "' not found on Tree."
                End If

                ' === CAZUL 2: PROPRIETATE A UNUI NOD ===
            ElseIf parts.Length = 3 Then
                Dim nodeID As String = parts(2)

                ' 1. Găsim nodul după ID (care e Key în VBA)
                Dim node As TreeItem = FindNodeByID(nodeID)

                If node IsNot Nothing Then
                    ' 2. Căutăm proprietatea în clasa TreeItem
                    Dim propInfo As PropertyInfo = node.GetType().GetProperty(propName, BindingFlags.Public Or BindingFlags.Instance Or BindingFlags.IgnoreCase)

                    If propInfo IsNot Nothing Then
                        Dim val = propInfo.GetValue(node, Nothing)
                        result = FormatValue(val)
                    Else
                        result = "ERROR: Property '" & propName & "' not found on Node."
                    End If
                Else
                    result = "ERROR: Node with ID '" & nodeID & "' not found."
                End If
            End If

        Catch ex As Exception
            result = "ERROR: " & ex.Message
        End Try

        Return result
    End Function

    ' Metodă publică pentru a seta starea checkbox-ului din exterior (VBA) cu propagare
    Public Sub SetItemCheckState(pItem As TreeItem, pState As TreeCheckState)
        SetNodeStateWithPropagation(pItem, pState)
        Me.Invalidate()
    End Sub

    ' Metodă publică pentru a goli toate elementele din control
    ' Metodă publică pentru a goli toate elementele din control
    Public Sub Clear()
        ' 1. Oprim orice desenare sau calcul de layout
        Me.SuspendLayout()

        ' 2. Golim datele
        Items.Clear()
        pSelectedItem = Nothing
        pHoveredItem = Nothing

        ' 3. Resetăm Scroll-ul la zero (CRITIC)
        Me.AutoScrollPosition = New Point(0, 0)
        Me.AutoScrollMinSize = Size.Empty

        ' 4. Repornim logica
        Me.ResumeLayout(False)
        Me.PerformLayout()

        ' 5. Forțăm redesenarea imediată a întregului control
        Me.Invalidate()
        Me.Update()
    End Sub
End Class
