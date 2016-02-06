Public Class Statistics

    Private Sub Statistics_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DatabaseInfoTotalTEXTBOX.Text = ItemObjects.Count
        InfoBaseBUTTON_Click_1(sender, e)
    End Sub

    Private Sub InfoQualityBUTTON_Click_1(sender As Object, e As EventArgs) Handles InfoQualityBUTTON.Click
        'DatabaseInfoTABCONTROL.SelectTab(0)
        DatabaseInfoDATAGRIDVIEW.Rows.Clear()
        ItemBaseList.Clear()
        ItemBaseValues.Clear()
        Dim arraypointer As Integer

        For index = 0 To ItemObjects.Count - 1
            If ItemBaseList.Contains(ItemObjects(index).ItemQuality) = False Then
                ItemBaseList.Add(ItemObjects(index).ItemQuality)
                ItemBaseValues.Add(1)
                Continue For
            End If
            arraypointer = ItemBaseList.IndexOf(ItemObjects(index).ItemQuality)
            If arraypointer >= 0 Then
                ItemBaseValues(arraypointer) = ItemBaseValues(arraypointer) + 1
            End If
        Next

        For index = 0 To ItemBaseList.Count - 1
            DatabaseInfoDATAGRIDVIEW.Rows.Add(ItemBaseList(index), ItemBaseValues(index), ItemBaseValues(index) / ItemObjects.Count * 100)
        Next
    End Sub

    Private Sub InfoRealmBUTTON_Click_1(sender As Object, e As EventArgs) Handles InfoRealmBUTTON.Click

        'DatabaseInfoTABCONTROL.SelectTab(0)
        DatabaseInfoDATAGRIDVIEW.Rows.Clear()
        ItemBaseList.Clear()
        ItemBaseValues.Clear()
        Dim arraypointer As Integer

        For index = 0 To ItemObjects.Count - 1
            If ItemBaseList.Contains(ItemObjects(index).ItemRealm) = False Then
                ItemBaseList.Add(ItemObjects(index).ItemRealm)
                ItemBaseValues.Add(1)
                Continue For
            End If
            arraypointer = ItemBaseList.IndexOf(ItemObjects(index).ItemRealm)
            If arraypointer >= 0 Then
                ItemBaseValues(arraypointer) = ItemBaseValues(arraypointer) + 1
            End If
        Next

        For index = 0 To ItemBaseList.Count - 1
            DatabaseInfoDATAGRIDVIEW.Rows.Add(ItemBaseList(index), ItemBaseValues(index), ItemBaseValues(index) / ItemObjects.Count * 100)
        Next
    End Sub

    Private Sub InfoBaseBUTTON_Click_1(sender As Object, e As EventArgs) Handles InfoBaseBUTTON.Click
        'DatabaseInfoTABCONTROL.SelectTab(0)
        DatabaseInfoDATAGRIDVIEW.Rows.Clear()
        ItemBaseList.Clear()
        ItemBaseValues.Clear()
        Dim arraypointer As Integer

        For index = 0 To ItemObjects.Count - 1
            If ItemBaseList.Contains(ItemObjects(index).ItemBase) = False Then
                ItemBaseList.Add(ItemObjects(index).ItemBase)
                ItemBaseValues.Add(1)
                Continue For
            End If
            arraypointer = ItemBaseList.IndexOf(ItemObjects(index).ItemBase)
            If arraypointer >= 0 Then
                ItemBaseValues(arraypointer) = ItemBaseValues(arraypointer) + 1
            End If
        Next

        For index = 0 To ItemBaseList.Count - 1
            DatabaseInfoDATAGRIDVIEW.Rows.Add(ItemBaseList(index), ItemBaseValues(index), ItemBaseValues(index) / ItemObjects.Count * 100)
        Next
    End Sub

    Private Sub IntoRuneBUTTON_Click_1(sender As Object, e As EventArgs) Handles IntoRuneBUTTON.Click
        'DatabaseInfoTABCONTROL.SelectTab(0)
        DatabaseInfoDATAGRIDVIEW.Rows.Clear()
        ItemBaseList.Clear()
        ItemBaseValues.Clear()
        Dim arraypointer As Integer

        For index = 0 To ItemObjects.Count - 1
            If ItemObjects(index).ItemBase = "Rune" Then
                If ItemBaseList.Contains(ItemObjects(index).ItemName) = False Then
                    ItemBaseList.Add(ItemObjects(index).ItemName)
                    ItemBaseValues.Add(1)
                    Continue For
                End If
                arraypointer = ItemBaseList.IndexOf(ItemObjects(index).ItemName)
                If arraypointer >= 0 Then
                    ItemBaseValues(arraypointer) = ItemBaseValues(arraypointer) + 1
                End If
            End If

        Next

        For index = 0 To ItemBaseList.Count - 1
            DatabaseInfoDATAGRIDVIEW.Rows.Add(ItemBaseList(index), ItemBaseValues(index), ItemBaseValues(index) / ItemObjects.Count * 100)
        Next
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub
End Class