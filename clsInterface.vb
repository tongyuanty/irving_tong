Public Class clsInterface
    Dim frm As New Form1
    '这个函数专门用来扫描
    Public Function Scan(ByVal dictPara As KFO.Dictionary)
        Return frm.Scan(dictPara)
    End Function
    '这个函数专门用来openscan
    Public Sub IntoScan()

    End Sub
    '这个函数专门用来closescan
    Public Sub CloseScan()

        frm.Close()
        frm.Dispose()
        frm = Nothing

    End Sub
End Class

