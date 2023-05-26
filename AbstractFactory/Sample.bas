Attribute VB_Name = "Sample"
Public Sub Main()
    Dim client As New client
    Dim factory1 As New ConcreteFactory1
    Dim factory2 As New ConcreteFactory2

    ' ファクトリ1を使用して実行
    client.Run factory1

    ' ファクトリ2を使用して実行
    client.Run factory2
End Sub
