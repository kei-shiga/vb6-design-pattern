Attribute VB_Name = "Sample"
Public Sub Main()
    Dim client As New client
    Dim factory1 As New ConcreteFactory1
    Dim factory2 As New ConcreteFactory2

    ' �t�@�N�g��1���g�p���Ď��s
    client.Run factory1

    ' �t�@�N�g��2���g�p���Ď��s
    client.Run factory2
End Sub
