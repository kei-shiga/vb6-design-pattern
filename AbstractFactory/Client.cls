VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Client"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' クライアントのクラス
Option Explicit

Private productA As IProductA
Private productB As IProductB

Public Sub Run(factory As IAbstractFactory)
    ' 製品Aを生成
    Set productA = factory.CreateProductA()
    ' 製品Bを生成
    Set productB = factory.CreateProductB()

    ' 生成した製品を使用
    productA.OperationA
    productB.OperationB
End Sub
