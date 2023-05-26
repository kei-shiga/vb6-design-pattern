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
' �N���C�A���g�̃N���X
Option Explicit

Private productA As IProductA
Private productB As IProductB

Public Sub Run(factory As IAbstractFactory)
    ' ���iA�𐶐�
    Set productA = factory.CreateProductA()
    ' ���iB�𐶐�
    Set productB = factory.CreateProductB()

    ' �����������i���g�p
    productA.OperationA
    productB.OperationB
End Sub