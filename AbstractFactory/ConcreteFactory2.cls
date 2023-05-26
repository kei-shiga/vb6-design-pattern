VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConcreteFactory2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' 具体的なファクトリ2
Option Explicit

Implements IAbstractFactory

Public Function IAbstractFactory_CreateProductA() As IProductA
    Set IAbstractFactory_CreateProductA = New ConcreteProductA2
End Function

Public Function IAbstractFactory_CreateProductB() As IProductB
    Set IAbstractFactory_CreateProductB = New ConcreteProductB2
End Function
