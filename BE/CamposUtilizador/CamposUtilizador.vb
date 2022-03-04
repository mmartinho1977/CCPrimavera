Imports Microsoft.VisualBasic

Namespace BE

    Public Class CamposUtilizadorCollection
        Inherits System.Collections.Specialized.NameObjectCollectionBase


        Public Overridable Function Add(ByVal CampoUtilizador As BE.CampoUtilizador, ByVal key As String) As Integer
            MyBase.BaseSet(key, CampoUtilizador)
        End Function


        Default Public Overridable Property Item(ByVal index As Integer) As BE.CampoUtilizador
            Get
                Return DirectCast(MyBase.BaseGet(index), BE.CampoUtilizador)
            End Get
            Set(ByVal value As BE.CampoUtilizador)
                MyBase.BaseSet(index, value)
            End Set
        End Property


        Default Public Overridable Property Item(ByVal key As String) As BE.CampoUtilizador
            Get
                Return DirectCast(MyBase.BaseGet(key), BE.CampoUtilizador)
            End Get
            Set(ByVal value As BE.CampoUtilizador)
                MyBase.BaseSet(key, value)
            End Set
        End Property


    End Class

End Namespace
