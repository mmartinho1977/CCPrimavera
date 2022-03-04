Imports Microsoft.VisualBasic


Namespace BE

    Public Class LinhasPendenteCollection

        Inherits System.Collections.Specialized.NameObjectCollectionBase


        Public Overridable Function Add(ByVal LinhaPendente As BE.LinhaPendente, ByVal key As String) As Integer
            MyBase.BaseSet(key, LinhaPendente)
        End Function


        Default Public Overridable Property Item(ByVal index As Integer) As BE.LinhaPendente
            Get
                Return DirectCast(MyBase.BaseGet(index), BE.LinhaPendente)
            End Get
            Set(ByVal value As BE.LinhaPendente)
                MyBase.BaseSet(index, value)
            End Set
        End Property


        Default Public Overridable Property Item(ByVal key As String) As BE.LinhaPendente
            Get
                Return DirectCast(MyBase.BaseGet(key), BE.LinhaPendente)
            End Get
            Set(ByVal value As BE.LinhaPendente)
                MyBase.BaseSet(key, value)
            End Set
        End Property

    End Class

End Namespace

