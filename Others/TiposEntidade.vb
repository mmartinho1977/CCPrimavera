Imports Microsoft.VisualBasic
Imports System.Collections.Generic



Namespace Others

    Public Class TiposEntidade
        Private Shared mTiposEntidade As List(Of TipoEntidade)


        Sub New()
            mTiposEntidade = New List(Of TipoEntidade)
            mTiposEntidade.Add(New TipoEntidade("C", "Cliente", "Clientes"))
            mTiposEntidade.Add(New TipoEntidade("F", "Fornecedor", "Fornecedores"))
            mTiposEntidade.Add(New TipoEntidade("D", "Outro devedor", "Outros devedores"))
            mTiposEntidade.Add(New TipoEntidade("R", "Outro credor", "Outros credores"))
        End Sub

        Public Shared ReadOnly Property TiposEntidade() As List(Of TipoEntidade)
            Get
                Return mTiposEntidade
            End Get
        End Property


        Public Shared ReadOnly Property Cliente() As String
            Get
                Return "C"
            End Get
        End Property


        Public Shared ReadOnly Property Fornecedor() As String
            Get
                Return "F"
            End Get
        End Property


        Public Shared ReadOnly Property OutroDevedor() As String
            Get
                Return "D"
            End Get
        End Property


        Public Shared ReadOnly Property OutroCredor() As String
            Get
                Return "R"
            End Get
        End Property

    End Class






    Public Structure TipoEntidade
        Dim mCodigo As String
        Dim mDescricaoSingular As String
        Dim mDescricaoPlural As String


        Sub New(ByVal codigo As String, ByVal descricaoSingular As String, ByVal descricaoPlural As String)
            mCodigo = codigo
            mDescricaoSingular = descricaoSingular
            mDescricaoPlural = descricaoPlural
        End Sub


        Public Property Codigo() As String
            Get
                Return mCodigo
            End Get
            Set(ByVal value As String)
                mCodigo = value
            End Set
        End Property


        Public Property DescricaoSingular() As String
            Get
                Return mDescricaoSingular
            End Get
            Set(ByVal value As String)
                mDescricaoSingular = value
            End Set
        End Property


        Public Property DescricaoPlural() As String
            Get
                Return mDescricaoPlural
            End Get
            Set(ByVal value As String)
                mDescricaoPlural = value
            End Set
        End Property


    End Structure

End Namespace

