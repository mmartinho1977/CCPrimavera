Imports Microsoft.VisualBasic


Namespace BE

    Public Class EntidadeAssociada
        Private mTipoEntidade As String
        Private mEntidade As String
        Private mTipoEntidadeAssociada As String
        Private mEntidadeAssociada As String
        Private mCamposUtilizador As BE.CamposUtilizadorCollection
        Private mEmModoEdicao As Boolean


        Sub New()
            mCamposUtilizador = New BE.CamposUtilizadorCollection
            mEmModoEdicao = False
        End Sub


        Public Property TipoEntidade() As String
            Get
                Return mTipoEntidade
            End Get
            Set(ByVal value As String)
                mTipoEntidade = value
            End Set
        End Property

        Public Property Entidade() As String
            Get
                Return mEntidade
            End Get
            Set(ByVal value As String)
                mEntidade = value
            End Set
        End Property

        Public Property TipoEntidadeAssociada() As String
            Get
                Return mTipoEntidadeAssociada
            End Get
            Set(ByVal value As String)
                mTipoEntidadeAssociada = value
            End Set
        End Property

        Public Property EntidadeAssociada() As String
            Get
                Return mEntidadeAssociada
            End Get
            Set(ByVal value As String)
                mEntidadeAssociada = value
            End Set
        End Property

        Public Property CamposUtilizador() As BE.CamposUtilizadorCollection
            Get
                Return mCamposUtilizador
            End Get
            Set(ByVal value As BE.CamposUtilizadorCollection)
                mCamposUtilizador = value
            End Set
        End Property

        Public Property EmModoEdicao() As Boolean
            Get
                Return mEmModoEdicao
            End Get
            Set(ByVal value As Boolean)
                mEmModoEdicao = value
            End Set
        End Property

    End Class

End Namespace

