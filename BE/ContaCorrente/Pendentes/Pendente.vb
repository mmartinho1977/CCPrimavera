Imports Microsoft.VisualBasic


Namespace BE

    Public Class Pendente
        Private mIdHistorico As String
        Private mFilial As String
        Private mModulo As String
        Private mTipoDocumento As String
        Private mSerieDocumento As String
        Private mNumeroDocumento As String
        Private mNumeroDocumentoInterno As System.Int16
        Private mTipoEntidade As String
        Private mEntidade As String
        Private mTipoConta As String
        Private mEstado As String
        Private mDataDocumento As Date
        Private mDataVencimento As Date
        Private mDataIntroducao As Date
        Private mCondicaoPagamento As String
        Private mModoPagamento As String
        Private mMoeda As String
        Private mValorTotal As Decimal
        Private mValorPendente As Decimal
        Private mObservacoes As String
        Private mUtilizador As String
        Private mLinhas As BE.LinhasPendenteCollection
        Private mCamposUtilizador As BE.CamposUtilizadorCollection
        Private mEmModoEdicao As Boolean 'False = Entidade nova a inserir, true = Entidade existente a editar


        Sub New()
            mIdHistorico = Guid.NewGuid.ToString
            mFilial = "000"
            mModulo = "M"
            mTipoDocumento = ""
            mSerieDocumento = ""
            mNumeroDocumento = ""
            mNumeroDocumentoInterno = 0
            mTipoEntidade = "C"
            mEntidade = ""
            mTipoConta = ""
            mEstado = ""
            mDataDocumento = Now().Date
            mDataVencimento = Now().Date
            mDataIntroducao = Now().Date
            mCondicaoPagamento = "1"
            mModoPagamento = "NUM"
            mMoeda = "EUR"
            mValorTotal = 0
            mValorPendente = 0
            mObservacoes = ""
            mUtilizador = ""
            mLinhas = New BE.LinhasPendenteCollection
            mCamposUtilizador = New BE.CamposUtilizadorCollection
        End Sub

        Public Property IDHistorico() As String
            Get
                Return mIdHistorico
            End Get
            Set(ByVal value As String)
                mIdHistorico = value
            End Set
        End Property

        Public Property Filial() As String
            Get
                Return mFilial
            End Get
            Set(ByVal value As String)
                mFilial = value
            End Set
        End Property

        Public Property Modulo() As String
            Get
                Return mModulo
            End Get
            Set(ByVal value As String)
                mModulo = value
            End Set
        End Property

        Public Property TipoDocumento() As String
            Get
                Return mTipoDocumento
            End Get
            Set(ByVal value As String)
                mTipoDocumento = value
            End Set
        End Property

        Public Property SerieDocumento() As String
            Get
                Return mSerieDocumento
            End Get
            Set(ByVal value As String)
                mSerieDocumento = value
            End Set
        End Property

        Public Property NumeroDocumento() As String
            Get
                Return mNumeroDocumento
            End Get
            Set(ByVal value As String)
                mNumeroDocumento = value
            End Set
        End Property

        Public Property NumeroDocumentoInterno() As System.Int64
            Get
                Return mNumeroDocumentoInterno
            End Get
            Set(ByVal value As System.Int64)
                mNumeroDocumentoInterno = value
            End Set
        End Property

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


        Public Property TipoConta() As String
            Get
                Return mTipoConta
            End Get
            Set(ByVal value As String)
                mTipoConta = value
            End Set
        End Property


        Public Property Estado() As String
            Get
                Return mEstado
            End Get
            Set(ByVal value As String)
                mEstado = value
            End Set
        End Property


        Public Property DataDocumento() As Date
            Get
                Return mDataDocumento
            End Get
            Set(ByVal value As Date)
                mDataDocumento = value
            End Set
        End Property

        Public Property DataVencimento() As Date
            Get
                Return mDataVencimento
            End Get
            Set(ByVal value As Date)
                mDataVencimento = value
            End Set
        End Property

        Public Property DataIntroducao() As Date
            Get
                Return mDataIntroducao
            End Get
            Set(ByVal value As Date)
                mDataIntroducao = value
            End Set
        End Property

        Public Property CondicaoPagamento() As String
            Get
                Return mCondicaoPagamento
            End Get
            Set(ByVal value As String)
                mCondicaoPagamento = value
            End Set
        End Property

        Public Property ModoPagamento() As String
            Get
                Return mModoPagamento
            End Get
            Set(ByVal value As String)
                mModoPagamento = value
            End Set
        End Property

        Public Property Moeda() As String
            Get
                Return mMoeda
            End Get
            Set(ByVal value As String)
                mMoeda = value
            End Set
        End Property

        Public Property ValorTotal() As Decimal
            Get
                Return mValorTotal
            End Get
            Set(ByVal value As Decimal)
                mValorTotal = value
            End Set
        End Property

        Public Property ValorPendente() As Decimal
            Get
                Return mValorPendente
            End Get
            Set(ByVal value As Decimal)
                mValorPendente = value
            End Set
        End Property

        Public Property Observacoes() As String
            Get
                Return mObservacoes
            End Get
            Set(ByVal value As String)
                mObservacoes = value
            End Set
        End Property

        Public Property Utilizador() As String
            Get
                Return mUtilizador
            End Get
            Set(ByVal value As String)
                mUtilizador = value
            End Set
        End Property

        Public Property Linhas() As BE.LinhasPendenteCollection
            Get
                Return mLinhas
            End Get
            Set(ByVal value As BE.LinhasPendenteCollection)
                mLinhas = value
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

        Protected Overrides Sub Finalize()
            mLinhas = Nothing
            mCamposUtilizador = Nothing
            MyBase.Finalize()
        End Sub

    End Class

End Namespace

