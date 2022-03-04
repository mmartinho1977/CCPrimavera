Imports Microsoft.VisualBasic


Namespace BE

    Public Class Artigo
        Private mArtigo As String
        Private mTipoArtigo As String
        Private mCodigoBarras As String
        Private mDescricao As String
        Private mDescricaoComercial As String
        Private mCaracteristicas As String
        Private mIva As String
        Private mArmazemSugestao As String
        Private mLocalizacaoSugestao As String
        Private mFamilia As String
        Private mSubFamilia As String
        Private mMarca As String
        Private mModelo As String
        Private mGarantia As String
        Private mFornecedorPrincipal As String
        Private mDataUltimaEntrada As Date
        Private mDataUltimaSaida As Date
        Private mPCM As Double
        Private mUPC As Double
        Private mStockActual As Double
        Private mUnidadeBase As String
        Private mUnidadeCompra As String
        Private mUnidadeVenda As String
        Private mUnidadeEntrada As String
        Private mUnidadeSaida As String
        Private mMovimentaStock As Boolean
        Private mObservacoes As String
        Private mDataUltimaActualizacao As Date
        Private mPercentagemIncidenciaIva As Double
        Private mPercentagemIvaDedutivel As Double
        Private mCamposUtilizador As BE.CamposUtilizadorCollection
        Private mEmModoEdicao As Boolean 'False = Entidade nova a inserir, true = Entidade existente a editar

        ' PROPRIEDADES INSERIDAS A 6 DE OUTUBRO DE 2012
        Private artMoeda As ArtigoMoeda
        Private metaInformacao As Dictionary(Of String, Object)


        Sub New()
            mArtigo = ""
            mTipoArtigo = ""
            mCodigoBarras = ""
            mDescricao = ""
            mDescricaoComercial = ""
            mCaracteristicas = ""
            mIva = ""
            mArmazemSugestao = ""
            mLocalizacaoSugestao = ""
            mFamilia = ""
            mSubFamilia = ""
            mMarca = ""
            mModelo = ""
            mGarantia = ""
            mFornecedorPrincipal = ""
            mDataUltimaEntrada = Now.Date
            mDataUltimaSaida = Now.Date
            mPCM = 0
            mUPC = 0
            mStockActual = 0
            mUnidadeBase = ""
            mUnidadeCompra = ""
            mUnidadeVenda = ""
            mUnidadeEntrada = ""
            mUnidadeSaida = ""
            mMovimentaStock = False
            mObservacoes = ""
            mDataUltimaActualizacao = Now.Date
            mPercentagemIncidenciaIva = 0
            mPercentagemIvaDedutivel = 0
            mCamposUtilizador = New BE.CamposUtilizadorCollection
            metaInformacao = Nothing
            artMoeda = Nothing
        End Sub


        Public Property Artigo() As String
            Get
                Return mArtigo
            End Get
            Set(ByVal value As String)
                mArtigo = value.TrimEnd
            End Set
        End Property

        Public Property TipoArtigo() As String
            Get
                Return mTipoArtigo
            End Get
            Set(ByVal value As String)
                mTipoArtigo = value
            End Set
        End Property

        Public Property CodigoBarras() As String
            Get
                Return mCodigoBarras
            End Get
            Set(ByVal value As String)
                mCodigoBarras = value
            End Set
        End Property

        Public Property Descricao() As String
            Get
                Return mDescricao
            End Get
            Set(ByVal value As String)
                mDescricao = value
            End Set
        End Property

        Public Property DescricaoComercial() As String
            Get
                Return mDescricaoComercial
            End Get
            Set(ByVal value As String)
                mDescricaoComercial = value
            End Set
        End Property

        Public Property Caracteristicas() As String
            Get
                Return mCaracteristicas
            End Get
            Set(ByVal value As String)
                mCaracteristicas = value
            End Set
        End Property

        Public Property ArmazemSugestao() As String
            Get
                Return mArmazemSugestao
            End Get
            Set(ByVal value As String)
                mArmazemSugestao = value
            End Set
        End Property

        Public Property LocalizacaoSugestao() As String
            Get
                Return mLocalizacaoSugestao
            End Get
            Set(ByVal value As String)
                mLocalizacaoSugestao = value
            End Set
        End Property

        Public Property Iva() As String
            Get
                Return mIva
            End Get
            Set(ByVal value As String)
                mIva = value
            End Set
        End Property

        Public Property Familia() As String
            Get
                Return mFamilia
            End Get
            Set(ByVal value As String)
                mFamilia = value
            End Set
        End Property

        Public Property SubFamilia() As String
            Get
                Return mSubFamilia
            End Get
            Set(ByVal value As String)
                mSubFamilia = value
            End Set
        End Property

        Public Property Marca() As String
            Get
                Return mMarca
            End Get
            Set(ByVal value As String)
                mMarca = value
            End Set
        End Property

        Public Property Modelo() As String
            Get
                Return mModelo
            End Get
            Set(ByVal value As String)
                mModelo = value
            End Set
        End Property

        Public Property Garantia() As String
            Get
                Return mGarantia
            End Get
            Set(ByVal value As String)
                mGarantia = value
            End Set
        End Property

        Public Property FornecedorPrincipal() As String
            Get
                Return mFornecedorPrincipal
            End Get
            Set(ByVal value As String)
                mFornecedorPrincipal = value
            End Set
        End Property

        Public Property DataUltimaEntrada() As Date
            Get
                Return mDataUltimaEntrada
            End Get
            Set(ByVal value As Date)
                mDataUltimaEntrada = value
            End Set
        End Property

        Public Property DataUltimaSaida() As Date
            Get
                Return mDataUltimaSaida
            End Get
            Set(ByVal value As Date)
                mDataUltimaSaida = value
            End Set
        End Property

        Public Property PCM() As Double
            Get
                Return mPCM
            End Get
            Set(ByVal value As Double)
                mPCM = value
            End Set
        End Property

        Public Property UPC() As Double
            Get
                Return mUPC
            End Get
            Set(ByVal value As Double)
                mUPC = value
            End Set
        End Property

        Public Property StockActual() As Double
            Get
                Return mStockActual
            End Get
            Set(ByVal value As Double)
                mStockActual = value
            End Set
        End Property

        Public Property UnidadeBase() As String
            Get
                Return mUnidadeBase
            End Get
            Set(ByVal value As String)
                mUnidadeBase = value
            End Set
        End Property

        Public Property UnidadeCompra() As String
            Get
                Return mUnidadeCompra
            End Get
            Set(ByVal value As String)
                mUnidadeCompra = value
            End Set
        End Property

        Public Property UnidadeVenda() As String
            Get
                Return mUnidadeVenda
            End Get
            Set(ByVal value As String)
                mUnidadeVenda = value
            End Set
        End Property

        Public Property UnidadeEntrada() As String
            Get
                Return mUnidadeEntrada
            End Get
            Set(ByVal value As String)
                mUnidadeEntrada = value
            End Set
        End Property

        Public Property UnidadeSaida() As String
            Get
                Return mUnidadeSaida
            End Get
            Set(ByVal value As String)
                mUnidadeSaida = value
            End Set
        End Property

        Public Property MovimentaStock() As Boolean
            Get
                Return mMovimentaStock
            End Get
            Set(ByVal value As Boolean)
                mMovimentaStock = value
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

        Public Property DataUltimaActualizacao() As Date
            Get
                Return mDataUltimaActualizacao
            End Get
            Set(ByVal value As Date)
                mDataUltimaActualizacao = value
            End Set
        End Property


        Public Property PercentagemIncidenciaIva() As Double
            Get
                Return mPercentagemIncidenciaIva
            End Get
            Set(ByVal value As Double)
                mPercentagemIncidenciaIva = value
            End Set
        End Property


        Public Property PercentagemIvaDedutivel() As Double
            Get
                Return mPercentagemIvaDedutivel
            End Get
            Set(ByVal value As Double)
                mPercentagemIvaDedutivel = value
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


        Public Property ArtigoMoeda() As ArtigoMoeda
            Get
                Return artMoeda
            End Get
            Set(ByVal value As ArtigoMoeda)
                artMoeda = value
            End Set
        End Property


        Public Property MetaInfo() As Dictionary(Of String, Object)
            Get
                Return MetaInformacao
            End Get
            Set(value As Dictionary(Of String, Object))
                metaInformacao = value
            End Set
        End Property


        Public Overrides Function ToString() As String
            Return mDescricao
        End Function


        Protected Overrides Sub Finalize()
            mCamposUtilizador = Nothing
            MyBase.Finalize()
        End Sub


    End Class

End Namespace

