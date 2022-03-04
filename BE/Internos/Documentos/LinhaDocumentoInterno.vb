Imports Microsoft.VisualBasic
Imports CCPrimavera.Specifications.LinhaDocumentoInterno_Specs


Namespace BE

    Public Class LinhaDocumentoInterno
        Private mId As String
        Private mNumeroLinha As System.Int32
        Private mTipoLinha As TiposLinhas
        Private mArtigo As String
        Private mArmazem As String
        Private mLote As String
        Private mLocalizacao As String
        Private mDescricao As String
        Private mCodigoIva As String
        Private mDesconto1 As Single
        Private mDesconto2 As Single
        Private mDesconto3 As Single
        Private mMovimentaStock As Boolean
        Private mQuantidade As Double
        Private mQuantidadeSatisfeita As Double
        Private mDataEntrega As Date
        Private mDataStock As Date
        Private mPrecoUnitario As Double
        Private mPrecoMedioCusto As Double
        Private mUnidade As String
        Private mTaxaIva As Single
        Private mPrecoLiquido As Double
        Private mPercentagemIncidenciaIva As Single
        Private mPercentagemIvaDedutivel As Double
        Private mIvaNaoDedutivel As Double
        Private mRegimeIva As String
        Private mTaxaProRata As Double
        Private mModuloOrigemCopia As String
        Private mIdLinhaOrigemCopia As String


        Private mCamposUtilizador As BE.CamposUtilizadorCollection


        Sub New()
            mId = Guid.NewGuid.ToString
            mNumeroLinha = 0
            mTipoLinha = TiposLinhas.Mercadoria_TipoArtigo_3_TipoLinha_10
            mArtigo = ""
            mArmazem = ""
            mLote = ""
            mLocalizacao = ""
            mDescricao = ""
            mCodigoIva = ""
            mDesconto1 = 0
            mDesconto2 = 0
            mDesconto3 = 0
            mMovimentaStock = False
            mQuantidade = 0
            mQuantidadeSatisfeita = 0
            mDataEntrega = Now
            mDataStock = Now
            mPrecoUnitario = 0
            mPrecoMedioCusto = 0
            mUnidade = ""
            mTaxaIva = 0
            mPrecoLiquido = 0
            mPercentagemIncidenciaIva = 0
            mPercentagemIvaDedutivel = 0
            mIvaNaoDedutivel = 0
            mRegimeIva = ""
            mModuloOrigemCopia = ""
            mIdLinhaOrigemCopia = ""

            mCamposUtilizador = New BE.CamposUtilizadorCollection
        End Sub

        Public Property Id() As String
            Get
                Return mId
            End Get
            Set(ByVal value As String)
                mId = value
            End Set
        End Property

        Public Property NumeroLinha() As System.Int32
            Get
                Return mNumeroLinha
            End Get
            Set(ByVal value As System.Int32)
                mNumeroLinha = value
            End Set
        End Property

        Public Property TipoLinha() As TiposLinhas
            Get
                Return mTipoLinha
            End Get
            Set(ByVal value As TiposLinhas)
                mTipoLinha = value
            End Set
        End Property

        Public Property Artigo() As String
            Get
                Return mArtigo
            End Get
            Set(ByVal value As String)
                mArtigo = value
            End Set
        End Property

        Public Property Armazem() As String
            Get
                Return mArmazem
            End Get
            Set(ByVal value As String)
                mArmazem = value
            End Set
        End Property

        Public Property Lote() As String
            Get
                Return mLote
            End Get
            Set(ByVal value As String)
                mLote = value
            End Set
        End Property

        Public Property Localizacao() As String
            Get
                Return mLocalizacao
            End Get
            Set(ByVal value As String)
                mLocalizacao = value
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

        Public Property CodigoIva() As String
            Get
                Return mCodigoIva
            End Get
            Set(ByVal value As String)
                mCodigoIva = value
            End Set
        End Property

        Public Property Desconto1() As Single
            Get
                Return mDesconto1
            End Get
            Set(ByVal value As Single)
                mDesconto1 = value
            End Set
        End Property

        Public Property Desconto2() As Single
            Get
                Return mDesconto2
            End Get
            Set(ByVal value As Single)
                mDesconto2 = value
            End Set
        End Property

        Public Property Desconto3() As Single
            Get
                Return mDesconto3
            End Get
            Set(ByVal value As Single)
                mDesconto3 = value
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

        Public Property Quantidade() As Double
            Get
                Return mQuantidade
            End Get
            Set(ByVal value As Double)
                mQuantidade = value
            End Set
        End Property

        Public Property QuantidadeSatisfeita() As Double
            Get
                Return mQuantidadeSatisfeita
            End Get
            Set(ByVal value As Double)
                mQuantidadeSatisfeita = value
            End Set
        End Property

        Public Property DataEntrega() As Date
            Get
                Return mDataEntrega
            End Get
            Set(ByVal value As Date)
                mDataEntrega = value
            End Set
        End Property

        Public Property DataStock() As Date
            Get
                Return mDataStock
            End Get
            Set(ByVal value As Date)
                mDataStock = value
            End Set
        End Property

        Public Property PrecoUnitario() As Double
            Get
                Return mPrecoUnitario
            End Get
            Set(ByVal value As Double)
                mPrecoUnitario = value
            End Set
        End Property

        Public Property PrecoMedioCusto() As Double
            Get
                Return mPrecoMedioCusto
            End Get
            Set(ByVal value As Double)
                mPrecoMedioCusto = value
            End Set
        End Property

        Public Property Unidade() As String
            Get
                Return mUnidade
            End Get
            Set(ByVal value As String)
                mUnidade = value
            End Set
        End Property

        Public Property TaxaIva() As Single
            Get
                Return mTaxaIva
            End Get
            Set(ByVal value As Single)
                mTaxaIva = value
            End Set
        End Property

        Public Property PrecoLiquido() As Double
            Get
                Return mPrecoLiquido
            End Get
            Set(ByVal value As Double)
                mPrecoLiquido = value
            End Set
        End Property

        Public Property PercentagemIncidenciaIva() As Single
            Get
                Return mPercentagemIncidenciaIva
            End Get
            Set(ByVal value As Single)
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

        Public Property IvaNaoDedutivel() As Double
            Get
                Return mIvaNaoDedutivel
            End Get
            Set(ByVal value As Double)
                mIvaNaoDedutivel = value
            End Set
        End Property

        Public Property RegimeIva() As String
            Get
                Return mRegimeIva
            End Get
            Set(ByVal value As String)
                mRegimeIva = value
            End Set
        End Property

        Public Property TaxaProRata() As Double
            Get
                Return mTaxaProRata
            End Get
            Set(ByVal value As Double)
                mTaxaProRata = value
            End Set
        End Property


        Public Property ModuloOrigemCopia() As String
            Get
                Return mModuloOrigemCopia
            End Get
            Set(ByVal value As String)
                mModuloOrigemCopia = value
            End Set
        End Property


        Public Property IdLinhaOrigemCopia() As String
            Get
                Return mIdLinhaOrigemCopia
            End Get
            Set(ByVal value As String)
                mIdLinhaOrigemCopia = value
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

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub

    End Class

End Namespace

