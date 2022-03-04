Imports Microsoft.VisualBasic


Namespace BE

    Public Class LinhaPendente
        Private mId As String
        Private mDescricao As String
        Private mPercentagemIvaDedutivel As Decimal
        Private mTaxaProRata As Decimal
        Private mValorRecargo As Decimal
        Private mCodigoIva As String
        Private mValorIncidencia As Decimal
        Private mValorIva As Decimal
        Private mValorTotal As Decimal
        Private mCBLLigacaoGeral As String
        Private mCBLLigacaoAnalitica As String
        Private mCBLLigacaoCentrosCusto As String
        Private mCBLLigacaoFuncional As String
        Private mDataUltimaActualizacao As Date


        Sub New()
            mId = Guid.NewGuid.ToString
            mDescricao = ""
            mPercentagemIvaDedutivel = 0
            mTaxaProRata = 0
            mValorRecargo = 0
            mCodigoIva = "00"
            mValorIncidencia = 0
            mValorIva = 0
            mValorTotal = 0
            mCBLLigacaoGeral = ""
            mCBLLigacaoAnalitica = ""
            mCBLLigacaoCentrosCusto = ""
            mCBLLigacaoFuncional = ""
            mDataUltimaActualizacao = Now().Date
        End Sub

        Public Property Id() As String
            Get
                Return mId
            End Get
            Set(ByVal value As String)
                mId = value
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

        Public Property PercentagemIvaDedutivel() As Decimal
            Get
                Return mPercentagemIvaDedutivel
            End Get
            Set(ByVal value As Decimal)
                mPercentagemIvaDedutivel = value
            End Set
        End Property

        Public Property TaxaProRata() As Decimal
            Get
                Return mTaxaProRata
            End Get
            Set(ByVal value As Decimal)
                mTaxaProRata = value
            End Set
        End Property

        Public Property ValorRecargo() As Decimal
            Get
                Return mValorRecargo
            End Get
            Set(ByVal value As Decimal)
                mValorRecargo = value
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

        Public Property ValorIncidencia() As Decimal
            Get
                Return mValorIncidencia
            End Get
            Set(ByVal value As Decimal)
                mValorIncidencia = value
            End Set
        End Property

        Public Property ValorIva() As Decimal
            Get
                Return mValorIva
            End Get
            Set(ByVal value As Decimal)
                mValorIva = value
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


        Public Property CBLLigacaoGeral() As String
            Get
                Return mCBLLigacaoGeral
            End Get
            Set(ByVal value As String)
                mCBLLigacaoGeral = value
            End Set
        End Property


        Public Property CBLLigacaoAnalitica() As String
            Get
                Return mCBLLigacaoAnalitica
            End Get
            Set(ByVal value As String)
                mCBLLigacaoAnalitica = value
            End Set
        End Property


        Public Property CBLLigacaoCentrosCusto() As String
            Get
                Return mCBLLigacaoCentrosCusto
            End Get
            Set(ByVal value As String)
                mCBLLigacaoCentrosCusto = value
            End Set
        End Property


        Public Property CBLLigacaoFuncional() As String
            Get
                Return mCBLLigacaoFuncional
            End Get
            Set(ByVal value As String)
                mCBLLigacaoFuncional = value
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


        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub

    End Class

End Namespace

