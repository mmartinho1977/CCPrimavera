Imports Microsoft.VisualBasic


Namespace BE

    Public Class DocumentoVenda
        Private mId As String
        Private mFilial As String
        Private mSeccao As String
        Private mDocumento_Tipo As String
        Private mDocumento_Serie As String
        Private mDocumento_Numero As System.Int32
        Private mDocumento_Moeda As String
        Private mDocumento_Cambio As Double
        Private mDocumento_CambioMoedaBase As Double
        Private mDocumento_CambioMoedaAlternativa As Double
        Private mDocumento_MoedaDaUEM As Boolean
        Private mDocumento_Arredondamento As Byte
        Private mDocumento_ArredondamentoIva As Byte
        Private mEntidade_Tipo As String
        Private mEntidade_Codigo As String
        Private mEntidade_Nome As String
        Private mEntidade_Morada As String
        Private mEntidade_Morada2 As String
        Private mEntidade_Localidade As String
        Private mEntidade_CodigoPostal As String
        Private mEntidade_LocalidadePostal As String
        Private mEntidade_Distrito As String
        Private mEntidade_Pais As String
        Private mEntidade_Contribuinte As String
        Private mEntidade_Desconto As Double
        Private mEntidade_Zona As String
        Private mEntidadeFacturacao_Tipo As String
        Private mEntidadeFacturacao_Codigo As String
        Private mEntidadeFacturacao_Nome As String
        Private mEntidadeFacturacao_Morada As String
        Private mEntidadeFacturacao_Morada2 As String
        Private mEntidadeFacturacao_Localidade As String
        Private mEntidadeFacturacao_CodigoPostal As String
        Private mEntidadeFacturacao_LocalidadePostal As String
        Private mEntidadeFacturacao_Distrito As String
        Private mEntidadeFacturacao_Pais As String
        Private mEntidadeFacturacao_Contribuinte As String
        Private mEntidadeEntrega_Tipo As String
        Private mEntidadeEntrega_Codigo As String
        Private mEntidadeEntrega_Nome As String
        Private mEntidadeEntrega_Morada As String
        Private mEntidadeEntrega_Morada2 As String
        Private mEntidadeEntrega_Localidade As String
        Private mEntidadeEntrega_CodigoPostal As String
        Private mEntidadeEntrega_LocalidadePostal As String
        Private mEntidadeEntrega_Distrito As String
        Private mCondicaoPagamento As String
        Private mModoExpedicao As String
        Private mModoPagamento As String
        Private mData As Date
        Private mVencimento As Date
        Private mRequisicao As String
        Private mLocalOperacao As String
        Private mTotal_Mercadoria As Double
        Private mTotal_Descontos As Double
        Private mTotal_Iva As Double
        Private mTotal_Outros As Double
        Private mTotal_Documento As Double
        Private mTransporte_CargaDescarga As CargaDescarga
        Private mTransporte_CargaLocal As String
        Private mTransporte_CargaData As String
        Private mTransporte_CargaHora As String
        Private mTransporte_DescargaLocal As String
        Private mTransporte_DescargaData As String
        Private mTransporte_DescargaHora As String
        Private mTransporte_Matricula As String
        Private mObservacoes As String
        Private mUtilizador As String
        Private mDataUltimaActualizacao As Date
        Private mCamposUtilizador As BE.CamposUtilizadorCollection
        Private mLinhas As BE.LinhasDocumentosVendaCollection
        Private mEmModoEdicao As Boolean


        Sub New()
            mId = ""
            mFilial = "000"
            mSeccao = "1"
            mDocumento_Tipo = ""
            mDocumento_Serie = ""
            mDocumento_Numero = 0
            mDocumento_Moeda = ""
            mDocumento_Cambio = 0
            mDocumento_CambioMoedaBase = 0
            mDocumento_CambioMoedaAlternativa = 0
            mDocumento_MoedaDaUEM = True
            mDocumento_Arredondamento = 0
            mDocumento_ArredondamentoIva = 0
            mEntidade_Tipo = ""
            mEntidade_Codigo = ""
            mEntidade_Nome = ""
            mEntidade_Morada = ""
            mEntidade_Morada2 = ""
            mEntidade_Localidade = ""
            mEntidade_CodigoPostal = ""
            mEntidade_LocalidadePostal = ""
            mEntidade_Distrito = ""
            mEntidade_Pais = ""
            mEntidade_Contribuinte = ""
            mEntidade_Desconto = 0
            mEntidade_Zona = ""
            mEntidadeFacturacao_Tipo = ""
            mEntidadeFacturacao_Codigo = ""
            mEntidadeFacturacao_Nome = ""
            mEntidadeFacturacao_Morada = ""
            mEntidadeFacturacao_Morada2 = ""
            mEntidadeFacturacao_Localidade = ""
            mEntidadeFacturacao_CodigoPostal = ""
            mEntidadeFacturacao_LocalidadePostal = ""
            mEntidadeFacturacao_Distrito = ""
            mEntidadeFacturacao_Pais = ""
            mEntidadeFacturacao_Contribuinte = ""
            mEntidadeEntrega_Tipo = ""
            mEntidadeEntrega_Codigo = ""
            mEntidadeEntrega_Nome = ""
            mEntidadeEntrega_Morada = ""
            mEntidadeEntrega_Morada2 = ""
            mEntidadeEntrega_Localidade = ""
            mEntidadeEntrega_CodigoPostal = ""
            mEntidadeEntrega_LocalidadePostal = ""
            mEntidadeEntrega_Distrito = ""
            mCondicaoPagamento = ""
            mModoExpedicao = ""
            mModoPagamento = ""
            mData = Now.Date
            mVencimento = Now.Date
            mRequisicao = ""
            mLocalOperacao = ""
            mTotal_Mercadoria = 0
            mTotal_Descontos = 0
            mTotal_Iva = 0
            mTotal_Outros = 0
            mTransporte_CargaDescarga = New CargaDescarga()
            mTransporte_CargaLocal = ""
            mTransporte_CargaData = ""
            mTransporte_CargaHora = ""
            mTransporte_DescargaLocal = ""
            mTransporte_DescargaData = ""
            mTransporte_DescargaHora = ""
            mTransporte_Matricula = ""
            mObservacoes = ""
            mUtilizador = ""
            mDataUltimaActualizacao = Now.Date
            mCamposUtilizador = New BE.CamposUtilizadorCollection
            mLinhas = New BE.LinhasDocumentosVendaCollection
            mEmModoEdicao = False
        End Sub

        Public Property Id() As String
            Get
                Return mId
            End Get
            Set(ByVal value As String)
                mId = value
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

        Public Property Seccao() As String
            Get
                Return mSeccao
            End Get
            Set(ByVal value As String)
                mSeccao = value
            End Set
        End Property

        Public Property Documento_Tipo() As String
            Get
                Return mDocumento_Tipo
            End Get
            Set(ByVal value As String)
                mDocumento_Tipo = value
            End Set
        End Property

        Public Property Documento_Serie() As String
            Get
                Return mDocumento_Serie
            End Get
            Set(ByVal value As String)
                mDocumento_Serie = value
            End Set
        End Property

        Public Property Documento_Numero() As System.Int32
            Get
                Return mDocumento_Numero
            End Get
            Set(ByVal value As System.Int32)
                mDocumento_Numero = value
            End Set
        End Property

        Public Property Documento_Moeda() As String
            Get
                Return mDocumento_Moeda
            End Get
            Set(ByVal value As String)
                mDocumento_Moeda = value
            End Set
        End Property

        Public Property Documento_Cambio() As Double
            Get
                Return mDocumento_Cambio
            End Get
            Set(ByVal value As Double)
                mDocumento_Cambio = value
            End Set
        End Property

        Public Property Documento_CambioMoedaBase() As Double
            Get
                Return mDocumento_CambioMoedaBase
            End Get
            Set(ByVal value As Double)
                mDocumento_CambioMoedaBase = value
            End Set
        End Property

        Public Property Documento_CambioMoedaAlternativa() As Double
            Get
                Return mDocumento_CambioMoedaAlternativa
            End Get
            Set(ByVal value As Double)
                mDocumento_CambioMoedaAlternativa = value
            End Set
        End Property

        Public Property Documento_MoedaDaUEM() As Boolean
            Get
                Return mDocumento_MoedaDaUEM
            End Get
            Set(ByVal value As Boolean)
                mDocumento_MoedaDaUEM = value
            End Set
        End Property

        Public Property Documento_Arredondamento() As Byte
            Get
                Return mDocumento_Arredondamento
            End Get
            Set(ByVal value As Byte)
                mDocumento_Arredondamento = value
            End Set
        End Property

        Public Property Documento_ArredondamentoIva() As Byte
            Get
                Return mDocumento_ArredondamentoIva
            End Get
            Set(ByVal value As Byte)
                mDocumento_ArredondamentoIva = value
            End Set
        End Property

        Public Property ENTIDADE_Tipo() As String
            Get
                Return mEntidade_Tipo
            End Get
            Set(ByVal value As String)
                mEntidade_Tipo = value
            End Set
        End Property

        Public Property ENTIDADE_Codigo() As String
            Get
                Return mEntidade_Codigo
            End Get
            Set(ByVal value As String)
                mEntidade_Codigo = value
            End Set
        End Property

        Public Property ENTIDADE_Nome() As String
            Get
                Return mEntidade_Nome
            End Get
            Set(ByVal value As String)
                mEntidade_Nome = value
            End Set
        End Property

        Public Property ENTIDADE_Morada() As String
            Get
                Return mEntidade_Morada
            End Get
            Set(ByVal value As String)
                mEntidade_Morada = value
            End Set
        End Property

        Public Property ENTIDADE_Morada2() As String
            Get
                Return mEntidade_Morada2
            End Get
            Set(ByVal value As String)
                mEntidade_Morada2 = value
            End Set
        End Property

        Public Property ENTIDADE_Localidade() As String
            Get
                Return mEntidade_Localidade
            End Get
            Set(ByVal value As String)
                mEntidade_Localidade = value
            End Set
        End Property

        Public Property ENTIDADE_CodigoPostal() As String
            Get
                Return mEntidade_CodigoPostal
            End Get
            Set(ByVal value As String)
                mEntidade_CodigoPostal = value
            End Set
        End Property

        Public Property ENTIDADE_LocalidadePostal() As String
            Get
                Return mEntidade_LocalidadePostal
            End Get
            Set(ByVal value As String)
                mEntidade_LocalidadePostal = value
            End Set
        End Property

        Public Property ENTIDADE_Distrito() As String
            Get
                Return mEntidade_Distrito
            End Get
            Set(ByVal value As String)
                mEntidade_Distrito = value
            End Set
        End Property

        Public Property ENTIDADE_Pais() As String
            Get
                Return mEntidade_Pais
            End Get
            Set(ByVal value As String)
                mEntidade_Pais = value
            End Set
        End Property

        Public Property ENTIDADE_Contribuinte() As String
            Get
                Return mEntidade_Contribuinte
            End Get
            Set(ByVal value As String)
                mEntidade_Contribuinte = value
            End Set
        End Property

        Public Property ENTIDADE_Desconto() As Double
            Get
                Return mEntidade_Desconto
            End Get
            Set(ByVal value As Double)
                mEntidade_Desconto = value
            End Set
        End Property

        Public Property ENTIDADE_Zona() As String
            Get
                Return mEntidade_Zona
            End Get
            Set(ByVal value As String)
                mEntidade_Zona = value
            End Set
        End Property

        Public Property EntidadeFacturacao_Tipo() As String
            Get
                Return mEntidadeFacturacao_Tipo
            End Get
            Set(ByVal value As String)
                mEntidadeFacturacao_Tipo = value
            End Set
        End Property

        Public Property EntidadeFacturacao_Codigo() As String
            Get
                Return mEntidadeFacturacao_Codigo
            End Get
            Set(ByVal value As String)
                mEntidadeFacturacao_Codigo = value
            End Set
        End Property

        Public Property EntidadeFacturacao_Nome() As String
            Get
                Return mEntidadeFacturacao_Nome
            End Get
            Set(ByVal value As String)
                mEntidadeFacturacao_Nome = value
            End Set
        End Property

        Public Property EntidadeFacturacao_Morada() As String
            Get
                Return mEntidadeFacturacao_Morada
            End Get
            Set(ByVal value As String)
                mEntidadeFacturacao_Morada = value
            End Set
        End Property

        Public Property EntidadeFacturacao_Morada2() As String
            Get
                Return mEntidadeFacturacao_Morada2
            End Get
            Set(ByVal value As String)
                mEntidadeFacturacao_Morada2 = value
            End Set
        End Property

        Public Property EntidadeFacturacao_Localidade() As String
            Get
                Return mEntidadeFacturacao_Localidade
            End Get
            Set(ByVal value As String)
                mEntidadeFacturacao_Localidade = value
            End Set
        End Property

        Public Property EntidadeFacturacao_CodigoPostal() As String
            Get
                Return mEntidadeFacturacao_CodigoPostal
            End Get
            Set(ByVal value As String)
                mEntidadeFacturacao_CodigoPostal = value
            End Set
        End Property

        Public Property EntidadeFacturacao_LocalidadePostal() As String
            Get
                Return mEntidadeFacturacao_LocalidadePostal
            End Get
            Set(ByVal value As String)
                mEntidadeFacturacao_LocalidadePostal = value
            End Set
        End Property

        Public Property EntidadeFacturacao_Distrito() As String
            Get
                Return mEntidadeFacturacao_Distrito
            End Get
            Set(ByVal value As String)
                mEntidadeFacturacao_Distrito = value
            End Set
        End Property

        Public Property EntidadeFacturacao_Pais() As String
            Get
                Return mEntidadeFacturacao_Pais
            End Get
            Set(ByVal value As String)
                mEntidadeFacturacao_Pais = value
            End Set
        End Property

        Public Property EntidadeFacturacao_Contribuinte() As String
            Get
                Return mEntidadeFacturacao_Contribuinte
            End Get
            Set(ByVal value As String)
                mEntidadeFacturacao_Contribuinte = value
            End Set
        End Property

        Public Property EntidadeEntrega_Tipo() As String
            Get
                Return mEntidadeEntrega_Tipo
            End Get
            Set(ByVal value As String)
                mEntidadeEntrega_Tipo = value
            End Set
        End Property

        Public Property EntidadeEntrega_Codigo() As String
            Get
                Return mEntidadeEntrega_Codigo
            End Get
            Set(ByVal value As String)
                mEntidadeEntrega_Codigo = value
            End Set
        End Property

        Public Property EntidadeEntrega_Nome() As String
            Get
                Return mEntidadeEntrega_Nome
            End Get
            Set(ByVal value As String)
                mEntidadeEntrega_Nome = value
            End Set
        End Property

        Public Property EntidadeEntrega_Morada() As String
            Get
                Return mEntidadeEntrega_Morada
            End Get
            Set(ByVal value As String)
                mEntidadeEntrega_Morada = value
            End Set
        End Property

        Public Property EntidadeEntrega_Morada2() As String
            Get
                Return mEntidadeEntrega_Morada2
            End Get
            Set(ByVal value As String)
                mEntidadeEntrega_Morada2 = value
            End Set
        End Property

        Public Property EntidadeEntrega_Localidade() As String
            Get
                Return mEntidadeEntrega_Localidade
            End Get
            Set(ByVal value As String)
                mEntidadeEntrega_Localidade = value
            End Set
        End Property

        Public Property EntidadeEntrega_CodigoPostal() As String
            Get
                Return mEntidadeEntrega_CodigoPostal
            End Get
            Set(ByVal value As String)
                mEntidadeEntrega_CodigoPostal = value
            End Set
        End Property

        Public Property EntidadeEntrega_LocalidadePostal() As String
            Get
                Return mEntidadeEntrega_LocalidadePostal
            End Get
            Set(ByVal value As String)
                mEntidadeEntrega_LocalidadePostal = value
            End Set
        End Property

        Public Property EntidadeEntrega_Distrito() As String
            Get
                Return mEntidadeEntrega_Distrito
            End Get
            Set(ByVal value As String)
                mEntidadeEntrega_Distrito = value
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

        Public Property ModoExpedicao() As String
            Get
                Return mModoExpedicao
            End Get
            Set(ByVal value As String)
                mModoExpedicao = value
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

        Public Property Data() As Date
            Get
                Return mData
            End Get
            Set(ByVal value As Date)
                mData = value
            End Set
        End Property

        Public Property Vencimento() As Date
            Get
                Return mVencimento
            End Get
            Set(ByVal value As Date)
                mVencimento = value
            End Set
        End Property

        Public Property Requisicao() As String
            Get
                Return mRequisicao
            End Get
            Set(ByVal value As String)
                mRequisicao = value
            End Set
        End Property

        Public Property LocalOperacao() As String
            Get
                Return mLocalOperacao
            End Get
            Set(ByVal value As String)
                mLocalOperacao = value
            End Set
        End Property

        Public Property TOTAL_Mercadoria() As Double
            Get
                Return mTotal_Mercadoria
            End Get
            Set(ByVal value As Double)
                mTotal_Mercadoria = value
            End Set
        End Property

        Public Property TOTAL_Descontos() As Double
            Get
                Return mTotal_Descontos
            End Get
            Set(ByVal value As Double)
                mTotal_Descontos = value
            End Set
        End Property

        Public Property TOTAL_Iva() As Double
            Get
                Return mTotal_Iva
            End Get
            Set(ByVal value As Double)
                mTotal_Iva = value
            End Set
        End Property

        Public Property TOTAL_Outros() As Double
            Get
                Return mTotal_Outros
            End Get
            Set(ByVal value As Double)
                mTotal_Outros = value
            End Set
        End Property

        Public Property TOTAL_Documento() As Double
            Get
                Return mTotal_Documento
            End Get
            Set(ByVal value As Double)
                mTotal_Documento = value
            End Set
        End Property


        Public Property TRANSPORTE_CargaDescarga() As CargaDescarga
            Get
                Return mTransporte_CargaDescarga
            End Get
            Set(ByVal value As CargaDescarga)
                mTransporte_CargaDescarga = value
            End Set
        End Property


        Public Property TRANSPORTE_Carga_Local() As String
            Get
                Return mTransporte_CargaLocal
            End Get
            Set(ByVal value As String)
                mTransporte_CargaLocal = value
            End Set
        End Property

        Public Property TRANSPORTE_Carga_Data() As String
            Get
                Return mTransporte_CargaData
            End Get
            Set(ByVal value As String)
                mTransporte_CargaData = value
            End Set
        End Property

        Public Property TRANSPORTE_Carga_Hora() As String
            Get
                Return mTransporte_CargaHora
            End Get
            Set(ByVal value As String)
                mTransporte_CargaHora = value
            End Set
        End Property

        Public Property TRANSPORTE_Descarga_Local() As String
            Get
                Return mTransporte_DescargaLocal
            End Get
            Set(ByVal value As String)
                mTransporte_DescargaLocal = value
            End Set
        End Property

        Public Property TRANSPORTE_Descarga_Data() As String
            Get
                Return mTransporte_DescargaData
            End Get
            Set(ByVal value As String)
                mTransporte_DescargaData = value
            End Set
        End Property

        Public Property TRANSPORTE_Descarga_Hora() As String
            Get
                Return mTransporte_DescargaHora
            End Get
            Set(ByVal value As String)
                mTransporte_DescargaHora = value
            End Set
        End Property

        Public Property TRANSPORTE_Matricula() As String
            Get
                Return mTransporte_Matricula
            End Get
            Set(ByVal value As String)
                mTransporte_Matricula = value
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

        Public Property DataUltimaActualizacao() As Date
            Get
                Return mDataUltimaActualizacao
            End Get
            Set(ByVal value As Date)
                mDataUltimaActualizacao = value
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

        Public Property Linhas() As BE.LinhasDocumentosVendaCollection
            Get
                Return mLinhas
            End Get
            Set(ByVal value As BE.LinhasDocumentosVendaCollection)
                mLinhas = value
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

