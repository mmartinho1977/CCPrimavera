Imports Microsoft.VisualBasic


Namespace BE

    Public Class DocumentoInterno
        Private mId As String
        Private mFilial As String
        Private mDocumento_Tipo As String
        Private mDocumento_Serie As String
        Private mDocumento_Numero As System.Int32
        Private mDocumento_Moeda As String
        Private mDocumento_Cambio As Double
        Private mDocumento_CambioMoedaBase As Double
        Private mDocumento_CambioMoedaAlternativa As Double
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
        Private mEntidade_Contribuinte As String
        Private mEntidade_Desconto As Double
        Private mCondicaoPagamento As String
        Private mModoExpedicao As String
        Private mModoPagamento As String
        Private mData As Date
        Private mVencimento As Date
        Private mRegimeIva As String
        Private mTotal_Mercadoria As Double
        Private mTotal_Descontos As Double
        Private mTotal_Iva As Double
        Private mTotal_Documento As Double
        Private mTransporte_CargaLocal As String
        Private mTransporte_CargaData As String
        Private mTransporte_CargaHora As String
        Private mTransporte_DescargaLocal As String
        Private mTransporte_DescargaData As String
        Private mTransporte_DescargaHora As String
        Private mTransporte_Matricula As String
        Private mEstado As String
        Private mObservacoes As String
        Private mUtilizador As String
        Private mDataUltimaActualizacao As Date
        Private mCamposUtilizador As BE.CamposUtilizadorCollection
        Private mLinhas As BE.LinhasDocumentosInternosCollection
        Private mEmModoEdicao As Boolean


        Sub New()
            mId = ""
            mFilial = "000"
            mDocumento_Tipo = ""
            mDocumento_Serie = ""
            mDocumento_Numero = 0
            mDocumento_Moeda = ""
            mDocumento_Cambio = 0
            mDocumento_CambioMoedaBase = 0
            mDocumento_CambioMoedaAlternativa = 0
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
            mEntidade_Contribuinte = ""
            mEntidade_Desconto = 0
            mCondicaoPagamento = ""
            mModoExpedicao = ""
            mModoPagamento = ""
            mData = Now.Date
            mVencimento = Now.Date
            mRegimeIva = ""
            mTotal_Mercadoria = 0
            mTotal_Descontos = 0
            mTotal_Iva = 0
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
            mLinhas = New BE.LinhasDocumentosInternosCollection
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

        Public Property RegimeIva() As String
            Get
                Return mRegimeIva
            End Get
            Set(ByVal value As String)
                mRegimeIva = value
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

        Public Property TOTAL_Documento() As Double
            Get
                Return mTotal_Documento
            End Get
            Set(ByVal value As Double)
                mTotal_Documento = value
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

        Public Property Estado() As String
            Get
                Return mEstado
            End Get
            Set(ByVal value As String)
                mEstado = value
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

        Public Property Linhas() As BE.LinhasDocumentosInternosCollection
            Get
                Return mLinhas
            End Get
            Set(ByVal value As BE.LinhasDocumentosInternosCollection)
                mLinhas = value
            End Set
        End Property


        Public Function LinhasResumo() As String
            Dim resumo As String = ""

            If IsNothing(mLinhas) Then
                Return "(linhas de DI não carregadas)"
            End If

            If (mLinhas.Count = 0) Then
                Return "(DI sem linhas)"
            End If

            ' CRIAR LISTA TEMPORARIA ORDENADA DESCENDENTEMENTE
            Dim lista = From linha In mLinhas.List()
                    Where linha.TipoLinha = Specifications.LinhaDocumentoInterno_Specs.TiposLinhas.Mercadoria_TipoArtigo_3_TipoLinha_10 _
                             Or linha.TipoLinha = Specifications.LinhaDocumentoInterno_Specs.TiposLinhas.MateriaPrima_TipoArtigo_6_TipoLinha_13 _
                             Or linha.TipoLinha = Specifications.LinhaDocumentoInterno_Specs.TiposLinhas.MateriaSubsidiaria_TipoArtigo_7_TipoLinha_14
                    Order By linha.PrecoLiquido Descending
                    Select linha

            ' CRIAR UM RESUMO TEXTUAL
            For i = 0 To lista.Count - 1
                resumo += String.Format("^{0}({1})$", lista(i).Artigo.TrimEnd(), lista(i).Quantidade.ToString())
            Next

            Return resumo
        End Function


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

