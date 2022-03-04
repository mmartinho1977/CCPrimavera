Imports Microsoft.VisualBasic
Imports CCUtils


Namespace Specifications

    Public Class DocumentoVenda_Specs
        Public Shared CampoObrigatorio_Id As Boolean = True
        Public Shared CampoObrigatorio_Filial As Boolean = True
        Public Shared CampoObrigatorio_Seccao As Boolean = True
        Public Shared CampoObrigatorio_Documento_Tipo As Boolean = True
        Public Shared CampoObrigatorio_Documento_Serie As Boolean = True
        Public Shared CampoObrigatorio_Documento_Numero As Boolean = True
        Public Shared CampoObrigatorio_Documento_Moeda As Boolean = True
        Public Shared CampoObrigatorio_Documento_Cambio As Boolean = True
        Public Shared CampoObrigatorio_Documento_CambioMoedaBase As Boolean = True
        Public Shared CampoObrigatorio_Documento_CambioMoedaAlternativa As Boolean = True
        Public Shared CampoObrigatorio_Documento_Arredondamento As Boolean = True
        Public Shared CampoObrigatorio_Documento_ArredondamentoIva As Boolean = True
        Public Shared CampoObrigatorio_Entidade_Tipo As Boolean = True
        Public Shared CampoObrigatorio_Entidade_Codigo As Boolean = True
        Public Shared CampoObrigatorio_Entidade_Nome As Boolean = True
        Public Shared CampoObrigatorio_Entidade_Morada As Boolean = False
        Public Shared CampoObrigatorio_Entidade_Morada2 As Boolean = False
        Public Shared CampoObrigatorio_Entidade_Localidade As Boolean = False
        Public Shared CampoObrigatorio_Entidade_CodigoPostal As Boolean = False
        Public Shared CampoObrigatorio_Entidade_LocalidadePostal As Boolean = False
        Public Shared CampoObrigatorio_Entidade_Contribuinte As Boolean = False
        Public Shared CampoObrigatorio_Entidade_Desconto As Boolean = False
        Public Shared CampoObrigatorio_Entidade_Zona As Boolean = False

        Public Shared CampoObrigatorio_EntidadeFacturacao_Tipo As Boolean = False
        Public Shared CampoObrigatorio_EntidadeFacturacao_Codigo As Boolean = False
        Public Shared CampoObrigatorio_EntidadeFacturacao_Nome As Boolean = False
        Public Shared CampoObrigatorio_EntidadeFacturacao_Morada As Boolean = False
        Public Shared CampoObrigatorio_EntidadeFacturacao_Morada2 As Boolean = False
        Public Shared CampoObrigatorio_EntidadeFacturacao_Localidade As Boolean = False
        Public Shared CampoObrigatorio_EntidadeFacturacao_CodigoPostal As Boolean = False
        Public Shared CampoObrigatorio_EntidadeFacturacao_LocalidadePostal As Boolean = False
        Public Shared CampoObrigatorio_EntidadeFacturacao_Contribuinte As Boolean = False

        Public Shared CampoObrigatorio_CondicaoPagamento As Boolean = True
        Public Shared CampoObrigatorio_ModoExpedicao As Boolean = False
        Public Shared CampoObrigatorio_ModoPagamento As Boolean = True
        Public Shared CampoObrigatorio_Data As Boolean = True
        Public Shared CampoObrigatorio_Vencimento As Boolean = True
        Public Shared CampoObrigatorio_Requisicao As Boolean = False
        Public Shared CampoObrigatorio_Total_Mercadoria As Boolean = False
        Public Shared CampoObrigatorio_Total_Descontos As Boolean = False
        Public Shared CampoObrigatorio_Total_Iva As Boolean = False
        Public Shared CampoObrigatorio_Total_Outros As Boolean = False
        Public Shared CampoObrigatorio_Transporte_CargaLocal As Boolean = False
        Public Shared CampoObrigatorio_Transporte_CargaData As Boolean = False
        Public Shared CampoObrigatorio_Transporte_CargaHora As Boolean = False
        Public Shared CampoObrigatorio_Transporte_DescargaLocal As Boolean = False
        Public Shared CampoObrigatorio_Transporte_DescargaData As Boolean = False
        Public Shared CampoObrigatorio_Transporte_DescargaHora As Boolean = False
        Public Shared CampoObrigatorio_Transporte_Matricula As Boolean = False
        Public Shared CampoObrigatorio_Observacoes As Boolean = False
        Public Shared CampoObrigatorio_Utilizador As Boolean = True
        Public Shared CampoObrigatorio_DataUltimaActualizacao As Boolean = True

        Public Shared ComprimentoMaximo_Filial As System.Int16 = 3
        Public Shared ComprimentoMaximo_Seccao As System.Int16 = 1
        Public Shared ComprimentoMaximo_Documento_Tipo As System.Int16 = 5
        Public Shared ComprimentoMaximo_Documento_Serie As System.Int16 = 5
        Public Shared ComprimentoMaximo_Documento_Moeda As System.Int16 = 3
        Public Shared ComprimentoMaximo_Entidade_Tipo As System.Int16 = 1
        Public Shared ComprimentoMaximo_Entidade_Codigo As System.Int16 = 12
        Public Shared ComprimentoMaximo_Entidade_Nome As System.Int16 = 50
        Public Shared ComprimentoMaximo_Entidade_Morada As System.Int16 = 50
        Public Shared ComprimentoMaximo_Entidade_Morada2 As System.Int16 = 50
        Public Shared ComprimentoMaximo_Entidade_Localidade As System.Int16 = 50
        Public Shared ComprimentoMaximo_Entidade_CodigoPostal As System.Int16 = 15
        Public Shared ComprimentoMaximo_Entidade_LocalidadePostal As System.Int16 = 50
        Public Shared ComprimentoMaximo_Entidade_Contribuinte As System.Int16 = 20
        Public Shared ComprimentoMaximo_Entidade_Zona As System.Int16 = 2
        Public Shared ComprimentoMaximo_EntidadeFacturacao_Tipo As System.Int16 = 1
        Public Shared ComprimentoMaximo_EntidadeFacturacao_Codigo As System.Int16 = 12
        Public Shared ComprimentoMaximo_EntidadeFacturacao_Nome As System.Int16 = 50
        Public Shared ComprimentoMaximo_EntidadeFacturacao_Morada As System.Int16 = 50
        Public Shared ComprimentoMaximo_EntidadeFacturacao_Morada2 As System.Int16 = 50
        Public Shared ComprimentoMaximo_EntidadeFacturacao_Localidade As System.Int16 = 50
        Public Shared ComprimentoMaximo_EntidadeFacturacao_CodigoPostal As System.Int16 = 15
        Public Shared ComprimentoMaximo_EntidadeFacturacao_LocalidadePostal As System.Int16 = 50
        Public Shared ComprimentoMaximo_EntidadeFacturacao_Contribuinte As System.Int16 = 20
        Public Shared ComprimentoMaximo_CondicaoPagamento As System.Int16 = 2
        Public Shared ComprimentoMaximo_ModoExpedicao As System.Int16 = 2
        Public Shared ComprimentoMaximo_ModoPagamento As System.Int16 = 5
        Public Shared ComprimentoMaximo_Requisicao As System.Int16 = 20
        Public Shared ComprimentoMaximo_Transporte_CargaLocal As System.Int16 = 50
        Public Shared ComprimentoMaximo_Transporte_CargaData As System.Int16 = 20
        Public Shared ComprimentoMaximo_Transporte_CargaHora As System.Int16 = 5
        Public Shared ComprimentoMaximo_Transporte_DescargaLocal As System.Int16 = 50
        Public Shared ComprimentoMaximo_Transporte_DescargaData As System.Int16 = 20
        Public Shared ComprimentoMaximo_Transporte_DescargaHora As System.Int16 = 5
        Public Shared ComprimentoMaximo_Transporte_Matricula As System.Int16 = 25
        Public Shared ComprimentoMaximo_Observacoes As System.Int16 = 512
        Public Shared ComprimentoMaximo_Utilizador As System.Int16 = 20

        Public Shared ComprimentoMinimo_Filial As System.Int16 = 3
        Public Shared ComprimentoMinimo_Seccao As System.Int16 = 1
        Public Shared ComprimentoMinimo_Documento_Tipo As System.Int16 = 2
        Public Shared ComprimentoMinimo_Documento_Serie As System.Int16 = 1
        Public Shared ComprimentoMinimo_Documento_Moeda As System.Int16 = 2
        Public Shared ComprimentoMinimo_Entidade_Tipo As System.Int16 = 1
        Public Shared ComprimentoMinimo_Entidade_Codigo As System.Int16 = 1
        Public Shared ComprimentoMinimo_Entidade_Nome As System.Int16 = 3
        Public Shared ComprimentoMinimo_Entidade_Morada As System.Int16 = 1
        Public Shared ComprimentoMinimo_Entidade_Morada2 As System.Int16 = 1
        Public Shared ComprimentoMinimo_Entidade_Localidade As System.Int16 = 0
        Public Shared ComprimentoMinimo_Entidade_CodigoPostal As System.Int16 = 4
        Public Shared ComprimentoMinimo_Entidade_LocalidadePostal As System.Int16 = 2
        Public Shared ComprimentoMinimo_Entidade_Contribuinte As System.Int16 = 9
        Public Shared ComprimentoMinimo_Entidade_Zona As System.Int16 = 1
        Public Shared ComprimentoMinimo_EntidadeFacturacao_Tipo As System.Int16 = 1
        Public Shared ComprimentoMinimo_EntidadeFacturacao_Codigo As System.Int16 = 1
        Public Shared ComprimentoMinimo_EntidadeFacturacao_Nome As System.Int16 = 3
        Public Shared ComprimentoMinimo_EntidadeFacturacao_Morada As System.Int16 = 1
        Public Shared ComprimentoMinimo_EntidadeFacturacao_Morada2 As System.Int16 = 1
        Public Shared ComprimentoMinimo_EntidadeFacturacao_Localidade As System.Int16 = 0
        Public Shared ComprimentoMinimo_EntidadeFacturacao_CodigoPostal As System.Int16 = 4
        Public Shared ComprimentoMinimo_EntidadeFacturacao_LocalidadePostal As System.Int16 = 2
        Public Shared ComprimentoMinimo_EntidadeFacturacao_Contribuinte As System.Int16 = 9
        Public Shared ComprimentoMinimo_CondicaoPagamento As System.Int16 = 1
        Public Shared ComprimentoMinimo_ModoExpedicao As System.Int16 = 1
        Public Shared ComprimentoMinimo_ModoPagamento As System.Int16 = 1
        Public Shared ComprimentoMinimo_Requisicao As System.Int16 = 1
        Public Shared ComprimentoMinimo_Transporte_CargaLocal As System.Int16 = 2
        Public Shared ComprimentoMinimo_Transporte_CargaData As System.Int16 = 2
        Public Shared ComprimentoMinimo_Transporte_CargaHora As System.Int16 = 2
        Public Shared ComprimentoMinimo_Transporte_DescargaLocal As System.Int16 = 2
        Public Shared ComprimentoMinimo_Transporte_DescargaData As System.Int16 = 2
        Public Shared ComprimentoMinimo_Transporte_DescargaHora As System.Int16 = 2
        Public Shared ComprimentoMinimo_Transporte_Matricula As System.Int16 = 2
        Public Shared ComprimentoMinimo_Observacoes As System.Int16 = 2
        Public Shared ComprimentoMinimo_Utilizador As System.Int16 = 1

        Public Shared LimiteInferior_Documento_Numero As System.Int32 = 0
        Public Shared LimiteInferior_Documento_Cambio As Double = 0.01
        Public Shared LimiteInferior_Documento_CambioMoedaBase As Double = 0.01
        Public Shared LimiteInferior_Documento_CambioMoedaAlternativa As Double = 0.01
        Public Shared LimiteInferior_Documento_Arredondamento As System.Int16 = 0
        Public Shared LimiteInferior_Documento_ArredondamentoIva As System.Int16 = 0
        Public Shared LimiteInferior_Entidade_Desconto As Double = 0
        Public Shared LimiteInferior_Total_Mercadoria As Double = 0
        Public Shared LimiteInferior_Total_Descontos As Double = 0
        Public Shared LimiteInferior_Total_Iva As Double = 0
        Public Shared LimiteInferior_Total_Outros As Double = 0

        Public Shared LimiteSuperior_Documento_Numero As System.Int32 = System.Int32.MaxValue
        Public Shared LimiteSuperior_Documento_Cambio As Double = System.Double.MaxValue
        Public Shared LimiteSuperior_Documento_CambioMoedaBase As Double = System.Double.MaxValue
        Public Shared LimiteSuperior_Documento_CambioMoedaAlternativa As Double = System.Double.MaxValue
        Public Shared LimiteSuperior_Documento_Arredondamento As System.Int16 = 4
        Public Shared LimiteSuperior_Documento_ArredondamentoIva As System.Int16 = 4
        Public Shared LimiteSuperior_Entidade_Desconto As Double = 100
        Public Shared LimiteSuperior_Total_Mercadoria As Double = 9999999
        Public Shared LimiteSuperior_Total_Descontos As Double = 999999
        Public Shared LimiteSuperior_Total_Iva As Double = 999999
        Public Shared LimiteSuperior_Total_Outros As Double = 999999

        Public Shared LimiteTemporalInferior_Data As Date = "01-01-2008"
        Public Shared LimiteTemporalSuperior_Data As Date = "31-12-2078"
        Public Shared LimiteTemporalInferior_Vencimento As Date = "01-01-2008"
        Public Shared LimiteTemporalSuperior_Vencimento As Date = "31-12-2078"
        Public Shared LimiteTemporalInferior_DataUltimaActualizacao As Date = "01-01-2008"
        Public Shared LimiteTemporalSuperior_DataUltimaActualizacao As Date = "31-12-2078"


        Sub New()

        End Sub




        Public Shared Function Certifica(ByRef DocumentoVenda As BE.DocumentoVenda, ByRef mensagem As String) As Boolean
            Dim mMensagem As String
            Dim i As System.Int16

            mMensagem = ""

            With DocumentoVenda
                CCValidation.Int64("Documento_Numero", .Documento_Numero, True, LimiteInferior_Documento_Numero, LimiteSuperior_Documento_Numero, mMensagem)
                CCValidation.Guid("Id", .Id, CampoObrigatorio_Id, mMensagem)
                CCValidation.Texto("Filial", .Filial, CampoObrigatorio_Filial, New Char() {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}, ComprimentoMinimo_Filial, ComprimentoMaximo_Filial, mMensagem)
                CCValidation.Texto("Seccao", .Seccao, CampoObrigatorio_Seccao, False, ComprimentoMinimo_Seccao, ComprimentoMaximo_Seccao, mMensagem)
                CCValidation.Texto("Documento_Tipo", .Documento_Tipo, CampoObrigatorio_Documento_Tipo, False, ComprimentoMinimo_Documento_Tipo, ComprimentoMaximo_Documento_Tipo, mMensagem)
                CCValidation.Texto("Documento_Serie", .Documento_Serie, CampoObrigatorio_Documento_Serie, False, ComprimentoMinimo_Documento_Serie, ComprimentoMaximo_Documento_Serie, mMensagem)
                CCValidation.Int64("Documento_Numero", .Documento_Numero, True, LimiteInferior_Documento_Numero, LimiteSuperior_Documento_Numero, mMensagem)
                CCValidation.Texto("Documento_Moeda", .Documento_Moeda, CampoObrigatorio_Documento_Moeda, False, ComprimentoMinimo_Documento_Moeda, ComprimentoMaximo_Documento_Moeda, mMensagem)
                CCValidation.Double("Documento_Cambio", .Documento_Cambio, True, LimiteInferior_Documento_Cambio, LimiteSuperior_Documento_Cambio, mMensagem)
                CCValidation.Double("Documento_CambioMoedaBase", .Documento_CambioMoedaBase, True, LimiteInferior_Documento_CambioMoedaBase, LimiteSuperior_Documento_CambioMoedaBase, mMensagem)
                CCValidation.Double("Documento_CambioMoedaAlternativa", .Documento_CambioMoedaAlternativa, True, LimiteInferior_Documento_CambioMoedaAlternativa, LimiteSuperior_Documento_CambioMoedaAlternativa, mMensagem)
                CCValidation.Int64("Documento_Arredondamento", .Documento_Arredondamento, True, LimiteInferior_Documento_Arredondamento, LimiteSuperior_Documento_Arredondamento, mMensagem)
                CCValidation.Int64("Documento_ArredondamentoIva", .Documento_ArredondamentoIva, True, LimiteInferior_Documento_ArredondamentoIva, LimiteSuperior_Documento_ArredondamentoIva, mMensagem)
                CCValidation.Texto("Entidade_Tipo", .ENTIDADE_Tipo, CampoObrigatorio_Entidade_Tipo, False, ComprimentoMinimo_Entidade_Tipo, ComprimentoMaximo_Entidade_Tipo, mMensagem)
                CCValidation.Texto("Entidade_Codigo", .ENTIDADE_Codigo, CampoObrigatorio_Entidade_Codigo, True, ComprimentoMinimo_Entidade_Codigo, ComprimentoMaximo_Entidade_Codigo, mMensagem)
                CCValidation.Texto("Entidade_Nome", .ENTIDADE_Nome, CampoObrigatorio_Entidade_Nome, True, ComprimentoMinimo_Entidade_Nome, ComprimentoMaximo_Entidade_Nome, mMensagem)
                CCValidation.Texto("Entidade_Morada", .ENTIDADE_Morada, CampoObrigatorio_Entidade_Morada, True, ComprimentoMinimo_Entidade_Morada, ComprimentoMaximo_Entidade_Morada, mMensagem)
                CCValidation.Texto("Entidade_Morada2", .ENTIDADE_Morada2, CampoObrigatorio_Entidade_Morada2, True, ComprimentoMinimo_Entidade_Morada2, ComprimentoMaximo_Entidade_Morada2, mMensagem)
                CCValidation.Texto("Entidade_Localidade", .ENTIDADE_Localidade, CampoObrigatorio_Entidade_Localidade, True, ComprimentoMinimo_Entidade_Localidade, ComprimentoMaximo_Entidade_Localidade, mMensagem)
                CCValidation.Texto("Entidade_CodigoPostal", .ENTIDADE_CodigoPostal, CampoObrigatorio_Entidade_CodigoPostal, True, ComprimentoMinimo_Entidade_CodigoPostal, ComprimentoMaximo_Entidade_CodigoPostal, mMensagem)
                CCValidation.Texto("Entidade_LocalidadePostal", .ENTIDADE_LocalidadePostal, CampoObrigatorio_Entidade_LocalidadePostal, True, ComprimentoMinimo_Entidade_LocalidadePostal, ComprimentoMaximo_Entidade_LocalidadePostal, mMensagem)
                CCValidation.Texto("Entidade_Contribuinte", .ENTIDADE_Contribuinte, CampoObrigatorio_Entidade_Contribuinte, True, ComprimentoMinimo_Entidade_Contribuinte, ComprimentoMaximo_Entidade_Contribuinte, mMensagem)
                CCValidation.Decimal("Entidade_Desconto", .ENTIDADE_Desconto, True, LimiteInferior_Entidade_Desconto, LimiteSuperior_Entidade_Desconto, mMensagem)
                CCValidation.Texto("Entidade_Zona", .ENTIDADE_Zona, CampoObrigatorio_Entidade_Zona, True, ComprimentoMinimo_Entidade_Zona, ComprimentoMaximo_Entidade_Zona, mMensagem)
                CCValidation.Texto("EntidadeFacturacao_Tipo", .EntidadeFacturacao_Tipo, CampoObrigatorio_EntidadeFacturacao_Tipo, True, ComprimentoMinimo_EntidadeFacturacao_Tipo, ComprimentoMaximo_EntidadeFacturacao_Tipo, mMensagem)
                CCValidation.Texto("EntidadeFacturacao_Codigo", .EntidadeFacturacao_Codigo, CampoObrigatorio_EntidadeFacturacao_Codigo, True, ComprimentoMinimo_EntidadeFacturacao_Codigo, ComprimentoMaximo_EntidadeFacturacao_Codigo, mMensagem)
                CCValidation.Texto("EntidadeFacturacao_Nome", .EntidadeFacturacao_Nome, CampoObrigatorio_EntidadeFacturacao_Nome, True, ComprimentoMinimo_EntidadeFacturacao_Nome, ComprimentoMaximo_EntidadeFacturacao_Nome, mMensagem)
                CCValidation.Texto("EntidadeFacturacao_Morada", .EntidadeFacturacao_Morada, CampoObrigatorio_EntidadeFacturacao_Morada, True, ComprimentoMinimo_EntidadeFacturacao_Morada, ComprimentoMaximo_EntidadeFacturacao_Morada, mMensagem)
                CCValidation.Texto("EntidadeFacturacao_Morada2", .EntidadeFacturacao_Morada2, CampoObrigatorio_EntidadeFacturacao_Morada2, True, ComprimentoMinimo_EntidadeFacturacao_Morada2, ComprimentoMaximo_EntidadeFacturacao_Morada2, mMensagem)
                CCValidation.Texto("EntidadeFacturacao_Localidade", .EntidadeFacturacao_Localidade, CampoObrigatorio_EntidadeFacturacao_Localidade, True, ComprimentoMinimo_EntidadeFacturacao_Localidade, ComprimentoMaximo_EntidadeFacturacao_Localidade, mMensagem)
                CCValidation.Texto("EntidadeFacturacao_CodigoPostal", .EntidadeFacturacao_CodigoPostal, CampoObrigatorio_EntidadeFacturacao_CodigoPostal, True, ComprimentoMinimo_EntidadeFacturacao_CodigoPostal, ComprimentoMaximo_EntidadeFacturacao_CodigoPostal, mMensagem)
                CCValidation.Texto("EntidadeFacturacao_LocalidadePostal", .EntidadeFacturacao_LocalidadePostal, CampoObrigatorio_EntidadeFacturacao_LocalidadePostal, True, ComprimentoMinimo_EntidadeFacturacao_LocalidadePostal, ComprimentoMaximo_EntidadeFacturacao_LocalidadePostal, mMensagem)
                CCValidation.Texto("EntidadeFacturacao_Contribuinte", .EntidadeFacturacao_Contribuinte, CampoObrigatorio_EntidadeFacturacao_Contribuinte, True, ComprimentoMinimo_EntidadeFacturacao_Contribuinte, ComprimentoMaximo_EntidadeFacturacao_Contribuinte, mMensagem)
                CCValidation.Texto("CondicaoPagamento", .CondicaoPagamento, CampoObrigatorio_CondicaoPagamento, True, ComprimentoMinimo_CondicaoPagamento, ComprimentoMaximo_CondicaoPagamento, mMensagem)
                CCValidation.Texto("ModoExpedicao", .ModoExpedicao, CampoObrigatorio_ModoExpedicao, True, ComprimentoMinimo_ModoExpedicao, ComprimentoMaximo_ModoExpedicao, mMensagem)
                CCValidation.Texto("ModoPagamento", .ModoPagamento, CampoObrigatorio_ModoPagamento, True, ComprimentoMinimo_ModoPagamento, ComprimentoMaximo_ModoPagamento, mMensagem)
                CCValidation.Data("Data", .Data, True, LimiteTemporalInferior_Data, LimiteTemporalSuperior_Data, mMensagem)
                CCValidation.Data("Vencimento", .Vencimento, True, LimiteTemporalInferior_Vencimento, LimiteTemporalSuperior_Vencimento, mMensagem)
                CCValidation.Texto("Requisicao", .Requisicao, CampoObrigatorio_Requisicao, True, ComprimentoMinimo_Requisicao, ComprimentoMaximo_Requisicao, mMensagem)
                CCValidation.Double("Total_Mercadoria", .TOTAL_Mercadoria, True, LimiteInferior_Total_Mercadoria, LimiteSuperior_Total_Mercadoria, mMensagem)
                CCValidation.Double("Total_Descontos", .TOTAL_Descontos, True, LimiteInferior_Total_Descontos, LimiteSuperior_Total_Descontos, mMensagem)
                CCValidation.Double("Total_Iva", .TOTAL_Iva, True, LimiteInferior_Total_Iva, LimiteSuperior_Total_Iva, mMensagem)
                CCValidation.Double("Total_Outros", .TOTAL_Outros, True, LimiteInferior_Total_Outros, LimiteSuperior_Total_Outros, mMensagem)
                CCValidation.Texto("Transporte_CargaLocal", .TRANSPORTE_Carga_Local, CampoObrigatorio_Transporte_CargaLocal, True, ComprimentoMinimo_Transporte_CargaLocal, ComprimentoMaximo_Transporte_CargaLocal, mMensagem)
                CCValidation.Texto("Transporte_CargaData", .TRANSPORTE_Carga_Data, CampoObrigatorio_Transporte_CargaData, True, ComprimentoMinimo_Transporte_CargaData, ComprimentoMaximo_Transporte_CargaData, mMensagem)
                CCValidation.Texto("Transporte_CargaHora", .TRANSPORTE_Carga_Hora, CampoObrigatorio_Transporte_CargaHora, True, ComprimentoMinimo_Transporte_CargaHora, ComprimentoMaximo_Transporte_CargaHora, mMensagem)
                CCValidation.Texto("Transporte_DescargaLocal", .TRANSPORTE_Descarga_Local, CampoObrigatorio_Transporte_DescargaLocal, True, ComprimentoMinimo_Transporte_DescargaLocal, ComprimentoMaximo_Transporte_DescargaLocal, mMensagem)
                CCValidation.Texto("Transporte_DescargaData", .TRANSPORTE_Descarga_Data, CampoObrigatorio_Transporte_DescargaData, True, ComprimentoMinimo_Transporte_DescargaData, ComprimentoMaximo_Transporte_DescargaData, mMensagem)
                CCValidation.Texto("Transporte_DescargaHora", .TRANSPORTE_Descarga_Hora, CampoObrigatorio_Transporte_DescargaHora, True, ComprimentoMinimo_Transporte_DescargaHora, ComprimentoMaximo_Transporte_DescargaHora, mMensagem)
                CCValidation.Texto("Transporte_Matricula", .TRANSPORTE_Matricula, CampoObrigatorio_Transporte_Matricula, True, ComprimentoMinimo_Transporte_Matricula, ComprimentoMaximo_Transporte_Matricula, mMensagem)
                CCValidation.Texto("Observacoes", .Observacoes, CampoObrigatorio_Observacoes, True, ComprimentoMinimo_Observacoes, ComprimentoMaximo_Observacoes, mMensagem)
                CCValidation.Texto("Utilizador", .Utilizador, CampoObrigatorio_Utilizador, True, ComprimentoMinimo_Utilizador, ComprimentoMaximo_Utilizador, mMensagem)
                CCValidation.Data("DataUltimaActualizacao", .DataUltimaActualizacao, True, LimiteTemporalInferior_DataUltimaActualizacao, LimiteTemporalSuperior_DataUltimaActualizacao, mMensagem)

                For i = 0 To DocumentoVenda.Linhas.Count - 1
                    LinhaDocumentoVenda_Specs.Certifica(.Linhas(i), mMensagem)
                Next

            End With


            If mMensagem.Trim.Length > 0 Then
                mensagem += mMensagem
                Return False
            Else
                Return True
            End If

        End Function

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub

    End Class

End Namespace

