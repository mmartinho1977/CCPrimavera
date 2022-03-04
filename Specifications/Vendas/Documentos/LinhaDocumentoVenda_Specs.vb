Imports Microsoft.VisualBasic
Imports CCUtils


Namespace Specifications

    Public Class LinhaDocumentoVenda_Specs

        Public Enum TiposLinhas
            Mercadoria_TipoArtigo_3_TipoLinha_10 = 10
            ServicoA_TipoArtigo_0_TipoLinha_20 = 20
            MateriaPrima_TipoArtigo_6_TipoLinha_13 = 13
            MateriaSubsidiaria_TipoArtigo_7_TipoLinha_14 = 14
            MaoDeObra_TipoArtigo_12_TipoLinha_23 = 23
            Comentario_60 = 60
            Portes_50 = 50
            Acerto_30 = 30
        End Enum

        Public Shared CampoObrigatorio_Id As Boolean
        Public Shared CampoObrigatorio_TipoLinha As Boolean
        Public Shared CampoObrigatorio_Artigo As Boolean
        Public Shared CampoObrigatorio_Armazem As Boolean
        Public Shared CampoObrigatorio_Localizacao As Boolean
        Public Shared CampoObrigatorio_Lote As Boolean
        Public Shared CampoObrigatorio_Descricao As Boolean
        Public Shared CampoObrigatorio_CodigoIva As Boolean
        Public Shared CampoObrigatorio_Desconto1 As Boolean
        Public Shared CampoObrigatorio_Desconto2 As Boolean
        Public Shared CampoObrigatorio_Desconto3 As Boolean
        Public Shared CampoObrigatorio_MovimentaStock As Boolean
        Public Shared CampoObrigatorio_Quantidade As Boolean
        Public Shared CampoObrigatorio_QuantidadeSatisfeita As Boolean
        Public Shared CampoObrigatorio_DataEntrega As Boolean
        Public Shared CampoObrigatorio_DataStock As Boolean
        Public Shared CampoObrigatorio_PrecoUnitario As Boolean
        Public Shared CampoObrigatorio_PrecoMedioCusto As Boolean
        Public Shared CampoObrigatorio_Unidade As Boolean
        Public Shared CampoObrigatorio_TaxaIva As Boolean
        Public Shared CampoObrigatorio_PrecoLiquido As Boolean

        Public Shared ComprimentoMaximo_Artigo As System.Int16
        Public Shared ComprimentoMaximo_Armazem As System.Int16
        Public Shared ComprimentoMaximo_Localizacao As System.Int16
        Public Shared ComprimentoMaximo_Lote As System.Int16
        Public Shared ComprimentoMaximo_Descricao As System.Int16
        Public Shared ComprimentoMaximo_CodigoIva As System.Int16
        Public Shared ComprimentoMaximo_Unidade As System.Int16

        Public Shared ComprimentoMinimo_Artigo As System.Int16
        Public Shared ComprimentoMinimo_Armazem As System.Int16
        Public Shared ComprimentoMinimo_Localizacao As System.Int16
        Public Shared ComprimentoMinimo_Lote As System.Int16
        Public Shared ComprimentoMinimo_Descricao As System.Int16
        Public Shared ComprimentoMinimo_CodigoIva As System.Int16
        Public Shared ComprimentoMinimo_Unidade As System.Int16

        Public Shared LimiteInferior_Desconto1 As Single
        Public Shared LimiteInferior_Desconto2 As Single
        Public Shared LimiteInferior_Desconto3 As Single
        Public Shared LimiteInferior_Quantidade As Double
        Public Shared LimiteInferior_QuantidadeSatisfeita As Double
        Public Shared LimiteInferior_PrecoUnitario As Double
        Public Shared LimiteInferior_PrecoMedioCusto As Double
        Public Shared LimiteInferior_TaxaIva As Double
        Public Shared LimiteInferior_PrecoLiquido As Double

        Public Shared LimiteSuperior_Desconto1 As Single
        Public Shared LimiteSuperior_Desconto2 As Single
        Public Shared LimiteSuperior_Desconto3 As Single
        Public Shared LimiteSuperior_Quantidade As Double
        Public Shared LimiteSuperior_QuantidadeSatisfeita As Double
        Public Shared LimiteSuperior_PrecoUnitario As Double
        Public Shared LimiteSuperior_PrecoMedioCusto As Double
        Public Shared LimiteSuperior_TaxaIva As Double
        Public Shared LimiteSuperior_PrecoLiquido As Double

        Public Shared LimiteTemporalInferior_DataEntrega As Date
        Public Shared LimiteTemporalSuperior_DataEntrega As Date
        Public Shared LimiteTemporalInferior_DataStock As Date
        Public Shared LimiteTemporalSuperior_DataStock As Date




        Shared Sub New()

            CampoObrigatorio_Id = True
            CampoObrigatorio_TipoLinha = True
            CampoObrigatorio_MovimentaStock = True
            CampoObrigatorio_Descricao = False 'POR CAUSA DAS LINHAS EM BRANCO QUE POSSAM INTERCALAR LINHAS DO DOCUMENTO
            CampoObrigatorio_Desconto1 = False
            CampoObrigatorio_Desconto2 = False
            CampoObrigatorio_Desconto3 = False
            CampoObrigatorio_DataEntrega = False
            CampoObrigatorio_DataStock = False

            ComprimentoMaximo_Artigo = 48
            ComprimentoMaximo_Armazem = 5
            ComprimentoMaximo_Localizacao = 30
            ComprimentoMaximo_Lote = 20
            ComprimentoMaximo_Descricao = 512
            ComprimentoMaximo_CodigoIva = 2
            ComprimentoMaximo_Unidade = 5

            ComprimentoMinimo_Artigo = 1
            ComprimentoMinimo_Armazem = 1
            ComprimentoMinimo_Armazem = 0
            ComprimentoMinimo_Localizacao = 0
            ComprimentoMinimo_Lote = 1
            ComprimentoMinimo_Descricao = 1
            ComprimentoMinimo_CodigoIva = 1
            ComprimentoMinimo_Unidade = 1

            LimiteInferior_Desconto1 = 0
            LimiteInferior_Desconto2 = 0
            LimiteInferior_Desconto3 = 0
            LimiteInferior_Quantidade = 0
            LimiteInferior_QuantidadeSatisfeita = 0
            LimiteInferior_PrecoUnitario = 0
            LimiteInferior_PrecoMedioCusto = 0
            LimiteInferior_TaxaIva = 0
            LimiteInferior_PrecoLiquido = 0

            LimiteSuperior_Desconto1 = 100
            LimiteSuperior_Desconto2 = 80
            LimiteSuperior_Desconto3 = 50
            LimiteSuperior_Quantidade = 999999
            LimiteSuperior_QuantidadeSatisfeita = 999999
            LimiteSuperior_PrecoUnitario = 999999
            LimiteSuperior_PrecoMedioCusto = 999999
            LimiteSuperior_TaxaIva = 35
            LimiteSuperior_PrecoLiquido = 99999999

            LimiteTemporalInferior_DataEntrega = "30-12-1899"
            LimiteTemporalSuperior_DataEntrega = Now.Date.AddDays(90)
            LimiteTemporalInferior_DataStock = "30-12-1899"
            LimiteTemporalSuperior_DataStock = Now.Date.AddDays(90)

        End Sub



        Private Shared Sub RefreshCamposObrigatorios(ByVal TipoLinha As TiposLinhas, ByVal MovimentaStock As Boolean)
            Select Case TipoLinha

                Case TiposLinhas.Comentario_60
                    CampoObrigatorio_Artigo = False
                    CampoObrigatorio_CodigoIva = False
                    CampoObrigatorio_Quantidade = False
                    CampoObrigatorio_QuantidadeSatisfeita = False
                    CampoObrigatorio_PrecoUnitario = False
                    CampoObrigatorio_Unidade = False
                    CampoObrigatorio_TaxaIva = False
                    CampoObrigatorio_PrecoLiquido = False
                    CampoObrigatorio_Armazem = False
                    CampoObrigatorio_Localizacao = False

                Case TiposLinhas.Acerto_30
                    CampoObrigatorio_Artigo = False
                    CampoObrigatorio_CodigoIva = False
                    CampoObrigatorio_Quantidade = True
                    CampoObrigatorio_QuantidadeSatisfeita = False
                    CampoObrigatorio_PrecoUnitario = True
                    CampoObrigatorio_Unidade = False
                    CampoObrigatorio_TaxaIva = False
                    CampoObrigatorio_PrecoLiquido = True
                    CampoObrigatorio_Armazem = False
                    CampoObrigatorio_Localizacao = False

                Case TiposLinhas.Portes_50
                    CampoObrigatorio_Artigo = False
                    CampoObrigatorio_CodigoIva = True
                    CampoObrigatorio_Quantidade = True
                    CampoObrigatorio_QuantidadeSatisfeita = False
                    CampoObrigatorio_PrecoUnitario = True
                    CampoObrigatorio_Unidade = False
                    CampoObrigatorio_TaxaIva = True
                    CampoObrigatorio_PrecoLiquido = True
                    CampoObrigatorio_Armazem = False
                    CampoObrigatorio_Localizacao = False

                Case Else
                    CampoObrigatorio_Artigo = True
                    CampoObrigatorio_CodigoIva = True
                    CampoObrigatorio_Quantidade = True
                    CampoObrigatorio_QuantidadeSatisfeita = False
                    CampoObrigatorio_PrecoUnitario = True
                    CampoObrigatorio_PrecoMedioCusto = True
                    CampoObrigatorio_Unidade = True
                    CampoObrigatorio_TaxaIva = True
                    CampoObrigatorio_PrecoLiquido = True
                    If MovimentaStock = True Then
                        CampoObrigatorio_Armazem = True
                        CampoObrigatorio_Localizacao = True
                        CampoObrigatorio_Lote = False
                    Else
                        CampoObrigatorio_Armazem = False
                        CampoObrigatorio_Localizacao = False
                        CampoObrigatorio_Lote = False
                    End If

            End Select
        End Sub



        Public Shared Function Certifica(ByRef LinhaDocumentoVenda As BE.LinhaDocumentoVenda, ByRef mensagem As String) As Boolean
            Dim mMensagem As String

            mMensagem = ""

            With LinhaDocumentoVenda

                RefreshCamposObrigatorios(.TipoLinha, .MovimentaStock)

                CCValidation.Guid("Id", .Id, CampoObrigatorio_Id, mMensagem)
                CCValidation.Texto("Artigo", .Artigo, CampoObrigatorio_Artigo, True, ComprimentoMinimo_Artigo, ComprimentoMaximo_Artigo, mMensagem)
                CCValidation.Texto("Armazem", .Armazem, CampoObrigatorio_Armazem, True, ComprimentoMinimo_Armazem, ComprimentoMaximo_Armazem, mMensagem)
                CCValidation.Texto("Localizacao", .Localizacao, CampoObrigatorio_Localizacao, True, ComprimentoMinimo_Localizacao, ComprimentoMaximo_Localizacao, mMensagem)
                CCValidation.Texto("Lote", .Lote, CampoObrigatorio_Lote, True, ComprimentoMinimo_Lote, ComprimentoMaximo_Lote, mMensagem)
                CCValidation.Texto("Descricao", .Descricao, CampoObrigatorio_Descricao, True, ComprimentoMinimo_Descricao, ComprimentoMaximo_Descricao, mMensagem)
                CCValidation.Texto("CodigoIva", .CodigoIva, CampoObrigatorio_CodigoIva, True, ComprimentoMinimo_CodigoIva, ComprimentoMaximo_CodigoIva, mMensagem)
                CCValidation.Decimal("Desconto1", .Desconto1, True, LimiteInferior_Desconto1, LimiteSuperior_Desconto1, mMensagem)
                CCValidation.Decimal("Desconto2", .Desconto2, True, LimiteInferior_Desconto2, LimiteSuperior_Desconto2, mMensagem)
                CCValidation.Decimal("Desconto3", .Desconto3, True, LimiteInferior_Desconto3, LimiteSuperior_Desconto3, mMensagem)
                CCValidation.Decimal("Quantidade", .Quantidade, True, LimiteInferior_Quantidade, LimiteSuperior_Quantidade, mMensagem)
                CCValidation.Decimal("QuantidadeSatisfeita", .QuantidadeSatisfeita, True, LimiteInferior_QuantidadeSatisfeita, LimiteSuperior_QuantidadeSatisfeita, mMensagem)
                CCValidation.Data("DataEntrega", .DataEntrega, True, LimiteTemporalInferior_DataEntrega, LimiteTemporalSuperior_DataEntrega, mMensagem)
                CCValidation.Data("DataStock", .DataStock, True, LimiteTemporalInferior_DataStock, LimiteTemporalSuperior_DataStock, mMensagem)
                CCValidation.Decimal("PrecoUnitario", .PrecoUnitario, True, LimiteInferior_PrecoUnitario, LimiteSuperior_PrecoUnitario, mMensagem)
                CCValidation.Decimal("PrecoMedioCusto", .PrecoMedioCusto, True, LimiteInferior_PrecoMedioCusto, LimiteSuperior_PrecoMedioCusto, mMensagem)
                CCValidation.Texto("Unidade", .Unidade, CampoObrigatorio_Unidade, True, ComprimentoMinimo_Unidade, ComprimentoMaximo_Unidade, mMensagem)
                CCValidation.Decimal("TaxaIva", .TaxaIva, True, LimiteInferior_TaxaIva, LimiteSuperior_TaxaIva, mMensagem)
                CCValidation.Decimal("PrecoLiquido", .PrecoLiquido, True, LimiteInferior_PrecoLiquido, LimiteSuperior_PrecoLiquido, mMensagem)

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

