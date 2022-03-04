Imports Microsoft.VisualBasic
Imports CCUtils


Namespace Specifications

    Public Class LinhaPendente_Specs
        Public Shared CampoObrigatorio_Descricao As Boolean
        Public Shared CampoObrigatorio_CodigoIva As Boolean
        Public Shared CampoObrigatorio_CBLLigacaoGeral As Boolean
        Public Shared CampoObrigatorio_CBLLigacaoAnalitica As Boolean
        Public Shared CampoObrigatorio_CBLLigacaoCentrosCusto As Boolean
        Public Shared CampoObrigatorio_CBLLigacaoFuncional As Boolean

        Public Shared ComprimentoMaximo_Descricao As System.Int16
        Public Shared ComprimentoMaximo_CodigoIva As System.Int16
        Public Shared ComprimentoMaximo_CBLLigacaoGeral As System.Int16
        Public Shared ComprimentoMaximo_CBLLigacaoAnalitica As System.Int16
        Public Shared ComprimentoMaximo_CBLLigacaoCentrosCusto As System.Int16
        Public Shared ComprimentoMaximo_CBLLigacaoFuncional As System.Int16

        Public Shared ComprimentoMinimo_Descricao As System.Int16
        Public Shared ComprimentoMinimo_CodigoIva As System.Int16
        Public Shared ComprimentoMinimo_CBLLigacaoGeral As System.Int16
        Public Shared ComprimentoMinimo_CBLLigacaoAnalitica As System.Int16
        Public Shared ComprimentoMinimo_CBLLigacaoCentrosCusto As System.Int16
        Public Shared ComprimentoMinimo_CBLLigacaoFuncional As System.Int16

        Public Shared LimiteInferior_PercentagemIvaDedutivel As Decimal
        Public Shared LimiteInferior_TaxaProRata As Decimal
        Public Shared LimiteInferior_ValorRecargo As Decimal
        Public Shared LimiteInferior_ValorIncidencia As Decimal
        Public Shared LimiteInferior_ValorIva As Decimal
        Public Shared LimiteInferior_ValorTotal As Decimal

        Public Shared LimiteSuperior_PercentagemIvaDedutivel As Decimal
        Public Shared LimiteSuperior_TaxaProRata As Decimal
        Public Shared LimiteSuperior_ValorRecargo As Decimal
        Public Shared LimiteSuperior_ValorIncidencia As Decimal
        Public Shared LimiteSuperior_ValorIva As Decimal
        Public Shared LimiteSuperior_ValorTotal As Decimal

        Public Shared LimiteTemporalInferior_DataUltimaActualizacao As Date
        Public Shared LimiteTemporalSuperior_DataUltimaActualizacao As Date



        Sub New()

            CampoObrigatorio_Descricao = True
            CampoObrigatorio_CodigoIva = True
            CampoObrigatorio_CBLLigacaoGeral = False
            CampoObrigatorio_CBLLigacaoAnalitica = False
            CampoObrigatorio_CBLLigacaoCentrosCusto = False
            CampoObrigatorio_CBLLigacaoFuncional = False

            ComprimentoMaximo_Descricao = 35
            ComprimentoMaximo_CodigoIva = 2
            ComprimentoMaximo_CBLLigacaoGeral = 20
            ComprimentoMaximo_CBLLigacaoAnalitica = 20
            ComprimentoMaximo_CBLLigacaoCentrosCusto = 15
            ComprimentoMaximo_CBLLigacaoFuncional = 15


            ComprimentoMinimo_Descricao = 0
            ComprimentoMinimo_CodigoIva = 1
            ComprimentoMinimo_CBLLigacaoGeral = 2
            ComprimentoMinimo_CBLLigacaoAnalitica = 2
            ComprimentoMinimo_CBLLigacaoCentrosCusto = 1
            ComprimentoMaximo_CBLLigacaoFuncional = 1

            LimiteInferior_PercentagemIvaDedutivel = 0
            LimiteInferior_TaxaProRata = 0
            LimiteInferior_ValorRecargo = 0
            LimiteInferior_ValorIncidencia = 0.01
            LimiteInferior_ValorIva = 0
            LimiteInferior_ValorTotal = 0.01

            LimiteSuperior_PercentagemIvaDedutivel = 100
            LimiteSuperior_TaxaProRata = 100
            LimiteSuperior_ValorRecargo = 9999
            LimiteSuperior_ValorIncidencia = 9999999
            LimiteSuperior_ValorIva = 9999999
            LimiteSuperior_ValorTotal = 9999999

            LimiteTemporalInferior_DataUltimaActualizacao = "01-07-2007"
            LimiteTemporalSuperior_DataUltimaActualizacao = "31-12-2078"

        End Sub

        Public Shared Function Certifica(ByRef LinhaPendente As BE.LinhaPendente, ByRef mensagem As String) As Boolean
            Dim mMensagem As String

            mMensagem = ""

            With LinhaPendente
                CCValidation.Texto("Descricao", .Descricao, CampoObrigatorio_Descricao, True, ComprimentoMinimo_Descricao, ComprimentoMaximo_Descricao, mMensagem)
                CCValidation.Texto("CodigoIva", .CodigoIva, CampoObrigatorio_CodigoIva, True, ComprimentoMinimo_CodigoIva, ComprimentoMaximo_CodigoIva, mMensagem)
                CCValidation.Decimal("PercentagemIvaDedutivel", .PercentagemIvaDedutivel, True, LimiteInferior_PercentagemIvaDedutivel, LimiteSuperior_PercentagemIvaDedutivel, mMensagem)
                CCValidation.Decimal("TaxaProRata", .TaxaProRata, True, LimiteInferior_TaxaProRata, LimiteSuperior_TaxaProRata, mMensagem)
                CCValidation.Decimal("ValorRecargo", .ValorRecargo, True, LimiteInferior_ValorRecargo, LimiteSuperior_ValorRecargo, mMensagem)
                CCValidation.Decimal("ValorIncidencia", .ValorIncidencia, True, LimiteInferior_ValorIncidencia, LimiteSuperior_ValorIncidencia, mMensagem)
                CCValidation.Decimal("ValorIva", .ValorIva, True, LimiteInferior_ValorIva, LimiteSuperior_ValorIva, mMensagem)
                CCValidation.Decimal("ValorTotal", .ValorTotal, True, LimiteInferior_ValorTotal, LimiteSuperior_ValorTotal, mMensagem)
                CCValidation.Texto("CBLLigacaoGeral", .CBLLigacaoGeral, CampoObrigatorio_CBLLigacaoGeral, New Char() {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}, ComprimentoMinimo_CBLLigacaoGeral, ComprimentoMaximo_CBLLigacaoGeral, mMensagem)
                CCValidation.Texto("CBLLigacaAnalitica", .CBLLigacaoAnalitica, CampoObrigatorio_CBLLigacaoAnalitica, New Char() {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}, ComprimentoMinimo_CBLLigacaoAnalitica, ComprimentoMaximo_CBLLigacaoAnalitica, mMensagem)
                CCValidation.Texto("CBLLigacaoCentrosCusto", .CBLLigacaoCentrosCusto, CampoObrigatorio_CBLLigacaoCentrosCusto, New Char() {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}, ComprimentoMinimo_CBLLigacaoCentrosCusto, ComprimentoMaximo_CBLLigacaoCentrosCusto, mMensagem)
                CCValidation.Texto("CBLLigacaoFuncional", .CBLLigacaoFuncional, CampoObrigatorio_CBLLigacaoFuncional, New Char() {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}, ComprimentoMinimo_CBLLigacaoFuncional, ComprimentoMaximo_CBLLigacaoFuncional, mMensagem)
                CCValidation.Data("DataUltimaActualizacao", .DataUltimaActualizacao, True, LimiteTemporalInferior_DataUltimaActualizacao, LimiteTemporalSuperior_DataUltimaActualizacao, mMensagem)
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

