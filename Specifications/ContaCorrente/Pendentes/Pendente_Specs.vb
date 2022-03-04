Imports Microsoft.VisualBasic
Imports CCUtils


Namespace Specifications

    Public Class Pendente_Specs
        Public Shared CampoObrigatorio_Filial As Boolean
        Public Shared CampoObrigatorio_Modulo As Boolean
        Public Shared CampoObrigatorio_TipoDocumento As Boolean
        Public Shared CampoObrigatorio_SerieDocumento As Boolean
        Public Shared CampoObrigatorio_NumeroDocumento As Boolean
        Public Shared CampoObrigatorio_TipoEntidade As Boolean
        Public Shared CampoObrigatorio_Entidade As Boolean
        Public Shared CampoObrigatorio_TipoConta As Boolean
        Public Shared CampoObrigatorio_Estado As Boolean
        Public Shared CampoObrigatorio_CondicaoPagamento As Boolean
        Public Shared CampoObrigatorio_ModoPagamento As Boolean
        Public Shared CampoObrigatorio_Moeda As Boolean
        Public Shared CampoObrigatorio_Observacoes As Boolean
        Public Shared CampoObrigatorio_Utilizador As Boolean

        Public Shared ComprimentoMaximo_Filial As System.Int16
        Public Shared ComprimentoMaximo_Modulo As System.Int16
        Public Shared ComprimentoMaximo_TipoDocumento As System.Int16
        Public Shared ComprimentoMaximo_SerieDocumento As System.Int16
        Public Shared ComprimentoMaximo_NumeroDocumento As System.Int16
        Public Shared ComprimentoMaximo_TipoEntidade As System.Int16
        Public Shared ComprimentoMaximo_Entidade As System.Int16
        Public Shared ComprimentoMaximo_TipoConta As System.Int16
        Public Shared ComprimentoMaximo_Estado As System.Int16
        Public Shared ComprimentoMaximo_CondicaoPagamento As System.Int16
        Public Shared ComprimentoMaximo_ModoPagamento As System.Int16
        Public Shared ComprimentoMaximo_Moeda As System.Int16
        Public Shared ComprimentoMaximo_Observacoes As System.Int16
        Public Shared ComprimentoMaximo_Utilizador As System.Int16

        Public Shared ComprimentoMinimo_Filial As System.Int16
        Public Shared ComprimentoMinimo_Modulo As System.Int16
        Public Shared ComprimentoMinimo_TipoDocumento As System.Int16
        Public Shared ComprimentoMinimo_SerieDocumento As System.Int16
        Public Shared ComprimentoMinimo_NumeroDocumento As System.Int16
        Public Shared ComprimentoMinimo_TipoEntidade As System.Int16
        Public Shared ComprimentoMinimo_Entidade As System.Int16
        Public Shared ComprimentoMinimo_TipoConta As System.Int16
        Public Shared ComprimentoMinimo_Estado As System.Int16
        Public Shared ComprimentoMinimo_CondicaoPagamento As System.Int16
        Public Shared ComprimentoMinimo_ModoPagamento As System.Int16
        Public Shared ComprimentoMinimo_Moeda As System.Int16
        Public Shared ComprimentoMinimo_Observacoes As System.Int16
        Public Shared ComprimentoMinimo_Utilizador As System.Int16

        Public Shared LimiteInferior_NumeroDocumentoInterno As System.Int64
        Public Shared LimiteSuperior_NumeroDocumentoInterno As System.Int64
        Public Shared LimiteInferior_ValorTotal As Decimal
        Public Shared LimiteSuperior_ValorTotal As Decimal
        Public Shared LimiteInferior_ValorPendente As Decimal
        Public Shared LimiteSuperior_ValorPendente As Decimal

        Public Shared LimiteTemporalInferior_DataDocumento As Date
        Public Shared LimiteTemporalSuperior_DataDocumento As Date
        Public Shared LimiteTemporalInferior_DataVencimento As Date
        Public Shared LimiteTemporalSuperior_DataVencimento As Date
        Public Shared LimiteTemporalInferior_DataIntroducao As Date
        Public Shared LimiteTemporalSuperior_DataIntroducao As Date



        Sub New()
            CampoObrigatorio_Filial = True
            CampoObrigatorio_Modulo = True
            CampoObrigatorio_TipoDocumento = True
            CampoObrigatorio_SerieDocumento = True
            CampoObrigatorio_NumeroDocumento = True
            CampoObrigatorio_TipoEntidade = True
            CampoObrigatorio_Entidade = True
            CampoObrigatorio_TipoConta = True
            CampoObrigatorio_Estado = True
            CampoObrigatorio_CondicaoPagamento = True
            CampoObrigatorio_ModoPagamento = True
            CampoObrigatorio_Moeda = True
            CampoObrigatorio_Observacoes = False
            CampoObrigatorio_Utilizador = True

            ComprimentoMaximo_Filial = 3
            ComprimentoMaximo_Modulo = 1
            ComprimentoMaximo_TipoDocumento = 5
            ComprimentoMaximo_SerieDocumento = 5
            ComprimentoMaximo_NumeroDocumento = 20
            ComprimentoMaximo_TipoEntidade = 1
            ComprimentoMaximo_Entidade = 12
            ComprimentoMaximo_TipoConta = 3
            ComprimentoMaximo_Estado = 4
            ComprimentoMaximo_CondicaoPagamento = 2
            ComprimentoMaximo_ModoPagamento = 5
            ComprimentoMaximo_Moeda = 3
            ComprimentoMaximo_Observacoes = 512
            ComprimentoMaximo_Utilizador = 20

            ComprimentoMinimo_Filial = 3
            ComprimentoMinimo_Modulo = 1
            ComprimentoMinimo_TipoDocumento = 2
            ComprimentoMinimo_SerieDocumento = 1
            ComprimentoMinimo_NumeroDocumento = 1
            ComprimentoMinimo_TipoEntidade = 1
            ComprimentoMinimo_Entidade = 1
            ComprimentoMinimo_TipoConta = 2
            ComprimentoMinimo_Estado = 2
            ComprimentoMinimo_CondicaoPagamento = 1
            ComprimentoMinimo_ModoPagamento = 2
            ComprimentoMinimo_Moeda = 2
            ComprimentoMinimo_Observacoes = 0
            ComprimentoMinimo_Utilizador = 3

            LimiteInferior_NumeroDocumentoInterno = 1
            LimiteSuperior_NumeroDocumentoInterno = 999999999

            LimiteInferior_ValorTotal = 0.01
            LimiteSuperior_ValorTotal = 9999999

            LimiteInferior_ValorPendente = 0.01
            LimiteSuperior_ValorPendente = 9999999

            LimiteTemporalInferior_DataDocumento = "01-01-2008"
            LimiteTemporalSuperior_DataDocumento = "01-01-2078"
            LimiteTemporalInferior_DataVencimento = "01-01-2008"
            LimiteTemporalSuperior_DataVencimento = "01-01-2078"
            LimiteTemporalInferior_DataIntroducao = "01-01-2008"
            LimiteTemporalSuperior_DataIntroducao = "31-12-2078"

        End Sub



        Public Shared Function Certifica(ByRef Pendente As BE.Pendente, ByRef mensagem As String) As Boolean
            Dim mMensagem As String
            Dim i As System.Int16

            mMensagem = ""

            With Pendente

                If .EmModoEdicao = False Then
                    If .NumeroDocumentoInterno > 0 Then
                        mMensagem += "NumeroDocumentoInterno: Numeração deve ser atribuida pelo sistema; "
                    End If
                Else
                    CCValidation.Int64("NumeroDocumentoInterno", .NumeroDocumentoInterno, True, LimiteInferior_NumeroDocumentoInterno, LimiteSuperior_NumeroDocumentoInterno, mMensagem)
                End If

                CCValidation.Texto("Filial", .Filial, CampoObrigatorio_Filial, New Char() {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}, ComprimentoMinimo_Filial, ComprimentoMaximo_Filial, mMensagem)
                CCValidation.Texto("Modulo", .Modulo, CampoObrigatorio_Modulo, True, ComprimentoMinimo_Modulo, ComprimentoMaximo_Modulo, mMensagem)
                CCValidation.Texto("TipoDocumento", .TipoDocumento, CampoObrigatorio_TipoDocumento, True, ComprimentoMinimo_TipoDocumento, ComprimentoMaximo_TipoDocumento, mMensagem)
                CCValidation.Texto("SerieDocumento", .SerieDocumento, CampoObrigatorio_SerieDocumento, True, ComprimentoMinimo_SerieDocumento, ComprimentoMaximo_SerieDocumento, mMensagem)
                CCValidation.Texto("NumeroDocumento", .NumeroDocumento, CampoObrigatorio_NumeroDocumento, True, ComprimentoMinimo_NumeroDocumento, ComprimentoMaximo_NumeroDocumento, mMensagem)
                CCValidation.Texto("TipoEntidade", .TipoEntidade, CampoObrigatorio_TipoEntidade, True, ComprimentoMinimo_TipoEntidade, ComprimentoMaximo_TipoEntidade, mMensagem)
                CCValidation.Texto("Entidade", .Entidade, CampoObrigatorio_Entidade, True, ComprimentoMinimo_Entidade, ComprimentoMaximo_Entidade, mMensagem)
                CCValidation.Texto("TipoConta", .TipoConta, CampoObrigatorio_TipoConta, True, ComprimentoMinimo_TipoConta, ComprimentoMaximo_TipoConta, mMensagem)
                CCValidation.Texto("Estado", .Estado, CampoObrigatorio_Estado, True, ComprimentoMinimo_Estado, ComprimentoMaximo_Estado, mMensagem)
                CCValidation.Texto("CondicaoPagamento", .CondicaoPagamento, CampoObrigatorio_CondicaoPagamento, True, ComprimentoMinimo_CondicaoPagamento, ComprimentoMaximo_CondicaoPagamento, mMensagem)
                CCValidation.Texto("ModoPagamento", .ModoPagamento, CampoObrigatorio_ModoPagamento, True, ComprimentoMinimo_ModoPagamento, ComprimentoMaximo_ModoPagamento, mMensagem)
                CCValidation.Texto("Moeda", .Moeda, CampoObrigatorio_Moeda, True, ComprimentoMinimo_Moeda, ComprimentoMaximo_Moeda, mMensagem)
                CCValidation.Texto("Observacoes", .Observacoes, CampoObrigatorio_Observacoes, True, ComprimentoMinimo_Observacoes, ComprimentoMaximo_Observacoes, mMensagem)
                CCValidation.Texto("Utilizador", .Utilizador, CampoObrigatorio_Utilizador, True, ComprimentoMinimo_Utilizador, ComprimentoMaximo_Utilizador, mMensagem)
                CCValidation.Decimal("ValorTotal", .ValorTotal, True, LimiteInferior_ValorTotal, LimiteSuperior_ValorTotal, mMensagem)
                CCValidation.Decimal("ValorPendente", .ValorPendente, True, LimiteInferior_ValorPendente, LimiteSuperior_ValorPendente, mMensagem)
                CCValidation.Data("DataDocumento", .DataDocumento, True, LimiteTemporalInferior_DataDocumento, LimiteTemporalSuperior_DataDocumento, mMensagem)
                CCValidation.Data("DataVencimento", .DataVencimento, True, LimiteTemporalInferior_DataVencimento, LimiteTemporalSuperior_DataVencimento, mMensagem)
                CCValidation.Data("DataIntroducao", .DataIntroducao, True, LimiteTemporalInferior_DataIntroducao, LimiteTemporalSuperior_DataIntroducao, mMensagem)

                For i = 0 To Pendente.Linhas.Count - 1
                    LinhaPendente_Specs.Certifica(.Linhas(i), mMensagem)
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

