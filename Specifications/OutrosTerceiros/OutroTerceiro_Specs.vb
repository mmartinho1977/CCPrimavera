Imports Microsoft.VisualBasic
Imports CCUtils


Namespace Specifications

    Public Class OutroTerceiro_Specs
        Public Shared CampoObrigatorio_Codigo As Boolean
        Public Shared CampoObrigatorio_Nome As Boolean
        Public Shared CampoObrigatorio_Morada As Boolean
        Public Shared CampoObrigatorio_Localidade As Boolean
        Public Shared CampoObrigatorio_CodigoPostal As Boolean
        Public Shared CampoObrigatorio_LocalidadePostal As Boolean
        Public Shared CampoObrigatorio_Telefone As Boolean
        Public Shared CampoObrigatorio_Fax As Boolean
        Public Shared CampoObrigatorio_Contribuinte As Boolean
        Public Shared CampoObrigatorio_CondicaoPagamento As Boolean
        Public Shared CampoObrigatorio_Moeda As Boolean
        Public Shared CampoObrigatorio_ModoPagamento As Boolean

        Public Shared ComprimentoMaximo_Codigo As System.Int16
        Public Shared ComprimentoMaximo_Nome As System.Int16
        Public Shared ComprimentoMaximo_Morada As System.Int16
        Public Shared ComprimentoMaximo_Localidade As System.Int16
        Public Shared ComprimentoMaximo_CodigoPostal As System.Int16
        Public Shared ComprimentoMaximo_LocalidadePostal As System.Int16
        Public Shared ComprimentoMaximo_Telefone As System.Int16
        Public Shared ComprimentoMaximo_Fax As System.Int16
        Public Shared ComprimentoMaximo_Contribuinte As System.Int16
        Public Shared ComprimentoMaximo_CondicaoPagamento As System.Int16
        Public Shared ComprimentoMaximo_Moeda As System.Int16
        Public Shared ComprimentoMaximo_ModoPagamento As System.Int16

        Public Shared ComprimentoMinimo_Codigo As System.Int16
        Public Shared ComprimentoMinimo_Nome As System.Int16
        Public Shared ComprimentoMinimo_Morada As System.Int16
        Public Shared ComprimentoMinimo_Localidade As System.Int16
        Public Shared ComprimentoMinimo_CodigoPostal As System.Int16
        Public Shared ComprimentoMinimo_LocalidadePostal As System.Int16
        Public Shared ComprimentoMinimo_Telefone As System.Int16
        Public Shared ComprimentoMinimo_Fax As System.Int16
        Public Shared ComprimentoMinimo_Contribuinte As System.Int16
        Public Shared ComprimentoMinimo_CondicaoPagamento As System.Int16
        Public Shared ComprimentoMinimo_Moeda As System.Int16
        Public Shared ComprimentoMinimo_ModoPagamento As System.Int16

        Public Shared LimiteTemporalInferior_DataCriacao As Date
        Public Shared LimiteTemporalSuperior_DataCriacao As Date
        Public Shared LimiteTemporalInferior_DataUltimaActualizacao As Date
        Public Shared LimiteTemporalSuperior_DataUltimaActualizacao As Date


        Sub New(ByVal TipoEntidade As String)

            CampoObrigatorio_Codigo = True
            CampoObrigatorio_Nome = True
            CampoObrigatorio_Morada = True
            CampoObrigatorio_Localidade = False
            CampoObrigatorio_CodigoPostal = True
            CampoObrigatorio_LocalidadePostal = True
            CampoObrigatorio_Telefone = False
            CampoObrigatorio_Fax = False
            CampoObrigatorio_Contribuinte = False
            CampoObrigatorio_CondicaoPagamento = True
            CampoObrigatorio_Moeda = True
            CampoObrigatorio_ModoPagamento = True

            ComprimentoMaximo_Codigo = 12
            ComprimentoMaximo_Nome = 50
            ComprimentoMaximo_Morada = 50
            ComprimentoMaximo_Localidade = 50
            ComprimentoMaximo_CodigoPostal = 15
            ComprimentoMaximo_LocalidadePostal = 50
            ComprimentoMaximo_Telefone = 15
            ComprimentoMaximo_Fax = 15
            ComprimentoMaximo_Contribuinte = 15
            ComprimentoMaximo_CondicaoPagamento = 2
            ComprimentoMaximo_Moeda = 3
            ComprimentoMaximo_ModoPagamento = 5

            ComprimentoMinimo_Codigo = 5
            ComprimentoMinimo_Nome = 10
            ComprimentoMinimo_Morada = 7
            ComprimentoMinimo_Localidade = 5
            ComprimentoMinimo_CodigoPostal = 8
            ComprimentoMinimo_LocalidadePostal = 4
            ComprimentoMinimo_Telefone = 9
            ComprimentoMinimo_Fax = 9
            ComprimentoMinimo_Contribuinte = 6
            ComprimentoMinimo_CondicaoPagamento = 1
            ComprimentoMinimo_Moeda = 2
            ComprimentoMinimo_ModoPagamento = 2

            LimiteTemporalInferior_DataCriacao = "01-01-1905"
            LimiteTemporalSuperior_DataCriacao = "01-01-2078"
            LimiteTemporalInferior_DataUltimaActualizacao = "01-07-2007"
            LimiteTemporalSuperior_DataUltimaActualizacao = "31-12-2078"

        End Sub


        Public Shared Function Certifica(ByRef OutroTerceiro As BE.OutroTerceiro, ByRef mensagem As String) As Boolean
            Dim mMensagem As String

            mMensagem = ""

            With OutroTerceiro

                If .EmModoEdicao = False Then
                    If .Codigo.TrimEnd.Length > 0 Then
                        mMensagem += "Codigo: no modo inserção o comprimento deste campo deve ser vazio; "
                    End If
                Else
                    CCValidation.Texto("Codigo", .Codigo, CampoObrigatorio_Codigo, True, ComprimentoMinimo_Codigo, ComprimentoMaximo_Codigo, mMensagem)
                End If
                CCValidation.Texto("Nome", .Nome, CampoObrigatorio_Nome, True, ComprimentoMinimo_Nome, ComprimentoMaximo_Nome, mMensagem)
                CCValidation.Texto("Morada", .Morada, CampoObrigatorio_Morada, True, ComprimentoMinimo_Morada, ComprimentoMaximo_Morada, mMensagem)
                CCValidation.Texto("Localidade", .Localidade, CampoObrigatorio_Localidade, True, ComprimentoMinimo_Localidade, ComprimentoMaximo_Localidade, mMensagem)
                CCValidation.Texto("CodigoPostal", .CodigoPostal, CampoObrigatorio_CodigoPostal, True, ComprimentoMinimo_CodigoPostal, ComprimentoMaximo_CodigoPostal, mMensagem)
                CCValidation.Texto("LocalidadePostal", .LocalidadePostal, CampoObrigatorio_LocalidadePostal, True, ComprimentoMinimo_LocalidadePostal, ComprimentoMaximo_LocalidadePostal, mMensagem)
                CCValidation.Texto("Telefone", .Telefone, CampoObrigatorio_Telefone, New Char() {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}, ComprimentoMinimo_Telefone, ComprimentoMaximo_Telefone, mMensagem)
                CCValidation.Texto("Fax", .Fax, CampoObrigatorio_Fax, New Char() {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}, ComprimentoMinimo_Fax, ComprimentoMaximo_Fax, mMensagem)
                CCValidation.Texto("Contribuinte", .Contribuinte, CampoObrigatorio_Contribuinte, New Char() {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}, ComprimentoMinimo_Contribuinte, ComprimentoMaximo_Contribuinte, mMensagem)
                CCValidation.Texto("CondicaoPagamento", .CondicaoPagamento, CampoObrigatorio_CondicaoPagamento, True, ComprimentoMinimo_CondicaoPagamento, ComprimentoMaximo_CondicaoPagamento, mMensagem)
                CCValidation.Texto("Moeda", .Moeda, CampoObrigatorio_Moeda, True, ComprimentoMinimo_Moeda, ComprimentoMaximo_Moeda, mMensagem)
                CCValidation.Texto("ModoPagamento", .ModoPagamento, CampoObrigatorio_ModoPagamento, True, ComprimentoMinimo_ModoPagamento, ComprimentoMaximo_ModoPagamento, mMensagem)
                CCValidation.Data("DataCriacao", .DataCriacao, True, LimiteTemporalInferior_DataCriacao, LimiteTemporalSuperior_DataCriacao, mMensagem)
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

