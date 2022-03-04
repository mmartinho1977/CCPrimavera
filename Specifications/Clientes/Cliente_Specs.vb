Imports Microsoft.VisualBasic
Imports CCUtils


Namespace Specifications


    Public Class Cliente_Specs
        Public Shared CampoObrigatorio_Codigo As Boolean = True
        Public Shared CampoObrigatorio_Nome As Boolean = True
        Public Shared CampoObrigatorio_Morada As Boolean = True
        Public Shared CampoObrigatorio_Morada2 As Boolean = False
        Public Shared CampoObrigatorio_Localidade As Boolean = False
        Public Shared CampoObrigatorio_CodigoPostal As Boolean = True
        Public Shared CampoObrigatorio_LocalidadePostal As Boolean = True
        Public Shared CampoObrigatorio_Telefone As Boolean = False
        Public Shared CampoObrigatorio_Telefone2 As Boolean = False
        Public Shared CampoObrigatorio_Fax As Boolean = False
        Public Shared CampoObrigatorio_Contribuinte As Boolean = True
        Public Shared CampoObrigatorio_CondicaoPagamento As Boolean = True
        Public Shared CampoObrigatorio_Moeda As Boolean = True
        Public Shared CampoObrigatorio_ModoPagamento As Boolean = True
        Public Shared CampoObrigatorio_TabelaPrecos As Boolean = True

        Public Shared ComprimentoMaximo_Codigo As System.Int16 = 12
        Public Shared ComprimentoMaximo_Nome As System.Int16 = 50
        Public Shared ComprimentoMaximo_Morada As System.Int16 = 50
        Public Shared ComprimentoMaximo_Morada2 As System.Int16 = 50
        Public Shared ComprimentoMaximo_Localidade As System.Int16 = 50
        Public Shared ComprimentoMaximo_CodigoPostal As System.Int16 = 15
        Public Shared ComprimentoMaximo_LocalidadePostal As System.Int16 = 50
        Public Shared ComprimentoMaximo_Telefone As System.Int16 = 15
        Public Shared ComprimentoMaximo_Telefone2 As System.Int16 = 15
        Public Shared ComprimentoMaximo_Fax As System.Int16 = 15
        Public Shared ComprimentoMaximo_Contribuinte As System.Int16 = 15
        Public Shared ComprimentoMaximo_CondicaoPagamento As System.Int16 = 2
        Public Shared ComprimentoMaximo_Moeda As System.Int16 = 3
        Public Shared ComprimentoMaximo_ModoPagamento As System.Int16 = 5
        Public Shared ComprimentoMaximo_TabelaPrecos As System.Int16 = 4

        Public Shared ComprimentoMinimo_Codigo As System.Int16 = 1
        Public Shared ComprimentoMinimo_Nome As System.Int16 = 10
        Public Shared ComprimentoMinimo_Morada As System.Int16 = 5
        Public Shared ComprimentoMinimo_Morada2 As System.Int16 = 5
        Public Shared ComprimentoMinimo_Localidade As System.Int16 = 5
        Public Shared ComprimentoMinimo_CodigoPostal As System.Int16 = 8
        Public Shared ComprimentoMinimo_LocalidadePostal As System.Int16 = 4
        Public Shared ComprimentoMinimo_Telefone As System.Int16 = 9
        Public Shared ComprimentoMinimo_Telefone2 As System.Int16 = 9
        Public Shared ComprimentoMinimo_Fax As System.Int16 = 9
        Public Shared ComprimentoMinimo_Contribuinte As System.Int16 = 6
        Public Shared ComprimentoMinimo_CondicaoPagamento As System.Int16 = 1
        Public Shared ComprimentoMinimo_Moeda As System.Int16 = 2
        Public Shared ComprimentoMinimo_ModoPagamento As System.Int16 = 2
        Public Shared ComprimentoMinimo_TabelaPrecos As System.Int16 = 3

        Public Shared LimiteTemporalInferior_DataCriacao As Date = "01-01-1905"
        Public Shared LimiteTemporalSuperior_DataCriacao As Date = "31-12-2078"
        Public Shared LimiteTemporalInferior_DataUltimaActualizacao As Date = "01-07-2007"
        Public Shared LimiteTemporalSuperior_DataUltimaActualizacao As Date = "31-12-2078"


        Sub New()

        End Sub




        Public Shared Function Certifica(ByRef Cliente As BE.Cliente, ByRef mensagem As String) As Boolean
            Dim mMensagem As String

            mMensagem = ""

            With Cliente

                If .EmModoEdicao = False Then
                    If .Codigo.TrimEnd.Length > 0 Then
                        mMensagem += "Codigo: no modo inserção o comprimento deste campo deve ser vazio; "
                    End If
                Else
                    CCValidation.Texto("Codigo", .Codigo, CampoObrigatorio_Codigo, True, ComprimentoMinimo_Codigo, ComprimentoMaximo_Codigo, mMensagem)
                End If

                CCValidation.Texto("Nome", .Nome, CampoObrigatorio_Nome, True, ComprimentoMinimo_Nome, ComprimentoMaximo_Nome, mMensagem)
                CCValidation.Texto("Morada", .Morada, CampoObrigatorio_Morada, True, ComprimentoMinimo_Morada, ComprimentoMaximo_Morada, mMensagem)
                CCValidation.Texto("Morada2", .Morada2, CampoObrigatorio_Morada2, True, ComprimentoMinimo_Morada2, ComprimentoMaximo_Morada2, mMensagem)
                CCValidation.Texto("Localidade", .Localidade, CampoObrigatorio_Localidade, True, ComprimentoMinimo_Localidade, ComprimentoMaximo_Localidade, mMensagem)
                CCValidation.Texto("CodigoPostal", .CodigoPostal, CampoObrigatorio_CodigoPostal, True, ComprimentoMinimo_CodigoPostal, ComprimentoMaximo_CodigoPostal, mMensagem)
                CCValidation.Texto("LocalidadePostal", .LocalidadePostal, CampoObrigatorio_LocalidadePostal, True, ComprimentoMinimo_LocalidadePostal, ComprimentoMaximo_LocalidadePostal, mMensagem)
                CCValidation.Texto("Telefone", .Telefone, CampoObrigatorio_Telefone, New Char() {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}, ComprimentoMinimo_Telefone, ComprimentoMaximo_Telefone, mMensagem)
                CCValidation.Texto("Telefone2", .Telefone2, CampoObrigatorio_Telefone2, New Char() {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}, ComprimentoMinimo_Telefone2, ComprimentoMaximo_Telefone2, mMensagem)
                CCValidation.Texto("Fax", .Fax, CampoObrigatorio_Fax, New Char() {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}, ComprimentoMinimo_Fax, ComprimentoMaximo_Fax, mMensagem)
                CCValidation.Texto("Contribuinte", .Contribuinte, CampoObrigatorio_Contribuinte, New Char() {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}, ComprimentoMinimo_Contribuinte, ComprimentoMaximo_Contribuinte, mMensagem)
                CCValidation.Texto("CondicaoPagamento", .CondicaoPagamento, CampoObrigatorio_CondicaoPagamento, True, ComprimentoMinimo_CondicaoPagamento, ComprimentoMaximo_CondicaoPagamento, mMensagem)
                CCValidation.Texto("Moeda", .Moeda, CampoObrigatorio_Moeda, True, ComprimentoMinimo_Moeda, ComprimentoMaximo_Moeda, mMensagem)
                CCValidation.Texto("ModoPagamento", .ModoPagamento, CampoObrigatorio_ModoPagamento, True, ComprimentoMinimo_ModoPagamento, ComprimentoMaximo_ModoPagamento, mMensagem)
                CCValidation.Texto("TabelaPrecos", .TabelaPrecos, CampoObrigatorio_TabelaPrecos, True, ComprimentoMinimo_TabelaPrecos, ComprimentoMaximo_TabelaPrecos, mMensagem)
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

