Imports Microsoft.VisualBasic
Imports CCUtils


Namespace Specifications


    Public Class CargaDescarga_Specs
        Public Shared CampoObrigatorio_CARGA_Morada As Boolean
        Public Shared CampoObrigatorio_CARGA_Morada2 As Boolean
        Public Shared CampoObrigatorio_CARGA_Localidade As Boolean
        Public Shared CampoObrigatorio_CARGA_CodigoPostal As Boolean
        Public Shared CampoObrigatorio_CARGA_LocalidadePostal As Boolean
        Public Shared CampoObrigatorio_CARGA_Distrito As Boolean
        Public Shared CampoObrigatorio_CARGA_Pais As Boolean
        Public Shared CampoObrigatorio_CARGA_Data As Boolean
        Public Shared CampoObrigatorio_CARGA_Hora As Boolean
        Public Shared CampoObrigatorio_DESCARGA_TipoEntidade As Boolean
        Public Shared CampoObrigatorio_DESCARGA_Entidade As Boolean
        Public Shared CampoObrigatorio_DESCARGA_Nome As Boolean
        Public Shared CampoObrigatorio_DESCARGA_Morada As Boolean
        Public Shared CampoObrigatorio_DESCARGA_Morada2 As Boolean
        Public Shared CampoObrigatorio_DESCARGA_Localidade As Boolean
        Public Shared CampoObrigatorio_DESCARGA_CodigoPostal As Boolean
        Public Shared CampoObrigatorio_DESCARGA_LocalidadePostal As Boolean
        Public Shared CampoObrigatorio_DESCARGA_Distrito As Boolean
        Public Shared CampoObrigatorio_DESCARGA_Pais As Boolean
        Public Shared CampoObrigatorio_DESCARGA_Data As Boolean
        Public Shared CampoObrigatorio_DESCARGA_Hora As Boolean


        Public Shared ComprimentoMaximo_CARGA_Morada As System.Int16
        Public Shared ComprimentoMaximo_CARGA_Morada2 As System.Int16
        Public Shared ComprimentoMaximo_CARGA_Localidade As System.Int16
        Public Shared ComprimentoMaximo_CARGA_CodigoPostal As System.Int16
        Public Shared ComprimentoMaximo_CARGA_LocalidadePostal As System.Int16
        Public Shared ComprimentoMaximo_CARGA_Distrito As System.Int16
        Public Shared ComprimentoMaximo_CARGA_Pais As System.Int16
        Public Shared ComprimentoMaximo_CARGA_Data As System.Int16
        Public Shared ComprimentoMaximo_CARGA_Hora As System.Int16
        Public Shared ComprimentoMaximo_DESCARGA_TipoEntidade As System.Int16
        Public Shared ComprimentoMaximo_DESCARGA_Entidade As System.Int16
        Public Shared ComprimentoMaximo_DESCARGA_Nome As System.Int16
        Public Shared ComprimentoMaximo_DESCARGA_Morada As System.Int16
        Public Shared ComprimentoMaximo_DESCARGA_Morada2 As System.Int16
        Public Shared ComprimentoMaximo_DESCARGA_Localidade As System.Int16
        Public Shared ComprimentoMaximo_DESCARGA_CodigoPostal As System.Int16
        Public Shared ComprimentoMaximo_DESCARGA_LocalidadePostal As System.Int16
        Public Shared ComprimentoMaximo_DESCARGA_Distrito As System.Int16
        Public Shared ComprimentoMaximo_DESCARGA_Pais As System.Int16
        Public Shared ComprimentoMaximo_DESCARGA_Data As System.Int16
        Public Shared ComprimentoMaximo_DESCARGA_Hora As System.Int16

        Public Shared ComprimentoMinimo_CARGA_Morada As System.Int16
        Public Shared ComprimentoMinimo_CARGA_Morada2 As System.Int16
        Public Shared ComprimentoMinimo_CARGA_Localidade As System.Int16
        Public Shared ComprimentoMinimo_CARGA_CodigoPostal As System.Int16
        Public Shared ComprimentoMinimo_CARGA_LocalidadePostal As System.Int16
        Public Shared ComprimentoMinimo_CARGA_Distrito As System.Int16
        Public Shared ComprimentoMinimo_CARGA_Pais As System.Int16
        Public Shared ComprimentoMinimo_CARGA_Data As System.Int16
        Public Shared ComprimentoMinimo_CARGA_Hora As System.Int16
        Public Shared ComprimentoMinimo_DESCARGA_TipoEntidade As System.Int16
        Public Shared ComprimentoMinimo_DESCARGA_Entidade As System.Int16
        Public Shared ComprimentoMinimo_DESCARGA_Nome As System.Int16
        Public Shared ComprimentoMinimo_DESCARGA_Morada As System.Int16
        Public Shared ComprimentoMinimo_DESCARGA_Morada2 As System.Int16
        Public Shared ComprimentoMinimo_DESCARGA_Localidade As System.Int16
        Public Shared ComprimentoMinimo_DESCARGA_CodigoPostal As System.Int16
        Public Shared ComprimentoMinimo_DESCARGA_LocalidadePostal As System.Int16
        Public Shared ComprimentoMinimo_DESCARGA_Distrito As System.Int16
        Public Shared ComprimentoMinimo_DESCARGA_Pais As System.Int16
        Public Shared ComprimentoMinimo_DESCARGA_Data As System.Int16
        Public Shared ComprimentoMinimo_DESCARGA_Hora As System.Int16


        Sub New()

            CampoObrigatorio_CARGA_Morada = True
            CampoObrigatorio_CARGA_Morada2 = False
            CampoObrigatorio_CARGA_Localidade = True
            CampoObrigatorio_CARGA_CodigoPostal = True
            CampoObrigatorio_CARGA_LocalidadePostal = True
            CampoObrigatorio_CARGA_Distrito = True
            CampoObrigatorio_CARGA_Pais = True
            CampoObrigatorio_CARGA_Data = True
            CampoObrigatorio_CARGA_Hora = True
            CampoObrigatorio_DESCARGA_TipoEntidade = True
            CampoObrigatorio_DESCARGA_Entidade = True
            CampoObrigatorio_DESCARGA_Nome = True
            CampoObrigatorio_DESCARGA_Morada = True
            CampoObrigatorio_DESCARGA_Morada2 = False
            CampoObrigatorio_DESCARGA_Localidade = True
            CampoObrigatorio_DESCARGA_CodigoPostal = True
            CampoObrigatorio_DESCARGA_LocalidadePostal = True
            CampoObrigatorio_DESCARGA_Distrito = True
            CampoObrigatorio_DESCARGA_Pais = True
            CampoObrigatorio_DESCARGA_Data = True
            CampoObrigatorio_DESCARGA_Hora = False


            ComprimentoMaximo_CARGA_Morada = 50
            ComprimentoMaximo_CARGA_Morada2 = 50
            ComprimentoMaximo_CARGA_Localidade = 50
            ComprimentoMaximo_CARGA_CodigoPostal = 15
            ComprimentoMaximo_CARGA_LocalidadePostal = 50
            ComprimentoMaximo_CARGA_Distrito = 2
            ComprimentoMaximo_CARGA_Pais = 2
            ComprimentoMaximo_CARGA_Data = 10
            ComprimentoMaximo_CARGA_Hora = 5
            ComprimentoMaximo_DESCARGA_TipoEntidade = 1
            ComprimentoMaximo_DESCARGA_Entidade = 12
            ComprimentoMaximo_DESCARGA_Nome = 50
            ComprimentoMaximo_DESCARGA_Morada = 50
            ComprimentoMaximo_DESCARGA_Morada2 = 50
            ComprimentoMaximo_DESCARGA_Localidade = 50
            ComprimentoMaximo_DESCARGA_CodigoPostal = 15
            ComprimentoMaximo_DESCARGA_LocalidadePostal = 50
            ComprimentoMaximo_DESCARGA_Distrito = 2
            ComprimentoMaximo_DESCARGA_Pais = 2
            ComprimentoMaximo_DESCARGA_Data = 10
            ComprimentoMaximo_DESCARGA_Hora = 5


            ComprimentoMinimo_CARGA_Morada = 5
            ComprimentoMinimo_CARGA_Morada2 = 3
            ComprimentoMinimo_CARGA_Localidade = 5
            ComprimentoMinimo_CARGA_CodigoPostal = 8
            ComprimentoMinimo_CARGA_LocalidadePostal = 4
            ComprimentoMinimo_CARGA_Distrito = 2
            ComprimentoMinimo_CARGA_Pais = 2
            ComprimentoMinimo_CARGA_Data = 10
            ComprimentoMinimo_CARGA_Hora = 5
            ComprimentoMinimo_DESCARGA_TipoEntidade = 1
            ComprimentoMinimo_DESCARGA_Entidade = 2
            ComprimentoMinimo_DESCARGA_Nome = 10
            ComprimentoMinimo_DESCARGA_Morada = 5
            ComprimentoMinimo_DESCARGA_Morada2 = 3
            ComprimentoMinimo_DESCARGA_Localidade = 5
            ComprimentoMinimo_DESCARGA_CodigoPostal = 8
            ComprimentoMinimo_DESCARGA_LocalidadePostal = 4
            ComprimentoMinimo_DESCARGA_Distrito = 2
            ComprimentoMinimo_DESCARGA_Pais = 2
            ComprimentoMinimo_DESCARGA_Data = 10
            ComprimentoMinimo_DESCARGA_Hora = 5


        End Sub




        Public Shared Function Certifica(ByRef cargaDescarga As BE.CargaDescarga, ByRef mensagem As String) As Boolean
            Dim mMensagem As String

            mMensagem = ""

            With cargaDescarga

                ' sem qualquer validação do (...)EmModoEdicao porque o objeto não é persistido na base de dados (gerado em runtime)

                CCValidation.Texto("CARGA_Morada", .CARGA_Morada, CampoObrigatorio_CARGA_Morada, True, ComprimentoMinimo_CARGA_Morada, ComprimentoMaximo_CARGA_Morada, mMensagem)
                CCValidation.Texto("CARGA_Morada2", .CARGA_Morada2, CampoObrigatorio_CARGA_Morada2, True, ComprimentoMinimo_CARGA_Morada2, ComprimentoMaximo_CARGA_Morada2, mMensagem)
                CCValidation.Texto("CARGA_Localidade", .CARGA_Localidade, CampoObrigatorio_CARGA_Localidade, True, ComprimentoMinimo_CARGA_Localidade, ComprimentoMaximo_CARGA_Localidade, mMensagem)
                CCValidation.Texto("CARGA_CodigoPostal", .CARGA_CodigoPostal, CampoObrigatorio_CARGA_CodigoPostal, True, ComprimentoMinimo_CARGA_CodigoPostal, ComprimentoMaximo_CARGA_CodigoPostal, mMensagem)
                CCValidation.Texto("CARGA_LocalidadePostal", .CARGA_LocalidadePostal, CampoObrigatorio_CARGA_LocalidadePostal, True, ComprimentoMinimo_CARGA_LocalidadePostal, ComprimentoMaximo_CARGA_LocalidadePostal, mMensagem)
                CCValidation.Texto("CARGA_Distrito", .CARGA_Distrito, CampoObrigatorio_CARGA_Distrito, True, ComprimentoMinimo_CARGA_Distrito, ComprimentoMaximo_CARGA_Distrito, mMensagem)
                CCValidation.Texto("CARGA_Pais", .CARGA_Pais, CampoObrigatorio_CARGA_Pais, True, ComprimentoMinimo_CARGA_Pais, ComprimentoMaximo_CARGA_Pais, mMensagem)
                CCValidation.Texto("CARGA_Data", .CARGA_Data, CampoObrigatorio_CARGA_Data, True, ComprimentoMinimo_CARGA_Data, ComprimentoMaximo_CARGA_Data, mMensagem)
                CCValidation.Texto("CARGA_Hora", .CARGA_Hora, CampoObrigatorio_CARGA_Hora, True, ComprimentoMinimo_CARGA_Hora, ComprimentoMaximo_CARGA_Hora, mMensagem)

                CCValidation.Texto("DESCARGA_TipoEntidade", .DESCARGA_TipoEntidade, CampoObrigatorio_DESCARGA_TipoEntidade, True, ComprimentoMinimo_DESCARGA_TipoEntidade, ComprimentoMaximo_DESCARGA_TipoEntidade, mMensagem)
                CCValidation.Texto("DESCARGA_Entidade", .DESCARGA_Entidade, CampoObrigatorio_DESCARGA_Entidade, True, ComprimentoMinimo_DESCARGA_Entidade, ComprimentoMaximo_DESCARGA_Entidade, mMensagem)
                CCValidation.Texto("DESCARGA_Nome", .DESCARGA_Nome, CampoObrigatorio_DESCARGA_Nome, True, ComprimentoMinimo_DESCARGA_Nome, ComprimentoMaximo_DESCARGA_Nome, mMensagem)
                CCValidation.Texto("DESCARGA_Morada", .DESCARGA_Morada, CampoObrigatorio_DESCARGA_Morada, True, ComprimentoMinimo_DESCARGA_Morada, ComprimentoMaximo_DESCARGA_Morada, mMensagem)
                CCValidation.Texto("DESCARGA_Morada2", .DESCARGA_Morada2, CampoObrigatorio_DESCARGA_Morada2, True, ComprimentoMinimo_DESCARGA_Morada2, ComprimentoMaximo_DESCARGA_Morada2, mMensagem)
                CCValidation.Texto("DESCARGA_Localidade", .DESCARGA_Localidade, CampoObrigatorio_DESCARGA_Localidade, True, ComprimentoMinimo_DESCARGA_Localidade, ComprimentoMaximo_DESCARGA_Localidade, mMensagem)
                CCValidation.Texto("DESCARGA_CodigoPostal", .DESCARGA_CodigoPostal, CampoObrigatorio_DESCARGA_CodigoPostal, True, ComprimentoMinimo_DESCARGA_CodigoPostal, ComprimentoMaximo_DESCARGA_CodigoPostal, mMensagem)
                CCValidation.Texto("DESCARGA_LocalidadePostal", .DESCARGA_LocalidadePostal, CampoObrigatorio_DESCARGA_LocalidadePostal, True, ComprimentoMinimo_DESCARGA_LocalidadePostal, ComprimentoMaximo_DESCARGA_LocalidadePostal, mMensagem)
                CCValidation.Texto("DESCARGA_Distrito", .DESCARGA_Distrito, CampoObrigatorio_DESCARGA_Distrito, True, ComprimentoMinimo_DESCARGA_Distrito, ComprimentoMaximo_DESCARGA_Distrito, mMensagem)
                CCValidation.Texto("DESCARGA_Pais", .DESCARGA_Pais, CampoObrigatorio_DESCARGA_Pais, True, ComprimentoMinimo_DESCARGA_Pais, ComprimentoMaximo_DESCARGA_Pais, mMensagem)
                CCValidation.Texto("DESCARGA_Data", .DESCARGA_Data, CampoObrigatorio_DESCARGA_Data, True, ComprimentoMinimo_DESCARGA_Data, ComprimentoMaximo_DESCARGA_Data, mMensagem)
                CCValidation.Texto("DESCARGA_Hora", .DESCARGA_Hora, CampoObrigatorio_DESCARGA_Hora, True, ComprimentoMinimo_DESCARGA_Hora, ComprimentoMaximo_DESCARGA_Hora, mMensagem)

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

