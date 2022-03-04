Imports Microsoft.VisualBasic
Imports CCUtils


Namespace Specifications

    Public Class EntidadeAssociada_Specs

        Public Shared CampoObrigatorio_TipoEntidade As Boolean
        Public Shared CampoObrigatorio_Entidade As Boolean
        Public Shared CampoObrigatorio_TipoEntidadeAssociada As Boolean
        Public Shared CampoObrigatorio_EntidadeAssociada As Boolean

        Public Shared ComprimentoMaximo_TipoEntidade As System.Int16
        Public Shared ComprimentoMaximo_Entidade As System.Int16
        Public Shared ComprimentoMaximo_TipoEntidadeAssociada As System.Int16
        Public Shared ComprimentoMaximo_EntidadeAssociada As System.Int16

        Public Shared ComprimentoMinimo_TipoEntidade As System.Int16
        Public Shared ComprimentoMinimo_Entidade As System.Int16
        Public Shared ComprimentoMinimo_TipoEntidadeAssociada As System.Int16
        Public Shared ComprimentoMinimo_EntidadeAssociada As System.Int16


        Sub New()
            CampoObrigatorio_TipoEntidade = True
            CampoObrigatorio_Entidade = True
            CampoObrigatorio_TipoEntidadeAssociada = True
            CampoObrigatorio_EntidadeAssociada = True

            ComprimentoMaximo_TipoEntidade = 1
            ComprimentoMaximo_Entidade = 12
            ComprimentoMaximo_TipoEntidadeAssociada = 1
            ComprimentoMaximo_EntidadeAssociada = 12

            ComprimentoMinimo_TipoEntidade = 1
            ComprimentoMinimo_Entidade = 5
            ComprimentoMinimo_TipoEntidadeAssociada = 1
            ComprimentoMinimo_EntidadeAssociada = 5
        End Sub


        Public Shared Function Certifica(ByRef EntidadeAssociada As BE.EntidadeAssociada, ByRef mensagem As String) As Boolean
            Dim mMensagem As String

            mMensagem = ""

            With EntidadeAssociada
                CCValidation.Texto("TipoEntidade", .TipoEntidade, CampoObrigatorio_TipoEntidade, True, ComprimentoMinimo_TipoEntidade, ComprimentoMaximo_TipoEntidade, mMensagem)
                CCValidation.Texto("Entidade", .Entidade, CampoObrigatorio_Entidade, True, ComprimentoMinimo_Entidade, ComprimentoMaximo_Entidade, mMensagem)
                CCValidation.Texto("TipoEntidadeAssociada", .TipoEntidadeAssociada, CampoObrigatorio_TipoEntidadeAssociada, True, ComprimentoMinimo_TipoEntidadeAssociada, ComprimentoMaximo_TipoEntidadeAssociada, mMensagem)
                CCValidation.Texto("EntidadeAssociada", .EntidadeAssociada, CampoObrigatorio_EntidadeAssociada, True, ComprimentoMinimo_EntidadeAssociada, ComprimentoMaximo_EntidadeAssociada, mMensagem)

            End With

            If mMensagem.TrimEnd.Length > 0 Then
                mensagem += mMensagem
                Return False
            Else
                Return True
            End If

        End Function


    End Class

End Namespace

