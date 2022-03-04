Imports Microsoft.VisualBasic
Imports Interop


Namespace BS

    Public Class Moedas
        Private objMotor As ErpBS900.ErpBS


        Sub New(ByRef Motor As ErpBS900.ErpBS)
            objMotor = Motor
        End Sub


        Public Function Edita(ByVal Moeda As String) As BE.Moeda
            Dim objMoeda As BE.Moeda

            If Len(Trim(Moeda)) = 0 Then
                Throw New Exception("O parametro Moeda da funcao de Edicao da classe Moedas deve conter uma string valida")
                Exit Function
            End If

            If Not Me.Existe(Moeda) Then
                Throw New Exception("O registo " & Moeda & " não existe na tabela de Moedas, pelo que nao pode ser editado.")
                Exit Function
            End If

            Try
                objMoeda = GetMoeda(objMotor.Comercial.Moedas.Edita(Moeda))
            Catch ex As Exception
                Throw ex
                Exit Function
            End Try

            Edita = objMoeda

            objMoeda = Nothing
        End Function



        Private Function GetMoeda(ByVal Moeda As GcpBE900.GcpBEMoeda) As BE.Moeda
            Dim objMoeda As BE.Moeda
            Dim objCampoUtilizador As BE.CampoUtilizador
            Dim i As System.Int16

            If IsNothing(Moeda) Then
                Throw New Exception("O argumento Moeda da Função privada GetMoeda não pode ser nulo.")
                Exit Function
            End If

            objMoeda = New BE.Moeda

            With objMoeda
                .Moeda = Moeda.Moeda
                .Descricao = Moeda.Descricao
                .DescricaoParteInteira = Moeda.DescParteInteira
                .DescricaoParteDecimal = Moeda.DescParteDecimal
                .ArredondamentoValores = Moeda.DecArredonda
                .ArredondamentoIva = Moeda.DecArredondaIVA
                .ArredondamentoPrecosUnitarios = Moeda.DecPrecUnit
                .EmModoEdicao = False

                For i = 1 To Moeda.CamposUtil.NumItens
                    If Not IsDBNull(Moeda.CamposUtil(i)) Then
                        objCampoUtilizador = New BE.CampoUtilizador
                        With objCampoUtilizador
                            Select Case Moeda.CamposUtil(i).TipoSimplificado
                                Case StdBE900.EnumTipoCampoSimplificado.tsTexto
                                    .Tipo = BE.CampoUtilizador.TiposDadosCampoUtilizador._String

                                Case StdBE900.EnumTipoCampoSimplificado.tsInteiro
                                    .Tipo = BE.CampoUtilizador.TiposDadosCampoUtilizador._Integer

                                Case StdBE900.EnumTipoCampoSimplificado.tsMonetario
                                    .Tipo = BE.CampoUtilizador.TiposDadosCampoUtilizador._Money

                                Case StdBE900.EnumTipoCampoSimplificado.tsData
                                    .Tipo = BE.CampoUtilizador.TiposDadosCampoUtilizador._Date

                                Case StdBE900.EnumTipoCampoSimplificado.tsBooleano
                                    .Tipo = BE.CampoUtilizador.TiposDadosCampoUtilizador._Boolean

                                Case StdBE900.EnumTipoCampoSimplificado.tsDouble
                                    .Tipo = BE.CampoUtilizador.TiposDadosCampoUtilizador._Double

                                Case Else
                                    Throw New Exception("Foi referenciado um tipo de dados em Iva.CamposUtil não esperado na classe Business.Moeda função GetMoeda")
                                    Exit Function

                            End Select
                            .Campo = Moeda.CamposUtil(i).Nome
                            If Not IsDBNull(Moeda.CamposUtil(i).Valor) Then
                                .Valor = CType(Moeda.CamposUtil(i).Valor, Object)
                            Else
                                .Valor = CType("", Object)
                            End If

                        End With
                        .CamposUtilizador.Add(objCampoUtilizador, objCampoUtilizador.Campo)
                        objCampoUtilizador = Nothing
                    End If
                Next
            End With

            GetMoeda = objMoeda

            objMoeda = Nothing
        End Function



        Public Function Existe(ByVal Moeda As String) As Boolean
            If Moeda.TrimEnd.Length = 0 Then
                Throw New Exception("O argumento string Moeda comunicado à função existe nao pode ser de tamanho zero.")
                Exit Function
            End If
            Return objMotor.Comercial.Moedas.Existe(Moeda)
        End Function



        Protected Overrides Sub Finalize()
            If Not IsNothing(objMotor) Then
                objMotor = Nothing
            End If
            MyBase.Finalize()
        End Sub

    End Class

End Namespace


