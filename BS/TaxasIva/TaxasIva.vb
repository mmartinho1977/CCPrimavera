Imports Microsoft.VisualBasic
Imports CCPrimavera.BE
Imports Interop


Namespace BS

    Public Class TaxasIva
        Private objMotor As ErpBS900.ErpBS

        Sub New(ByRef Motor As ErpBS900.ErpBS)
            objMotor = Motor
        End Sub


        Public Function Edita(ByVal Iva As String) As TaxaIva
            Dim objIva As TaxaIva

            If Len(Trim(Iva)) = 0 Then
                Throw New Exception("O parametro Iva da funcao de Edicao da classe Iva deve conter uma string valida")
                Exit Function
            End If

            If Not Me.Existe(Iva) Then
                Throw New Exception("O registo " & Iva & " não existe na tabela de Iva, pelo que nao pode ser editado.")
                Exit Function
            End If

            Try
                objIva = GetIva(objMotor.Comercial.Iva.Edita(Iva))
            Catch ex As Exception
                Throw ex
                Exit Function
            End Try

            Edita = objIva

            objIva = Nothing
        End Function



        Private Function GetIva(ByVal Iva As GcpBE900.GcpBEIva) As TaxaIva
            Dim objIva As TaxaIva
            Dim objCampoUtilizador As BE.CampoUtilizador
            Dim i As System.Int16

            If IsNothing(Iva) Then
                Throw New Exception("O argumento Iva da Função privada GetIva não pode ser nulo.")
                Exit Function
            End If

            objIva = New TaxaIva

            With objIva
                .Iva = Iva.IVA
                .Descricao = Iva.Descricao
                .Taxa = Iva.Taxa
                .EmModoEdicao = False

                For i = 1 To Iva.CamposUtil.NumItens
                    If Not IsDBNull(Iva.CamposUtil(i)) Then
                        objCampoUtilizador = New BE.CampoUtilizador
                        With objCampoUtilizador
                            Select Case Iva.CamposUtil(i).TipoSimplificado
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
                                    Throw New Exception("Foi referenciado um tipo de dados em Iva.CamposUtil não esperado na classe Business.Iva função GetIva")
                                    Exit Function

                            End Select
                            .Campo = Iva.CamposUtil(i).Nome
                            If Not IsDBNull(Iva.CamposUtil(i).Valor) Then
                                .Valor = CType(Iva.CamposUtil(i).Valor, Object)
                            Else
                                .Valor = CType("", Object)
                            End If

                        End With
                        .CamposUtilizador.Add(objCampoUtilizador, objCampoUtilizador.Campo)
                        objCampoUtilizador = Nothing
                    End If
                Next
            End With

            GetIva = objIva

            objIva = Nothing
        End Function


        Public Function Existe(ByVal Iva As String) As Boolean
            If Iva.TrimEnd.Length = 0 Then
                Throw New Exception("O argumento string iva comunicado à função existe nao pode ser de tamanho zero.")
                Exit Function
            End If
            Return objMotor.Comercial.Iva.Existe(Iva)
        End Function


        Protected Overrides Sub Finalize()
            If Not IsNothing(objMotor) Then
                objMotor = Nothing
            End If
            MyBase.Finalize()
        End Sub

    End Class

End Namespace


