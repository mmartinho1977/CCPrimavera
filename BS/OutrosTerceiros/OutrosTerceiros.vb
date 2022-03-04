Imports Microsoft.VisualBasic
Imports CCPrimavera.Others
Imports CCPrimavera.Others.Contadores
Imports CCPrimavera.Specifications.OutroTerceiro_Specs
Imports Interop



Namespace BS

    Public Class OutrosTerceiros
        Private objMotor As ErpBS900.ErpBS
        Private connectionString As String
        ''''''''''''''''''''Private objOutroTerceiro As GcpBE900.GcpBEOutroTerceiro


        Sub New(ByRef Motor As ErpBS900.ErpBS, ByVal connectionString As String)
            objMotor = Motor
            Me.connectionString = connectionString
        End Sub


        Public Sub Actualiza(ByRef OutroTerceiro As BE.OutroTerceiro)
            Dim strMensagem As String = ""
            Dim objOutroTerceiro As GcpBE900.GcpBEOutroTerceiro
            Dim objContadores As Contadores


            If IsNothing(OutroTerceiro) Then
                Throw New Exception("O objecto OutroTerceiro passado como parametro tem um valor null")
                Exit Sub
            End If

            If Certifica(OutroTerceiro, strMensagem) = False Then
                Throw New Exception(strMensagem)
                Exit Sub
            End If

            Try
                objOutroTerceiro = GetGcpBEOutroTerceiro(OutroTerceiro)
            Catch ex As Exception
                Throw ex
                Exit Sub
            End Try

            Select Case objOutroTerceiro.EmModoEdicao

                Case False 'Ou seja esta a inserir

                    strMensagem = objMotor.Comercial.OutrosTerceiros.ExisteContribuinte(objOutroTerceiro.NumContribuinte)
                    If Len(Trim(strMensagem)) > 0 Then
                        Throw New Exception("O OutroTerceiro não pode ser inserido porque o contribuinte " & objOutroTerceiro.NumContribuinte & " já existe associado ao OutroTerceiro " & strMensagem & ".")
                        Exit Sub
                    End If

                Case True 'Ou seja esta a alterar

                    If objMotor.Comercial.OutrosTerceiros.Existe(objOutroTerceiro.Terceiro) = False Then
                        Throw New Exception("O OutroTerceiro que pretende actualizar não existe ou foi removido")
                        Exit Sub
                    End If

                    If objMotor.Comercial.OutrosTerceiros.ValidaActualizacao(objOutroTerceiro, strMensagem) = False Then
                        Throw New Exception(strMensagem)
                        Exit Sub
                    End If

            End Select

            Try
                objMotor.IniciaTransaccao()

                If objOutroTerceiro.EmModoEdicao = False Then
                    ' Portanto a inserir tem que atribuir codigo antes de ser enviado contra a base de dados
                    Try
                        objContadores = New Contadores(objMotor)
                        If OutroTerceiro.TipoEntidade = "D" Then        ' D = outro devedor em primavera
                            objOutroTerceiro.Terceiro = objContadores.IncrementaDevolveString(ContadoresTipo.OutrosDevedores)
                        ElseIf OutroTerceiro.TipoEntidade = "R" Then    ' R = outro credor em primavera
                            objOutroTerceiro.Terceiro = objContadores.IncrementaDevolveString(ContadoresTipo.OutrosCredores)
                        Else
                            Throw New Exception("CC - Só estão previstos o tipo de terceiros: OutroCredor(R) e OutroDevedor(D). É favor corrigir a aplicação.")
                        End If
                        objContadores = Nothing


                        If objMotor.Comercial.OutrosTerceiros.Existe(objOutroTerceiro.Terceiro) Then
                            Throw New Exception("O registo " & objOutroTerceiro.Terceiro & " não pode ser inserido porque já existe.")
                            Exit Sub
                        End If
                        If objMotor.Comercial.OutrosTerceiros.ValidaActualizacao(objOutroTerceiro, strMensagem) = False Then
                            Throw New Exception(strMensagem)
                            Exit Sub
                        End If
                    Catch ex As Exception
                        objContadores = Nothing
                        Throw ex
                    End Try
                End If

                objMotor.Comercial.OutrosTerceiros.Actualiza(objOutroTerceiro)

                objMotor.TerminaTransaccao()

            Catch ex As Exception
                objMotor.DesfazTransaccao()
                Throw ex
                Exit Sub
            End Try

            OutroTerceiro.Codigo = objOutroTerceiro.Terceiro
            OutroTerceiro.EmModoEdicao = True

            objOutroTerceiro = Nothing
        End Sub



        Public Function Edita(ByVal OutroTerceiro As String) As BE.OutroTerceiro
            Dim objOutroTerceiro As BE.OutroTerceiro

            If Len(Trim(OutroTerceiro)) = 0 Then
                Throw New Exception("O parametro OutroTerceiro da funcao de edicao da classe OutroTerceiro deve conter uma string valida")
                Exit Function
            End If

            If objMotor.Comercial.OutrosTerceiros.Existe(OutroTerceiro) = False Then
                Throw New Exception("O registo " & OutroTerceiro & " não existe na tabela de OutrosTerceiros, pelo que nao pode ser editado.")
                Exit Function
            End If

            Try
                objOutroTerceiro = GetOutroTerceiro(objMotor.Comercial.OutrosTerceiros.Edita(OutroTerceiro))
            Catch ex As Exception
                Throw ex
                Exit Function
            End Try

            Edita = objOutroTerceiro
            objOutroTerceiro = Nothing

        End Function



        Public Function Existe(ByVal OutroTerceiro As String) As Boolean
            Return objMotor.Comercial.OutrosTerceiros.Existe(OutroTerceiro)
        End Function



        Public Sub Remove(ByVal TipoTerceiro As String, ByVal OutroTerceiro As String)
            Dim strMensagem As String = ""
            Dim objPrimavera As Motor

            If objMotor.Comercial.OutrosTerceiros.Existe(OutroTerceiro) = False Then
                Throw New Exception("O OutroTerceiro especificado não existe, pelo que não pode ser removido.")
                Exit Sub
            End If

            objPrimavera = New Motor(connectionString, objMotor.Contexto.EmpresaAberta, objMotor.Contexto.UtilizadorActual, objMotor.Contexto.PasswordUtilizadorActual)
            If objPrimavera.EntidadesAssociadas.Existe(TipoTerceiro, OutroTerceiro, strMensagem) = True Then
                Throw New Exception(strMensagem)
                Exit Sub
            End If
            objPrimavera = Nothing

            If objMotor.Comercial.OutrosTerceiros.ValidaRemocao(OutroTerceiro, strMensagem) = False Then
                Throw New Exception(strMensagem)
                Exit Sub
            End If

            Try
                objMotor.Comercial.OutrosTerceiros.Remove(OutroTerceiro)
            Catch ex As Exception
                Throw ex
                Exit Sub
            End Try

        End Sub


        Private Function GetOutroTerceiro(ByVal OutroTerceiro As GcpBE900.GcpBEOutroTerceiro) As BE.OutroTerceiro
            Dim objOutroTerceiro As New BE.OutroTerceiro(OutroTerceiro.TipoEntidade)
            Dim objCampoUtilizador As BE.CampoUtilizador = Nothing
            Dim i As System.Int16 = 0

            With objOutroTerceiro
                .Codigo = OutroTerceiro.Terceiro
                .Nome = OutroTerceiro.Nome
                .Morada = OutroTerceiro.Morada
                .Localidade = OutroTerceiro.Localidade
                .CodigoPostal = OutroTerceiro.CodigoPostal
                .LocalidadePostal = OutroTerceiro.LocalidadeCodigoPostal
                .Telefone = OutroTerceiro.Telefone
                .Fax = OutroTerceiro.Fax
                .Contribuinte = OutroTerceiro.NumContribuinte
                .DataCriacao = OutroTerceiro.DataCriacao
                .DataUltimaActualizacao = OutroTerceiro.DataUltimaActualizacao
                .ModoPagamento = OutroTerceiro.ModoPag
                .CondicaoPagamento = OutroTerceiro.CondPag
                .Moeda = OutroTerceiro.Moeda
                .PessoaSingular = OutroTerceiro.PessoaSingular
                .EmModoEdicao = OutroTerceiro.EmModoEdicao

                For i = 1 To OutroTerceiro.CamposUtil.NumItens
                    If Not IsDBNull(OutroTerceiro.CamposUtil(i)) Then
                        objCampoUtilizador = New BE.CampoUtilizador
                        With objCampoUtilizador
                            Select Case OutroTerceiro.CamposUtil(i).TipoSimplificado
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
                                    Throw New Exception("Foi referenciado um tipo de dados em OutroTerceiro.CamposUtil não esperado na classe Business.OutroTerceiro função GetOutroTerceiro")
                                    Exit Function

                            End Select
                            .Campo = OutroTerceiro.CamposUtil(i).Nome
                            If Not IsDBNull(OutroTerceiro.CamposUtil(i).Valor) Then
                                .Valor = CType(OutroTerceiro.CamposUtil(i).Valor, Object)
                            Else
                                .Valor = CType("", Object)
                            End If

                        End With
                        .CamposUtilizador.Add(objCampoUtilizador, objCampoUtilizador.Campo)
                        objCampoUtilizador = Nothing
                    End If
                Next

            End With

            GetOutroTerceiro = objOutroTerceiro
            OutroTerceiro = Nothing

        End Function


        Private Function GetGcpBEOutroTerceiro(ByVal OutroTerceiro As BE.OutroTerceiro) As GcpBE900.GcpBEOutroTerceiro
            Dim objOutroTerceiro As New GcpBE900.GcpBEOutroTerceiro
            Dim i As System.Int16 = 0

            With objOutroTerceiro
                .TipoEntidade = OutroTerceiro.TipoEntidade
                .NaturezaTerceiro = OutroTerceiro.NaturezaTerceiro
                .Terceiro = OutroTerceiro.Codigo
                .Nome = OutroTerceiro.Nome
                .Morada = OutroTerceiro.Morada
                .Localidade = OutroTerceiro.Localidade
                .CodigoPostal = OutroTerceiro.CodigoPostal
                .LocalidadeCodigoPostal = OutroTerceiro.LocalidadePostal
                .Telefone = OutroTerceiro.Telefone
                .Fax = OutroTerceiro.Fax
                .NumContribuinte = OutroTerceiro.Contribuinte
                .DataCriacao = OutroTerceiro.DataCriacao
                .DataUltimaActualizacao = OutroTerceiro.DataUltimaActualizacao
                .ModoPag = OutroTerceiro.ModoPagamento
                .CondPag = OutroTerceiro.CondicaoPagamento
                .Moeda = OutroTerceiro.Moeda
                .PessoaSingular = OutroTerceiro.PessoaSingular
                .EmModoEdicao = OutroTerceiro.EmModoEdicao

                If OutroTerceiro.CamposUtilizador.Count > 0 Then
                    For i = 0 To OutroTerceiro.CamposUtilizador.Count - 1

                        Select Case OutroTerceiro.CamposUtilizador(i).Tipo
                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._String
                                .CamposUtil.Item(OutroTerceiro.CamposUtilizador(i).Campo).Valor = CType(OutroTerceiro.CamposUtilizador(i).Valor, String)

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Integer
                                .CamposUtil.Item(OutroTerceiro.CamposUtilizador(i).Campo).Valor = CType(OutroTerceiro.CamposUtilizador(i).Valor, System.Int32)

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Money
                                .CamposUtil.Item(OutroTerceiro.CamposUtilizador(i).Campo).Valor = CType(OutroTerceiro.CamposUtilizador(i).Valor, System.Decimal)

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Date
                                .CamposUtil.Item(OutroTerceiro.CamposUtilizador(i).Campo).Valor = CType(OutroTerceiro.CamposUtilizador(i).Valor, System.DateTime)

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Boolean
                                .CamposUtil.Item(OutroTerceiro.CamposUtilizador(i).Campo).Valor = CType(OutroTerceiro.CamposUtilizador(i).Valor, Boolean)

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Double
                                .CamposUtil.Item(OutroTerceiro.CamposUtilizador(i).Campo).Valor = CType(OutroTerceiro.CamposUtilizador(i).Valor, Double)


                            Case Else
                                Throw New Exception("Foi referenciado um tipo de dados em OutroTerceiro.CamposUtil não esperado na classe Business.OutroTerceiro função GetGcpBEOutroTerceiro")
                                Exit Function

                        End Select

                    Next
                End If
            End With

            GetGcpBEOutroTerceiro = objOutroTerceiro
            objOutroTerceiro = Nothing

        End Function


        Protected Overrides Sub Finalize()
            ' '' '' '' '' '' ''If Not IsNothing(objOutroTerceiro) Then
            ' '' '' '' '' '' ''    objOutroTerceiro = Nothing
            ' '' '' '' '' '' ''End If
            If Not IsNothing(objMotor) Then
                objMotor = Nothing
            End If
            MyBase.Finalize()
        End Sub


    End Class

End Namespace
