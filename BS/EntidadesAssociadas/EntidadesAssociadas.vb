Imports Microsoft.VisualBasic
Imports CCPrimavera.Specifications.EntidadeAssociada_Specs
Imports Interop



Namespace BS

    Public Class EntidadesAssociadas
        Private objMotor As ErpBS900.ErpBS
        '''''''''''''''''Private objEntidadeAssociada As GcpBE900.GcpBEEntidadeAssociada

        Sub New(ByRef Motor As ErpBS900.ErpBS)
            objMotor = Motor
        End Sub


        Public Sub Actualiza(ByVal EntidadeAssociada As BE.EntidadeAssociada)
            Dim strMensagem As String = ""
            Dim objEntidadeAssociada As GcpBE900.GcpBEEntidadeAssociada

            If IsNothing(EntidadeAssociada) Then
                Throw New Exception("O objecto EntidadeAssociada passado como parametro tem um valor null")
                Exit Sub
            End If

            If Certifica(EntidadeAssociada, strMensagem) = False Then
                Throw New Exception(strMensagem)
                Exit Sub
            End If

            Try
                objEntidadeAssociada = GetGcpBEEntidadeAssociada(EntidadeAssociada)
            Catch ex As Exception
                Throw ex
                Exit Sub
            End Try

            With objEntidadeAssociada
                Select Case .EmModoEdicao
                    Case False 'INSERCAO
                        If objMotor.Comercial.EntidadesAssociadas.Existe(.TipoEntidade, .Entidade, .TipoEntidadeAssociada, .EntidadeAssociada) = True Then
                            Throw New Exception("A entidade associada " & .TipoEntidade & "/" & .Entidade & "/" & .TipoEntidadeAssociada & "/" & .EntidadeAssociada & " não pode ser inserida porque já existe")
                            Exit Sub
                        End If

                        If objMotor.Comercial.EntidadesAssociadas.ValidaActualizacao(objEntidadeAssociada, strMensagem) = False Then
                            Throw New Exception(strMensagem)
                            Exit Sub
                        End If

                    Case True 'ALTERACAO
                        Throw New Exception("Não foi preparado o tratamento do objecto EntidadeAssociada.EmModoEdicao=true")
                        Exit Sub

                End Select
            End With

            Try
                objMotor.Comercial.EntidadesAssociadas.Actualiza(objEntidadeAssociada)
            Catch ex As Exception
                Throw ex
            End Try

            objEntidadeAssociada = Nothing

        End Sub



        Public Function Edita(ByVal TipoEntidade As String, ByVal Entidade As String, ByVal TipoEntidadeAssociada As String, ByVal EntidadeAssociada As String) As BE.EntidadeAssociada
            Dim objEntidadeAssociada As BE.EntidadeAssociada

            If Len(Trim(EntidadeAssociada)) = 0 Then
                Throw New Exception("O parametro EntidadeAssociada da funcao de edicao da classe EntidadeAssociada deve conter uma string valida")
                Exit Function
            End If

            If objMotor.Comercial.EntidadesAssociadas.Existe(TipoEntidade, Entidade, TipoEntidadeAssociada, EntidadeAssociada) = False Then
                Throw New Exception("A entidade associada " & TipoEntidade & "/" & Entidade & "/" & TipoEntidadeAssociada & "/" & EntidadeAssociada & " não pode ser editada porque não existe.")
                Exit Function
            End If

            Try
                objEntidadeAssociada = GetEntidadeAssociada(objMotor.Comercial.EntidadesAssociadas.Edita(TipoEntidade, Entidade, TipoEntidadeAssociada, EntidadeAssociada))
            Catch ex As Exception
                Throw ex
            End Try

            Edita = objEntidadeAssociada

            objEntidadeAssociada = Nothing

        End Function



        Public Function Existe(ByVal TipoEntidade As String, ByVal Entidade As String, Optional ByRef strMensagem As String = "") As Boolean
            Dim mStdBELista As StdBE900.StdBELista
            Dim mGcpBEEntidadesAssociadas As GcpBE900.GcpBEEntidadesAssociadas
            Dim i As System.Int32

            If TipoEntidade.Trim.Length = 0 Then
                Throw New Exception("O argumento TipoEntidade não pode ser uma string vazia.")
                Exit Function
            End If

            If Entidade.Trim.Length = 0 Then
                Throw New Exception("O argumento Entidade não pode ser uma string vazia.")
                Exit Function
            End If

            If objMotor.Comercial.EntidadesAssociadas.ListaEntidadesAssociadas(TipoEntidade, Entidade).NumItens > 0 Then

                mGcpBEEntidadesAssociadas = New GcpBE900.GcpBEEntidadesAssociadas
                mGcpBEEntidadesAssociadas = objMotor.Comercial.EntidadesAssociadas.ListaEntidadesAssociadas(TipoEntidade, Entidade)

                strMensagem = "A entidade " & TipoEntidade & "/" & Entidade & " está referenciada em "
                strMensagem += mGcpBEEntidadesAssociadas.NumItens.ToString & " entidades como Entidade Associada."
                strMensagem += " " & "As entidades referenciadas são: "

                For i = 1 To mGcpBEEntidadesAssociadas.NumItens
                    strMensagem += "(" & mGcpBEEntidadesAssociadas(i).TipoEntidadeAssociada & "/" & mGcpBEEntidadesAssociadas(i).EntidadeAssociada & ")"
                    strMensagem += "; "
                Next

                mGcpBEEntidadesAssociadas = Nothing
                Return True
            End If

            mStdBELista = New StdBE900.StdBELista

            Try
                mStdBELista = objMotor.Consulta("Select TipoEntidade, Entidade From EntidadesAssociadas Where TipoEntidadeAssociada='" & TipoEntidade & "' And EntidadeAssociada='" & Entidade & "'")
            Catch ex As Exception
                Throw ex
                Exit Function
            End Try

            If mStdBELista.NumLinhas > 0 Then
                strMensagem = "A entidade " & TipoEntidade & "/" & Entidade & " está referenciada em "
                strMensagem += mStdBELista.NumLinhas.ToString & " entidades como EntidadeAssociada."
                strMensagem += " "
                strMensagem += "As entidades referenciadas são: "

                mStdBELista.Inicio()
                While (Not mStdBELista.NoFim)
                    strMensagem += "(" & mStdBELista.Valor(0).ToString & "/" & mStdBELista.Valor(1).ToString & ")"
                    strMensagem += "; "
                    mStdBELista.Seguinte()
                End While

                mStdBELista = Nothing
                Return True

            End If

            mStdBELista = Nothing

            strMensagem = "A EntidadeAssociada não existe"
            Return False

        End Function



        Public Sub Remove(ByVal TipoEntidade As String, ByVal Entidade As String, ByVal TipoEntidadeAssociada As String, ByVal EntidadeAssociada As String)
            Dim strMensagem As String = ""

            If Existe(TipoEntidade, Entidade, strMensagem) = False Then
                Throw New Exception(strMensagem)
                Exit Sub
            End If

            If objMotor.Comercial.EntidadesAssociadas.ValidaRemocao(TipoEntidade, Entidade, TipoEntidadeAssociada, EntidadeAssociada, strMensagem) = False Then
                Throw New Exception(strMensagem)
                Exit Sub
            End If

            objMotor.Comercial.EntidadesAssociadas.RemoveRelacao(TipoEntidade, Entidade, TipoEntidadeAssociada, EntidadeAssociada)

        End Sub



        Private Function GetEntidadeAssociada(ByVal EntidadeAssociada As GcpBE900.GcpBEEntidadeAssociada) As BE.EntidadeAssociada
            Dim objEntidadeAssociada As New BE.EntidadeAssociada
            Dim objCampoUtilizador As BE.CampoUtilizador = Nothing
            Dim i As System.Int16 = 0

            With objEntidadeAssociada
                .TipoEntidade = EntidadeAssociada.TipoEntidade
                .Entidade = EntidadeAssociada.Entidade
                .TipoEntidadeAssociada = EntidadeAssociada.TipoEntidadeAssociada
                .EntidadeAssociada = EntidadeAssociada.EntidadeAssociada
                .EmModoEdicao = EntidadeAssociada.EmModoEdicao
                For i = 1 To EntidadeAssociada.CamposUtil.NumItens
                    If Not IsDBNull(EntidadeAssociada.CamposUtil(i)) Then
                        objCampoUtilizador = New BE.CampoUtilizador
                        With objCampoUtilizador
                            Select Case EntidadeAssociada.CamposUtil(i).TipoSimplificado
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

                                Case Else
                                    Throw New Exception("Foi referenciado um tipo de dados em EntidadeAssociada.CamposUtil não esperado na classe business.entidadeassociada função GetEntidadeAssociada")
                                    Exit Function
                            End Select

                            .Campo = EntidadeAssociada.CamposUtil(i).Nome
                            If Not IsDBNull(EntidadeAssociada.CamposUtil(i).Valor) Then
                                .Valor = CType(EntidadeAssociada.CamposUtil(i).Valor, Object)
                            Else
                                .Valor = CType("", Object)
                            End If

                        End With
                        .CamposUtilizador.Add(objCampoUtilizador, objCampoUtilizador.Campo)
                        objCampoUtilizador = Nothing

                    End If

                Next
            End With

            GetEntidadeAssociada = objEntidadeAssociada
            objEntidadeAssociada = Nothing

        End Function


        Private Function GetGcpBEEntidadeAssociada(ByVal EntidadeAssociada As BE.EntidadeAssociada) As GcpBE900.GcpBEEntidadeAssociada
            Dim objEntidadeAssociada As New GcpBE900.GcpBEEntidadeAssociada
            Dim i As System.Int16 = 0

            With objEntidadeAssociada
                .TipoEntidade = EntidadeAssociada.TipoEntidade
                .Entidade = EntidadeAssociada.Entidade
                .TipoEntidadeAssociada = EntidadeAssociada.TipoEntidadeAssociada
                .EntidadeAssociada = EntidadeAssociada.EntidadeAssociada
                .EmModoEdicao = EntidadeAssociada.EmModoEdicao

                If EntidadeAssociada.CamposUtilizador.Count > 0 Then
                    For i = 0 To EntidadeAssociada.CamposUtilizador.Count - 1
                        Select Case EntidadeAssociada.CamposUtilizador(i).Tipo
                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._String
                                .CamposUtil(EntidadeAssociada.CamposUtilizador(i).Campo).Valor = CType(EntidadeAssociada.CamposUtilizador(i).Valor, String)

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Integer
                                .CamposUtil(EntidadeAssociada.CamposUtilizador(i).Campo).Valor = CType(EntidadeAssociada.CamposUtilizador(i).Valor, System.Int32)

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Money
                                .CamposUtil(EntidadeAssociada.CamposUtilizador(i).Campo).Valor = CType(EntidadeAssociada.CamposUtilizador(i).Valor, System.Decimal)

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Date
                                .CamposUtil(EntidadeAssociada.CamposUtilizador(i).Campo).Valor = CType(EntidadeAssociada.CamposUtilizador(i).Valor, System.DateTime)

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Boolean
                                .CamposUtil(EntidadeAssociada.CamposUtilizador(i).Campo).Valor = CType(EntidadeAssociada.CamposUtilizador(i).Valor, Boolean)

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Double
                                .CamposUtil(EntidadeAssociada.CamposUtilizador(i).Campo).Valor = CType(EntidadeAssociada.CamposUtilizador(i).Valor, Double)

                            Case Else
                                Throw New Exception("Foi referenciado um tipo de dados em EntidadeAssociada.CamposUtil não esperado na classe Business.EntidadeAssociada função GetGCPBEEntidadeAssociada")
                                Exit Function
                        End Select
                    Next
                End If
            End With

            GetGcpBEEntidadeAssociada = objEntidadeAssociada
            objEntidadeAssociada = Nothing

        End Function


        Protected Overrides Sub Finalize()
            ' '' '' '' '' '' '' ''If Not IsNothing(objEntidadeAssociada) Then
            ' '' '' '' '' '' '' ''    objEntidadeAssociada = Nothing
            ' '' '' '' '' '' '' ''End If
            If Not IsNothing(objMotor) Then
                objMotor = Nothing
            End If
            MyBase.Finalize()
        End Sub


    End Class

End Namespace

