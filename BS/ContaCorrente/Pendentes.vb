Imports Microsoft.VisualBasic
Imports CCPrimavera.Specifications.Pendente_Specs
Imports CCUtils
Imports System.Data.SqlClient
Imports CCUtils.CCSQLServer
Imports Interop




Namespace BS

    Public Class Pendentes
        Private objMotor As ErpBS900.ErpBS
        Private connectionString As String


        Sub New(ByRef Motor As ErpBS900.ErpBS, ByVal connectionString As String)
            objMotor = Motor
            Me.connectionString = connectionString
        End Sub



        Public Sub Actualiza(ByRef Pendente As BE.Pendente)
            Dim strMensagem As String = ""
            Dim objPendente As GcpBE900.GcpBEPendente


            If IsNothing(Pendente) Then
                Throw New Exception("O objecto Pendente passado como parametro tem um valor null;")
                Exit Sub
            End If

            If Pendente.Linhas.Count = 0 Then
                Throw New Exception("Um documento Pendente de ContaCorrente, deve conter pelo menos uma linha;")
                objPendente = Nothing
                Exit Sub
            End If

            If Certifica(Pendente, strMensagem) = False Then
                Throw New Exception(strMensagem)
                Exit Sub
            End If

            With Pendente
                If .EmModoEdicao = True Then
                    If Me.Existe(.Filial, .Modulo, .TipoDocumento, .SerieDocumento, .NumeroDocumentoInterno) = False Then
                        Throw New Exception("O Pendente que pretende actualizar não existe ou foi removido;")
                        Exit Sub
                    End If
                End If
            End With

            Try
                objPendente = GetGcpBEPendente(Pendente)
            Catch ex As Exception
                Throw ex
                Exit Sub
            End Try

            If objPendente.EmModoEdicao = True Then
                If objMotor.Comercial.Pendentes.ValidaActualizacao(objPendente, strMensagem) = False Then
                    Throw New Exception(strMensagem)
                    Exit Sub
                End If
            End If

            Try
                Throw New Exception("ATENÇÃO: Com a migração para a V755 não foi possivel testar a mudança do metodo ActualizaEX para o ActualizaPendente. Deve ser agora testado. Célio Carvalho.")
                objMotor.Comercial.Pendentes.ActualizaPendente(objPendente)
            Catch ex As Exception
                Throw ex
                Exit Sub
            End Try

            objPendente = Nothing
        End Sub



        Public Function Edita(ByVal Filial As String, ByVal Modulo As String, ByVal TipoDocumento As String, ByVal SerieDocumento As String, ByVal NumeroDocumentoInterno As System.Int64) As BE.Pendente
            Dim objPendente As BE.Pendente

            If Me.Existe(Filial, Modulo, TipoDocumento, SerieDocumento, NumeroDocumentoInterno) = False Then
                Throw New Exception("O documento pendente " & Filial & "/" & Modulo & "/" & TipoDocumento & "/" & SerieDocumento & "/" & NumeroDocumentoInterno & " não existe na tabela de Pendentes, pelo que nao pode ser editado;")
                Exit Function
            End If

            Try
                objPendente = GetPendente(objMotor.Comercial.Pendentes.Edita(Filial, Modulo, TipoDocumento, SerieDocumento, NumeroDocumentoInterno))
            Catch ex As Exception
                Throw ex
                Exit Function
            End Try

            Edita = objPendente
            objPendente = Nothing

        End Function



        Public Function Existe(ByVal Filial As String, ByVal Modulo As String, ByVal TipoDocumento As String, ByVal SerieDocumento As String, ByVal NumeroDocumentoInterno As System.Int64) As Boolean

            If Len(Trim(Filial)) = 0 Then
                Throw New Exception("O parametro Filial da funcao de edicao da classe Pendente deve conter uma string valida;")
                Exit Function
            End If

            If Len(Trim(Modulo)) = 0 Then
                Throw New Exception("O parametro Modulo da funcao de edicao da classe Pendente deve conter uma string valida;")
                Exit Function
            End If

            If Len(Trim(TipoDocumento)) = 0 Then
                Throw New Exception("O parametro TipoDocumento da funcao de edicao da classe Pendente deve conter uma string valida;")
                Exit Function
            End If

            If NumeroDocumentoInterno = 0 Then
                Throw New Exception("O parametro NumeroDocumentoInterno da funcao de edicao da classe Pendente deve conter im inteiro diferente de zero;")
                Exit Function
            End If

            Return objMotor.Comercial.Pendentes.Existe(Filial, Modulo, TipoDocumento, SerieDocumento, NumeroDocumentoInterno)

        End Function



        Public Sub Remove(ByVal Filial As String, ByVal Modulo As String, ByVal TipoDocumento As String, ByVal SerieDocumento As String, ByVal NumeroDocumentoInterno As System.Int64)
            Dim strMensagem As String = ""

            If Me.Existe(Filial, Modulo, TipoDocumento, SerieDocumento, NumeroDocumentoInterno) = False Then
                Throw New Exception("O Pendente especificado não existe, pelo que não pode ser removido.")
                Exit Sub
            End If

            If objMotor.Comercial.Pendentes.ValidaRemocao(Filial, Modulo, TipoDocumento, SerieDocumento, NumeroDocumentoInterno, strMensagem) = False Then
                Throw New Exception(strMensagem)
                Exit Sub
            End If

            Try
                objMotor.Comercial.Pendentes.Remove(Filial, Modulo, TipoDocumento, SerieDocumento, NumeroDocumentoInterno)
            Catch ex As Exception
                Throw ex
                Exit Sub
            End Try

        End Sub



        Private Function GetPendente(ByVal Pendente As GcpBE900.GcpBEPendente) As BE.Pendente
            Dim objPendente As New BE.Pendente
            Dim objLinhaPendente As BE.LinhaPendente
            Dim objCampoUtilizador As BE.CampoUtilizador = Nothing
            Dim i As System.Int16 = 0

            With objPendente
                .IDHistorico = Pendente.IDHistorico
                .Filial = Pendente.Filial
                .Modulo = Pendente.Modulo
                .TipoDocumento = Pendente.Tipodoc
                .SerieDocumento = Pendente.Serie
                .NumeroDocumentoInterno = Pendente.NumDocInt
                .NumeroDocumento = Pendente.NumDoc
                .TipoEntidade = Pendente.TipoEntidade
                .Entidade = Pendente.Entidade
                .TipoConta = Pendente.TipoConta
                .Estado = Pendente.Estado
                .DataDocumento = Pendente.DataDoc
                .DataVencimento = Pendente.DataVenc
                .DataIntroducao = Pendente.DataIntroducao
                .CondicaoPagamento = Pendente.CondPag
                .ModoPagamento = Pendente.ModoPag
                .Moeda = Pendente.Moeda
                .ValorTotal = Pendente.ValorTotal
                .ValorPendente = Pendente.ValorPendente
                .Observacoes = Pendente.Observacoes
                .Utilizador = Pendente.Utilizador
                .EmModoEdicao = Pendente.EmModoEdicao


                For i = 1 To Pendente.Linhas.NumItens
                    objLinhaPendente = New BE.LinhaPendente
                    With objLinhaPendente
                        .Id = Pendente.Linhas(i).ID
                        .Descricao = Pendente.Linhas(i).Descricao
                        .CodigoIva = Pendente.Linhas(i).CodIva
                        .PercentagemIvaDedutivel = Pendente.Linhas(i).PercIvaDedutivel
                        .TaxaProRata = Pendente.Linhas(i).TaxaProRata
                        .ValorRecargo = Pendente.Linhas(i).ValorRecargo
                        .ValorIncidencia = Pendente.Linhas(i).Incidencia
                        .ValorIva = Pendente.Linhas(i).ValorIva
                        .ValorTotal = Pendente.Linhas(i).Total
                        .CBLLigacaoGeral = Pendente.Linhas(i).CBLContaLigacaoGeral
                        .CBLLigacaoAnalitica = Pendente.Linhas(i).CBLContaLigacaoAnalitica
                        .CBLLigacaoCentrosCusto = Pendente.Linhas(i).CBLContaLigacaoCentrosCusto
                        .CBLLigacaoFuncional = Pendente.Linhas(i).CBLContaLigacaoFuncional
                        .DataUltimaActualizacao = Pendente.Linhas(i).DataUltimaActualizacao
                    End With
                    objPendente.Linhas.Add(objLinhaPendente, objLinhaPendente.Id)
                    objLinhaPendente = Nothing
                Next

                For i = 1 To Pendente.CamposUtil.NumItens
                    If Not IsDBNull(Pendente.CamposUtil(i)) Then
                        objCampoUtilizador = New BE.CampoUtilizador
                        With objCampoUtilizador
                            Select Case Pendente.CamposUtil(i).TipoSimplificado
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
                                    Throw New Exception("Foi referenciado um tipo de dados em Pendente.CamposUtil não esperado na classe Business.Pendente função GetPendente")
                                    Exit Function

                            End Select
                            .Campo = Pendente.CamposUtil(i).Nome
                            If Not IsDBNull(Pendente.CamposUtil(i).Valor) Then
                                .Valor = CType(Pendente.CamposUtil(i).Valor, Object)
                            Else
                                .Valor = CType("", Object)
                            End If

                        End With
                        .CamposUtilizador.Add(objCampoUtilizador, objCampoUtilizador.Campo)
                        objCampoUtilizador = Nothing
                    End If
                Next

            End With

            GetPendente = objPendente
            objPendente = Nothing

        End Function



        Private Function GetGcpBEPendente(ByVal Pendente As BE.Pendente) As GcpBE900.GcpBEPendente
            Dim objPendente As New GcpBE900.GcpBEPendente
            Dim objLinhaPendente As GcpBE900.GCPBELinhaPendente
            Dim i As System.Int16 = 0

            If Pendente.EmModoEdicao = True Then 'EM EDIÇÃO
                With Pendente
                    objPendente = objMotor.Comercial.Pendentes.Edita(.Filial, .Modulo, .TipoDocumento, .SerieDocumento, .NumeroDocumentoInterno)
                End With
            End If

            With objPendente
                .Filial = Pendente.Filial
                .Modulo = Pendente.Modulo
                .Tipodoc = Pendente.TipoDocumento.TrimEnd
                .Serie = Pendente.SerieDocumento.TrimEnd
                .NumDocInt = Pendente.NumeroDocumentoInterno
                .TipoEntidade = Pendente.TipoEntidade
                .Entidade = Pendente.Entidade
                .TipoConta = Pendente.TipoConta
                .Estado = Pendente.Estado
                .DataDoc = Pendente.DataDocumento
                .ValorTotal = Pendente.ValorTotal
                .ValorPendente = Pendente.ValorPendente
            End With


            If Pendente.EmModoEdicao = False Then
                objPendente = objMotor.Comercial.Pendentes.PreencheDadosRelacionados(objPendente)
            Else
                objPendente = objMotor.Comercial.Pendentes.PreencheDadosRelacionados(objPendente, GcpBE900.PreencheRelacaoPendentes.pdDadosEntidade)
            End If


            With objPendente
                If Pendente.NumeroDocumento <> "0" Then
                    .NumDoc = Pendente.NumeroDocumento
                End If
                .Observacoes = Pendente.Observacoes
                .Utilizador = Pendente.Utilizador
                .Moeda = Pendente.Moeda
                .ModoPag = Pendente.ModoPagamento
                .CondPag = Pendente.CondicaoPagamento
                .DataDoc = Pendente.DataDocumento
                .DataIntroducao = Pendente.DataIntroducao
                .DataVenc = objMotor.Comercial.Pendentes.CalculaDataVencimento(.DataDoc, .CondPag)
                .TipoConta = Pendente.TipoConta
                .Estado = Pendente.Estado
            End With


            If Pendente.EmModoEdicao = False Then
                objPendente.IDHistorico = Pendente.IDHistorico
            Else
                objPendente.Linhas.RemoveTodos()
            End If


            With objPendente

                For i = 0 To Pendente.Linhas.Count - 1
                    objLinhaPendente = New GcpBE900.GCPBELinhaPendente
                    With objLinhaPendente
                        .ID = Pendente.Linhas(i).Id
                        .Descricao = Pendente.Linhas(i).Descricao
                        .CodIva = Pendente.Linhas(i).CodigoIva
                        .PercIvaDedutivel = Pendente.Linhas(i).PercentagemIvaDedutivel
                        .TaxaProRata = Pendente.Linhas(i).TaxaProRata
                        .ValorRecargo = Pendente.Linhas(i).ValorRecargo
                        .Incidencia = Pendente.Linhas(i).ValorIncidencia
                        .ValorIva = Pendente.Linhas(i).ValorIva()
                        .Total = Pendente.Linhas(i).ValorTotal
                        .DataUltimaActualizacao = Pendente.Linhas(i).DataUltimaActualizacao
                    End With
                    CBLClassificaLinha(objLinhaPendente)
                    .Linhas.Insere(objLinhaPendente)
                    objLinhaPendente = Nothing
                Next


                If Pendente.CamposUtilizador.Count > 0 Then
                    For i = 0 To Pendente.CamposUtilizador.Count - 1

                        Select Case Pendente.CamposUtilizador(i).Tipo
                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._String
                                .CamposUtil.Item(Pendente.CamposUtilizador(i).Campo).Valor = CType(Pendente.CamposUtilizador(i).Valor, String)

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Integer
                                .CamposUtil.Item(Pendente.CamposUtilizador(i).Campo).Valor = CType(Pendente.CamposUtilizador(i).Valor, System.Int32)

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Money
                                .CamposUtil.Item(Pendente.CamposUtilizador(i).Campo).Valor = CType(Pendente.CamposUtilizador(i).Valor, System.Decimal)

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Date
                                .CamposUtil.Item(Pendente.CamposUtilizador(i).Campo).Valor = CType(Pendente.CamposUtilizador(i).Valor, System.DateTime)

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Boolean
                                .CamposUtil.Item(Pendente.CamposUtilizador(i).Campo).Valor = CType(Pendente.CamposUtilizador(i).Valor, Boolean)

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Double
                                .CamposUtil.Item(Pendente.CamposUtilizador(i).Campo).Valor = CType(Pendente.CamposUtilizador(i).Valor, Double)

                            Case Else
                                Throw New Exception("Foi referenciado um tipo de dados em Pendente.CamposUtil não esperado na classe Business.Pendente função GetGcpBEPendente")
                                Exit Function

                        End Select

                    Next
                End If
            End With

            GetGcpBEPendente = objPendente
            objPendente = Nothing

        End Function



        Private Sub CBLClassificaLinha(ByRef objLinhaPendente As GcpBE900.GCPBELinhaPendente)
            Dim objDescritivoCC As GcpBE900.GcpBETabDescritivosCC

            If Not objMotor.Comercial.TabDescritivosCC.Existe(objLinhaPendente.Descricao.TrimEnd) Then
                Throw New Exception("A descrição -" & objLinhaPendente.Descricao.TrimEnd & "- não está definida na tabela TabDescritivosCC.")
                Exit Sub
            End If

            objDescritivoCC = objMotor.Comercial.TabDescritivosCC.Edita(objLinhaPendente.Descricao.TrimEnd)
            With objLinhaPendente
                .CBLContaLigacaoGeral = objDescritivoCC.CBLContaLigacaoGeral
                .CBLContaLigacaoAnalitica = objDescritivoCC.CBLContaLigacaoAnalitica
                .CBLContaLigacaoCentrosCusto = objDescritivoCC.CBLContaLigacaoCentrosCusto
                .CBLContaLigacaoFuncional = objDescritivoCC.CBLContaLigacaoFuncional
            End With

            objDescritivoCC = Nothing
        End Sub



        Public Function List(ByVal entidade As String) As System.Collections.Generic.List(Of BE.PendenteCCT)
            Dim query As String
            Dim mList As System.Collections.Generic.List(Of BE.PendenteCCT)
            Dim mTbl As System.Data.DataTable
            Dim objSQL As CCSQLServer
            Dim objPendenteCCT As BE.PendenteCCT
            Dim parameters As List(Of SqlParameter)
            Dim sqlParameter As System.Data.SqlClient.SqlParameter
            Dim i As System.Int16

            ' pré-validações
            If (entidade.Trim().Length = 0) Then
                Throw New Exception("O argumento entidade do método Pendentes.List() não pode ser vazio!")
            End If



            query = <sql><![CDATA[
--DECLARE @entidade AS NVARCHAR(12);
--SET @entidade = '114387';

SELECT TipoConta, 
	   Estado, 
	   DataDoc, 
	   DataVenc, 
	   TipoDoc, 
	   Serie, 
	   NumDocInt,
	   NumDoc, 
	   ValorTotal AS Total, 
	   ValorPendente AS Pendente

FROM Pendentes p
WHERE p.TipoEntidade = 'C'
	  AND p.Entidade = @entidade
ORDER BY p.DataDoc ASC
                         ]]>
                    </sql>


            parameters = New List(Of SqlParameter)

            ' @Entidade
            sqlParameter = New System.Data.SqlClient.SqlParameter
            With sqlParameter
                .DbType = Data.DbType.String
                .ParameterName = "entidade"
                .Value = entidade.TrimEnd
            End With
            parameters.Add(sqlParameter)
            sqlParameter = Nothing


            ' Inicializar tabela para albergar dados
            mTbl = New System.Data.DataTable


            Try
                ' Executar query
                objSQL = New CCSQLServer(connectionString, ModosOpenCloseConnection.auto)
                mTbl = objSQL.GetDataTable(Data.CommandType.Text, query, parameters)

            Catch ex As Exception
                objSQL = Nothing
                mTbl = Nothing
                Throw ex
                Exit Function
            End Try

            ' Preparar lista para ficar com os dados
            mList = New System.Collections.Generic.List(Of BE.PendenteCCT)

            ' Percorrer cada linha da tabela populada para gerar um objeto LinhaExtratoCCT
            For i = 0 To mTbl.Rows.Count - 1
                objPendenteCCT = New BE.PendenteCCT
                With objPendenteCCT

                    If Not IsDBNull(mTbl.Rows(i).Item("TipoConta")) Then
                        .TipoConta = CType(mTbl.Rows(i).Item("TipoConta"), System.String)
                    End If

                    If Not IsDBNull(mTbl.Rows(i).Item("Estado")) Then
                        .Estado = CType(mTbl.Rows(i).Item("Estado"), System.String)
                    End If

                    If Not IsDBNull(mTbl.Rows(i).Item("DataDoc")) Then
                        .Data = CType(mTbl.Rows(i).Item("DataDoc"), Date)
                    End If

                    If Not IsDBNull(mTbl.Rows(i).Item("DataVenc")) Then
                        .Vencimento = CType(mTbl.Rows(i).Item("DataVenc"), Date)
                    End If

                    If Not IsDBNull(mTbl.Rows(i).Item("TipoDoc")) Then
                        .TipoDocumento = CType(mTbl.Rows(i).Item("TipoDoc"), System.String)
                    End If

                    If Not IsDBNull(mTbl.Rows(i).Item("Serie")) Then
                        .Serie = CType(mTbl.Rows(i).Item("Serie"), System.String)
                    End If

                    If Not IsDBNull(mTbl.Rows(i).Item("NumDocInt")) Then
                        .Numero = CType(mTbl.Rows(i).Item("NumDocInt"), System.Int32)
                    End If

                    If Not IsDBNull(mTbl.Rows(i).Item("NumDoc")) Then
                        .NumeroExterno = CType(mTbl.Rows(i).Item("NumDoc"), String)
                    End If

                    If Not IsDBNull(mTbl.Rows(i).Item("Total")) Then
                        .Total = CType(mTbl.Rows(i).Item("Total"), Decimal)
                    End If

                    If Not IsDBNull(mTbl.Rows(i).Item("Pendente")) Then
                        .Pendente = CType(mTbl.Rows(i).Item("Pendente"), Decimal)
                    End If

                End With
                mList.Add(objPendenteCCT)
                objPendenteCCT = Nothing
            Next

            List = mList

            objSQL = Nothing
            mTbl = Nothing
            mList = Nothing
        End Function




        Protected Overrides Sub Finalize()
            If Not IsNothing(objMotor) Then
                objMotor = Nothing
            End If
            MyBase.Finalize()
        End Sub


    End Class

End Namespace
