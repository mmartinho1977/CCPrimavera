Imports Microsoft.VisualBasic
Imports CCUtils
Imports CCUtils.CCSQLServer
Imports System.Data.SqlClient
Imports System.Collections.Generic
Imports System.Text
Imports CCPrimavera.Others
Imports CCPrimavera.Others.Contadores
Imports CCPrimavera.Specifications.Cliente_Specs
Imports Interop



Namespace BS

    Public Class Clientes
        Private objMotor As ErpBS900.ErpBS
        Private connectionString As String


        Sub New(ByRef Motor As ErpBS900.ErpBS, ByVal connectionString As String)
            objMotor = Motor
            Me.connectionString = connectionString
        End Sub



        Public Sub Actualiza(ByRef Cliente As BE.Cliente)
            Dim strMensagem As String = ""
            Dim objCliente As GcpBE900.GcpBECliente
            Dim objContadores As Contadores

            If IsNothing(Cliente) Then
                Throw New Exception("O objecto Cliente passado como parametro tem um valor null")
                Exit Sub
            End If

            If Certifica(Cliente, strMensagem) = False Then
                Throw New Exception(strMensagem)
                Exit Sub
            End If

            Try
                objCliente = GetGcpBECliente(Cliente)
            Catch ex As Exception
                Throw ex
                Exit Sub
            End Try

            Select Case objCliente.EmModoEdicao

                Case False 'Ou seja esta a inserir
                    strMensagem = objMotor.Comercial.Clientes.ExisteContribuinte(objCliente.NumContribuinte)
                    If Len(Trim(strMensagem)) > 0 Then
                        Throw New Exception("O Cliente não pode ser inserido porque o contribuinte " & objCliente.NumContribuinte & " já existe associado ao Cliente " & strMensagem & ".")
                        Exit Sub
                    End If

                Case True 'Ou seja esta a alterar
                    If objMotor.Comercial.Clientes.Existe(objCliente.Cliente) = False Then
                        Throw New Exception("O Cliente que pretende alterar não existe ou foi removido")
                        Exit Sub
                    End If
                    If objMotor.Comercial.Clientes.ValidaActualizacao(objCliente, strMensagem) = False Then
                        Throw New Exception(strMensagem)
                        Exit Sub
                    End If

            End Select

            Try

                objMotor.IniciaTransaccao()

                If objCliente.EmModoEdicao = False Then
                    Try
                        objContadores = New Contadores(objMotor)
                        objCliente.Cliente = objContadores.IncrementaDevolveString(ContadoresTipo.Clientes)
                        objContadores = Nothing

                        If objMotor.Comercial.Clientes.Existe(objCliente.Cliente) Then
                            Throw New Exception("O registo " & objCliente.Cliente & " não pode ser inserido porque já existe.")
                        End If

                        If objMotor.Comercial.Clientes.ValidaActualizacao(objCliente, strMensagem) = False Then
                            Throw New Exception(strMensagem)
                        End If

                    Catch ex As Exception
                        objContadores = Nothing
                        Throw ex
                    End Try
                End If

                objMotor.Comercial.Clientes.Actualiza(objCliente)
                objMotor.TerminaTransaccao()

            Catch ex As Exception

                objMotor.DesfazTransaccao()

                Throw ex
                Exit Sub
            End Try

            Cliente.Codigo = objCliente.Cliente
            Cliente.EmModoEdicao = True

            objCliente = Nothing
        End Sub



        Public Function Edita(ByVal Cliente As String) As BE.Cliente
            Dim objCliente As BE.Cliente

            If Len(Trim(Cliente)) = 0 Then
                Throw New Exception("O parametro Cliente da funcao de edicao da classe cliente deve conter uma string valida")
                Exit Function
            End If

            If objMotor.Comercial.Clientes.Existe(Cliente) = False Then
                Throw New Exception("O registo " & Cliente & " não existe na tabela de clientes, pelo que nao pode ser editado.")
                Exit Function
            End If

            Try
                objCliente = GetCliente(objMotor.Comercial.Clientes.Edita(Cliente))
            Catch ex As Exception
                Throw ex
                Exit Function
            End Try

            Edita = objCliente

            objCliente = Nothing

        End Function



        Public Function List(ByVal expressoes() As String, ByVal orderByClause As String) As System.Collections.Generic.List(Of BE.Cliente)
            ' Optional ByVal texto As String = ""

            Dim mList As System.Collections.Generic.List(Of BE.Cliente)
            Dim mTbl As System.Data.DataTable
            Dim objSQL As CCSQLServer
            Dim objCliente As BE.Cliente
            Dim parameters As List(Of SqlParameter)
            Dim mSBSelect As StringBuilder
            Dim mSBWhere As StringBuilder
            Dim mSBOrder As StringBuilder
            Dim mSBSQLQuery As StringBuilder
            Dim expressao As String
            Dim i As System.Int16

            ' lista que passará vazia, porque necessitei de criar a query em runtime por causa do numero variável de filtros (parametros?)
            parameters = New List(Of SqlParameter)

            objSQL = New CCSQLServer(connectionString, ModosOpenCloseConnection.auto)
            mTbl = New System.Data.DataTable
            mSBSelect = New StringBuilder
            mSBWhere = New StringBuilder
            mSBOrder = New StringBuilder
            mSBSQLQuery = New StringBuilder

            Try
                With mSBSelect
                    .Append("SELECT Cliente AS Codigo, Nome, Fac_Mor AS Morada, Fac_Mor2 AS Morada2, Fac_Local AS Localidade, Fac_Cp AS CodigoPostal, Fac_Cploc AS LocalidadePostal,")
                    .Append(" Fac_Tel AS Telefone, Telefone2, Fac_Fax AS Fax, EnderecoWeb, NumContrib AS Contribuinte, Distrito, Pais, PessoaSingular, TipoPrec AS TabelaPrecos,")
                    .Append(" CondPag AS CondicaoPagamento, ModoExp AS ModoExpedicao, ModoPag AS ModoPagamento, Moeda, DataCriacao, DataUltimaActualizacao AS DataUltimaActualizacao")
                    .Append(" FROM Clientes WHERE ClienteAnulado = 0")
                End With


                For index = 0 To expressoes.Length - 1
                    If expressoes(index).Trim.Length > 0 Then
                        expressao = CCStrings.CleanDangerousText(expressoes(index))
                        mSBWhere.Append(String.Format(" AND (Cliente COLLATE LATIN1_GENERAL_CI_AI LIKE '%{0}%'", expressao))
                        mSBWhere.Append(String.Format(" OR Nome COLLATE LATIN1_GENERAL_CI_AI LIKE '%{0}%'", expressao))
                        mSBWhere.Append(String.Format(" OR NumContrib COLLATE LATIN1_GENERAL_CI_AI LIKE '%{0}%'", expressao))
                        mSBWhere.Append(")")
                    End If
                Next

                If orderByClause.TrimEnd().Length = 0 Then
                    With mSBOrder
                        .Append(" ORDER BY Nome ASC")
                    End With
                Else
                    With mSBOrder
                        .Append(String.Format(" ORDER BY {0}", orderByClause.TrimEnd()))
                    End With
                End If
                

                With mSBSQLQuery
                    .Append(mSBSelect.ToString.TrimEnd)
                    If mSBWhere.ToString.Trim.Length > 0 Then
                        .Append(mSBWhere.ToString.TrimEnd)
                    End If
                    .Append(mSBOrder.ToString.TrimEnd)
                End With

                mSBSelect = Nothing
                mSBWhere = Nothing
                mSBOrder = Nothing

                mTbl = objSQL.GetDataTable(Data.CommandType.Text, mSBSQLQuery.ToString.TrimEnd, parameters)

            Catch ex As Exception
                objSQL = Nothing
                mSBSQLQuery = Nothing
                mTbl = Nothing
                Throw ex
                Exit Function
            End Try

            ' inicializar a lista
            mList = New System.Collections.Generic.List(Of BE.Cliente)

            ' sair com a lista com zero elementos
            If mTbl.Rows.Count = 0 Then
                objSQL = Nothing
                mSBSQLQuery = Nothing
                mTbl = Nothing
                List = mList    ' devolve a lista que, no minimo, terá zero elementos.
                Exit Function
            End If

            ' converter as linhas do datatable numa lista de objetos
            For i = 0 To mTbl.Rows.Count - 1
                objCliente = New BE.Cliente
                With objCliente

                    If Not IsDBNull(mTbl.Rows(i).Item("Codigo")) Then
                        .Codigo = CType(mTbl.Rows(i).Item("Codigo"), System.String)
                    End If
                    If Not IsDBNull(mTbl.Rows(i).Item("Nome")) Then
                        .Nome = CType(mTbl.Rows(i).Item("Nome"), System.String)
                    End If
                    If Not IsDBNull(mTbl.Rows(i).Item("Morada")) Then
                        .Morada = CType(mTbl.Rows(i).Item("Morada"), System.String)
                    End If
                    If Not IsDBNull(mTbl.Rows(i).Item("Morada2")) Then
                        .Morada2 = CType(mTbl.Rows(i).Item("Morada2"), System.String)
                    End If
                    If Not IsDBNull(mTbl.Rows(i).Item("Localidade")) Then
                        .Localidade = CType(mTbl.Rows(i).Item("Localidade"), System.String)
                    End If
                    If Not IsDBNull(mTbl.Rows(i).Item("CodigoPostal")) Then
                        .CodigoPostal = CType(mTbl.Rows(i).Item("CodigoPostal"), System.String)
                    End If
                    If Not IsDBNull(mTbl.Rows(i).Item("LocalidadePostal")) Then
                        .LocalidadePostal = CType(mTbl.Rows(i).Item("LocalidadePostal"), System.String)
                    End If
                    If Not IsDBNull(mTbl.Rows(i).Item("Telefone")) Then
                        .Telefone = CType(mTbl.Rows(i).Item("Telefone"), System.String)
                    End If
                    If Not IsDBNull(mTbl.Rows(i).Item("Telefone2")) Then
                        .Telefone2 = CType(mTbl.Rows(i).Item("Telefone2"), System.String)
                    End If
                    If Not IsDBNull(mTbl.Rows(i).Item("Fax")) Then
                        .Fax = CType(mTbl.Rows(i).Item("Fax"), System.String)
                    End If
                    If Not IsDBNull(mTbl.Rows(i).Item("EnderecoWeb")) Then
                        .EnderecoWeb = CType(mTbl.Rows(i).Item("EnderecoWeb"), System.String)
                    End If
                    If Not IsDBNull(mTbl.Rows(i).Item("Contribuinte")) Then
                        .Contribuinte = CType(mTbl.Rows(i).Item("Contribuinte"), System.String)
                    End If
                    If Not IsDBNull(mTbl.Rows(i).Item("Distrito")) Then
                        .Distrito = CType(mTbl.Rows(i).Item("Distrito"), System.String)
                    End If
                    If Not IsDBNull(mTbl.Rows(i).Item("Pais")) Then
                        .Pais = CType(mTbl.Rows(i).Item("Pais"), System.String)
                    End If
                    If Not IsDBNull(mTbl.Rows(i).Item("PessoaSingular")) Then
                        .PessoaSingular = CType(mTbl.Rows(i).Item("PessoaSingular"), Boolean)
                    End If
                    If Not IsDBNull(mTbl.Rows(i).Item("TabelaPrecos")) Then
                        .TabelaPrecos = CType(mTbl.Rows(i).Item("TabelaPrecos"), System.String)
                    End If
                    If Not IsDBNull(mTbl.Rows(i).Item("CondicaoPagamento")) Then
                        .CondicaoPagamento = CType(mTbl.Rows(i).Item("CondicaoPagamento"), System.String)
                    End If
                    If Not IsDBNull(mTbl.Rows(i).Item("ModoExpedicao")) Then
                        .ModoExpedicao = CType(mTbl.Rows(i).Item("ModoExpedicao"), System.String)
                    End If
                    If Not IsDBNull(mTbl.Rows(i).Item("ModoPagamento")) Then
                        .ModoPagamento = CType(mTbl.Rows(i).Item("ModoPagamento"), System.String)
                    End If
                    If Not IsDBNull(mTbl.Rows(i).Item("Moeda")) Then
                        .Moeda = CType(mTbl.Rows(i).Item("Moeda"), System.String)
                    End If
                    If Not IsDBNull(mTbl.Rows(i).Item("DataCriacao")) Then
                        .DataCriacao = CType(mTbl.Rows(i).Item("DataCriacao"), Date)
                    End If
                    If Not IsDBNull(mTbl.Rows(i).Item("DataUltimaActualizacao")) Then
                        .DataUltimaActualizacao = CType(mTbl.Rows(i).Item("DataUltimaActualizacao"), Date)
                    End If
                    .EmModoEdicao = False
                End With
                mList.Add(objCliente)
                objCliente = Nothing
            Next

            List = mList

            objSQL = Nothing
            mSBSQLQuery = Nothing
            mTbl = Nothing
            mList = Nothing
        End Function



        Public Function Existe(ByVal Cliente As String) As Boolean
            Return objMotor.Comercial.Clientes.Existe(Cliente)
        End Function



        Public Sub Remove(ByVal Cliente As String)
            Dim strMensagem As String = ""
            Dim objPrimavera As Motor

            If objMotor.Comercial.Clientes.Existe(Cliente) = False Then
                Throw New Exception("A entidade especificada não existe, pelo que não pode ser removida.")
                Exit Sub
            End If

            objPrimavera = New Motor(connectionString, objMotor.Contexto.EmpresaAberta, objMotor.Contexto.UtilizadorActual, objMotor.Contexto.PasswordUtilizadorActual)

            If objPrimavera.EntidadesAssociadas.Existe(TiposEntidade.Cliente, Cliente, strMensagem) = True Then
                Throw New Exception(strMensagem)
                Exit Sub
            End If
            objPrimavera = Nothing


            If objMotor.Comercial.Clientes.ValidaRemocao(Cliente, strMensagem) = False Then
                Throw New Exception(strMensagem)
                Exit Sub
            End If

            Try
                objMotor.Comercial.Clientes.Remove(Cliente)
            Catch ex As Exception
                Throw ex
                Exit Sub
            End Try

        End Sub



        Private Function GetCliente(ByVal Cliente As GcpBE900.GcpBECliente) As BE.Cliente
            Dim objCliente As New BE.Cliente
            Dim objCampoUtilizador As BE.CampoUtilizador = Nothing
            Dim i As System.Int16 = 0

            With objCliente
                .Codigo = Cliente.Cliente
                .Nome = Cliente.Nome
                .Morada = Cliente.Morada
                .Morada2 = Cliente.Morada2
                .Localidade = Cliente.Localidade
                .CodigoPostal = Cliente.CodigoPostal
                .LocalidadePostal = Cliente.LocalidadeCodigoPostal
                .Telefone = Cliente.Telefone
                .Telefone2 = Cliente.Telefone2
                .Fax = Cliente.Fax
                .Contribuinte = Cliente.NumContribuinte
                .Distrito = Cliente.Distrito
                .Pais = Cliente.Pais
                .EnderecoWeb = Cliente.EnderecoWeb
                .DataCriacao = Cliente.DataCriacao
                .DataUltimaActualizacao = Cliente.DataUltimaActualizacao
                .ModoPagamento = Cliente.ModoPag
                .ModoExpedicao = Cliente.ModoExp
                .CondicaoPagamento = Cliente.CondPag
                .Moeda = Cliente.Moeda
                Select Case Cliente.LinhaPrecos
                    Case "0"
                        .TabelaPrecos = "PVP1"
                    Case "1"
                        .TabelaPrecos = "PVP2"
                    Case "2"
                        .TabelaPrecos = "PVP3"
                    Case "3"
                        .TabelaPrecos = "PVP4"
                    Case "4"
                        .TabelaPrecos = "PVP5"
                    Case "5"
                        .TabelaPrecos = "PVP6"
                End Select
                .PessoaSingular = Cliente.PessoaSingular
                .EmModoEdicao = Cliente.EmModoEdicao

                For i = 1 To Cliente.CamposUtil.NumItens
                    If Not IsDBNull(Cliente.CamposUtil(i)) Then
                        objCampoUtilizador = New BE.CampoUtilizador
                        With objCampoUtilizador
                            Select Case Cliente.CamposUtil(i).TipoSimplificado
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
                                    Throw New Exception("Foi referenciado um tipo de dados em Cliente.CamposUtil não esperado na classe Business.Cliente função GetCliente")
                                    Exit Function

                            End Select
                            .Campo = Cliente.CamposUtil(i).Nome
                            If Not IsDBNull(Cliente.CamposUtil(i).Valor) Then
                                .Valor = CType(Cliente.CamposUtil(i).Valor, Object)
                            Else
                                .Valor = CType("", Object)
                            End If

                        End With
                        .CamposUtilizador.Add(objCampoUtilizador, objCampoUtilizador.Campo)
                        objCampoUtilizador = Nothing
                    End If
                Next

            End With

            GetCliente = objCliente
            objCliente = Nothing

        End Function



        Private Function GetGcpBECliente(ByVal Cliente As BE.Cliente) As GcpBE900.GcpBECliente
            Dim objCliente As New GcpBE900.GcpBECliente
            Dim i As System.Int16 = 0

            With objCliente
                .Cliente = Cliente.Codigo
                .Nome = Cliente.Nome
                .Morada = Cliente.Morada
                .Morada2 = Cliente.Morada2
                .Localidade = Cliente.Localidade
                .CodigoPostal = Cliente.CodigoPostal
                .LocalidadeCodigoPostal = Cliente.LocalidadePostal
                .Telefone = Cliente.Telefone
                .Telefone2 = Cliente.Telefone2
                .Fax = Cliente.Fax
                .NumContribuinte = Cliente.Contribuinte
                .Distrito = Cliente.Distrito
                .Pais = Cliente.Pais
                .EnderecoWeb = Cliente.EnderecoWeb
                .DataCriacao = Cliente.DataCriacao
                .DataUltimaActualizacao = Cliente.DataUltimaActualizacao
                .ModoPag = Cliente.ModoPagamento
                .ModoExp = Cliente.ModoExpedicao
                .CondPag = Cliente.CondicaoPagamento
                .Moeda = Cliente.Moeda
                Select Case Cliente.TabelaPrecos
                    Case "PVP1"
                        .LinhaPrecos = "0"
                    Case "PVP2"
                        .LinhaPrecos = "1"
                    Case "PVP3"
                        .LinhaPrecos = "2"
                    Case "PVP4"
                        .LinhaPrecos = "3"
                    Case "PVP5"
                        .LinhaPrecos = "4"
                    Case "PVP6"
                        .LinhaPrecos = "5"
                End Select
                .PessoaSingular = Cliente.PessoaSingular
                .EmModoEdicao = Cliente.EmModoEdicao
                If Cliente.CamposUtilizador.Count > 0 Then
                    For i = 0 To Cliente.CamposUtilizador.Count - 1

                        Select Case Cliente.CamposUtilizador(i).Tipo
                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._String
                                .CamposUtil.Item(Cliente.CamposUtilizador(i).Campo).Valor = CType(Cliente.CamposUtilizador(i).Valor, String)

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Integer
                                .CamposUtil.Item(Cliente.CamposUtilizador(i).Campo).Valor = CType(Cliente.CamposUtilizador(i).Valor, System.Int32)

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Money
                                .CamposUtil.Item(Cliente.CamposUtilizador(i).Campo).Valor = CType(Cliente.CamposUtilizador(i).Valor, System.Decimal)

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Date
                                .CamposUtil.Item(Cliente.CamposUtilizador(i).Campo).Valor = CType(Cliente.CamposUtilizador(i).Valor, System.DateTime)

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Boolean
                                .CamposUtil.Item(Cliente.CamposUtilizador(i).Campo).Valor = CType(Cliente.CamposUtilizador(i).Valor, Boolean)

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Boolean
                                .CamposUtil.Item(Cliente.CamposUtilizador(i).Campo).Valor = CType(Cliente.CamposUtilizador(i).Valor, Double)


                            Case Else
                                Throw New Exception("Foi referenciado um tipo de dados em Cliente.CamposUtil não esperado na classe Business.Cliente função GetGcpBECliente")
                                Exit Function

                        End Select

                    Next
                End If
            End With

            GetGcpBECliente = objCliente
            objCliente = Nothing

        End Function



        Protected Overrides Sub Finalize()
            If Not IsNothing(objMotor) Then
                objMotor = Nothing
            End If
            MyBase.Finalize()
        End Sub


    End Class

End Namespace


