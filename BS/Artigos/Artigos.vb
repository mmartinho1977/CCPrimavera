Imports Microsoft.VisualBasic
Imports Interop


Namespace BS

    Public Class Artigos
        Private objMotor As ErpBS900.ErpBS

        Sub New(ByRef Motor As ErpBS900.ErpBS)
            objMotor = Motor
        End Sub


        Public Function EditaParaVendas(ByVal Artigo As String) As BE.Artigo
            Dim objArtigo As BE.Artigo

            If Len(Trim(Artigo)) = 0 Then
                Throw New Exception("O parametro Artigo da funcao de EdicaoParaVendas da classe Artigo deve conter uma string valida")
                Exit Function
            End If

            If Not Me.Existe(Artigo) Then
                Throw New Exception("O registo " & Artigo & " não existe na tabela de artigos, pelo que nao pode ser editado.")
                Exit Function
            End If

            Try
                objArtigo = GetArtigo(objMotor.Comercial.Artigos.EditaParaVendas(Artigo, "EUR"))
            Catch ex As Exception
                Throw ex
                Exit Function
            End Try

            EditaParaVendas = objArtigo

            objArtigo = Nothing
        End Function


        Private Function GetArtigo(ByVal Artigo As GcpBE900.GcpBEArtigo) As BE.Artigo
            Dim objArtigo As BE.Artigo
            Dim objCampoUtilizador As BE.CampoUtilizador
            Dim i As System.Int16

            If IsNothing(Artigo) Then
                Throw New Exception("O argumento Artigo da Função privada GetArtigo não pode ser nulo.")
                Exit Function
            End If

            objArtigo = New BE.Artigo

            With objArtigo
                .Artigo = Artigo.Artigo
                .TipoArtigo = Artigo.TipoArtigo
                .CodigoBarras = Artigo.CodBarras
                .Descricao = Artigo.Descricao
                .DescricaoComercial = Artigo.DescricaoComercial
                .Caracteristicas = Artigo.Caracteristicas
                .Iva = Artigo.IVA
                .ArmazemSugestao = Artigo.ArmazemSugestao
                .LocalizacaoSugestao = Artigo.LocalizacaoSugestao
                .Familia = Artigo.Familia
                .SubFamilia = Artigo.SubFamilia
                .Marca = Artigo.Marca
                .Modelo = Artigo.Modelo
                .Garantia = Artigo.Garantia
                .FornecedorPrincipal = Artigo.FornecedorPrincipal
                .DataUltimaEntrada = Artigo.DataUltimaEntrada
                .DataUltimaSaida = Artigo.DataUltimaSaida
                .PCM = Artigo.PCMedio
                .UPC = Artigo.PCUltimo
                .StockActual = Artigo.StkActual
                .UnidadeBase = Artigo.UnidadeBase
                .UnidadeCompra = Artigo.UnidadeCompra
                .UnidadeVenda = Artigo.UnidadeVenda
                .UnidadeEntrada = Artigo.UnidadeEntrada
                .UnidadeSaida = Artigo.UnidadeSaida
                If Artigo.MovStock = "S" Then
                    .MovimentaStock = True
                Else
                    .MovimentaStock = False
                End If
                .Observacoes = Artigo.Observacoes
                .DataUltimaActualizacao = Artigo.DataUltimaActualizacao
                .PercentagemIncidenciaIva = Artigo.PercIncidenciaIVA
                .PercentagemIvaDedutivel = Artigo.PercIvaDedutivel
                .EmModoEdicao = False

                For i = 1 To Artigo.CamposUtil.NumItens
                    If Not IsDBNull(Artigo.CamposUtil(i)) Then
                        objCampoUtilizador = New BE.CampoUtilizador
                        With objCampoUtilizador
                            Select Case Artigo.CamposUtil(i).TipoSimplificado
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
                                    Throw New Exception("Foi referenciado um tipo de dados em Artigo.CamposUtil não esperado na classe Business.Artigo função GetArtigo")
                                    Exit Function

                            End Select
                            .Campo = Artigo.CamposUtil(i).Nome
                            If Not IsDBNull(Artigo.CamposUtil(i).Valor) Then
                                .Valor = CType(Artigo.CamposUtil(i).Valor, Object)
                            Else
                                .Valor = CType("", Object)
                            End If

                        End With
                        .CamposUtilizador.Add(objCampoUtilizador, objCampoUtilizador.Campo)
                        objCampoUtilizador = Nothing
                    End If
                Next
            End With

            GetArtigo = objArtigo

            objArtigo = Nothing
        End Function


        Public Function Existe(ByVal Artigo As String) As Boolean
            If Artigo.TrimEnd.Length = 0 Then
                Throw New Exception("O argumento string artigo comunicado à função existe nao pode ser de tamanho zero.")
                Exit Function
            End If
            Return objMotor.Comercial.Artigos.Existe(Artigo)
        End Function


        Public Function ExisteCodigoBarras(ByVal codigoBarras As String) As Boolean
            Dim codigoArtigo As String

            If codigoBarras.TrimEnd.Length = 0 Then
                Throw New Exception("O argumento string codigoBarras comunicado à função existe nao pode ser de tamanho zero.")
                Exit Function
            End If

            ' VERIFICAR SE O TEXTO RECEBIDO É UM CODIGO DE BARRAS
            codigoArtigo = objMotor.Comercial.Artigos.DaArtigoComCodBarras(codigoBarras)

            ' SE NÃO FOI RETORNADO NENHUM CÓDIGO DE ARTIGO, ENTÃO O VALOR RECEBIDO NÃO É UM CODIGO DE BARRAS VÁLIDO.
            If IsNothing(codigoArtigo) Then
                Return False
            Else
                Return True
            End If

        End Function


        Public Function GetArtigoComCodigoBarras(ByVal CodigoBarras As String) As String
            If Not objMotor.Comercial.Artigos.ExisteCodBarras(CodigoBarras) Then
                Throw New Exception("O Codigo de barras especificado não existe")
                Exit Function
            End If

            Return objMotor.Comercial.Artigos.DaArtigoComCodBarras(CodigoBarras)
        End Function


        Public Function GetArtigos(ByVal moeda As String, ByVal sqlOrdenacao As String, Optional ByVal codigoArtigoFiltro As String = "", Optional ByVal descricaoFiltro As String = "", Optional ByVal codigoCliente As String = "") As List(Of BE.Artigo)
            Dim lista As StdBE900.StdBELista
            Dim artigos As New List(Of BE.Artigo)
            Dim artigo = New BE.Artigo
            Dim sqlQuery As String


            If codigoCliente.Trim().Length > 0 Then
                sqlQuery = "DECLARE @precoCliente nvarchar(1) "
                sqlQuery += String.Format("Select Top 1 @precoCliente = tipoPrec From Clientes where Cliente = '{0}' ", codigoCliente.TrimEnd())
            Else
                sqlQuery = "DECLARE @precoCliente nvarchar(1) "
                sqlQuery += String.Format("Set @precoCliente='0' ", codigoCliente.TrimEnd())
            End If

            sqlQuery += "SELECT TOP(200) artigo.Artigo, artigo.Descricao, artigo.CodBarras, "
            sqlQuery += "artigo.UnidadeBase, artigo.UnidadeVenda, artigo.UnidadeCompra, artigo.UnidadeEntrada, artigo.UnidadeSaida, "
            sqlQuery += "artigo.Iva, Iva.Taxa, artigo.PCMedio, artigo.PCUltimo, artigo.STKActual, "
            sqlQuery += "artigoMoeda.PVP1, artigoMoeda.PVP2, artigoMoeda.PVP3, artigoMoeda.PVP4, artigoMoeda.PVP5, artigoMoeda.PVP6, "
            sqlQuery += "artigoMoeda.PVP1IvaIncluido, artigoMoeda.PVP2IvaIncluido, artigoMoeda.PVP3IvaIncluido, artigoMoeda.PVP4IvaIncluido, "
            sqlQuery += "artigoMoeda.PVP5IvaIncluido, artigoMoeda.PVP6IvaIncluido, "
            sqlQuery += "Case @precoCliente "
            sqlQuery += "When '0' then artigoMoeda.PVP1 "
            sqlQuery += "When '1' then artigoMoeda.PVP2 "
            sqlQuery += "When '2' then artigoMoeda.PVP3 "
            sqlQuery += "When '3' then artigoMoeda.PVP4 "
            sqlQuery += "When '4' then artigoMoeda.PVP5 "
            sqlQuery += "When '5' then artigoMoeda.PVP6 "
            sqlQuery += "Else artigoMoeda.PVP1 end as precoCliente "
            sqlQuery += "FROM Artigo artigo "
            sqlQuery += "INNER JOIN Iva iva ON artigo.Iva = iva.Iva "
            sqlQuery += "LEFT JOIN ArtigoMoeda artigoMoeda ON artigo.Artigo = artigoMoeda.Artigo "
            sqlQuery += String.Format("WHERE artigoMoeda.Moeda='{0}' and artigo.ArtigoAnulado=0", moeda.TrimEnd().ToUpper())

            ' SE FOR PEDIDO O FILTRO POR ARTIGO
            If (codigoArtigoFiltro.Trim().Length > 0) Then
                sqlQuery += String.Format("and artigo.Artigo LIKE '%{0}%' ", codigoArtigoFiltro.TrimEnd())
            End If

            ' SE FOR PEDIDO O FILTRO PELA DESCRICAO
            If (descricaoFiltro.Trim().Length > 0) Then
                sqlQuery += String.Format("and artigo.Descricao LIKE '%{0}%' ", descricaoFiltro.TrimEnd())
            End If

            ' ORDENACAO
            sqlQuery += String.Format("ORDER BY {0} ", sqlOrdenacao)


            ' EXECUTAR PESQUISA
            lista = objMotor.Consulta(sqlQuery)

            lista.Inicio()

            While Not lista.NoFim
                artigo = Nothing
                artigo = New BE.Artigo

                artigo.Artigo = lista.Valor("Artigo").ToString()
                artigo.Descricao = lista.Valor("Descricao").ToString()
                artigo.CodigoBarras = lista.Valor("CodBarras").ToString()
                artigo.UnidadeBase = lista.Valor("UnidadeBase").ToString()
                artigo.UnidadeVenda = lista.Valor("UnidadeVenda").ToString()
                artigo.UnidadeCompra = lista.Valor("UnidadeCompra").ToString()
                artigo.UnidadeEntrada = lista.Valor("UnidadeEntrada").ToString()
                artigo.UnidadeSaida = lista.Valor("UnidadeSaida").ToString()
                artigo.Iva = lista.Valor("Iva").ToString()
                artigo.PCM = CType(lista.Valor("PCMedio"), Double)
                artigo.UPC = CType(lista.Valor("PCUltimo"), Double)
                artigo.PCM = CType(lista.Valor("PCMedio"), Double)
                artigo.StockActual = CType(lista.Valor("STKActual"), Double)

                'Meta Informacao
                artigo.MetaInfo = New Dictionary(Of String, Object)
                artigo.MetaInfo.Add("Iva.Taxa", lista.Valor("Taxa"))
                artigo.MetaInfo.Add("ArtigoMoeda.PVP1", lista.Valor("PVP1"))
                artigo.MetaInfo.Add("ArtigoMoeda.PVP2", lista.Valor("PVP2"))
                artigo.MetaInfo.Add("ArtigoMoeda.PVP3", lista.Valor("PVP3"))
                artigo.MetaInfo.Add("ArtigoMoeda.PVP4", lista.Valor("PVP4"))
                artigo.MetaInfo.Add("ArtigoMoeda.PVP5", lista.Valor("PVP5"))
                artigo.MetaInfo.Add("ArtigoMoeda.PVP6", lista.Valor("PVP6"))
                artigo.MetaInfo.Add("ArtigoMoeda.PVP1IvaIncluido", lista.Valor("PVP1IvaIncluido"))
                artigo.MetaInfo.Add("ArtigoMoeda.PVP2IvaIncluido", lista.Valor("PVP2IvaIncluido"))
                artigo.MetaInfo.Add("ArtigoMoeda.PVP3IvaIncluido", lista.Valor("PVP3IvaIncluido"))
                artigo.MetaInfo.Add("ArtigoMoeda.PVP4IvaIncluido", lista.Valor("PVP4IvaIncluido"))
                artigo.MetaInfo.Add("ArtigoMoeda.PVP5IvaIncluido", lista.Valor("PVP5IvaIncluido"))
                artigo.MetaInfo.Add("ArtigoMoeda.PVP6IvaIncluido", lista.Valor("PVP6IvaIncluido"))
                artigo.MetaInfo.Add("ArtigoMoeda.PrecoCliente", lista.Valor("precoCliente"))
                artigos.Add(artigo)
                lista.Seguinte()
            End While


            Return artigos

        End Function


        Public Function Edita(ByVal codigoArtigoCodigoBarras As String) As BE.Artigo
            Dim artigo As BE.Artigo
            Dim codigoArtigo As String
            Dim objArtigosMoedas As BS.ArtigosPrecos


            If codigoArtigoCodigoBarras.Trim().Length = 0 Then
                Throw New Exception("o codigo recebido tem que ter cumprimento superior a zero caracteres!")
            End If


            ' SE NÃO EXISTE ESTE CODIGO COMO CODIGO DE ARTIGO ENTÃO TENTA VER SE EXISTE ESSE CODIGO COMO CODIGO DE BARRAS
            If Not Existe(codigoArtigoCodigoBarras) Then
                codigoArtigo = objMotor.Comercial.Artigos.DaArtigoComCodBarras(codigoArtigoCodigoBarras)

                ' SE NÃO FOI RETORNADO NENHUM CÓDIGO DE ARTIGO, ISTO NÃO É UM CODIGO DE BARRAS VÁLIDO.
                If IsNothing(codigoArtigo) Then
                    Throw New Exception("o codigo recebido não corresponde a um código de artigo nem a um código de barras!")
                End If
            Else
                codigoArtigo = codigoArtigoCodigoBarras
            End If


            ' ABRE A FICHA DE ARTIGO (aqui tenho a certeza que codigoArtigo corresponde efetivamente a um codigo de artigo existente)
            artigo = GetArtigo(objMotor.Comercial.Artigos.Edita(codigoArtigo))

            ' PREENCHER O ARTIGOMOEDA
            objArtigosMoedas = New BS.ArtigosPrecos(objMotor)

            If objArtigosMoedas.Existe(codigoArtigo, "EUR", artigo.UnidadeVenda) Then
                artigo.ArtigoMoeda = objArtigosMoedas.Edita(codigoArtigo, "EUR", artigo.UnidadeVenda)
            End If

            Return artigo

        End Function



        Protected Overrides Sub Finalize()
            If Not IsNothing(objMotor) Then
                objMotor = Nothing
            End If
            MyBase.Finalize()
        End Sub

    End Class

End Namespace


