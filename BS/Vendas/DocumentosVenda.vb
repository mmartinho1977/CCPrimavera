Imports Microsoft.VisualBasic
Imports CCPrimavera.Specifications
Imports CCPrimavera.Specifications.DocumentoVenda_Specs
Imports Interop



Namespace BS

    Public Class DocumentosVenda
        Private objMotor As ErpBS900.ErpBS
        Private objCargasDescargas As BS.CargasDescargas



        Sub New(ByRef Motor As ErpBS900.ErpBS)
            objMotor = Motor
            objCargasDescargas = New BS.CargasDescargas(objMotor)
        End Sub



        Public Sub PreencheDadosRelacionados_Todos(ByRef DocumentoVenda As BE.DocumentoVenda)
            Dim objDocumentoVenda As GcpBE900.GcpBEDocumentoVenda

            If IsNothing(DocumentoVenda) Then
                Throw New Exception("O objecto DocumentoVenda passado como parametro na função PreencheDadosRelacionados tem um valor null;")
                Exit Sub
            End If

            objDocumentoVenda = Me.GetGcpBEDocumentoVenda(DocumentoVenda)

            Try
                ' preencher dados relacionados (preenche também os dados do objeto CargaDescarga)
                objMotor.Comercial.Vendas.PreencheDadosRelacionados(objDocumentoVenda, GcpBE900.PreencheRelacaoVendas.vdDadosTodos)

            Catch ex As Exception
                objDocumentoVenda = Nothing
                Throw ex
                Exit Sub
            End Try

            DocumentoVenda = GetDocumentoVenda(objDocumentoVenda)

            ' TWEAK: como a primavera não preenche corretamente os dados de transporte (antigos), tem que se forçar o metodo PreencheDadosRelacionados_CargaDescarga
            '        existe aqui um passo de conversão de tipos Be.DocumentoVenda <-> GcpBEDocumentoVenda desnecessário, mas mantenho por questões de organização
            Me.PreencheDadosRelacionados_CargaDescarga(DocumentoVenda)
            ' há aqui qualquer problema porque os dados são preenchidos pelo preencheDadosRelacionados //////////////////// VERIFICAR

            objDocumentoVenda = Nothing
        End Sub



        Public Sub PreencheDadosRelacionados_Cliente(ByRef DocumentoVenda As BE.DocumentoVenda)
            Dim objDocumentoVenda As GcpBE900.GcpBEDocumentoVenda

            If IsNothing(DocumentoVenda) Then
                Throw New Exception("O objecto DocumentoVenda passado como parametro na função PreencheDadosRelacionados tem um valor null;")
                Exit Sub
            End If

            objDocumentoVenda = Me.GetGcpBEDocumentoVenda(DocumentoVenda)

            Try
                objMotor.Comercial.Vendas.PreencheDadosRelacionados(objDocumentoVenda, GcpBE900.PreencheRelacaoVendas.vdDadosCliente)
            Catch ex As Exception
                objDocumentoVenda = Nothing
                Throw ex
                Exit Sub
            End Try

            DocumentoVenda = GetDocumentoVenda(objDocumentoVenda)
            objDocumentoVenda = Nothing
        End Sub



        Public Sub PreencheDadosRelacionados_CargaDescarga(ByRef DocumentoVenda As BE.DocumentoVenda)
            Dim objDocumentoVenda As BE.DocumentoVenda
            Dim objCargaDescarga As BE.CargaDescarga


            If IsNothing(DocumentoVenda) Then
                Throw New Exception("O objecto DocumentoVenda passado como parametro na função PreencheDadosRelacionados tem um valor null;")
                Exit Sub
            End If

            objDocumentoVenda = DocumentoVenda

            Try

                ' preenche objeto da CargaDescarga
                objCargaDescarga = objCargasDescargas.GeraCargaDescarga(objDocumentoVenda.ENTIDADE_Tipo, objDocumentoVenda.ENTIDADE_Codigo)
                objDocumentoVenda.TRANSPORTE_CargaDescarga = objCargaDescarga

                ' preenche o resto dos dados (antigos do objeto)
                objDocumentoVenda.TRANSPORTE_Carga_Local = objCargaDescarga.CARGA_Localidade
                objDocumentoVenda.TRANSPORTE_Carga_Data = objCargaDescarga.CARGA_Data
                objDocumentoVenda.TRANSPORTE_Carga_Hora = objCargaDescarga.CARGA_Hora
                objDocumentoVenda.TRANSPORTE_Descarga_Local = objCargaDescarga.DESCARGA_Localidade
                objDocumentoVenda.TRANSPORTE_Descarga_Data = objCargaDescarga.DESCARGA_Data
                objDocumentoVenda.TRANSPORTE_Descarga_Hora = objCargaDescarga.DESCARGA_Hora

            Catch ex As Exception
                objDocumentoVenda = Nothing
                Throw ex
                Exit Sub
            End Try

            DocumentoVenda = objDocumentoVenda
            objCargaDescarga = Nothing
            objDocumentoVenda = Nothing
        End Sub



        Public Sub PreencheDadosRelacionados_TipoDocumento(ByRef DocumentoVenda As BE.DocumentoVenda)
            Dim objDocumentoVenda As GcpBE900.GcpBEDocumentoVenda

            If IsNothing(DocumentoVenda) Then
                Throw New Exception("O objecto DocumentoVenda passado como parametro na função PreencheDadosRelacionados tem um valor null;")
                Exit Sub
            End If

            objDocumentoVenda = Me.GetGcpBEDocumentoVenda(DocumentoVenda)

            Try
                objMotor.Comercial.Vendas.PreencheDadosRelacionados(objDocumentoVenda, GcpBE900.PreencheRelacaoVendas.vdDadosTipoDoc)
            Catch ex As Exception
                objDocumentoVenda = Nothing
                Throw ex
                Exit Sub
            End Try

            DocumentoVenda = GetDocumentoVenda(objDocumentoVenda)
            objDocumentoVenda = Nothing
        End Sub



        Public Sub CalculaDataVencimento(ByRef DocumentoVenda As BE.DocumentoVenda)
            Dim objDocumentoVenda As GcpBE900.GcpBEDocumentoVenda

            If IsNothing(DocumentoVenda) Then
                Throw New Exception("O objecto DocumentoVenda passado como parametro na função CalculaDataVencimento tem um valor null;")
                Exit Sub
            End If

            objDocumentoVenda = Me.GetGcpBEDocumentoVenda(DocumentoVenda)

            Try
                With objDocumentoVenda
                    objDocumentoVenda.DataVenc = objMotor.Comercial.Vendas.CalculaDataVencimento(.DataDoc, .CondPag)
                End With
            Catch ex As Exception
                objDocumentoVenda = Nothing
                Throw ex
                Exit Sub
            End Try

            DocumentoVenda = GetDocumentoVenda(objDocumentoVenda)
            objDocumentoVenda = Nothing
        End Sub



        Public Sub ActualizaDatasStock(ByRef DocumentoVenda As BE.DocumentoVenda)
            Dim objDocumentoVenda As BE.DocumentoVenda
            Dim i As System.Int16

            If IsNothing(DocumentoVenda) Then
                Throw New Exception("O objecto DocumentoVenda passado como parametro na função CalculaDataVencimento tem um valor null;")
                Exit Sub
            End If

            objDocumentoVenda = DocumentoVenda

            Try
                For i = 0 To objDocumentoVenda.Linhas.Count - 1
                    If objDocumentoVenda.Linhas(i).MovimentaStock = True Then
                        objDocumentoVenda.Linhas(i).DataStock = objDocumentoVenda.Data
                    End If
                Next
            Catch ex As Exception
                objDocumentoVenda = Nothing
                Throw ex
                Exit Sub
            End Try

            DocumentoVenda = objDocumentoVenda
            objDocumentoVenda = Nothing
        End Sub



        Public Sub CalculaValoresTotais(ByRef DocumentoVenda As BE.DocumentoVenda)
            Dim objDocumentoVenda As GcpBE900.GcpBEDocumentoVenda

            If IsNothing(DocumentoVenda) Then
                Throw New Exception("O objecto DocumentoVenda passado como parametro na função CalculaValoresTotais tem um valor null;")
                Exit Sub
            End If

            objDocumentoVenda = Me.GetGcpBEDocumentoVenda(DocumentoVenda)

            Try
                objMotor.Comercial.Vendas.CalculaValoresTotais(objDocumentoVenda)
            Catch ex As Exception
                objDocumentoVenda = Nothing
                Throw ex
                Exit Sub
            End Try

            DocumentoVenda = GetDocumentoVenda(objDocumentoVenda)
            objDocumentoVenda = Nothing
        End Sub



        Public Sub Actualiza(ByRef DocumentoVenda As BE.DocumentoVenda, Optional ByVal GuardaComoRascunho As Boolean = False)
            Dim strMensagem As String = ""
            Dim objDocumentoVenda As GcpBE900.GcpBEDocumentoVenda
            Dim i As System.Int32


            If IsNothing(DocumentoVenda) Then
                Throw New Exception("O objecto DocumentoVenda passado como parametro tem um valor null;")
                Exit Sub
            End If

            With DocumentoVenda
                If (.TOTAL_Mercadoria - .TOTAL_Descontos + .TOTAL_Iva + .TOTAL_Outros) < 0 Then
                    Throw New Exception("O documento de venda não pode ter um total negativo;")
                    objDocumentoVenda = Nothing
                    Exit Sub
                End If
            End With

            If Certifica(DocumentoVenda, strMensagem) = False Then
                Throw New Exception(strMensagem)
                Exit Sub
            End If

            With DocumentoVenda
                If .EmModoEdicao = True Then
                    If Me.Existe(.Filial, .Documento_Tipo, .Documento_Serie, .Documento_Numero) = False Then
                        Throw New Exception("O DocumentoVenda que pretende actualizar não existe ou foi removido;")
                        Exit Sub
                    End If
                End If
            End With

            With DocumentoVenda
                For i = .Linhas.Count - 1 To 0 Step -1
                    If .Linhas(i).TipoLinha = LinhaDocumentoVenda_Specs.TiposLinhas.Comentario_60 And .Linhas(i).Descricao.TrimEnd.Length = 0 Then
                        .Linhas.RemoveAt(i)
                    Else
                        Exit For
                    End If
                Next
            End With

            objDocumentoVenda = GetGcpBEDocumentoVenda(DocumentoVenda)

            Try
                If objDocumentoVenda.EmModoEdicao = False Then
                    'O NUMERADOR DO DOCUMENTO NOVO DEVE SER ATRIBUIDO SOMENTE NO ACTO DE GRAVAÇÃO DO DOCUMENTO
                    objDocumentoVenda.NumDoc = 0
                End If

                If GuardaComoRascunho = False Then
                    objMotor.Comercial.Vendas.Actualiza(objDocumentoVenda)
                Else
                    objMotor.Comercial.Vendas.ActualizaRascunho(objDocumentoVenda)
                End If

            Catch ex As Exception
                Throw ex
                Exit Sub
            End Try

            objDocumentoVenda = Nothing
        End Sub



        Public Function Edita(ByVal Filial As String, ByVal Documento_Tipo As String, ByVal Documento_Serie As String, ByVal Documento_Numero As System.Int64) As BE.DocumentoVenda
            Dim objDocumentoVenda As BE.DocumentoVenda

            If Me.Existe(Filial, Documento_Tipo, Documento_Serie, Documento_Numero) = False Then
                Throw New Exception("O documento de venda " & Filial & "/" & Documento_Tipo & "/" & Documento_Serie & "/" & Documento_Numero & " não existe na tabela CabecDoc, pelo que nao pode ser editado;")
                Exit Function
            End If

            Try
                objDocumentoVenda = GetDocumentoVenda(objMotor.Comercial.Vendas.Edita(Filial, Documento_Tipo, Documento_Serie, Documento_Numero))
            Catch ex As Exception
                Throw ex
                Exit Function
            End Try

            Edita = objDocumentoVenda
            objDocumentoVenda = Nothing

        End Function



        Public Function Existe(ByVal Filial As String, ByVal Documento_Tipo As String, ByVal Documento_Serie As String, ByVal Documento_Numero As System.Int64) As Boolean

            If Len(Trim(Filial)) = 0 Then
                Throw New Exception("O parametro Filial da funcao Existe da classe DocumentoVenda deve conter uma string valida;")
                Exit Function
            End If

            If Len(Trim(Documento_Tipo)) = 0 Then
                Throw New Exception("O parametro Documento_Tipo da funcao Existe da classe DocumentoVenda deve conter uma string valida;")
                Exit Function
            End If

            If Len(Trim(Documento_Serie)) = 0 Then
                Throw New Exception("O parametro Documento_Serie da funcao Existe da classe DocumentoVenda deve conter uma string valida;")
                Exit Function
            End If

            If Documento_Numero = 0 Then
                Throw New Exception("O parametro NumeroDocumento da funcao Existe da classe DocumentoVenda deve conter im inteiro diferente de zero;")
                Exit Function
            End If

            Return objMotor.Comercial.Vendas.Existe(Filial, Documento_Tipo, Documento_Serie, Documento_Numero)

        End Function



        Public Sub Remove(ByVal Filial As String, ByVal Documento_Tipo As String, ByVal Documento_Serie As String, ByVal Documento_Numero As System.Int64)
            Dim strMensagem As String = ""

            If Me.Existe(Filial, Documento_Tipo, Documento_Serie, Documento_Numero) = False Then
                Throw New Exception("O DocumentoVenda especificado não existe, pelo que não pode ser removido.")
                Exit Sub
            End If

            If objMotor.Comercial.Vendas.ValidaRemocao(Filial, Documento_Tipo, Documento_Serie, Documento_Numero, strMensagem) = False Then
                Throw New Exception(strMensagem)
                Exit Sub
            End If

            Try
                objMotor.Comercial.Vendas.Remove(Filial, Documento_Tipo, Documento_Serie, Documento_Numero)
            Catch ex As Exception
                Throw ex
                Exit Sub
            End Try

        End Sub



        Public Function Duplica(ByVal DocumentoVendaOrigem As BE.DocumentoVenda, _
                                ByVal Destino_Documento_Tipo As String, ByVal Destino_Documento_Serie As String, _
                                ByVal Destino_Entidade_Tipo As String, ByVal Destino_Entidade_Codigo As String) As BE.DocumentoVenda
            Dim objDocumentoVendaDestino As BE.DocumentoVenda
            Dim objLinhaDocumentoVendaDestino As BE.LinhaDocumentoVenda
            Dim i As System.Int32

            If IsNothing(DocumentoVendaOrigem) Then
                Throw New Exception("O argumento DocumentoVenda do metodo Duplica não pode ser nothing.")
                Exit Function
            End If

            objDocumentoVendaDestino = New BE.DocumentoVenda

            With objDocumentoVendaDestino
                .Id = Guid.NewGuid.ToString
                .Filial = DocumentoVendaOrigem.Filial
                .Seccao = DocumentoVendaOrigem.Seccao
                .Documento_Tipo = Destino_Documento_Tipo
                .Documento_Serie = Destino_Documento_Serie
                .Documento_Numero = 0
                .Documento_Moeda = DocumentoVendaOrigem.Documento_Moeda
                .Documento_Cambio = DocumentoVendaOrigem.Documento_Cambio
                .Documento_CambioMoedaBase = DocumentoVendaOrigem.Documento_CambioMoedaBase
                .Documento_CambioMoedaAlternativa = DocumentoVendaOrigem.Documento_CambioMoedaAlternativa
                .Documento_Arredondamento = DocumentoVendaOrigem.Documento_Arredondamento
                .Documento_ArredondamentoIva = DocumentoVendaOrigem.Documento_ArredondamentoIva
                .ENTIDADE_Tipo = Destino_Entidade_Tipo
                .ENTIDADE_Codigo = Destino_Entidade_Codigo
                .ENTIDADE_Nome = DocumentoVendaOrigem.ENTIDADE_Nome
                .ENTIDADE_Morada = DocumentoVendaOrigem.ENTIDADE_Morada
                .ENTIDADE_Morada2 = DocumentoVendaOrigem.ENTIDADE_Morada2
                .ENTIDADE_Localidade = DocumentoVendaOrigem.ENTIDADE_Localidade
                .ENTIDADE_CodigoPostal = DocumentoVendaOrigem.ENTIDADE_CodigoPostal
                .ENTIDADE_LocalidadePostal = DocumentoVendaOrigem.ENTIDADE_LocalidadePostal
                .ENTIDADE_Contribuinte = DocumentoVendaOrigem.ENTIDADE_Contribuinte
                .ENTIDADE_Desconto = DocumentoVendaOrigem.ENTIDADE_Desconto
                .ENTIDADE_Zona = DocumentoVendaOrigem.ENTIDADE_Zona
                .EntidadeFacturacao_Tipo = DocumentoVendaOrigem.EntidadeFacturacao_Tipo
                .EntidadeFacturacao_Codigo = DocumentoVendaOrigem.EntidadeFacturacao_Codigo
                .EntidadeFacturacao_Nome = DocumentoVendaOrigem.EntidadeFacturacao_Nome
                .EntidadeFacturacao_Morada = DocumentoVendaOrigem.EntidadeFacturacao_Morada
                .EntidadeFacturacao_Morada2 = DocumentoVendaOrigem.EntidadeFacturacao_Morada2
                .EntidadeFacturacao_Localidade = DocumentoVendaOrigem.EntidadeFacturacao_Localidade
                .EntidadeFacturacao_CodigoPostal = DocumentoVendaOrigem.EntidadeFacturacao_CodigoPostal
                .EntidadeFacturacao_LocalidadePostal = DocumentoVendaOrigem.EntidadeFacturacao_LocalidadePostal
                .EntidadeFacturacao_Contribuinte = DocumentoVendaOrigem.EntidadeFacturacao_Contribuinte
                .CondicaoPagamento = DocumentoVendaOrigem.CondicaoPagamento
                .ModoExpedicao = DocumentoVendaOrigem.ModoExpedicao
                .ModoPagamento = DocumentoVendaOrigem.ModoPagamento
                .Data = Now.Date
                .Vencimento = Now.Date
                .Requisicao = DocumentoVendaOrigem.Requisicao
                .TOTAL_Mercadoria = DocumentoVendaOrigem.TOTAL_Mercadoria
                .TOTAL_Descontos = DocumentoVendaOrigem.TOTAL_Descontos
                .TOTAL_Iva = DocumentoVendaOrigem.TOTAL_Iva
                .TOTAL_Outros = DocumentoVendaOrigem.TOTAL_Outros
                .TOTAL_Documento = DocumentoVendaOrigem.TOTAL_Documento
                .TRANSPORTE_CargaDescarga = DocumentoVendaOrigem.TRANSPORTE_CargaDescarga
                .TRANSPORTE_Carga_Local = DocumentoVendaOrigem.TRANSPORTE_Carga_Local
                .TRANSPORTE_Carga_Data = DocumentoVendaOrigem.TRANSPORTE_Carga_Data
                .TRANSPORTE_Carga_Hora = DocumentoVendaOrigem.TRANSPORTE_Carga_Hora
                .TRANSPORTE_Descarga_Local = DocumentoVendaOrigem.TRANSPORTE_Descarga_Local
                .TRANSPORTE_Descarga_Data = DocumentoVendaOrigem.TRANSPORTE_Descarga_Data
                .TRANSPORTE_Descarga_Hora = DocumentoVendaOrigem.TRANSPORTE_Descarga_Hora
                .TRANSPORTE_Matricula = DocumentoVendaOrigem.TRANSPORTE_Matricula
                .Observacoes = DocumentoVendaOrigem.Observacoes
                .Utilizador = DocumentoVendaOrigem.Utilizador
                .DataUltimaActualizacao = DocumentoVendaOrigem.DataUltimaActualizacao
                .CamposUtilizador = DocumentoVendaOrigem.CamposUtilizador
                .EmModoEdicao = False
            End With

            For i = 0 To DocumentoVendaOrigem.Linhas.Count - 1
                objLinhaDocumentoVendaDestino = New BE.LinhaDocumentoVenda
                With objLinhaDocumentoVendaDestino
                    .Armazem = DocumentoVendaOrigem.Linhas(i).Armazem
                    .Artigo = DocumentoVendaOrigem.Linhas(i).Artigo
                    .Lote = DocumentoVendaOrigem.Linhas(i).Lote
                    .CamposUtilizador = DocumentoVendaOrigem.Linhas(i).CamposUtilizador
                    .CodigoIva = DocumentoVendaOrigem.Linhas(i).CodigoIva
                    .DataEntrega = DocumentoVendaOrigem.Linhas(i).DataEntrega
                    .DataStock = DocumentoVendaOrigem.Linhas(i).DataStock
                    .Desconto1 = DocumentoVendaOrigem.Linhas(i).Desconto1
                    .Desconto2 = DocumentoVendaOrigem.Linhas(i).Desconto2
                    .Desconto3 = DocumentoVendaOrigem.Linhas(i).Desconto3
                    .Descricao = DocumentoVendaOrigem.Linhas(i).Descricao
                    .Id = Guid.NewGuid.ToString
                    .Localizacao = DocumentoVendaOrigem.Linhas(i).Localizacao
                    .MovimentaStock = DocumentoVendaOrigem.Linhas(i).MovimentaStock
                    .NumeroLinha = DocumentoVendaOrigem.Linhas(i).NumeroLinha
                    .PrecoLiquido = DocumentoVendaOrigem.Linhas(i).PrecoLiquido
                    .PrecoUnitario = DocumentoVendaOrigem.Linhas(i).PrecoUnitario
                    .PrecoMedioCusto = DocumentoVendaOrigem.Linhas(i).PrecoMedioCusto
                    .Quantidade = DocumentoVendaOrigem.Linhas(i).Quantidade
                    .QuantidadeSatisfeita = DocumentoVendaOrigem.Linhas(i).QuantidadeSatisfeita
                    .TaxaIva = DocumentoVendaOrigem.Linhas(i).TaxaIva
                    .TipoLinha = DocumentoVendaOrigem.Linhas(i).TipoLinha
                    .Unidade = DocumentoVendaOrigem.Linhas(i).Unidade
                    .PercentagemIncidenciaIva = DocumentoVendaOrigem.Linhas(i).PercentagemIncidenciaIva
                    .PercentagemIvaDedutivel = DocumentoVendaOrigem.Linhas(i).PercentagemIvaDedutivel
                    .IvaNaoDedutivel = DocumentoVendaOrigem.Linhas(i).IvaNaoDedutivel
                End With
                objDocumentoVendaDestino.Linhas.Add(objLinhaDocumentoVendaDestino, objLinhaDocumentoVendaDestino.Id)
                objLinhaDocumentoVendaDestino = Nothing
            Next

            Me.PreencheDadosRelacionados_Todos(objDocumentoVendaDestino)
            Me.CalculaDataVencimento(objDocumentoVendaDestino)
            Me.CalculaValoresTotais(objDocumentoVendaDestino)

            Duplica = objDocumentoVendaDestino

            objDocumentoVendaDestino = Nothing
        End Function



        Private Function GetDocumentoVenda(ByVal DocumentoVenda As GcpBE900.GcpBEDocumentoVenda) As BE.DocumentoVenda
            Dim objDocumentoVenda As New BE.DocumentoVenda
            Dim objLinhaDocumentoVenda As BE.LinhaDocumentoVenda
            Dim objCampoUtilizador As BE.CampoUtilizador = Nothing
            Dim i As System.Int32 = 0
            Dim ii As System.Int32 = 0

            With objDocumentoVenda
                .Id = DocumentoVenda.ID
                .Filial = DocumentoVenda.Filial
                .Seccao = DocumentoVenda.Seccao
                .Documento_Tipo = DocumentoVenda.Tipodoc
                .Documento_Serie = DocumentoVenda.Serie
                .Documento_Numero = DocumentoVenda.NumDoc
                .Documento_Moeda = DocumentoVenda.Moeda
                .Documento_Cambio = DocumentoVenda.Cambio
                .Documento_CambioMoedaBase = DocumentoVenda.CambioMBase
                .Documento_CambioMoedaAlternativa = DocumentoVenda.CambioMAlt
                .Documento_MoedaDaUEM = DocumentoVenda.MoedaDaUEM
                .Documento_Arredondamento = DocumentoVenda.Arredondamento
                .Documento_ArredondamentoIva = DocumentoVenda.ArredondamentoIva
                .ENTIDADE_Tipo = DocumentoVenda.TipoEntidade
                .ENTIDADE_Codigo = DocumentoVenda.Entidade
                .ENTIDADE_Nome = DocumentoVenda.Nome
                .ENTIDADE_Morada = DocumentoVenda.Morada
                .ENTIDADE_Morada2 = DocumentoVenda.Morada2
                .ENTIDADE_Localidade = DocumentoVenda.Localidade
                .ENTIDADE_CodigoPostal = DocumentoVenda.CodigoPostal
                .ENTIDADE_LocalidadePostal = DocumentoVenda.LocalidadeCodigoPostal
                .ENTIDADE_Distrito = DocumentoVenda.Distrito
                .ENTIDADE_Pais = DocumentoVenda.Pais
                .ENTIDADE_Contribuinte = DocumentoVenda.NumContribuinte
                .ENTIDADE_Desconto = DocumentoVenda.DescEntidade
                .ENTIDADE_Zona = DocumentoVenda.Zona
                .EntidadeFacturacao_Tipo = DocumentoVenda.TipoEntidadeFac
                .EntidadeFacturacao_Codigo = DocumentoVenda.EntidadeFac
                .EntidadeFacturacao_Nome = DocumentoVenda.NomeFac
                .EntidadeFacturacao_Morada = DocumentoVenda.MoradaFac
                .EntidadeFacturacao_Morada2 = DocumentoVenda.Morada2Fac
                .EntidadeFacturacao_Localidade = DocumentoVenda.LocalidadeFac
                .EntidadeFacturacao_CodigoPostal = DocumentoVenda.CodigoPostalFac
                .EntidadeFacturacao_LocalidadePostal = DocumentoVenda.LocalidadeCodigoPostalFac
                .EntidadeFacturacao_Distrito = DocumentoVenda.DistritoFac
                .EntidadeFacturacao_Pais = DocumentoVenda.PaisFac
                .EntidadeFacturacao_Contribuinte = DocumentoVenda.NumContribuinteFac
                .EntidadeEntrega_Tipo = DocumentoVenda.TipoEntidadeEntrega
                .EntidadeEntrega_Codigo = DocumentoVenda.EntidadeEntrega
                .EntidadeEntrega_Nome = DocumentoVenda.NomeEntrega
                .EntidadeEntrega_Morada = DocumentoVenda.MoradaEntrega
                .EntidadeEntrega_Morada2 = DocumentoVenda.Morada2Entrega
                .EntidadeEntrega_Localidade = DocumentoVenda.LocalidadeEntrega
                .EntidadeEntrega_CodigoPostal = DocumentoVenda.CodPostalEntrega
                .EntidadeEntrega_LocalidadePostal = DocumentoVenda.CodPostalLocalidadeEntrega
                .EntidadeEntrega_Distrito = DocumentoVenda.DistritoEntrega
                .CondicaoPagamento = DocumentoVenda.CondPag
                .ModoExpedicao = DocumentoVenda.ModoExp
                .ModoPagamento = DocumentoVenda.ModoPag
                .Data = DocumentoVenda.DataDoc
                .Vencimento = DocumentoVenda.DataVenc
                .Requisicao = DocumentoVenda.Requisicao
                .LocalOperacao = DocumentoVenda.LocalOperacao
                .TOTAL_Mercadoria = DocumentoVenda.TotalMerc
                .TOTAL_Descontos = DocumentoVenda.TotalDesc
                .TOTAL_Iva = DocumentoVenda.TotalIva
                .TOTAL_Outros = DocumentoVenda.TotalOutros
                .TOTAL_Documento = DocumentoVenda.TotalDocumento
                .TRANSPORTE_CargaDescarga = objCargasDescargas.GetCargaDescarga(DocumentoVenda.CargaDescarga)
                .TRANSPORTE_Carga_Local = DocumentoVenda.LocalCarga
                .TRANSPORTE_Carga_Data = DocumentoVenda.DataCarga
                .TRANSPORTE_Carga_Hora = DocumentoVenda.HoraCarga
                .TRANSPORTE_Descarga_Local = DocumentoVenda.LocalDescarga
                .TRANSPORTE_Descarga_Data = DocumentoVenda.DataDescarga
                .TRANSPORTE_Descarga_Hora = DocumentoVenda.HoraDescarga
                .TRANSPORTE_Matricula = DocumentoVenda.Matricula
                .Observacoes = DocumentoVenda.Observacoes
                .Utilizador = DocumentoVenda.Utilizador
                .DataUltimaActualizacao = DocumentoVenda.DataUltimaActualizacao
                .EmModoEdicao = DocumentoVenda.EmModoEdicao

                For i = 1 To DocumentoVenda.Linhas.NumItens
                    objLinhaDocumentoVenda = New BE.LinhaDocumentoVenda
                    With objLinhaDocumentoVenda

                        .Id = DocumentoVenda.Linhas(i).IdLinha
                        .NumeroLinha = i

                        Select Case DocumentoVenda.Linhas(i).TipoLinha

                            Case "60"
                                .TipoLinha = LinhaDocumentoVenda_Specs.TiposLinhas.Comentario_60
                                .Descricao = DocumentoVenda.Linhas(i).Descricao

                            Case "30"
                                .TipoLinha = LinhaDocumentoVenda_Specs.TiposLinhas.Acerto_30
                                .Descricao = DocumentoVenda.Linhas(i).Descricao
                                .Quantidade = DocumentoVenda.Linhas(i).Quantidade
                                .PrecoUnitario = DocumentoVenda.Linhas(i).PrecUnit
                                .PrecoLiquido = DocumentoVenda.Linhas(i).PrecoLiquido

                            Case "50"
                                .TipoLinha = LinhaDocumentoVenda_Specs.TiposLinhas.Portes_50
                                .Descricao = DocumentoVenda.Linhas(i).Descricao
                                .Quantidade = DocumentoVenda.Linhas(i).Quantidade
                                .PrecoUnitario = DocumentoVenda.Linhas(i).PrecUnit
                                .CodigoIva = DocumentoVenda.Linhas(i).CodIva
                                .TaxaIva = DocumentoVenda.Linhas(i).TaxaIva
                                .PrecoLiquido = DocumentoVenda.Linhas(i).PrecoLiquido

                            Case Else
                                Select Case DocumentoVenda.Linhas(i).TipoLinha
                                    Case "10"
                                        .TipoLinha = LinhaDocumentoVenda_Specs.TiposLinhas.Mercadoria_TipoArtigo_3_TipoLinha_10
                                    Case "20"
                                        .TipoLinha = LinhaDocumentoVenda_Specs.TiposLinhas.ServicoA_TipoArtigo_0_TipoLinha_20
                                    Case "13"
                                        .TipoLinha = LinhaDocumentoVenda_Specs.TiposLinhas.MateriaPrima_TipoArtigo_6_TipoLinha_13
                                    Case "14"
                                        .TipoLinha = LinhaDocumentoVenda_Specs.TiposLinhas.MateriaSubsidiaria_TipoArtigo_7_TipoLinha_14
                                    Case "23"
                                        .TipoLinha = LinhaDocumentoVenda_Specs.TiposLinhas.MaoDeObra_TipoArtigo_12_TipoLinha_23
                                End Select
                                .Artigo = DocumentoVenda.Linhas(i).Artigo
                                If DocumentoVenda.Linhas(i).MovStock = "S" Then
                                    .MovimentaStock = True
                                    .Armazem = DocumentoVenda.Linhas(i).Armazem
                                    .Localizacao = DocumentoVenda.Linhas(i).Localizacao
                                    If DocumentoVenda.Linhas(i).Lote.Contains("<") Then
                                        .Lote = ""
                                    Else
                                        .Lote = DocumentoVenda.Linhas(i).Lote
                                    End If
                                Else
                                    .MovimentaStock = False
                                    .Armazem = ""
                                    .Localizacao = ""
                                    .Lote = ""
                                End If
                                .Quantidade = DocumentoVenda.Linhas(i).Quantidade
                                .QuantidadeSatisfeita = DocumentoVenda.Linhas(i).QuantSatisfeita
                                .Unidade = DocumentoVenda.Linhas(i).Unidade
                                .DataEntrega = DocumentoVenda.Linhas(i).DataEntrega
                                .DataStock = DocumentoVenda.Linhas(i).DataStock
                                .Descricao = DocumentoVenda.Linhas(i).Descricao
                                .CodigoIva = DocumentoVenda.Linhas(i).CodIva
                                .TaxaIva = DocumentoVenda.Linhas(i).TaxaIva
                                .PrecoUnitario = DocumentoVenda.Linhas(i).PrecUnit
                                .PrecoMedioCusto = DocumentoVenda.Linhas(i).PCM     ' 20130803 adicionado tratamento PMC
                                .Desconto1 = DocumentoVenda.Linhas(i).Desconto1
                                .Desconto2 = DocumentoVenda.Linhas(i).Desconto2
                                .Desconto3 = DocumentoVenda.Linhas(i).Desconto3

                                ' ERRO DA PRIMAVERA PORQUE GRAVA UM PRECO LIQUIDO COM UM VALOR 0.000000000234323
                                If DocumentoVenda.Linhas(i).PrecoLiquido < 0.001 Then
                                    .PrecoLiquido = 0
                                Else
                                    .PrecoLiquido = DocumentoVenda.Linhas(i).PrecoLiquido
                                End If

                                .PercentagemIncidenciaIva = DocumentoVenda.Linhas(i).PercIncidenciaIVA
                                .PercentagemIvaDedutivel = DocumentoVenda.Linhas(i).PercIvaDedutivel
                                .IvaNaoDedutivel = DocumentoVenda.Linhas(i).IvaNaoDedutivel
                        End Select
                    End With

                    If DocumentoVenda.Linhas(i).CamposUtil.NumItens > 0 Then
                        For ii = 1 To DocumentoVenda.Linhas(i).CamposUtil.NumItens
                            If Not IsDBNull(DocumentoVenda.Linhas(i).CamposUtil(ii)) Then
                                objCampoUtilizador = New BE.CampoUtilizador
                                With objCampoUtilizador
                                    Select Case DocumentoVenda.Linhas(i).CamposUtil(ii).TipoSimplificado
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
                                            Throw New Exception("Foi referenciado um tipo de dados em DocumentoVenda.linhas(i).CamposUtil não esperado na classe Business.Vendas.DocumentoVenda função GetDocumentoVenda")
                                            Exit Function

                                    End Select
                                    .Campo = DocumentoVenda.Linhas(i).CamposUtil(ii).Nome
                                    If Not IsDBNull(DocumentoVenda.Linhas(i).CamposUtil(ii).Valor) Then
                                        .Valor = CType(DocumentoVenda.Linhas(i).CamposUtil(ii).Valor, Object)
                                    Else
                                        .Valor = CType("", Object)
                                    End If

                                End With
                                objLinhaDocumentoVenda.CamposUtilizador.Add(objCampoUtilizador, objCampoUtilizador.Campo)
                                objCampoUtilizador = Nothing
                            End If
                        Next
                    End If

                    objDocumentoVenda.Linhas.Add(objLinhaDocumentoVenda, objLinhaDocumentoVenda.Id)
                    objLinhaDocumentoVenda = Nothing
                Next

                For i = 1 To DocumentoVenda.CamposUtil.NumItens
                    If Not IsDBNull(DocumentoVenda.CamposUtil(i)) Then
                        objCampoUtilizador = New BE.CampoUtilizador
                        With objCampoUtilizador
                            Select Case DocumentoVenda.CamposUtil(i).TipoSimplificado
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
                                    Throw New Exception("Foi referenciado um tipo de dados em DocumentoVenda.Linhas.CamposUtil não esperado na classe Business.DocumentoVenda função GetDocumentoVenda")
                                    Exit Function

                            End Select
                            .Campo = DocumentoVenda.CamposUtil(i).Nome
                            If Not IsDBNull(DocumentoVenda.CamposUtil(i).Valor) Then
                                .Valor = CType(DocumentoVenda.CamposUtil(i).Valor, Object)
                            Else
                                .Valor = CType("", Object)
                            End If

                        End With
                        .CamposUtilizador.Add(objCampoUtilizador, objCampoUtilizador.Campo)
                        objCampoUtilizador = Nothing
                    End If
                Next

            End With

            GetDocumentoVenda = objDocumentoVenda
            objDocumentoVenda = Nothing

        End Function



        Private Function GetGcpBEDocumentoVenda(ByVal DocumentoVenda As BE.DocumentoVenda) As GcpBE900.GcpBEDocumentoVenda
            Dim objDocumentoVenda As New GcpBE900.GcpBEDocumentoVenda
            Dim objLinhaDocumentoVenda As GcpBE900.GcpBELinhaDocumentoVenda
            Dim i As System.Int32 = 0
            Dim ii As System.Int32 = 0

            If DocumentoVenda.EmModoEdicao = True Then 'EM EDIÇÃO
                With DocumentoVenda
                    objDocumentoVenda = objMotor.Comercial.Vendas.Edita(.Filial, .Documento_Tipo, .Documento_Serie, .Documento_Numero)
                End With
            End If

            With objDocumentoVenda
                .ID = DocumentoVenda.Id
                .Filial = DocumentoVenda.Filial
                .Seccao = DocumentoVenda.Seccao
                .Tipodoc = DocumentoVenda.Documento_Tipo
                .Serie = DocumentoVenda.Documento_Serie
                .Moeda = DocumentoVenda.Documento_Moeda
                .Cambio = DocumentoVenda.Documento_Cambio
                .CambioMBase = DocumentoVenda.Documento_CambioMoedaBase
                .CambioMAlt = DocumentoVenda.Documento_CambioMoedaAlternativa
                .MoedaDaUEM = DocumentoVenda.Documento_MoedaDaUEM
                .Arredondamento = DocumentoVenda.Documento_Arredondamento
                .ArredondamentoIva = DocumentoVenda.Documento_ArredondamentoIva
                .TipoEntidade = DocumentoVenda.ENTIDADE_Tipo
                .Entidade = DocumentoVenda.ENTIDADE_Codigo
                .NumDoc = DocumentoVenda.Documento_Numero
                .Nome = DocumentoVenda.ENTIDADE_Nome
                .Morada = DocumentoVenda.ENTIDADE_Morada
                .Morada2 = DocumentoVenda.ENTIDADE_Morada2
                .Localidade = DocumentoVenda.ENTIDADE_Localidade
                .CodigoPostal = DocumentoVenda.ENTIDADE_CodigoPostal
                .LocalidadeCodigoPostal = DocumentoVenda.ENTIDADE_LocalidadePostal
                .Distrito = DocumentoVenda.ENTIDADE_Distrito
                .Pais = DocumentoVenda.ENTIDADE_Pais
                .NumContribuinte = DocumentoVenda.ENTIDADE_Contribuinte
                .DescEntidade = DocumentoVenda.ENTIDADE_Desconto
                .Zona = DocumentoVenda.ENTIDADE_Zona
                .TipoEntidadeFac = DocumentoVenda.EntidadeFacturacao_Tipo
                .EntidadeFac = DocumentoVenda.EntidadeFacturacao_Codigo
                .NomeFac = DocumentoVenda.EntidadeFacturacao_Nome
                .MoradaFac = DocumentoVenda.EntidadeFacturacao_Morada
                .Morada2Fac = DocumentoVenda.EntidadeFacturacao_Morada2
                .LocalidadeFac = DocumentoVenda.EntidadeFacturacao_Localidade
                .CodigoPostalFac = DocumentoVenda.EntidadeFacturacao_CodigoPostal
                .LocalidadeCodigoPostalFac = DocumentoVenda.EntidadeFacturacao_LocalidadePostal
                .DistritoFac = DocumentoVenda.EntidadeFacturacao_Distrito
                .PaisFac = DocumentoVenda.EntidadeFacturacao_Pais
                .NumContribuinteFac = DocumentoVenda.EntidadeFacturacao_Contribuinte
                .TipoEntidadeEntrega = DocumentoVenda.EntidadeEntrega_Tipo
                .EntidadeEntrega = DocumentoVenda.EntidadeEntrega_Codigo
                .NomeEntrega = DocumentoVenda.EntidadeEntrega_Nome
                .MoradaEntrega = DocumentoVenda.EntidadeEntrega_Morada
                .Morada2Entrega = DocumentoVenda.EntidadeEntrega_Morada2
                .LocalidadeEntrega = DocumentoVenda.EntidadeEntrega_Localidade
                .CodPostalEntrega = DocumentoVenda.EntidadeEntrega_CodigoPostal
                .CodPostalLocalidadeEntrega = DocumentoVenda.EntidadeEntrega_LocalidadePostal
                .DistritoEntrega = DocumentoVenda.EntidadeEntrega_Distrito
                .CondPag = DocumentoVenda.CondicaoPagamento
                .ModoExp = DocumentoVenda.ModoExpedicao
                .ModoPag = DocumentoVenda.ModoPagamento
                .DataDoc = DocumentoVenda.Data
                .DataVenc = DocumentoVenda.Vencimento
                .LocalOperacao = DocumentoVenda.LocalOperacao
                .Requisicao = DocumentoVenda.Requisicao
                .TotalMerc = DocumentoVenda.TOTAL_Mercadoria
                .TotalDesc = DocumentoVenda.TOTAL_Descontos
                .TotalIva = DocumentoVenda.TOTAL_Iva
                .TotalOutros = DocumentoVenda.TOTAL_Outros
                .TotalDocumento = DocumentoVenda.TOTAL_Documento
                .CargaDescarga = objCargasDescargas.GetGcpBECargaDescarga(DocumentoVenda.TRANSPORTE_CargaDescarga)
                .LocalCarga = DocumentoVenda.TRANSPORTE_Carga_Local
                .DataCarga = DocumentoVenda.TRANSPORTE_Carga_Data
                .HoraCarga = DocumentoVenda.TRANSPORTE_Carga_Hora
                .LocalDescarga = DocumentoVenda.TRANSPORTE_Descarga_Local
                .DataDescarga = DocumentoVenda.TRANSPORTE_Descarga_Data
                .HoraDescarga = DocumentoVenda.TRANSPORTE_Descarga_Hora
                .Matricula = DocumentoVenda.TRANSPORTE_Matricula
                .Observacoes = DocumentoVenda.Observacoes
                .Utilizador = DocumentoVenda.Utilizador
                .DataUltimaActualizacao = DocumentoVenda.DataUltimaActualizacao
                .EmModoEdicao = DocumentoVenda.EmModoEdicao
            End With

            objDocumentoVenda.Linhas.RemoveTodos()

            With objDocumentoVenda

                For i = 0 To DocumentoVenda.Linhas.Count - 1
                    objLinhaDocumentoVenda = New GcpBE900.GcpBELinhaDocumentoVenda
                    With objLinhaDocumentoVenda

                        .IdLinha = DocumentoVenda.Linhas(i).Id

                        Select Case DocumentoVenda.Linhas(i).TipoLinha

                            Case LinhaDocumentoVenda_Specs.TiposLinhas.Comentario_60
                                .TipoLinha = "60"
                                .Descricao = DocumentoVenda.Linhas(i).Descricao

                            Case LinhaDocumentoVenda_Specs.TiposLinhas.Acerto_30
                                .TipoLinha = "30"
                                .Descricao = DocumentoVenda.Linhas(i).Descricao
                                .Quantidade = DocumentoVenda.Linhas(i).Quantidade
                                .PrecUnit = DocumentoVenda.Linhas(i).PrecoUnitario
                                .PrecoLiquido = DocumentoVenda.Linhas(i).PrecoLiquido

                            Case LinhaDocumentoVenda_Specs.TiposLinhas.Portes_50
                                .TipoLinha = "50"
                                .Descricao = DocumentoVenda.Linhas(i).Descricao
                                .Quantidade = DocumentoVenda.Linhas(i).Quantidade
                                .PrecUnit = DocumentoVenda.Linhas(i).PrecoUnitario
                                .CodIva = DocumentoVenda.Linhas(i).CodigoIva
                                .TaxaIva = DocumentoVenda.Linhas(i).TaxaIva
                                .PrecoLiquido = DocumentoVenda.Linhas(i).PrecoLiquido

                            Case Else
                                Select Case DocumentoVenda.Linhas(i).TipoLinha
                                    Case LinhaDocumentoVenda_Specs.TiposLinhas.Mercadoria_TipoArtigo_3_TipoLinha_10
                                        .TipoLinha = "10"
                                    Case LinhaDocumentoVenda_Specs.TiposLinhas.ServicoA_TipoArtigo_0_TipoLinha_20
                                        .TipoLinha = "20"
                                    Case LinhaDocumentoVenda_Specs.TiposLinhas.MateriaPrima_TipoArtigo_6_TipoLinha_13
                                        .TipoLinha = "13"
                                    Case LinhaDocumentoVenda_Specs.TiposLinhas.MateriaSubsidiaria_TipoArtigo_7_TipoLinha_14
                                        .TipoLinha = "14"
                                    Case LinhaDocumentoVenda_Specs.TiposLinhas.MaoDeObra_TipoArtigo_12_TipoLinha_23
                                        .TipoLinha = "23"
                                End Select
                                .Artigo = DocumentoVenda.Linhas(i).Artigo
                                If DocumentoVenda.Linhas(i).MovimentaStock = True Then
                                    .MovStock = "S"
                                    .Armazem = DocumentoVenda.Linhas(i).Armazem
                                    .Localizacao = DocumentoVenda.Linhas(i).Localizacao
                                    .Lote = DocumentoVenda.Linhas(i).Lote
                                Else
                                    .MovStock = "N"
                                    .Armazem = ""
                                    .Localizacao = ""
                                    .Lote = ""
                                End If
                                .Quantidade = DocumentoVenda.Linhas(i).Quantidade
                                .QuantSatisfeita = DocumentoVenda.Linhas(i).QuantidadeSatisfeita
                                .Unidade = DocumentoVenda.Linhas(i).Unidade
                                .DataEntrega = DocumentoVenda.Linhas(i).DataEntrega
                                .DataStock = DocumentoVenda.Linhas(i).DataStock
                                .Descricao = DocumentoVenda.Linhas(i).Descricao
                                .CodIva = DocumentoVenda.Linhas(i).CodigoIva
                                .TaxaIva = DocumentoVenda.Linhas(i).TaxaIva
                                .PrecUnit = DocumentoVenda.Linhas(i).PrecoUnitario
                                .PCM = DocumentoVenda.Linhas(i).PrecoMedioCusto ' 20130803 adicionado tratamento PMC
                                .Desconto1 = DocumentoVenda.Linhas(i).Desconto1
                                .Desconto2 = DocumentoVenda.Linhas(i).Desconto2
                                .Desconto3 = DocumentoVenda.Linhas(i).Desconto3
                                .PrecoLiquido = DocumentoVenda.Linhas(i).PrecoLiquido
                                .PercIncidenciaIVA = DocumentoVenda.Linhas(i).PercentagemIncidenciaIva
                                .PercIvaDedutivel = DocumentoVenda.Linhas(i).PercentagemIvaDedutivel
                                .IvaNaoDedutivel = DocumentoVenda.Linhas(i).IvaNaoDedutivel

                        End Select
                    End With


                    If DocumentoVenda.Linhas(i).CamposUtilizador.Count > 0 Then
                        For ii = 0 To DocumentoVenda.Linhas(i).CamposUtilizador.Count - 1

                            Select Case DocumentoVenda.Linhas(i).CamposUtilizador(ii).Tipo
                                Case BE.CampoUtilizador.TiposDadosCampoUtilizador._String
                                    objLinhaDocumentoVenda.CamposUtil.Item(DocumentoVenda.Linhas(i).CamposUtilizador(ii).Campo).Valor = CType(DocumentoVenda.Linhas(i).CamposUtilizador(ii).Valor, String)
                                    objLinhaDocumentoVenda.CamposUtil.Item(DocumentoVenda.Linhas(i).CamposUtilizador(ii).Campo).Tipo = StdBE900.EnumTipoCampo.tcNVarchar

                                Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Integer
                                    objLinhaDocumentoVenda.CamposUtil.Item(DocumentoVenda.Linhas(i).CamposUtilizador(ii).Campo).Valor = CType(DocumentoVenda.Linhas(i).CamposUtilizador(ii).Valor, System.Int32)
                                    objLinhaDocumentoVenda.CamposUtil.Item(DocumentoVenda.Linhas(i).CamposUtilizador(ii).Campo).Tipo = StdBE900.EnumTipoCampo.tcInt

                                Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Money
                                    objLinhaDocumentoVenda.CamposUtil.Item(DocumentoVenda.Linhas(i).CamposUtilizador(ii).Campo).Valor = CType(DocumentoVenda.Linhas(i).CamposUtilizador(ii).Valor, System.Decimal)
                                    objLinhaDocumentoVenda.CamposUtil.Item(DocumentoVenda.Linhas(i).CamposUtilizador(ii).Campo).Tipo = StdBE900.EnumTipoCampo.tcMoney

                                Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Date
                                    objLinhaDocumentoVenda.CamposUtil.Item(DocumentoVenda.Linhas(i).CamposUtilizador(ii).Campo).Valor = CType(DocumentoVenda.Linhas(i).CamposUtilizador(ii).Valor, System.DateTime)
                                    objLinhaDocumentoVenda.CamposUtil.Item(DocumentoVenda.Linhas(i).CamposUtilizador(ii).Campo).Tipo = StdBE900.EnumTipoCampo.tcDateTime

                                Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Boolean
                                    objLinhaDocumentoVenda.CamposUtil.Item(DocumentoVenda.Linhas(i).CamposUtilizador(ii).Campo).Valor = CType(DocumentoVenda.Linhas(i).CamposUtilizador(ii).Valor, Boolean)
                                    objLinhaDocumentoVenda.CamposUtil.Item(DocumentoVenda.Linhas(i).CamposUtilizador(ii).Campo).Tipo = StdBE900.EnumTipoCampo.tcBit

                                Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Double
                                    objLinhaDocumentoVenda.CamposUtil.Item(DocumentoVenda.Linhas(i).CamposUtilizador(ii).Campo).Valor = CType(DocumentoVenda.Linhas(i).CamposUtilizador(ii).Valor, Double)
                                    objLinhaDocumentoVenda.CamposUtil.Item(DocumentoVenda.Linhas(i).CamposUtilizador(ii).Campo).Tipo = StdBE900.EnumTipoCampo.tcDecimal


                                Case Else
                                    Throw New Exception("Foi referenciado um tipo de dados em DocumentoVenda.Linhas(i).CamposUtil não esperado na classe Business.DocumentoVenda função GetGcpBEDocumentoVenda")
                                    Exit Function

                            End Select

                        Next
                    End If

                    .Linhas.Insere(objLinhaDocumentoVenda)
                    objLinhaDocumentoVenda = Nothing
                Next


                If DocumentoVenda.CamposUtilizador.Count > 0 Then
                    For i = 0 To DocumentoVenda.CamposUtilizador.Count - 1

                        Select Case DocumentoVenda.CamposUtilizador(i).Tipo
                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._String
                                .CamposUtil.Item(DocumentoVenda.CamposUtilizador(i).Campo).Valor = CType(DocumentoVenda.CamposUtilizador(i).Valor, String)
                                .CamposUtil.Item(DocumentoVenda.CamposUtilizador(i).Campo).Tipo = StdBE900.EnumTipoCampo.tcNVarchar

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Integer
                                .CamposUtil.Item(DocumentoVenda.CamposUtilizador(i).Campo).Valor = CType(DocumentoVenda.CamposUtilizador(i).Valor, System.Int32)
                                .CamposUtil.Item(DocumentoVenda.CamposUtilizador(i).Campo).Tipo = StdBE900.EnumTipoCampo.tcInt

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Money
                                .CamposUtil.Item(DocumentoVenda.CamposUtilizador(i).Campo).Valor = CType(DocumentoVenda.CamposUtilizador(i).Valor, System.Decimal)
                                .CamposUtil.Item(DocumentoVenda.CamposUtilizador(i).Campo).Tipo = StdBE900.EnumTipoCampo.tcMoney

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Date
                                .CamposUtil.Item(DocumentoVenda.CamposUtilizador(i).Campo).Valor = CType(DocumentoVenda.CamposUtilizador(i).Valor, System.DateTime)
                                .CamposUtil.Item(DocumentoVenda.CamposUtilizador(i).Campo).Tipo = StdBE900.EnumTipoCampo.tcDateTime

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Boolean
                                .CamposUtil.Item(DocumentoVenda.CamposUtilizador(i).Campo).Valor = CType(DocumentoVenda.CamposUtilizador(i).Valor, Boolean)
                                .CamposUtil.Item(DocumentoVenda.CamposUtilizador(i).Campo).Tipo = StdBE900.EnumTipoCampo.tcBit

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Double
                                .CamposUtil.Item(DocumentoVenda.CamposUtilizador(i).Campo).Valor = CType(DocumentoVenda.CamposUtilizador(i).Valor, Double)
                                .CamposUtil.Item(DocumentoVenda.CamposUtilizador(i).Campo).Tipo = StdBE900.EnumTipoCampo.tcDecimal


                            Case Else
                                Throw New Exception("Foi referenciado um tipo de dados em DocumentoVenda.CamposUtil não esperado na classe Business.DocumentoVenda função GetGcpBEDocumentoVenda")
                                Exit Function

                        End Select

                    Next
                End If
            End With

            GetGcpBEDocumentoVenda = objDocumentoVenda
            objDocumentoVenda = Nothing

        End Function



        Protected Overrides Sub Finalize()
            If Not IsNothing(objMotor) Then
                objMotor = Nothing
            End If
            MyBase.Finalize()
        End Sub


    End Class

End Namespace
