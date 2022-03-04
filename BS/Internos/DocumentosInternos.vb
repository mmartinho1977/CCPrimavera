Imports Microsoft.VisualBasic
Imports CCPrimavera.Specifications
Imports CCPrimavera.Specifications.DocumentoInterno_Specs
Imports Interop



Namespace BS

    Public Class DocumentosInternos
        Private objMotor As ErpBS900.ErpBS


        Sub New(ByRef Motor As ErpBS900.ErpBS)
            objMotor = Motor
        End Sub



        Public Sub PreencheDadosRelacionados_Todos(ByRef DocumentoInterno As BE.DocumentoInterno)
            Dim objDocumentoInterno As GcpBE900.GcpBEDocumentoInterno

            If IsNothing(DocumentoInterno) Then
                Throw New Exception("O objecto DocumentoInterno passado como parametro na função PreencheDadosRelacionados tem um valor null;")
                Exit Sub
            End If

            objDocumentoInterno = Me.GetGcpBEDocumentoInterno(DocumentoInterno)

            Try
                objMotor.Comercial.Internos.PreencheDadosRelacionados(objDocumentoInterno, GcpBE900.PreencheDados.enuDadosTodos)
            Catch ex As Exception
                objDocumentoInterno = Nothing
                Throw ex
                Exit Sub
            End Try

            DocumentoInterno = GetDocumentoInterno(objDocumentoInterno)
            objDocumentoInterno = Nothing
        End Sub



        Public Sub PreencheDadosRelacionados_Cliente(ByRef DocumentoInterno As BE.DocumentoInterno)
            Dim objDocumentoInterno As GcpBE900.GcpBEDocumentoInterno

            If IsNothing(DocumentoInterno) Then
                Throw New Exception("O objecto DocumentoInterno passado como parametro na função PreencheDadosRelacionados tem um valor null;")
                Exit Sub
            End If

            objDocumentoInterno = Me.GetGcpBEDocumentoInterno(DocumentoInterno)

            Try
                objMotor.Comercial.Internos.PreencheDadosRelacionados(objDocumentoInterno, GcpBE900.PreencheDados.enuDadosEntidade)
            Catch ex As Exception
                objDocumentoInterno = Nothing
                Throw ex
                Exit Sub
            End Try

            DocumentoInterno = GetDocumentoInterno(objDocumentoInterno)
            objDocumentoInterno = Nothing
        End Sub



        Public Sub PreencheDadosRelacionados_TipoDocumento(ByRef DocumentoInterno As BE.DocumentoInterno)
            Dim objDocumentoInterno As GcpBE900.GcpBEDocumentoInterno

            If IsNothing(DocumentoInterno) Then
                Throw New Exception("O objecto DocumentoInterno passado como parametro na função PreencheDadosRelacionados tem um valor null;")
                Exit Sub
            End If

            objDocumentoInterno = Me.GetGcpBEDocumentoInterno(DocumentoInterno)

            Try
                objMotor.Comercial.Internos.PreencheDadosRelacionados(objDocumentoInterno, GcpBE900.PreencheDados.enuDadosTipoDoc)

                ' ESTE BLOCO APENAS EXISTE PORQUE A PRIMAVERA NÃO ATRIBUI "ID" AO OBJETO DOCUMENTOINTERNO NO PREENCHIMENTO DE DADOS RELACIONADOS
                If objDocumentoInterno.ID.Trim.Length = 0 Then
                    objDocumentoInterno.ID = Guid.NewGuid.ToString("B")     'Este formato (B) coloca as chavetas curvas
                End If


            Catch ex As Exception
                objDocumentoInterno = Nothing
                Throw ex
                Exit Sub
            End Try

            DocumentoInterno = GetDocumentoInterno(objDocumentoInterno)
            objDocumentoInterno = Nothing
        End Sub



        Public Sub CalculaDataVencimento(ByRef DocumentoInterno As BE.DocumentoInterno)
            Dim objDocumentoInterno As GcpBE900.GcpBEDocumentoInterno

            If IsNothing(DocumentoInterno) Then
                Throw New Exception("O objecto DocumentoInterno passado como parametro na função CalculaDataVencimento tem um valor null;")
                Exit Sub
            End If

            objDocumentoInterno = Me.GetGcpBEDocumentoInterno(DocumentoInterno)

            Try
                With objDocumentoInterno
                    'NOTAR QUE PARA OBTER A DATA DE VENCIMENTO NUM DOCUMENTO INTERNO, FAREI USO DAS VENDAS
                    objDocumentoInterno.DataVencimento = objMotor.Comercial.Vendas.CalculaDataVencimento(.Data, .CondicaoPagamento)
                End With
            Catch ex As Exception
                objDocumentoInterno = Nothing
                Throw ex
                Exit Sub
            End Try

            DocumentoInterno = GetDocumentoInterno(objDocumentoInterno)
            objDocumentoInterno = Nothing
        End Sub



        Public Sub ActualizaDatasStock(ByRef DocumentoInterno As BE.DocumentoInterno)
            Dim objDocumentoInterno As BE.DocumentoInterno
            Dim i As System.Int16

            If IsNothing(DocumentoInterno) Then
                Throw New Exception("O objecto DocumentoInterno passado como parametro na função CalculaDataVencimento tem um valor null;")
                Exit Sub
            End If

            objDocumentoInterno = DocumentoInterno

            Try
                For i = 0 To objDocumentoInterno.Linhas.Count - 1
                    If objDocumentoInterno.Linhas(i).MovimentaStock = True Then
                        objDocumentoInterno.Linhas(i).DataStock = objDocumentoInterno.Data
                    End If
                Next
            Catch ex As Exception
                objDocumentoInterno = Nothing
                Throw ex
                Exit Sub
            End Try

            DocumentoInterno = objDocumentoInterno
            objDocumentoInterno = Nothing
        End Sub



        Public Sub CalculaValoresTotais(ByRef DocumentoInterno As BE.DocumentoInterno)
            Dim objDocumentoInterno As GcpBE900.GcpBEDocumentoInterno

            If IsNothing(DocumentoInterno) Then
                Throw New Exception("O objecto DocumentoInterno passado como parametro na função CalculaValoresTotais tem um valor null;")
                Exit Sub
            End If

            objDocumentoInterno = Me.GetGcpBEDocumentoInterno(DocumentoInterno)

            Try
                objMotor.Comercial.Internos.CalculaValoresTotais(objDocumentoInterno)
            Catch ex As Exception
                objDocumentoInterno = Nothing
                Throw ex
                Exit Sub
            End Try

            DocumentoInterno = GetDocumentoInterno(objDocumentoInterno)
            objDocumentoInterno = Nothing
        End Sub



        Public Function ValidaActualizacao(ByVal DocumentoInterno As BE.DocumentoInterno, ByRef info As String) As Boolean
            Dim strMensagem As String = ""
            Dim objDocumentoInterno As GcpBE900.GcpBEDocumentoInterno
            Dim i As System.Int32


            If IsNothing(DocumentoInterno) Then
                Throw New Exception("O objecto DocumentoInterno passado como parametro tem um valor null;")
                Exit Function
            End If

            With DocumentoInterno
                If (.TOTAL_Mercadoria - .TOTAL_Descontos + .TOTAL_Iva) < 0 Then
                    Throw New Exception("O documento interno não pode ter um total negativo;")
                    objDocumentoInterno = Nothing
                    Exit Function
                End If
            End With

            If Certifica(DocumentoInterno, strMensagem) = False Then
                Throw New Exception(strMensagem)
                Exit Function
            End If

            With DocumentoInterno
                If .EmModoEdicao = True Then
                    If Me.Existe(.Filial, .Documento_Tipo, .Documento_Serie, .Documento_Numero) = False Then
                        Throw New Exception("O DocumentoInterno que pretende actualizar não existe ou foi removido;")
                        Exit Function
                    End If
                End If
            End With

            With DocumentoInterno
                For i = .Linhas.Count - 1 To 0 Step -1
                    If .Linhas(i).TipoLinha = LinhaDocumentoInterno_Specs.TiposLinhas.Comentario_60 And .Linhas(i).Descricao.TrimEnd.Length = 0 Then
                        .Linhas.RemoveAt(i)
                    Else
                        Exit For
                    End If
                Next
            End With

            objDocumentoInterno = GetGcpBEDocumentoInterno(DocumentoInterno)

            Try
                If objDocumentoInterno.EmModoEdicao = False Then
                    'O NUMERADOR DO DOCUMENTO NOVO DEVE SER ATRIBUIDO SOMENTE NO ACTO DE GRAVAÇÃO DO DOCUMENTO
                    objDocumentoInterno.NumDoc = 0
                End If

                If Not objMotor.Comercial.Internos.ValidaActualizacao(objDocumentoInterno, strMensagem) Then
                    info = strMensagem
                    Return False
                Else
                    Return True
                End If

            Catch ex As Exception
                Throw ex
                Exit Function
            End Try

            objDocumentoInterno = Nothing
        End Function



        Public Sub Actualiza(ByRef DocumentoInterno As BE.DocumentoInterno)
            Dim strMensagem As String = ""
            Dim objDocumentoInterno As GcpBE900.GcpBEDocumentoInterno
            Dim i As System.Int32


            If IsNothing(DocumentoInterno) Then
                Throw New Exception("O objecto DocumentoInterno passado como parametro tem um valor null;")
                Exit Sub
            End If

            With DocumentoInterno
                If (.TOTAL_Mercadoria - .TOTAL_Descontos + .TOTAL_Iva) < 0 Then
                    Throw New Exception("O documento interno não pode ter um total negativo;")
                    objDocumentoInterno = Nothing
                    Exit Sub
                End If
            End With

            If Certifica(DocumentoInterno, strMensagem) = False Then
                Throw New Exception(strMensagem)
                Exit Sub
            End If

            With DocumentoInterno
                If .EmModoEdicao = True Then
                    If Me.Existe(.Filial, .Documento_Tipo, .Documento_Serie, .Documento_Numero) = False Then
                        Throw New Exception("O DocumentoInterno que pretende actualizar não existe ou foi removido;")
                        Exit Sub
                    End If
                End If
            End With

            With DocumentoInterno
                For i = .Linhas.Count - 1 To 0 Step -1
                    If .Linhas(i).TipoLinha = LinhaDocumentoInterno_Specs.TiposLinhas.Comentario_60 And .Linhas(i).Descricao.TrimEnd.Length = 0 Then
                        .Linhas.RemoveAt(i)
                    Else
                        Exit For
                    End If
                Next
            End With

            objDocumentoInterno = GetGcpBEDocumentoInterno(DocumentoInterno)

            Try
                If objDocumentoInterno.EmModoEdicao = False Then
                    'O NUMERADOR DO DOCUMENTO NOVO DEVE SER ATRIBUIDO SOMENTE NO ACTO DE GRAVAÇÃO DO DOCUMENTO
                    objDocumentoInterno.NumDoc = 0
                End If

                objMotor.Comercial.Internos.Actualiza(objDocumentoInterno)
            Catch ex As Exception
                Throw ex
                Exit Sub
            End Try

            objDocumentoInterno = Nothing
        End Sub



        Public Function Edita(ByVal Filial As String, ByVal Documento_Tipo As String, ByVal Documento_Serie As String, ByVal Documento_Numero As System.Int64) As BE.DocumentoInterno
            Dim objDocumentoInterno As BE.DocumentoInterno

            If Me.Existe(Filial, Documento_Tipo, Documento_Serie, Documento_Numero) = False Then
                Throw New Exception("O documento interno " & Filial & "/" & Documento_Tipo & "/" & Documento_Serie & "/" & Documento_Numero & " não existe na tabela CabecInternos, pelo que nao pode ser editado;")
                Exit Function
            End If

            Try
                objDocumentoInterno = GetDocumentoInterno(objMotor.Comercial.Internos.Edita(Documento_Tipo, Documento_Numero, Documento_Serie, Filial))
            Catch ex As Exception
                Throw ex
                Exit Function
            End Try

            Edita = objDocumentoInterno
            objDocumentoInterno = Nothing

        End Function



        Public Function Existe(ByVal Filial As String, ByVal Documento_Tipo As String, ByVal Documento_Serie As String, ByVal Documento_Numero As System.Int64) As Boolean

            If Len(Trim(Filial)) = 0 Then
                Throw New Exception("O parametro Filial da funcao Existe da classe DocumentoInterno deve conter uma string valida;")
                Exit Function
            End If

            If Len(Trim(Documento_Tipo)) = 0 Then
                Throw New Exception("O parametro Documento_Tipo da funcao Existe da classe DocumentoInterno deve conter uma string valida;")
                Exit Function
            End If

            If Len(Trim(Documento_Serie)) = 0 Then
                Throw New Exception("O parametro Documento_Serie da funcao Existe da classe DocumentoInterno deve conter uma string valida;")
                Exit Function
            End If

            If Documento_Numero = 0 Then
                Throw New Exception("O parametro NumeroDocumento da funcao Existe da classe DocumentoInterno deve conter im inteiro diferente de zero;")
                Exit Function
            End If

            Return objMotor.Comercial.Internos.Existe(Documento_Tipo, Documento_Numero, Documento_Serie, Filial)

        End Function



        Public Sub Remove(ByVal Filial As String, ByVal Documento_Tipo As String, ByVal Documento_Serie As String, ByVal Documento_Numero As System.Int64)
            Dim strMensagem As String = ""

            If Me.Existe(Filial, Documento_Tipo, Documento_Serie, Documento_Numero) = False Then
                Throw New Exception("O DocumentoInterno especificado não existe, pelo que não pode ser removido.")
                Exit Sub
            End If

            If objMotor.Comercial.Internos.ValidaRemocao(Documento_Tipo, Documento_Numero, Documento_Serie, Filial, strMensagem) = False Then
                Throw New Exception(strMensagem)
                Exit Sub
            End If

            Try
                objMotor.Comercial.Internos.Remove(Filial, Documento_Tipo, Documento_Serie, Documento_Numero)
            Catch ex As Exception
                Throw ex
                Exit Sub
            End Try

        End Sub



        Public Function Duplica(ByVal DocumentoInternoOrigem As BE.DocumentoInterno, _
                                ByVal Destino_Documento_Tipo As String, ByVal Destino_Documento_Serie As String, _
                                ByVal Destino_Entidade_Tipo As String, ByVal Destino_Entidade_Codigo As String) As BE.DocumentoInterno
            Dim objDocumentoInternoDestino As BE.DocumentoInterno
            Dim objLinhaDocumentoInternoDestino As BE.LinhaDocumentoInterno
            Dim i As System.Int32

            If IsNothing(DocumentoInternoOrigem) Then
                Throw New Exception("O argumento DocumentoInterno do metodo Duplica não pode ser nothing.")
                Exit Function
            End If

            objDocumentoInternoDestino = New BE.DocumentoInterno

            With objDocumentoInternoDestino
                .Id = Guid.NewGuid.ToString
                .Filial = DocumentoInternoOrigem.Filial
                .Documento_Tipo = Destino_Documento_Tipo
                .Documento_Serie = Destino_Documento_Serie
                .Documento_Numero = 0
                .Documento_Moeda = DocumentoInternoOrigem.Documento_Moeda
                .Documento_Cambio = DocumentoInternoOrigem.Documento_Cambio
                .Documento_CambioMoedaBase = DocumentoInternoOrigem.Documento_CambioMoedaBase
                .Documento_CambioMoedaAlternativa = DocumentoInternoOrigem.Documento_CambioMoedaAlternativa
                .Documento_Arredondamento = DocumentoInternoOrigem.Documento_Arredondamento
                .Documento_ArredondamentoIva = DocumentoInternoOrigem.Documento_ArredondamentoIva
                .ENTIDADE_Tipo = Destino_Entidade_Tipo
                .ENTIDADE_Codigo = Destino_Entidade_Codigo
                .ENTIDADE_Nome = DocumentoInternoOrigem.ENTIDADE_Nome
                .ENTIDADE_Morada = DocumentoInternoOrigem.ENTIDADE_Morada
                .ENTIDADE_Morada2 = DocumentoInternoOrigem.ENTIDADE_Morada2
                .ENTIDADE_Localidade = DocumentoInternoOrigem.ENTIDADE_Localidade
                .ENTIDADE_CodigoPostal = DocumentoInternoOrigem.ENTIDADE_CodigoPostal
                .ENTIDADE_LocalidadePostal = DocumentoInternoOrigem.ENTIDADE_LocalidadePostal
                .ENTIDADE_Contribuinte = DocumentoInternoOrigem.ENTIDADE_Contribuinte
                .ENTIDADE_Desconto = DocumentoInternoOrigem.ENTIDADE_Desconto
                .CondicaoPagamento = DocumentoInternoOrigem.CondicaoPagamento
                .ModoExpedicao = DocumentoInternoOrigem.ModoExpedicao
                .ModoPagamento = DocumentoInternoOrigem.ModoPagamento
                .Data = Now.Date
                .Vencimento = Now.Date
                .TOTAL_Mercadoria = DocumentoInternoOrigem.TOTAL_Mercadoria
                .TOTAL_Descontos = DocumentoInternoOrigem.TOTAL_Descontos
                .TOTAL_Iva = DocumentoInternoOrigem.TOTAL_Iva
                .TOTAL_Documento = DocumentoInternoOrigem.TOTAL_Documento
                .TRANSPORTE_Carga_Local = DocumentoInternoOrigem.TRANSPORTE_Carga_Local
                .TRANSPORTE_Carga_Data = DocumentoInternoOrigem.TRANSPORTE_Carga_Data
                .TRANSPORTE_Carga_Hora = DocumentoInternoOrigem.TRANSPORTE_Carga_Hora
                .TRANSPORTE_Descarga_Local = DocumentoInternoOrigem.TRANSPORTE_Descarga_Local
                .TRANSPORTE_Descarga_Data = DocumentoInternoOrigem.TRANSPORTE_Descarga_Data
                .TRANSPORTE_Descarga_Hora = DocumentoInternoOrigem.TRANSPORTE_Descarga_Hora
                .TRANSPORTE_Matricula = DocumentoInternoOrigem.TRANSPORTE_Matricula
                .Estado = DocumentoInternoOrigem.Estado
                .Observacoes = DocumentoInternoOrigem.Observacoes
                .Utilizador = DocumentoInternoOrigem.Utilizador
                .DataUltimaActualizacao = DocumentoInternoOrigem.DataUltimaActualizacao
                .CamposUtilizador = DocumentoInternoOrigem.CamposUtilizador
                .EmModoEdicao = False
            End With

            For i = 0 To DocumentoInternoOrigem.Linhas.Count - 1
                objLinhaDocumentoInternoDestino = New BE.LinhaDocumentoInterno
                With objLinhaDocumentoInternoDestino
                    .Armazem = DocumentoInternoOrigem.Linhas(i).Armazem
                    .Artigo = DocumentoInternoOrigem.Linhas(i).Artigo
                    .Lote = DocumentoInternoOrigem.Linhas(i).Lote
                    .CamposUtilizador = DocumentoInternoOrigem.Linhas(i).CamposUtilizador
                    .CodigoIva = DocumentoInternoOrigem.Linhas(i).CodigoIva
                    .DataEntrega = DocumentoInternoOrigem.Linhas(i).DataEntrega
                    .DataStock = DocumentoInternoOrigem.Linhas(i).DataStock
                    .Desconto1 = DocumentoInternoOrigem.Linhas(i).Desconto1
                    .Desconto2 = DocumentoInternoOrigem.Linhas(i).Desconto2
                    .Desconto3 = DocumentoInternoOrigem.Linhas(i).Desconto3
                    .Descricao = DocumentoInternoOrigem.Linhas(i).Descricao
                    .Id = Guid.NewGuid.ToString
                    .Localizacao = DocumentoInternoOrigem.Linhas(i).Localizacao
                    .MovimentaStock = DocumentoInternoOrigem.Linhas(i).MovimentaStock
                    .NumeroLinha = DocumentoInternoOrigem.Linhas(i).NumeroLinha
                    .PrecoLiquido = DocumentoInternoOrigem.Linhas(i).PrecoLiquido
                    .PrecoUnitario = DocumentoInternoOrigem.Linhas(i).PrecoUnitario
                    .PrecoMedioCusto = DocumentoInternoOrigem.Linhas(i).PrecoMedioCusto ' 20130803 adicionado tratamento PMC
                    .Quantidade = DocumentoInternoOrigem.Linhas(i).Quantidade
                    .QuantidadeSatisfeita = DocumentoInternoOrigem.Linhas(i).QuantidadeSatisfeita
                    .TaxaIva = DocumentoInternoOrigem.Linhas(i).TaxaIva
                    .TipoLinha = DocumentoInternoOrigem.Linhas(i).TipoLinha
                    .Unidade = DocumentoInternoOrigem.Linhas(i).Unidade
                    .PercentagemIncidenciaIva = DocumentoInternoOrigem.Linhas(i).PercentagemIncidenciaIva
                    .PercentagemIvaDedutivel = DocumentoInternoOrigem.Linhas(i).PercentagemIvaDedutivel
                    .IvaNaoDedutivel = DocumentoInternoOrigem.Linhas(i).IvaNaoDedutivel
                End With
                objDocumentoInternoDestino.Linhas.Add(objLinhaDocumentoInternoDestino, objLinhaDocumentoInternoDestino.Id)
                objLinhaDocumentoInternoDestino = Nothing
            Next

            Me.PreencheDadosRelacionados_Todos(objDocumentoInternoDestino)
            Me.CalculaDataVencimento(objDocumentoInternoDestino)
            Me.CalculaValoresTotais(objDocumentoInternoDestino)

            Duplica = objDocumentoInternoDestino

            objDocumentoInternoDestino = Nothing
        End Function



        Public Function Duplica(ByVal DocumentoInterno As BE.DocumentoInterno, _
                        ByVal DocumentoVenda_Tipo As String, ByVal DocumentoVenda_Serie As String, _
                        ByVal DocumentoVenda_Entidade_Tipo As String, ByVal DocumentoVenda_Entidade_Codigo As String, _
                        ByVal ConsideraCamposUtilizador As Boolean) As BE.DocumentoVenda

            Dim objDocumentosVenda As BS.DocumentosVenda
            Dim objDocumentoVenda As BE.DocumentoVenda
            Dim objLinhaDocumentoVenda As BE.LinhaDocumentoVenda
            Dim i As System.Int32

            If IsNothing(DocumentoInterno) Then
                Throw New Exception("O argumento DocumentoInterno do metodo Duplica (p/ documento venda) não pode ser nothing.")
                Exit Function
            End If

            objDocumentosVenda = New BS.DocumentosVenda(objMotor)
            objDocumentoVenda = New BE.DocumentoVenda

            ' Dados principais do documentoVenda
            With objDocumentoVenda
                .EmModoEdicao = False
                .Documento_Tipo = DocumentoVenda_Tipo
                .Documento_Serie = DocumentoVenda_Serie
                .Documento_Numero = 0
                .Filial = DocumentoInterno.Filial
                .Documento_Moeda = DocumentoInterno.Documento_Moeda
                .ENTIDADE_Tipo = DocumentoVenda_Entidade_Tipo
                .ENTIDADE_Codigo = DocumentoVenda_Entidade_Codigo
                .Data = Now.Date
            End With

            ' Usar o motor para preencher os restantes dados automaticamente
            objDocumentosVenda.PreencheDadosRelacionados_Todos(objDocumentoVenda)

            ' Prencher os restantes dados do documento
            With objDocumentoVenda
                .TOTAL_Mercadoria = DocumentoInterno.TOTAL_Mercadoria
                .TOTAL_Descontos = DocumentoInterno.TOTAL_Descontos
                .TOTAL_Iva = DocumentoInterno.TOTAL_Iva
                .TOTAL_Outros = 0   ' Não implementei o conceito de outros nos documentos internos
                .TOTAL_Documento = DocumentoInterno.TOTAL_Documento
                .Observacoes = DocumentoInterno.Observacoes
                .Utilizador = DocumentoInterno.Utilizador
                .DataUltimaActualizacao = Now.Date

                If ConsideraCamposUtilizador = True Then
                    .CamposUtilizador = DocumentoInterno.CamposUtilizador
                End If

            End With


            ' Copiar as linhas
            For i = 0 To DocumentoInterno.Linhas.Count - 1
                objLinhaDocumentoVenda = New BE.LinhaDocumentoVenda
                With objLinhaDocumentoVenda
                    .Id = Guid.NewGuid.ToString("B")
                    .Armazem = DocumentoInterno.Linhas(i).Armazem
                    .Artigo = DocumentoInterno.Linhas(i).Artigo
                    .Lote = DocumentoInterno.Linhas(i).Lote
                    .CodigoIva = DocumentoInterno.Linhas(i).CodigoIva
                    .DataEntrega = Now.Date     'DocumentoInterno.Linhas(i).DataEntrega
                    .DataStock = Now.Date       'DocumentoInterno.Linhas(i).DataStock
                    .Desconto1 = DocumentoInterno.Linhas(i).Desconto1
                    .Desconto2 = DocumentoInterno.Linhas(i).Desconto2
                    .Desconto3 = DocumentoInterno.Linhas(i).Desconto3
                    .Descricao = DocumentoInterno.Linhas(i).Descricao
                    .Localizacao = DocumentoInterno.Linhas(i).Localizacao
                    .MovimentaStock = DocumentoInterno.Linhas(i).MovimentaStock
                    .NumeroLinha = DocumentoInterno.Linhas(i).NumeroLinha
                    .PrecoLiquido = DocumentoInterno.Linhas(i).PrecoLiquido
                    .PrecoUnitario = DocumentoInterno.Linhas(i).PrecoUnitario
                    .PrecoMedioCusto = DocumentoInterno.Linhas(i).PrecoMedioCusto   ' 20130803 adicionado tratamento PMC
                    .Quantidade = DocumentoInterno.Linhas(i).Quantidade
                    .QuantidadeSatisfeita = DocumentoInterno.Linhas(i).QuantidadeSatisfeita
                    .TaxaIva = DocumentoInterno.Linhas(i).TaxaIva
                    .TipoLinha = DocumentoInterno.Linhas(i).TipoLinha
                    .Unidade = DocumentoInterno.Linhas(i).Unidade
                    .PercentagemIncidenciaIva = DocumentoInterno.Linhas(i).PercentagemIncidenciaIva
                    .PercentagemIvaDedutivel = DocumentoInterno.Linhas(i).PercentagemIvaDedutivel
                    .IvaNaoDedutivel = DocumentoInterno.Linhas(i).IvaNaoDedutivel

                    If ConsideraCamposUtilizador = True Then
                        .CamposUtilizador = DocumentoInterno.Linhas(i).CamposUtilizador
                    End If

                End With

                objDocumentoVenda.Linhas.Add(objLinhaDocumentoVenda, objLinhaDocumentoVenda.Id)
                objLinhaDocumentoVenda = Nothing
            Next

            objDocumentosVenda.CalculaValoresTotais(objDocumentoVenda)

            Duplica = objDocumentoVenda

            objDocumentoVenda = Nothing
            objDocumentosVenda = Nothing
        End Function



        Private Function GetDocumentoInterno(ByVal DocumentoInterno As GcpBE900.GcpBEDocumentoInterno) As BE.DocumentoInterno
            Dim objDocumentoInterno As New BE.DocumentoInterno
            Dim objLinhaDocumentoInterno As BE.LinhaDocumentoInterno
            Dim objCampoUtilizador As BE.CampoUtilizador = Nothing
            Dim i As System.Int32 = 0
            Dim ii As System.Int32 = 0

            With objDocumentoInterno
                .Id = DocumentoInterno.ID
                .Filial = DocumentoInterno.Filial
                .Documento_Tipo = DocumentoInterno.Tipodoc
                .Documento_Serie = DocumentoInterno.Serie
                .Documento_Numero = DocumentoInterno.NumDoc
                .Documento_Moeda = DocumentoInterno.Moeda
                .Documento_Cambio = DocumentoInterno.Cambio
                .Documento_CambioMoedaBase = DocumentoInterno.CambioMBase
                .Documento_CambioMoedaAlternativa = DocumentoInterno.CambioMAlt
                .Documento_Arredondamento = DocumentoInterno.Arredondamento
                .Documento_ArredondamentoIva = DocumentoInterno.ArredondamentoIva
                .ENTIDADE_Tipo = DocumentoInterno.TipoEntidade
                .ENTIDADE_Codigo = DocumentoInterno.Entidade
                .ENTIDADE_Nome = DocumentoInterno.Nome
                .ENTIDADE_Morada = DocumentoInterno.Morada
                .ENTIDADE_Morada2 = DocumentoInterno.Morada2
                .ENTIDADE_Localidade = DocumentoInterno.Localidade
                .ENTIDADE_CodigoPostal = DocumentoInterno.CodPostal
                .ENTIDADE_LocalidadePostal = DocumentoInterno.CodPostalLocalidade
                .ENTIDADE_Contribuinte = DocumentoInterno.NumContribuinte
                .ENTIDADE_Desconto = DocumentoInterno.DescEntidade
                .CondicaoPagamento = DocumentoInterno.CondicaoPagamento
                .ModoExpedicao = DocumentoInterno.ModoExpedicao
                .ModoPagamento = DocumentoInterno.ModoPagamento
                .Data = DocumentoInterno.Data
                .Vencimento = DocumentoInterno.DataVencimento
                .RegimeIva = DocumentoInterno.RegimeIva
                .TOTAL_Mercadoria = DocumentoInterno.TotalMercadoria
                .TOTAL_Descontos = DocumentoInterno.TotalDesconto
                .TOTAL_Iva = DocumentoInterno.TotalIva
                .TOTAL_Documento = DocumentoInterno.TotalDocumento
                .TRANSPORTE_Carga_Local = DocumentoInterno.LocalCarga
                .TRANSPORTE_Carga_Data = DocumentoInterno.DataCarga
                .TRANSPORTE_Carga_Hora = DocumentoInterno.HoraCarga
                .TRANSPORTE_Descarga_Local = DocumentoInterno.LocalDescarga
                .TRANSPORTE_Descarga_Data = DocumentoInterno.DataDescarga
                .TRANSPORTE_Descarga_Hora = DocumentoInterno.HoraDescarga
                .TRANSPORTE_Matricula = DocumentoInterno.Matricula
                .Estado = DocumentoInterno.Estado
                .Observacoes = DocumentoInterno.Observacoes
                .Utilizador = DocumentoInterno.Utilizador
                .DataUltimaActualizacao = DocumentoInterno.DataUltimaActualizacao
                .EmModoEdicao = DocumentoInterno.EmModoEdicao

                For i = 1 To DocumentoInterno.Linhas.NumItens
                    objLinhaDocumentoInterno = New BE.LinhaDocumentoInterno
                    With objLinhaDocumentoInterno

                        .Id = DocumentoInterno.Linhas(i).ID
                        .NumeroLinha = i

                        Select Case DocumentoInterno.Linhas(i).TipoLinha

                            Case "60"
                                .TipoLinha = LinhaDocumentoInterno_Specs.TiposLinhas.Comentario_60
                                .Descricao = DocumentoInterno.Linhas(i).Descricao

                            Case "30"
                                .TipoLinha = LinhaDocumentoInterno_Specs.TiposLinhas.Acerto_30
                                .Descricao = DocumentoInterno.Linhas(i).Descricao
                                .Quantidade = DocumentoInterno.Linhas(i).Quantidade
                                .PrecoUnitario = DocumentoInterno.Linhas(i).PrecoUnitario
                                .PrecoLiquido = DocumentoInterno.Linhas(i).PrecoLiquido

                            Case "50"
                                .TipoLinha = LinhaDocumentoInterno_Specs.TiposLinhas.Portes_50
                                .Descricao = DocumentoInterno.Linhas(i).Descricao
                                .Quantidade = DocumentoInterno.Linhas(i).Quantidade
                                .PrecoUnitario = DocumentoInterno.Linhas(i).PrecoUnitario
                                .CodigoIva = DocumentoInterno.Linhas(i).CodigoIva
                                .TaxaIva = DocumentoInterno.Linhas(i).TaxaIva
                                .PrecoLiquido = DocumentoInterno.Linhas(i).PrecoLiquido

                            Case Else
                                Select Case DocumentoInterno.Linhas(i).TipoLinha
                                    Case "10"
                                        .TipoLinha = LinhaDocumentoInterno_Specs.TiposLinhas.Mercadoria_TipoArtigo_3_TipoLinha_10
                                    Case "20"
                                        .TipoLinha = LinhaDocumentoInterno_Specs.TiposLinhas.ServicoA_TipoArtigo_0_TipoLinha_20
                                    Case "13"
                                        .TipoLinha = LinhaDocumentoInterno_Specs.TiposLinhas.MateriaPrima_TipoArtigo_6_TipoLinha_13
                                    Case "14"
                                        .TipoLinha = LinhaDocumentoInterno_Specs.TiposLinhas.MateriaSubsidiaria_TipoArtigo_7_TipoLinha_14
                                    Case "23"
                                        .TipoLinha = LinhaDocumentoInterno_Specs.TiposLinhas.MaoDeObra_TipoArtigo_12_TipoLinha_23
                                End Select
                                .Artigo = DocumentoInterno.Linhas(i).Artigo
                                If DocumentoInterno.Linhas(i).MovimentaStock = True Then
                                    .MovimentaStock = True
                                    .Armazem = DocumentoInterno.Linhas(i).Armazem
                                    .Localizacao = DocumentoInterno.Linhas(i).Localizacao
                                    If DocumentoInterno.Linhas(i).Lote.Contains("<") Then
                                        .Lote = ""
                                    Else
                                        .Lote = DocumentoInterno.Linhas(i).Lote
                                    End If
                                Else
                                    .MovimentaStock = False
                                    .Armazem = ""
                                    .Localizacao = ""
                                    .Lote = ""
                                End If
                                .Quantidade = DocumentoInterno.Linhas(i).Quantidade
                                .QuantidadeSatisfeita = DocumentoInterno.Linhas(i).QntSatisfeita
                                .Unidade = DocumentoInterno.Linhas(i).Unidade
                                .DataEntrega = DocumentoInterno.Linhas(i).DataEntrega
                                .DataStock = DocumentoInterno.Linhas(i).DataDocStock
                                .Descricao = DocumentoInterno.Linhas(i).Descricao
                                .CodigoIva = DocumentoInterno.Linhas(i).CodigoIva
                                .TaxaIva = DocumentoInterno.Linhas(i).TaxaIva
                                .PrecoUnitario = DocumentoInterno.Linhas(i).PrecoUnitario
                                ' 20130803 - PMC (unitário) é o CustoPreviso com o TipoCustoPrevisto = 0 (0=PMC, 1=PCUltimo, 2 PCPadrao)
                                .PrecoMedioCusto = DocumentoInterno.Linhas(i).CustoPrevisto
                                .Desconto1 = DocumentoInterno.Linhas(i).Desconto1
                                .Desconto2 = DocumentoInterno.Linhas(i).Desconto2
                                .Desconto3 = DocumentoInterno.Linhas(i).Desconto3

                                ' ERRO DA PRIMAVERA PORQUE GRAVA UM PRECO LIQUIDO COM UM VALOR 0.000000000234323
                                If DocumentoInterno.Linhas(i).PrecoLiquido < 0.001 Then
                                    .PrecoLiquido = 0
                                Else
                                    .PrecoLiquido = DocumentoInterno.Linhas(i).PrecoLiquido
                                End If

                                .PercentagemIncidenciaIva = DocumentoInterno.Linhas(i).PercIncidenciaIVA
                                .PercentagemIvaDedutivel = DocumentoInterno.Linhas(i).PercIvaDedutivel
                                .IvaNaoDedutivel = DocumentoInterno.Linhas(i).IvaNaoDedutivel
                                .RegimeIva = DocumentoInterno.Linhas(i).RegimeIva
                                .TaxaProRata = DocumentoInterno.Linhas(i).TaxaProRata
                                .ModuloOrigemCopia = DocumentoInterno.Linhas(i).ModuloOrigemCopia
                                .IdLinhaOrigemCopia = DocumentoInterno.Linhas(i).IdLinhaOrigemCopia

                        End Select
                    End With

                    If DocumentoInterno.Linhas(i).CamposUtil.NumItens > 0 Then
                        For ii = 1 To DocumentoInterno.Linhas(i).CamposUtil.NumItens
                            If Not IsDBNull(DocumentoInterno.Linhas(i).CamposUtil(ii)) Then
                                objCampoUtilizador = New BE.CampoUtilizador
                                With objCampoUtilizador
                                    Select Case DocumentoInterno.Linhas(i).CamposUtil(ii).TipoSimplificado
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
                                            Throw New Exception("Foi referenciado um tipo de dados em DocumentoInterno.linhas(i).CamposUtil não esperado na classe Business.Internos.DocumentoInterno função GetDocumentoInterno")
                                            Exit Function

                                    End Select
                                    .Campo = DocumentoInterno.Linhas(i).CamposUtil(ii).Nome
                                    If Not IsDBNull(DocumentoInterno.Linhas(i).CamposUtil(ii).Valor) Then
                                        .Valor = CType(DocumentoInterno.Linhas(i).CamposUtil(ii).Valor, Object)
                                    Else
                                        .Valor = CType("", Object)
                                    End If

                                End With
                                objLinhaDocumentoInterno.CamposUtilizador.Add(objCampoUtilizador, objCampoUtilizador.Campo)
                                objCampoUtilizador = Nothing
                            End If
                        Next
                    End If

                    objDocumentoInterno.Linhas.Add(objLinhaDocumentoInterno, objLinhaDocumentoInterno.Id)
                    objLinhaDocumentoInterno = Nothing
                Next

                For i = 1 To DocumentoInterno.CamposUtil.NumItens
                    If Not IsDBNull(DocumentoInterno.CamposUtil(i)) Then
                        objCampoUtilizador = New BE.CampoUtilizador
                        With objCampoUtilizador
                            Select Case DocumentoInterno.CamposUtil(i).TipoSimplificado
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
                                    Throw New Exception("Foi referenciado um tipo de dados em DocumentoInterno.Linhas.CamposUtil não esperado na classe Business.DocumentoInterno função GetDocumentoInterno")
                                    Exit Function

                            End Select
                            .Campo = DocumentoInterno.CamposUtil(i).Nome
                            If Not IsDBNull(DocumentoInterno.CamposUtil(i).Valor) Then
                                .Valor = CType(DocumentoInterno.CamposUtil(i).Valor, Object)
                            Else
                                .Valor = CType("", Object)
                            End If

                        End With
                        .CamposUtilizador.Add(objCampoUtilizador, objCampoUtilizador.Campo)
                        objCampoUtilizador = Nothing
                    End If
                Next

            End With

            GetDocumentoInterno = objDocumentoInterno
            objDocumentoInterno = Nothing

        End Function



        Private Function GetGcpBEDocumentoInterno(ByVal DocumentoInterno As BE.DocumentoInterno) As GcpBE900.GcpBEDocumentoInterno
            Dim objDocumentoInterno As New GcpBE900.GcpBEDocumentoInterno
            Dim objLinhaDocumentoInterno As GcpBE900.GcpBELinhaDocumentoInterno
            Dim i As System.Int32 = 0
            Dim ii As System.Int32 = 0

            If DocumentoInterno.EmModoEdicao = True Then 'EM EDIÇÃO
                With DocumentoInterno
                    objDocumentoInterno = objMotor.Comercial.Internos.Edita(.Documento_Tipo, .Documento_Numero, .Documento_Serie, .Filial)
                End With
            End If

            With objDocumentoInterno
                .ID = DocumentoInterno.Id
                .Filial = DocumentoInterno.Filial
                .Tipodoc = DocumentoInterno.Documento_Tipo
                .Serie = DocumentoInterno.Documento_Serie
                .Moeda = DocumentoInterno.Documento_Moeda
                .Cambio = DocumentoInterno.Documento_Cambio
                .CambioMBase = DocumentoInterno.Documento_CambioMoedaBase
                .CambioMAlt = DocumentoInterno.Documento_CambioMoedaAlternativa
                .Arredondamento = DocumentoInterno.Documento_Arredondamento
                .ArredondamentoIva = DocumentoInterno.Documento_ArredondamentoIva
                .TipoEntidade = DocumentoInterno.ENTIDADE_Tipo
                .Entidade = DocumentoInterno.ENTIDADE_Codigo
                .NumDoc = DocumentoInterno.Documento_Numero
                .Nome = DocumentoInterno.ENTIDADE_Nome
                .Morada = DocumentoInterno.ENTIDADE_Morada
                .Morada2 = DocumentoInterno.ENTIDADE_Morada2
                .Localidade = DocumentoInterno.ENTIDADE_Localidade
                .CodPostal = DocumentoInterno.ENTIDADE_CodigoPostal
                .CodPostalLocalidade = DocumentoInterno.ENTIDADE_LocalidadePostal
                .NumContribuinte = DocumentoInterno.ENTIDADE_Contribuinte
                .DescEntidade = DocumentoInterno.ENTIDADE_Desconto
                .CondicaoPagamento = DocumentoInterno.CondicaoPagamento
                .ModoExpedicao = DocumentoInterno.ModoExpedicao
                .ModoPagamento = DocumentoInterno.ModoPagamento
                .Data = DocumentoInterno.Data
                .DataVencimento = DocumentoInterno.Vencimento
                .RegimeIva = DocumentoInterno.RegimeIva
                .TotalMercadoria = DocumentoInterno.TOTAL_Mercadoria
                .TotalDesconto = DocumentoInterno.TOTAL_Descontos
                .TotalIva = DocumentoInterno.TOTAL_Iva
                .TotalDocumento = DocumentoInterno.TOTAL_Documento
                .LocalCarga = DocumentoInterno.TRANSPORTE_Carga_Local
                .DataCarga = DocumentoInterno.TRANSPORTE_Carga_Data
                .HoraCarga = DocumentoInterno.TRANSPORTE_Carga_Hora
                .LocalDescarga = DocumentoInterno.TRANSPORTE_Descarga_Local
                .DataDescarga = DocumentoInterno.TRANSPORTE_Descarga_Data
                .HoraDescarga = DocumentoInterno.TRANSPORTE_Descarga_Hora
                .Matricula = DocumentoInterno.TRANSPORTE_Matricula
                .Estado = DocumentoInterno.Estado
                .Observacoes = DocumentoInterno.Observacoes
                .Utilizador = DocumentoInterno.Utilizador
                .DataUltimaActualizacao = DocumentoInterno.DataUltimaActualizacao
                .EmModoEdicao = DocumentoInterno.EmModoEdicao
            End With

            objDocumentoInterno.Linhas.RemoveTodos()

            With objDocumentoInterno

                For i = 0 To DocumentoInterno.Linhas.Count - 1
                    objLinhaDocumentoInterno = New GcpBE900.GcpBELinhaDocumentoInterno
                    With objLinhaDocumentoInterno

                        .ID = DocumentoInterno.Linhas(i).Id

                        Select Case DocumentoInterno.Linhas(i).TipoLinha

                            Case LinhaDocumentoInterno_Specs.TiposLinhas.Comentario_60
                                .TipoLinha = "60"
                                .Descricao = DocumentoInterno.Linhas(i).Descricao

                            Case LinhaDocumentoInterno_Specs.TiposLinhas.Acerto_30
                                .TipoLinha = "30"
                                .Descricao = DocumentoInterno.Linhas(i).Descricao
                                .Quantidade = DocumentoInterno.Linhas(i).Quantidade
                                .PrecoUnitario = DocumentoInterno.Linhas(i).PrecoUnitario
                                .PrecoLiquido = DocumentoInterno.Linhas(i).PrecoLiquido

                            Case LinhaDocumentoInterno_Specs.TiposLinhas.Portes_50
                                .TipoLinha = "50"
                                .Descricao = DocumentoInterno.Linhas(i).Descricao
                                .Quantidade = DocumentoInterno.Linhas(i).Quantidade
                                .PrecoUnitario = DocumentoInterno.Linhas(i).PrecoUnitario
                                .CodigoIva = DocumentoInterno.Linhas(i).CodigoIva
                                .TaxaIva = DocumentoInterno.Linhas(i).TaxaIva
                                .PrecoLiquido = DocumentoInterno.Linhas(i).PrecoLiquido

                            Case Else
                                Select Case DocumentoInterno.Linhas(i).TipoLinha
                                    Case LinhaDocumentoInterno_Specs.TiposLinhas.Mercadoria_TipoArtigo_3_TipoLinha_10
                                        .TipoLinha = "10"
                                    Case LinhaDocumentoInterno_Specs.TiposLinhas.ServicoA_TipoArtigo_0_TipoLinha_20
                                        .TipoLinha = "20"
                                    Case LinhaDocumentoInterno_Specs.TiposLinhas.MateriaPrima_TipoArtigo_6_TipoLinha_13
                                        .TipoLinha = "13"
                                    Case LinhaDocumentoInterno_Specs.TiposLinhas.MateriaSubsidiaria_TipoArtigo_7_TipoLinha_14
                                        .TipoLinha = "14"
                                    Case LinhaDocumentoInterno_Specs.TiposLinhas.MaoDeObra_TipoArtigo_12_TipoLinha_23
                                        .TipoLinha = "23"
                                End Select
                                .Artigo = DocumentoInterno.Linhas(i).Artigo
                                If DocumentoInterno.Linhas(i).MovimentaStock = True Then
                                    .MovimentaStock = True
                                    .Armazem = DocumentoInterno.Linhas(i).Armazem
                                    .Localizacao = DocumentoInterno.Linhas(i).Localizacao
                                    .Lote = DocumentoInterno.Linhas(i).Lote
                                Else
                                    .MovimentaStock = False
                                    .Armazem = ""
                                    .Localizacao = ""
                                    .Lote = ""
                                End If
                                .Quantidade = DocumentoInterno.Linhas(i).Quantidade
                                .QntSatisfeita = DocumentoInterno.Linhas(i).QuantidadeSatisfeita
                                .Unidade = DocumentoInterno.Linhas(i).Unidade
                                .DataEntrega = DocumentoInterno.Linhas(i).DataEntrega

                                ' CONSIDERO ESTA DATA IGUAL À DATASTOCK
                                .Data = DocumentoInterno.Linhas(i).DataStock

                                .DataDocStock = DocumentoInterno.Linhas(i).DataStock
                                .Descricao = DocumentoInterno.Linhas(i).Descricao
                                .CodigoIva = DocumentoInterno.Linhas(i).CodigoIva
                                .TaxaIva = DocumentoInterno.Linhas(i).TaxaIva
                                .PrecoUnitario = DocumentoInterno.Linhas(i).PrecoUnitario
                                .TipoCustoPrevisto = 0                                      '20130803  (0=PMC, 1=PCUltimo, 2 PCPadrao)
                                .CustoPrevisto = DocumentoInterno.Linhas(i).PrecoMedioCusto '20130803  CustoPreviso = PMC (unitário) com o TipoCustoPreviso = 0
                                .Desconto1 = DocumentoInterno.Linhas(i).Desconto1
                                .Desconto2 = DocumentoInterno.Linhas(i).Desconto2
                                .Desconto3 = DocumentoInterno.Linhas(i).Desconto3
                                .PrecoLiquido = DocumentoInterno.Linhas(i).PrecoLiquido
                                .PercIncidenciaIVA = DocumentoInterno.Linhas(i).PercentagemIncidenciaIva
                                .PercIvaDedutivel = DocumentoInterno.Linhas(i).PercentagemIvaDedutivel
                                .IvaNaoDedutivel = DocumentoInterno.Linhas(i).IvaNaoDedutivel
                                .RegimeIva = DocumentoInterno.Linhas(i).RegimeIva
                                .TaxaProRata = DocumentoInterno.Linhas(i).TaxaProRata
                                .ModuloOrigemCopia = DocumentoInterno.Linhas(i).ModuloOrigemCopia
                                .IdLinhaOrigemCopia = DocumentoInterno.Linhas(i).IdLinhaOrigemCopia

                                ' ESTA FLAG É OBRIGATÓRIA NOS DOCUMENTOS INTERNOS PARA QUE O VALOR DA LINHA SEJA CONSIDERADA NO CALCULO DOS TOTAIS DO DOCUMENTO
                                .ContabilizaTotais = True

                        End Select
                    End With


                    If DocumentoInterno.Linhas(i).CamposUtilizador.Count > 0 Then
                        For ii = 0 To DocumentoInterno.Linhas(i).CamposUtilizador.Count - 1

                            Select Case DocumentoInterno.Linhas(i).CamposUtilizador(ii).Tipo
                                Case BE.CampoUtilizador.TiposDadosCampoUtilizador._String
                                    objLinhaDocumentoInterno.CamposUtil.Item(DocumentoInterno.Linhas(i).CamposUtilizador(ii).Campo).Valor = CType(DocumentoInterno.Linhas(i).CamposUtilizador(ii).Valor, String)
                                    objLinhaDocumentoInterno.CamposUtil.Item(DocumentoInterno.Linhas(i).CamposUtilizador(ii).Campo).Tipo = StdBE900.EnumTipoCampo.tcNVarchar

                                Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Integer
                                    objLinhaDocumentoInterno.CamposUtil.Item(DocumentoInterno.Linhas(i).CamposUtilizador(ii).Campo).Valor = CType(DocumentoInterno.Linhas(i).CamposUtilizador(ii).Valor, System.Int32)
                                    objLinhaDocumentoInterno.CamposUtil.Item(DocumentoInterno.Linhas(i).CamposUtilizador(ii).Campo).Tipo = StdBE900.EnumTipoCampo.tcInt

                                Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Money
                                    objLinhaDocumentoInterno.CamposUtil.Item(DocumentoInterno.Linhas(i).CamposUtilizador(ii).Campo).Valor = CType(DocumentoInterno.Linhas(i).CamposUtilizador(ii).Valor, System.Decimal)
                                    objLinhaDocumentoInterno.CamposUtil.Item(DocumentoInterno.Linhas(i).CamposUtilizador(ii).Campo).Tipo = StdBE900.EnumTipoCampo.tcMoney

                                Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Date
                                    objLinhaDocumentoInterno.CamposUtil.Item(DocumentoInterno.Linhas(i).CamposUtilizador(ii).Campo).Valor = CType(DocumentoInterno.Linhas(i).CamposUtilizador(ii).Valor, System.DateTime)
                                    objLinhaDocumentoInterno.CamposUtil.Item(DocumentoInterno.Linhas(i).CamposUtilizador(ii).Campo).Tipo = StdBE900.EnumTipoCampo.tcDateTime

                                Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Boolean
                                    objLinhaDocumentoInterno.CamposUtil.Item(DocumentoInterno.Linhas(i).CamposUtilizador(ii).Campo).Valor = CType(DocumentoInterno.Linhas(i).CamposUtilizador(ii).Valor, Boolean)
                                    objLinhaDocumentoInterno.CamposUtil.Item(DocumentoInterno.Linhas(i).CamposUtilizador(ii).Campo).Tipo = StdBE900.EnumTipoCampo.tcBit

                                Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Double
                                    objLinhaDocumentoInterno.CamposUtil.Item(DocumentoInterno.Linhas(i).CamposUtilizador(ii).Campo).Valor = CType(DocumentoInterno.Linhas(i).CamposUtilizador(ii).Valor, Double)
                                    objLinhaDocumentoInterno.CamposUtil.Item(DocumentoInterno.Linhas(i).CamposUtilizador(ii).Campo).Tipo = StdBE900.EnumTipoCampo.tcDecimal


                                Case Else
                                    Throw New Exception("Foi referenciado um tipo de dados em DocumentoInterno.Linhas(i).CamposUtil não esperado na classe Business.DocumentoInterno função GetGcpBEDocumentoInterno")
                                    Exit Function

                            End Select

                        Next
                    End If

                    .Linhas.Insere(objLinhaDocumentoInterno)
                    objLinhaDocumentoInterno = Nothing
                Next


                If DocumentoInterno.CamposUtilizador.Count > 0 Then
                    For i = 0 To DocumentoInterno.CamposUtilizador.Count - 1

                        Select Case DocumentoInterno.CamposUtilizador(i).Tipo
                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._String
                                .CamposUtil.Item(DocumentoInterno.CamposUtilizador(i).Campo).Valor = CType(DocumentoInterno.CamposUtilizador(i).Valor, String)
                                .CamposUtil.Item(DocumentoInterno.CamposUtilizador(i).Campo).Tipo = StdBE900.EnumTipoCampo.tcNVarchar

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Integer
                                .CamposUtil.Item(DocumentoInterno.CamposUtilizador(i).Campo).Valor = CType(DocumentoInterno.CamposUtilizador(i).Valor, System.Int32)
                                .CamposUtil.Item(DocumentoInterno.CamposUtilizador(i).Campo).Tipo = StdBE900.EnumTipoCampo.tcInt

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Money
                                .CamposUtil.Item(DocumentoInterno.CamposUtilizador(i).Campo).Valor = CType(DocumentoInterno.CamposUtilizador(i).Valor, System.Decimal)
                                .CamposUtil.Item(DocumentoInterno.CamposUtilizador(i).Campo).Tipo = StdBE900.EnumTipoCampo.tcMoney

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Date
                                .CamposUtil.Item(DocumentoInterno.CamposUtilizador(i).Campo).Valor = CType(DocumentoInterno.CamposUtilizador(i).Valor, System.DateTime)
                                .CamposUtil.Item(DocumentoInterno.CamposUtilizador(i).Campo).Tipo = StdBE900.EnumTipoCampo.tcDateTime

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Boolean
                                .CamposUtil.Item(DocumentoInterno.CamposUtilizador(i).Campo).Valor = CType(DocumentoInterno.CamposUtilizador(i).Valor, Boolean)
                                .CamposUtil.Item(DocumentoInterno.CamposUtilizador(i).Campo).Tipo = StdBE900.EnumTipoCampo.tcBit

                            Case BE.CampoUtilizador.TiposDadosCampoUtilizador._Double
                                .CamposUtil.Item(DocumentoInterno.CamposUtilizador(i).Campo).Valor = CType(DocumentoInterno.CamposUtilizador(i).Valor, Double)
                                .CamposUtil.Item(DocumentoInterno.CamposUtilizador(i).Campo).Tipo = StdBE900.EnumTipoCampo.tcDecimal


                            Case Else
                                Throw New Exception("Foi referenciado um tipo de dados em DocumentoInterno.CamposUtil não esperado na classe Business.DocumentoInterno função GetGcpBEDocumentoInterno")
                                Exit Function

                        End Select

                    Next
                End If
            End With

            GetGcpBEDocumentoInterno = objDocumentoInterno
            objDocumentoInterno = Nothing

        End Function



        Protected Overrides Sub Finalize()
            If Not IsNothing(objMotor) Then
                objMotor = Nothing
            End If
            MyBase.Finalize()
        End Sub



    End Class

End Namespace
