Imports Microsoft.VisualBasic
Imports CCUtils
Imports System.Data.SqlClient
Imports System.Text
Imports CCUtils.CCSQLServer
Imports Interop



Namespace BS

    Public Class ExtratosCCT
        Private objMotor As ErpBS900.ErpBS
        Private connectionString As String


        Sub New(ByRef Motor As ErpBS900.ErpBS, ByVal connectionString As String)
            objMotor = Motor
            Me.connectionString = connectionString
        End Sub



        Public Function List(ByVal entidade As String, ByVal dataInicial As Date, ByVal dataFinal As Date) As System.Collections.Generic.List(Of BE.LinhaExtratoCCT)
            Dim query As String
            Dim mList As System.Collections.Generic.List(Of BE.LinhaExtratoCCT)
            Dim mTbl As System.Data.DataTable
            Dim objSQL As CCSQLServer
            Dim objLinhaExtrato As BE.LinhaExtratoCCT
            Dim parameters As List(Of SqlParameter)
            Dim sqlParameter As System.Data.SqlClient.SqlParameter
            Dim i As System.Int16

            ' pré-validações
            If (entidade.Trim().Length = 0) Then
                Throw New Exception("O argumento entidade do método ExtratosCCT.List() não pode ser vazio!")
            End If

            If (IsNothing(dataInicial) Or IsNothing(dataFinal)) Then
                Throw New Exception("Os argumentos que delimitam as datas do extrato em ExtratosCCT.List() devem estar preenchidas!")
            End If

            If (dataInicial > dataFinal) Then
                Throw New Exception("Os argumentos que delimitam as datas do extrato em ExtratosCCT.List() devem delimitar um periodo temporal válido!")
            End If


            query = <sql><![CDATA[
--DECLARE @entidade AS NVARCHAR(12), @dataInicial AS DATETIME, @dataFinal AS DATETIME;
--SET @entidade = '115119';
--SET @dataInicial = DATEADD(YEAR, -2, GETDATE());
--SET @dataFinal = GETDATE();
  
SELECT * 
FROM
	(
		-- ESTA QUERY CRIA UM EXTRATO COM RUNNING TOTAL À LINHA
		SELECT CAST(h.DataExtracto AS DATE) AS Data, h.TipoDoc, h.TipoDocumento, h.Serie, h.NumDoc, h.Total,
		-- RUNNING TOTAL
		ROUND(SUM(h.Total) OVER (ORDER BY h.DataExtracto ASC ROWS UNBOUNDED PRECEDING), 2) AS Saldo

		FROM 
			(
			    -- FILTRA SOMENTE OS DOCUMENTOS DE VENDA
				SELECT h.DataExtracto, h.TipoDoc, 'V' + RTRIM(dv.TipoDocumento) AS TipoDocumento, h.Serie, h.NumDoc, h.ValorTotal AS Total
				FROM Historico h
					INNER JOIN DocumentosVenda dv
						ON h.TipoDoc = dv.Documento
					LEFT JOIN Pendentes p
						ON h.Id = p.IdHistorico
				WHERE h.Modulo = 'V'
					  AND h.TipoConta = 'CCC'
					  AND h.TipoEntidade = 'C' 
					  AND h.Entidade = @Entidade
					  AND dv.TipoDocumento = '4'

				UNION ALL

				SELECT h.DataExtracto, h.TipoDoc, 'M' + RTRIM(dcc.TipoDocumento) AS TipoDocumento, h.Serie, h.NumDoc, 
				-- SÓ OS DOCUMENTOS DE LIQUIDAÇÃO É QUE OCUPAM A COLUNA DO DESCONTO 
				-- (ASSIM O TESTE AO TIPO DE DOCUMENTO É DESNECESSÁRIO)
				(h.ValorTotal + h.ValorDesconto) AS Total
				FROM Historico h
					INNER JOIN DocumentosCCT dcc
						ON h.TipoDoc = dcc.Documento
					LEFT JOIN Pendentes p
						ON h.Id = p.IdHistorico
				WHERE h.Modulo = 'M'
					  AND h.TipoConta = 'CCC'
					  AND h.TipoEntidade = 'C' 
					  AND h.Entidade = @Entidade
				      AND dcc.TipoDocumento <= '3'
		   ) h
	) h
WHERE h.Data >= @DataInicial AND h.Data <= @DataFinal
                         ]]>
                    </sql>


            parameters = New List(Of SqlParameter)

            ' @Entidade
            sqlParameter = New System.Data.SqlClient.SqlParameter
            With sqlParameter
                .DbType = Data.DbType.String
                .ParameterName = "Entidade"
                .Value = entidade.TrimEnd
            End With
            parameters.Add(sqlParameter)
            sqlParameter = Nothing

            ' @DataInicial
            sqlParameter = New System.Data.SqlClient.SqlParameter
            With sqlParameter
                .DbType = Data.DbType.Date
                .ParameterName = "DataInicial"
                .Value = dataInicial
            End With
            parameters.Add(sqlParameter)
            sqlParameter = Nothing

            ' @DataFinal
            sqlParameter = New System.Data.SqlClient.SqlParameter
            With sqlParameter
                .DbType = Data.DbType.Date
                .ParameterName = "DataFinal"
                .Value = dataFinal
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
            mList = New System.Collections.Generic.List(Of BE.LinhaExtratoCCT)

            ' Percorrer cada linha da tabela populada para gerar um objeto LinhaExtratoCCT
            For i = 0 To mTbl.Rows.Count - 1
                objLinhaExtrato = New BE.LinhaExtratoCCT
                With objLinhaExtrato

                    If Not IsDBNull(mTbl.Rows(i).Item("Data")) Then
                        .Documento_Data = CType(mTbl.Rows(i).Item("Data"), Date)
                    End If

                    If Not IsDBNull(mTbl.Rows(i).Item("TipoDoc")) Then
                        .Documento_Tipo = CType(mTbl.Rows(i).Item("TipoDoc"), System.String)
                    End If

                    If Not IsDBNull(mTbl.Rows(i).Item("TipoDocumento")) Then
                        .Documento_Tipo_TipoDocumento = CType(mTbl.Rows(i).Item("TipoDocumento"), System.String)
                    End If

                    If Not IsDBNull(mTbl.Rows(i).Item("Serie")) Then
                        .Documento_Serie = CType(mTbl.Rows(i).Item("Serie"), System.String)
                    End If

                    If Not IsDBNull(mTbl.Rows(i).Item("NumDoc")) Then
                        .Documento_Numero = CType(mTbl.Rows(i).Item("NumDoc"), System.Int32)
                    End If

                    If Not IsDBNull(mTbl.Rows(i).Item("Total")) Then
                        .Documento_Total = CType(mTbl.Rows(i).Item("Total"), Decimal)
                    End If

                    If Not IsDBNull(mTbl.Rows(i).Item("Saldo")) Then
                        .Saldo = CType(mTbl.Rows(i).Item("Saldo"), Decimal)
                    End If

                End With
                mList.Add(objLinhaExtrato)
                objLinhaExtrato = Nothing
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
