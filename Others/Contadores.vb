Imports Microsoft.VisualBasic
Imports Interop


'
'NOTAS IMPORTANTES (mais informação em Notas.txt)
'   O uso desta classe está suiportado pela tabela TDU_Contadores, que tem estrutura e conteudo (linhas) fixos
'   ATENÇÃO.


Namespace Others

    Public Class Contadores
        Private objMotor As ErpBS900.ErpBS

        ' Contadores e TiposEntidade não é a mesma coisa já que eu posso ter por exemplo um contador para numerar automaticamente os artigos, familias, codigos de barra, etc.
        Public Enum ContadoresTipo
            Clientes
            Fornecedores
            OutrosDevedores
            OutrosCredores
            Artigos
            EAN13
            Familias
        End Enum


        Sub New(ByVal Motor As ErpBS900.ErpBS)
            objMotor = Motor
        End Sub



        Public Function IncrementaDevolveString(ByVal contador As ContadoresTipo) As String
            Dim mNumero As System.Int32
            Dim mStrSelect As System.String
            Dim mMascara As System.String
            Dim mQtdCaracteres_Interrogacoes As System.Int16
            Dim mQtdCaracteres_Prefixo As System.Int16

            If (Not Existe(contador)) Then
                Throw New Exception("CC - O contador " + contador.ToString + " não existe na tabela TDU_Contadores.")
            End If

            Try
                mNumero = IncrementaDevolveInteiro(contador)
            Catch ex As Exception
                Throw ex
                IncrementaDevolveString = ""
                Exit Function
            End Try

            mStrSelect = "Select Top 1 CDU_Mascara From TDU_Contadores Where CDU_Contador ="
            mStrSelect += " '" & contador.ToString & "'"

            Try
                mMascara = objMotor.Consulta(mStrSelect).Valor(0).ToString
            Catch ex As Exception
                Throw ex
                Exit Function
            End Try

            If mMascara.Contains("?") = False Then
                Throw New Exception("O Prefixo do Contador " & contador.ToString & " não está devidamente configurado.")
                Exit Function
            End If

            mQtdCaracteres_Interrogacoes = mMascara.LastIndexOf("?") - mMascara.IndexOf("?") + 1
            mQtdCaracteres_Prefixo = mMascara.TrimEnd.Length - mQtdCaracteres_Interrogacoes

            mMascara = mMascara.Substring(0, mQtdCaracteres_Prefixo)

            Return Format(mNumero, "'" & mMascara & "'" & New String("0", mQtdCaracteres_Interrogacoes))

        End Function




        Public Function IncrementaDevolveInteiro(ByVal contador As ContadoresTipo) As System.Int32
            Dim mStrUpdate As String
            Dim mStrSelect As String
            Dim mRecordsAffected As System.Int16 = 0

            If (Not Existe(contador)) Then
                Throw New Exception("CC - O contador " + contador.ToString + " não existe na tabela TDU_Contadores.")
            End If

            mStrUpdate = "Update TDU_Contadores set CDU_UltimoNumero = (CDU_UltimoNumero+1) Where CDU_Contador ="
            mStrUpdate += " '" & contador.ToString & "'"

            mStrSelect = " Select Top 1 CDU_UltimoNumero From TDU_Contadores Where CDU_Contador ="
            mStrSelect += " '" & contador.ToString & "'"

            Try
                objMotor.IniciaTransaccao()
                objMotor.DSO.BDAPL.Execute(mStrUpdate, mRecordsAffected)
                If mRecordsAffected = 0 Then
                    Throw New Exception("O Contador " & contador.ToString & " não existe na tabela de contadores.")
                End If
                IncrementaDevolveInteiro = objMotor.DSO.BDAPL.Execute(mStrSelect).Fields("CDU_UltimoNumero").Value
                objMotor.TerminaTransaccao()
            Catch ex As Exception
                objMotor.DesfazTransaccao()
                Throw ex
            End Try
        End Function



        Private Function Existe(ByVal contador As ContadoresTipo) As Boolean
            Dim mStrSelect As String

            mStrSelect = "Select count('-') From TDU_Contadores Where CDU_Contador ="
            mStrSelect += "'" & contador.ToString & "'"
            Try
                If objMotor.DSO.BDAPL.Execute(mStrSelect).Fields(0).Value > 0 Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                Throw ex
            End Try

        End Function

    End Class

End Namespace

