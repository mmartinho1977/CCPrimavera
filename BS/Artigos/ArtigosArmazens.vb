Imports Microsoft.VisualBasic
Imports Interop


Namespace BS

    Public Class ArtigosArmazens
        Private objMotor As ErpBS900.ErpBS

        Sub New(ByRef Motor As ErpBS900.ErpBS)
            objMotor = Motor
        End Sub


        Public Function GetStockDisponivel(ByVal Artigo As String, ByVal Armazem As String, ByVal Lote As String, Optional ByVal Localizacao As String = "") As Double

            If Not Me.Existe(Artigo, Armazem, Lote, Localizacao) Then
                Throw New Exception("Não existe stock para a conjugação, " & Artigo & "/" & Armazem & "/" & Lote & "/" & Localizacao & ".")
                Exit Function
            End If

            Return objMotor.Comercial.ArtigosArmazens.DaStockDisponivelArtigoArmazem(Artigo, Armazem, Lote, Localizacao)
        End Function



        Public Function Existe(ByVal Artigo As String, ByVal Armazem As String, ByVal Lote As String, Optional ByVal Localizacao As String = "") As Boolean
            If Artigo.TrimEnd.Length = 0 Or Armazem.TrimEnd.Length = 0 Or Lote.TrimEnd.Length = 0 Then
                Throw New Exception("Um dos argumentos: Artigo, Armazem ou Lote foi comunicado com o tamanho zero, na função Existe.")
                Exit Function
            End If

            Return objMotor.Comercial.ArtigosArmazens.Existe(Artigo, Armazem, Lote, Localizacao)
        End Function


        Protected Overrides Sub Finalize()
            If Not IsNothing(objMotor) Then
                objMotor = Nothing
            End If
            MyBase.Finalize()
        End Sub

    End Class

End Namespace


