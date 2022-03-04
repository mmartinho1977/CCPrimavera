Imports Microsoft.VisualBasic
Imports Interop


Namespace BS

    Public Class ArtigosPrecos
        Private objMotor As ErpBS900.ErpBS

        Sub New(ByRef Motor As ErpBS900.ErpBS)
            objMotor = Motor
        End Sub


        Public Function Edita(ByVal Artigo As String, ByVal Moeda As String, ByVal Unidade As String) As BE.ArtigoMoeda
            Dim objArtigoMoeda As BE.ArtigoMoeda

            If Artigo.TrimEnd.Length = 0 Or Moeda.TrimEnd.Length = 0 Or Unidade.TrimEnd.Length = 0 Then
                Throw New Exception("Um dos argumentos: Artigo, Moeda ou Unidade da funcao de Edicao da classe ArtigoMoeda foi comunicada com o tamanho zero.")
                Exit Function
            End If

            If Me.Existe(Artigo, Moeda, Unidade) = False Then
                Throw New Exception("O registo " & Artigo & "/" & Moeda & "/" & Unidade & " não existe na tabela de artigosMoedas, pelo que nao pode ser editado.")
                Exit Function
            End If

            Try
                objArtigoMoeda = GetArtigoMoeda(objMotor.Comercial.ArtigosPrecos.Edita(Artigo, Moeda, Unidade))
            Catch ex As Exception
                Throw ex
                Exit Function
            End Try

            Edita = objArtigoMoeda

            objArtigoMoeda = Nothing
        End Function


        Private Function GetArtigoMoeda(ByVal ArtigoMoeda As GcpBE900.GcpBEArtigoMoeda) As BE.ArtigoMoeda
            Dim objArtigoMoeda As BE.ArtigoMoeda

            If IsNothing(ArtigoMoeda) Then
                Throw New Exception("O argumento ArtigoMoeda da Função privada GetArtigoMoeda não pode ser nulo.")
                Exit Function
            End If

            objArtigoMoeda = New BE.ArtigoMoeda

            With objArtigoMoeda

                .Artigo = ArtigoMoeda.Artigo
                .Moeda = ArtigoMoeda.Moeda
                .Unidade = ArtigoMoeda.Unidade
                .PVP1 = ArtigoMoeda.PVP1
                .PVP2 = ArtigoMoeda.PVP2
                .PVP3 = ArtigoMoeda.PVP3
                .PVP4 = ArtigoMoeda.PVP4
                .PVP5 = ArtigoMoeda.PVP5
                .PVP6 = ArtigoMoeda.PVP6
                .PVP1IvaIncluido = ArtigoMoeda.PVP1IvaIncluido
                .PVP2IvaIncluido = ArtigoMoeda.PVP2IvaIncluido
                .PVP3IvaIncluido = ArtigoMoeda.PVP3IvaIncluido
                .PVP4IvaIncluido = ArtigoMoeda.PVP4IvaIncluido
                .PVP5IvaIncluido = ArtigoMoeda.PVP5IvaIncluido
                .PVP6IvaIncluido = ArtigoMoeda.PVP6IvaIncluido
                .EmModoEdicao = False
            End With

            GetArtigoMoeda = objArtigoMoeda

            objArtigoMoeda = Nothing
        End Function


        Public Function Existe(ByVal Artigo As String, ByVal Moeda As String, ByVal Unidade As String) As Boolean
            If Artigo.TrimEnd.Length = 0 Or Moeda.TrimEnd.Length = 0 Or Unidade.TrimEnd.Length = 0 Then
                Throw New Exception("Um dos argumentos: Artigo, Moeda ou Unidade foi comunicado com o tamanho zero, na função Existe.")
                Exit Function
            End If

            Return objMotor.Comercial.ArtigosPrecos.Existe(Artigo, Moeda, Unidade)
        End Function


        Protected Overrides Sub Finalize()
            If Not IsNothing(objMotor) Then
                objMotor = Nothing
            End If
            MyBase.Finalize()
        End Sub

    End Class

End Namespace


