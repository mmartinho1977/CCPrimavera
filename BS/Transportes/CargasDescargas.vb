Imports Microsoft.VisualBasic
Imports CCPrimavera.Specifications
Imports CCPrimavera.Specifications.CargaDescarga_Specs
Imports Interop



Namespace BS

    Public Class CargasDescargas
        Private objMotor As ErpBS900.ErpBS


        Sub New(ByRef Motor As ErpBS900.ErpBS)
            objMotor = Motor
        End Sub



        Public Function GeraCargaDescarga(ByVal tipoEntidade As String, ByVal entidade As String) As BE.CargaDescarga
            Dim objCargaDescarga As New GcpBE900.GcpBECargaDescarga
            Dim entidadeNome As String
            Dim entidadeMorada As String
            Dim entidadeMorada2 As String
            Dim entidadeLocalidade As String
            Dim entidadeCodigoPostal As String
            Dim entidadeLocalidadePostal As String
            Dim entidadeDistrito As String
            Dim entidadePais As String


            If (tipoEntidade.Trim().Length = 0) Then
                Throw New Exception("A string tipoEntidade passada como parametro na função GeraCargaDescarga não pode ser vazia;")
                Exit Function
            End If
            If (entidade.Trim().Length = 0) Then
                Throw New Exception("A string entidade passada como parametro na função GeraCargaDescarga não pode ser vazia;")
                Exit Function
            End If

            Try

                ' recolher dados a partir da ficha da entidade
                Select Case tipoEntidade
                    Case "C"
                        Dim cliente As GcpBE900.GcpBECliente
                        cliente = objMotor.Comercial.Clientes.Edita(entidade)
                        entidadeNome = cliente.Nome
                        entidadeMorada = cliente.Morada
                        entidadeMorada2 = cliente.Morada2
                        entidadeLocalidade = cliente.Localidade
                        entidadeCodigoPostal = cliente.CodigoPostal
                        entidadeLocalidadePostal = cliente.LocalidadeCodigoPostal
                        entidadeDistrito = cliente.Distrito
                        entidadePais = cliente.Pais
                        Exit Select
                    Case Else
                        Throw New Exception("CCPrimavera / CargaDescarga - TipoEntidade não previsto (!=C)")
                End Select

                ' carga
                objCargaDescarga.MoradaCarga = objMotor.Contexto.IDMorada
                objCargaDescarga.Morada2Carga = ""
                If (objMotor.Contexto.IDNumPorta <> "") Then
                    objCargaDescarga.Morada2Carga += String.Format("Numero: {0}", objMotor.Contexto.IDNumPorta)
                End If
                If (objMotor.Contexto.IDPorta <> "") Then
                    objCargaDescarga.Morada2Carga += String.Format("  Porta: {0}", objMotor.Contexto.IDPorta)
                End If
                objCargaDescarga.LocalidadeCarga = objMotor.Contexto.IDLocalidade
                objCargaDescarga.CodPostalCarga = objMotor.Contexto.IDCodPostal
                objCargaDescarga.CodPostalLocalidadeCarga = objMotor.Contexto.IDCodPostalLocal
                objCargaDescarga.DistritoCarga = objMotor.Contexto.IDDistritoCod
                objCargaDescarga.PaisCarga = "PT"
                objCargaDescarga.DataCarga = DateTime.Now.ToString("dd-MM-yyyy")
                objCargaDescarga.HoraCarga = DateTime.Now.ToString("HH:mm")

                ' SUPOSTAMENTE: o editor de vendas ao gravar adiciona uns minutos para a entrega
                ' no futuro pode usar-se a seguinte instrução: 
                ' objCargaDescarga.HoraCarga = DateTime.Now.AddMinutes(10).ToString("HH:mm")

                ' descarga
                objCargaDescarga.TipoEntidadeEntrega = tipoEntidade
                objCargaDescarga.EntidadeEntrega = entidade
                objCargaDescarga.NomeEntrega = entidadeNome
                objCargaDescarga.MoradaEntrega = entidadeMorada
                objCargaDescarga.Morada2Entrega = entidadeMorada2
                objCargaDescarga.LocalidadeEntrega = entidadeLocalidade
                objCargaDescarga.CodPostalEntrega = entidadeCodigoPostal
                objCargaDescarga.CodPostalLocalidadeEntrega = entidadeLocalidadePostal
                objCargaDescarga.DistritoEntrega = entidadeDistrito
                objCargaDescarga.PaisEntrega = entidadePais
                objCargaDescarga.DataDescarga = DateTime.Now.ToString("dd-MM-yyyy")
                objCargaDescarga.HoraDescarga = ""

            Catch ex As Exception
                Throw ex
                Exit Function
            End Try

            GeraCargaDescarga = Me.GetCargaDescarga(objCargaDescarga)
            objCargaDescarga = Nothing

        End Function



        Public Function GetCargaDescarga(ByVal cargaDescarga As GcpBE900.GcpBECargaDescarga) As BE.CargaDescarga
            Dim objCargaDescarga As New BE.CargaDescarga

            If (IsNothing(cargaDescarga)) Then
                Throw New Exception("CCPrimavera / CargasDescargas, o argumento cargaDescarga não pode ser nulo no metodo GetCargaDescarga.")
            End If

            With objCargaDescarga
                .CARGA_Morada = cargaDescarga.MoradaCarga
                .CARGA_Morada2 = cargaDescarga.Morada2Carga
                .CARGA_Localidade = cargaDescarga.LocalidadeCarga
                .CARGA_CodigoPostal = cargaDescarga.CodPostalCarga
                .CARGA_LocalidadePostal = cargaDescarga.CodPostalLocalidadeCarga
                .CARGA_Distrito = cargaDescarga.DistritoCarga
                .CARGA_Pais = cargaDescarga.PaisCarga
                .CARGA_Data = cargaDescarga.DataCarga
                .CARGA_Hora = cargaDescarga.HoraCarga
                .DESCARGA_TipoEntidade = cargaDescarga.TipoEntidadeEntrega
                .DESCARGA_Entidade = cargaDescarga.EntidadeEntrega
                .DESCARGA_Nome = cargaDescarga.NomeEntrega
                .DESCARGA_Morada = cargaDescarga.MoradaEntrega
                .DESCARGA_Morada2 = cargaDescarga.Morada2Entrega
                .DESCARGA_Localidade = cargaDescarga.LocalidadeEntrega
                .DESCARGA_CodigoPostal = cargaDescarga.CodPostalEntrega
                .DESCARGA_LocalidadePostal = cargaDescarga.CodPostalLocalidadeEntrega
                .DESCARGA_Distrito = cargaDescarga.DistritoEntrega
                .DESCARGA_Pais = cargaDescarga.PaisEntrega
                .DESCARGA_Data = cargaDescarga.DataDescarga
                .DESCARGA_Hora = cargaDescarga.HoraDescarga
            End With

            GetCargaDescarga = objCargaDescarga
            objCargaDescarga = Nothing

        End Function



        Public Function GetGcpBECargaDescarga(ByVal cargaDescarga As BE.CargaDescarga) As GcpBE900.GcpBECargaDescarga
            Dim objCargaDescarga As New GcpBE900.GcpBECargaDescarga


            If (IsNothing(cargaDescarga)) Then
                Throw New Exception("CCPrimavera / CargasDescargas, o argumento cargaDescarga não pode ser nulo no metodo GetGcpBECargaDescarga.")
            End If

            With objCargaDescarga
                .MoradaCarga = cargaDescarga.CARGA_Morada
                .Morada2Carga = cargaDescarga.CARGA_Morada2
                .LocalidadeCarga = cargaDescarga.CARGA_Localidade
                .CodPostalCarga = cargaDescarga.CARGA_CodigoPostal
                .CodPostalLocalidadeCarga = cargaDescarga.CARGA_LocalidadePostal
                .DistritoCarga = cargaDescarga.CARGA_Distrito
                .PaisCarga = cargaDescarga.CARGA_Pais
                .DataCarga = cargaDescarga.CARGA_Data
                .HoraCarga = cargaDescarga.CARGA_Hora
                .TipoEntidadeEntrega = cargaDescarga.DESCARGA_TipoEntidade
                .EntidadeEntrega = cargaDescarga.DESCARGA_Entidade
                .NomeEntrega = cargaDescarga.DESCARGA_Nome
                .MoradaEntrega = cargaDescarga.DESCARGA_Morada
                .Morada2Entrega = cargaDescarga.DESCARGA_Morada2
                .LocalidadeEntrega = cargaDescarga.DESCARGA_Localidade
                .CodPostalEntrega = cargaDescarga.DESCARGA_CodigoPostal
                .CodPostalLocalidadeEntrega = cargaDescarga.DESCARGA_LocalidadePostal
                .DistritoEntrega = cargaDescarga.DESCARGA_Distrito
                .PaisEntrega = cargaDescarga.DESCARGA_Pais
                .DataDescarga = cargaDescarga.DESCARGA_Data
                .HoraDescarga = cargaDescarga.DESCARGA_Hora
            End With

            GetGcpBECargaDescarga = objCargaDescarga
            objCargaDescarga = Nothing

        End Function



        Protected Overrides Sub Finalize()
            If Not IsNothing(objMotor) Then
                objMotor = Nothing
            End If
            MyBase.Finalize()
        End Sub


    End Class

End Namespace
