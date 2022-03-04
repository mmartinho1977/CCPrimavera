Imports Microsoft.VisualBasic
Imports Interop


Namespace BS

    Public Class Nacionalidades
        Private objMotor As ErpBS900.ErpBS


        Sub New(ByRef Motor As ErpBS900.ErpBS)
            objMotor = Motor
        End Sub


        Public Function DevolveValorAtributo(ByRef Nacionalidade As String, ByRef Atributo As String) As String
            Return objMotor.RecursosHumanos.Nacionalidades.DaValorAtributo(Nacionalidade, Atributo)
        End Function



        Protected Overrides Sub Finalize()
            If Not IsNothing(objMotor) Then
                objMotor = Nothing
            End If
            MyBase.Finalize()
        End Sub

    End Class

End Namespace


