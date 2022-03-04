Imports Microsoft.VisualBasic

Namespace BE

    Public Class CampoUtilizador
        Public Enum TiposDadosCampoUtilizador
            _String
            _Integer
            _Money
            _Double
            _Date
            _Boolean
        End Enum
        Private mCampo As String
        Private mTipo As TiposDadosCampoUtilizador
        Private mValor As Object

        Sub New()
            mCampo = ""
            mTipo = Nothing
            mValor = ""
        End Sub

        Sub New(ByVal Campo As String, ByVal Tipo As TiposDadosCampoUtilizador, ByVal Valor As Object)
            mCampo = Campo
            mTipo = Tipo
            mValor = Valor
        End Sub


        Public Property Campo() As String
            Get
                Return mCampo
            End Get
            Set(ByVal value As String)
                mCampo = value
            End Set
        End Property

        Public Property Tipo() As TiposDadosCampoUtilizador
            Get
                Return mTipo
            End Get
            Set(ByVal value As TiposDadosCampoUtilizador)
                mTipo = value
            End Set
        End Property


        Public Property Valor() As Object
            Get
                If mTipo = TiposDadosCampoUtilizador._String Then
                    Return mValor.TrimEnd
                Else
                    Return mValor
                End If
            End Get
            Set(ByVal value As Object)
                mValor = value
            End Set
        End Property
    End Class

End Namespace

