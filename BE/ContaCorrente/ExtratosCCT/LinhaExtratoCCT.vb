Imports Microsoft.VisualBasic


Namespace BE

    Public Class LinhaExtratoCCT
        Private mId As String

        Private mDocumento_Data As Date
        Private mDocumento_Tipo As String
        Private mDocumento_Tipo_TipoDocumento As String   'permite distinguir os grupo do documento (DocumentoVenda.TipoDocumento)
        Private mDocumento_Serie As String
        Private mDocumento_Numero As System.Int32
        Private mDocumento_Total As Decimal
        Private mSaldo As Decimal


        Sub New()
            mId = ""
            mDocumento_Data = Nothing
            mDocumento_Tipo = ""
            mDocumento_Serie = ""
            mDocumento_Numero = 0
            mDocumento_Total = 0
            mSaldo = 0
        End Sub


        Public ReadOnly Property Id() As String
            Get
                Return String.Format("{0} {1}/{2}", mDocumento_Tipo.TrimEnd(), mDocumento_Serie.TrimEnd(), mDocumento_Numero.ToString().TrimEnd())
            End Get
        End Property


        Public Property Documento_Data() As Date
            Get
                Return mDocumento_Data
            End Get
            Set(ByVal value As Date)
                mDocumento_Data = value
            End Set
        End Property


        Public Property Documento_Tipo() As String
            Get
                Return mDocumento_Tipo
            End Get
            Set(ByVal value As String)
                mDocumento_Tipo = value
            End Set
        End Property


        Public Property Documento_Tipo_TipoDocumento() As String
            Get
                Return mDocumento_Tipo_TipoDocumento
            End Get
            Set(ByVal value As String)
                mDocumento_Tipo_TipoDocumento = value
            End Set
        End Property


        Public Property Documento_Serie() As String
            Get
                Return mDocumento_Serie
            End Get
            Set(ByVal value As String)
                mDocumento_Serie = value
            End Set
        End Property


        Public Property Documento_Numero() As System.Int32
            Get
                Return mDocumento_Numero
            End Get
            Set(ByVal value As System.Int32)
                mDocumento_Numero = value
            End Set
        End Property


        Public Property Documento_Total() As Decimal
            Get
                Return mDocumento_Total
            End Get
            Set(ByVal value As Decimal)
                mDocumento_Total = value
            End Set
        End Property


        Public Property Saldo() As Decimal
            Get
                Return mSaldo
            End Get
            Set(ByVal value As Decimal)
                mSaldo = value
            End Set
        End Property


        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub


    End Class

End Namespace

