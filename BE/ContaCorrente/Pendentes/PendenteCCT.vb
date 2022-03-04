Imports Microsoft.VisualBasic


Namespace BE

    Public Class PendenteCCT
        Private mId As String
        Private mTipoConta As String
        Private mEstado As String
        Private mData As Date
        Private mVencimento As Date
        Private mTipoDocumento As String
        Private mSerie As String
        Private mNumero As System.Int32
        Private mNumeroExterno As String
        Private mTotal As Decimal
        Private mPendente As Decimal


        Sub New()
            mId = ""
            mTipoConta = ""
            mEstado = ""
            mData = Nothing
            mVencimento = Nothing
            mTipoDocumento = ""
            mSerie = ""
            mNumero = 0
            mNumeroExterno = ""
            mTotal = 0
            mPendente = 0
        End Sub




        Public ReadOnly Property Id() As String
            Get
                Return String.Format("{0} {1}/{2}", mTipoDocumento.TrimEnd(), mSerie.TrimEnd(), mNumero.ToString().TrimEnd())
            End Get
        End Property



        Public Property TipoConta() As String
            Get
                Return mTipoConta
            End Get
            Set(ByVal value As String)
                mTipoConta = value
            End Set
        End Property



        Public Property Estado() As String
            Get
                Return mEstado
            End Get
            Set(ByVal value As String)
                mEstado = value
            End Set
        End Property



        Public Property Data() As Date
            Get
                Return mData
            End Get
            Set(ByVal value As Date)
                mData = value
            End Set
        End Property



        Public Property Vencimento() As Date
            Get
                Return mVencimento
            End Get
            Set(ByVal value As Date)
                mVencimento = value
            End Set
        End Property



        Public Property TipoDocumento() As String
            Get
                Return mTipoDocumento
            End Get
            Set(ByVal value As String)
                mTipoDocumento = value
            End Set
        End Property



        Public Property Serie() As String
            Get
                Return mSerie
            End Get
            Set(ByVal value As String)
                mSerie = value
            End Set
        End Property



        Public Property Numero() As System.Int32
            Get
                Return mNumero
            End Get
            Set(ByVal value As System.Int32)
                mNumero = value
            End Set
        End Property



        Public Property NumeroExterno() As String
            Get
                Return mNumeroExterno
            End Get
            Set(ByVal value As String)
                mNumeroExterno = value
            End Set
        End Property



        Public Property Total() As Decimal
            Get
                Return mTotal
            End Get
            Set(ByVal value As Decimal)
                mTotal = value
            End Set
        End Property



        Public Property Pendente() As Decimal
            Get
                Return mPendente
            End Get
            Set(ByVal value As Decimal)
                mPendente = value
            End Set
        End Property


        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub


    End Class

End Namespace

