Imports Microsoft.VisualBasic


Namespace BE

    Public Class ArtigoMoeda
        Private mArtigo As String
        Private mMoeda As String
        Private mUnidade As String
        Private mPVP1 As Double
        Private mPVP2 As Double
        Private mPVP3 As Double
        Private mPVP4 As Double
        Private mPVP5 As Double
        Private mPVP6 As Double
        Private mPVP1IvaIncluido As Boolean
        Private mPVP2IvaIncluido As Boolean
        Private mPVP3IvaIncluido As Boolean
        Private mPVP4IvaIncluido As Boolean
        Private mPVP5IvaIncluido As Boolean
        Private mPVP6IvaIncluido As Boolean
        Private mEmModoEdicao As Boolean 'False = Entidade nova a inserir, true = Entidade existente a editar


        Sub New()
            mArtigo = ""
            mMoeda = ""
            mUnidade = ""
            mPVP1 = 0
            mPVP2 = 0
            mPVP3 = 0
            mPVP4 = 0
            mPVP5 = 0
            mPVP6 = 0
            mPVP1IvaIncluido = False
            mPVP2IvaIncluido = False
            mPVP3IvaIncluido = False
            mPVP4IvaIncluido = False
            mPVP5IvaIncluido = False
            mPVP6IvaIncluido = False
        End Sub

        Public Property Artigo() As String
            Get
                Return mArtigo
            End Get
            Set(ByVal value As String)
                mArtigo = value.TrimEnd
            End Set
        End Property

        Public Property Moeda() As String
            Get
                Return mMoeda
            End Get
            Set(ByVal value As String)
                mMoeda = value
            End Set
        End Property

        Public Property Unidade() As String
            Get
                Return mUnidade
            End Get
            Set(ByVal value As String)
                mUnidade = value
            End Set
        End Property

        Public Property PVP1() As Double
            Get
                Return mPVP1
            End Get
            Set(ByVal value As Double)
                mPVP1 = value
            End Set
        End Property

        Public Property PVP2() As Double
            Get
                Return mPVP2
            End Get
            Set(ByVal value As Double)
                mPVP2 = value
            End Set
        End Property

        Public Property PVP3() As Double
            Get
                Return mPVP3
            End Get
            Set(ByVal value As Double)
                mPVP3 = value
            End Set
        End Property

        Public Property PVP4() As Double
            Get
                Return mPVP4
            End Get
            Set(ByVal value As Double)
                mPVP4 = value
            End Set
        End Property

        Public Property PVP5() As Double
            Get
                Return mPVP5
            End Get
            Set(ByVal value As Double)
                mPVP5 = value
            End Set
        End Property

        Public Property PVP6() As Double
            Get
                Return mPVP6
            End Get
            Set(ByVal value As Double)
                mPVP6 = value
            End Set
        End Property

        Public Property PVP1IvaIncluido() As Boolean
            Get
                Return mPVP1IvaIncluido
            End Get
            Set(ByVal value As Boolean)
                mPVP1IvaIncluido = value
            End Set
        End Property

        Public Property PVP2IvaIncluido() As Boolean
            Get
                Return mPVP2IvaIncluido
            End Get
            Set(ByVal value As Boolean)
                mPVP2IvaIncluido = value
            End Set
        End Property

        Public Property PVP3IvaIncluido() As Boolean
            Get
                Return mPVP3IvaIncluido
            End Get
            Set(ByVal value As Boolean)
                mPVP3IvaIncluido = value
            End Set
        End Property

        Public Property PVP4IvaIncluido() As Boolean
            Get
                Return mPVP4IvaIncluido
            End Get
            Set(ByVal value As Boolean)
                mPVP4IvaIncluido = value
            End Set
        End Property

        Public Property PVP5IvaIncluido() As Boolean
            Get
                Return mPVP5IvaIncluido
            End Get
            Set(ByVal value As Boolean)
                mPVP5IvaIncluido = value
            End Set
        End Property

        Public Property PVP6IvaIncluido() As Boolean
            Get
                Return mPVP6IvaIncluido
            End Get
            Set(ByVal value As Boolean)
                mPVP6IvaIncluido = value
            End Set
        End Property

        Public Property EmModoEdicao() As Boolean
            Get
                Return mEmModoEdicao
            End Get
            Set(ByVal value As Boolean)
                mEmModoEdicao = value
            End Set
        End Property

        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub

    End Class

End Namespace

