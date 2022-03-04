Imports Microsoft.VisualBasic


Namespace BE

    Public Class Moeda
        Private mMoeda As String
        Private mDescricao As String
        Private mDescricaoParteInteira As String
        Private mDescricaoParteDecimal As String
        Private mArredondamentoValores As System.Int16
        Private mArredondamentoIva As System.Int16
        Private mArredondamentoPrecosUnitarios As System.Int16
        Private mCamposUtilizador As BE.CamposUtilizadorCollection
        Private mEmModoEdicao As Boolean 'False = Entidade nova a inserir, true = Entidade existente a editar


        Sub New()
            mMoeda = ""
            mDescricao = ""
            mDescricaoParteInteira = ""
            mDescricaoParteDecimal = ""
            mArredondamentoValores = 0
            mArredondamentoIva = 0
            mArredondamentoPrecosUnitarios = 0
            mEmModoEdicao = False
            mCamposUtilizador = New BE.CamposUtilizadorCollection
        End Sub


        Public Property Moeda() As String
            Get
                Return mMoeda
            End Get
            Set(ByVal value As String)
                mMoeda = value
            End Set
        End Property


        Public Property Descricao() As String
            Get
                Return mDescricao
            End Get
            Set(ByVal value As String)
                mDescricao = value
            End Set
        End Property


        Public Property DescricaoParteInteira() As String
            Get
                Return mDescricaoParteInteira
            End Get
            Set(ByVal value As String)
                mDescricaoParteInteira = value
            End Set
        End Property


        Public Property DescricaoParteDecimal() As String
            Get
                Return mDescricaoParteDecimal
            End Get
            Set(ByVal value As String)
                mDescricaoParteDecimal = value
            End Set
        End Property


        Public Property ArredondamentoValores() As System.Int16
            Get
                Return mArredondamentoValores
            End Get
            Set(ByVal value As System.Int16)
                mArredondamentoValores = value
            End Set
        End Property


        Public Property ArredondamentoIva() As System.Int16
            Get
                Return mArredondamentoIva
            End Get
            Set(ByVal value As System.Int16)
                mArredondamentoIva = value
            End Set
        End Property


        Public Property ArredondamentoPrecosUnitarios() As System.Int16
            Get
                Return mArredondamentoPrecosUnitarios
            End Get
            Set(ByVal value As System.Int16)
                mArredondamentoPrecosUnitarios = value
            End Set
        End Property


        Public Property CamposUtilizador() As BE.CamposUtilizadorCollection
            Get
                Return mCamposUtilizador
            End Get
            Set(ByVal value As BE.CamposUtilizadorCollection)
                mCamposUtilizador = value
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
            mCamposUtilizador = Nothing
            MyBase.Finalize()
        End Sub

    End Class

End Namespace

