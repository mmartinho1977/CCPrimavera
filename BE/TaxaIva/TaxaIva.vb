Imports Microsoft.VisualBasic


Namespace BE

    Public Class TaxaIva
        Private mIva As String
        Private mDescricao As String
        Private mTaxa As Double
        Private mCamposUtilizador As BE.CamposUtilizadorCollection
        Private mEmModoEdicao As Boolean 'False = Entidade nova a inserir, true = Entidade existente a editar


        Sub New()
            mIva = ""
            mDescricao = ""
            mTaxa = 0
            mEmModoEdicao = False
            mCamposUtilizador = New BE.CamposUtilizadorCollection
        End Sub


        Public Property Iva() As String
            Get
                Return mIva
            End Get
            Set(ByVal value As String)
                mIva = value.TrimEnd
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


        Public Property Taxa() As Double
            Get
                Return mTaxa
            End Get
            Set(ByVal value As Double)
                mTaxa = value
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

