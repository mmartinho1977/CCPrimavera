Imports Microsoft.VisualBasic


Namespace BE

    Public Class OutroTerceiro
        Private mTipoEntidade As String
        Private mNaturezaTerceiro As String
        Private mCondicaoPagamento As String
        Private mMoeda As String
        Private mModoPagamento As String
        Private mCodigo As String
        Private mNome As String
        Private mMorada As String
        Private mLocalidade As String
        Private mCodigoPostal As String
        Private mLocalidadePostal As String
        Private mTelefone As String
        Private mFax As String
        Private mContribuinte As String
        Private mPessoaSingular As Boolean
        Private mDataCriacao As Date
        Private mDataUltimaActualizacao As Date
        Private mCamposUtilizador As BE.CamposUtilizadorCollection
        Private mEmModoEdicao As Boolean 'False = Entidade nova a inserir, true = Entidade existente a editar


        Sub New(ByVal TipoEntidade As String)

            If TipoEntidade.Trim.Length < 0 Or TipoEntidade.Trim.Length > 1 Then
                Throw New Exception("O argumento TipoEntidade do construtor da classe OutroTerceiro, deve ter um caracter.")
                Exit Sub
            End If

            Select Case TipoEntidade.ToString
                Case "D"
                    mTipoEntidade = "D"
                    mNaturezaTerceiro = "D"

                Case Else
                    Throw New Exception("O argumento TipoEntidade do construtor da Classe OutroTerceiro, deve ser válido e programado.")
                    Exit Sub
            End Select

            mCondicaoPagamento = "1"
            mMoeda = "EUR"
            mModoPagamento = "NUM"
            mPessoaSingular = True
            mDataCriacao = Now.Date
            mDataUltimaActualizacao = Now.Date
            mCamposUtilizador = New BE.CamposUtilizadorCollection

        End Sub

        Public ReadOnly Property TipoEntidade() As String
            Get
                Return mTipoEntidade
            End Get
        End Property

        Public ReadOnly Property NaturezaTerceiro() As String
            Get
                Return mNaturezaTerceiro
            End Get
        End Property

        Public Property Codigo() As String
            Get
                Return mCodigo
            End Get
            Set(ByVal value As String)
                mCodigo = value.TrimEnd
            End Set
        End Property

        Public Property Nome() As String
            Get
                Return mNome
            End Get
            Set(ByVal value As String)
                mNome = value.TrimEnd
            End Set
        End Property

        Public Property Morada() As String
            Get
                Return mMorada
            End Get
            Set(ByVal value As String)
                mMorada = value.TrimEnd
            End Set
        End Property

        Public Property Localidade() As String
            Get
                Return mLocalidade
            End Get
            Set(ByVal value As String)
                mLocalidade = value.TrimEnd
            End Set
        End Property

        Public Property CodigoPostal() As String
            Get
                Return mCodigoPostal
            End Get
            Set(ByVal value As String)
                mCodigoPostal = value.TrimEnd
            End Set
        End Property

        Public Property LocalidadePostal() As String
            Get
                Return mLocalidadePostal
            End Get
            Set(ByVal value As String)
                mLocalidadePostal = value.TrimEnd
            End Set
        End Property

        Public Property Telefone() As String
            Get
                Return mTelefone
            End Get
            Set(ByVal value As String)
                mTelefone = value
            End Set
        End Property

        Public Property Fax() As String
            Get
                Return mFax
            End Get
            Set(ByVal value As String)
                mFax = value
            End Set
        End Property

        Public Property Contribuinte() As String
            Get
                Return mContribuinte
            End Get
            Set(ByVal value As String)
                mContribuinte = value.TrimEnd
            End Set
        End Property

        Public Property CondicaoPagamento() As String
            Get
                Return mCondicaoPagamento
            End Get
            Set(ByVal value As String)
                mCondicaoPagamento = value.TrimEnd
            End Set
        End Property

        Public Property Moeda() As String
            Get
                Return mMoeda
            End Get
            Set(ByVal value As String)
                mMoeda = value.TrimEnd
            End Set
        End Property

        Public Property ModoPagamento() As String
            Get
                Return mModoPagamento
            End Get
            Set(ByVal value As String)
                mModoPagamento = value.TrimEnd
            End Set
        End Property

        Public Property PessoaSingular() As Boolean
            Get
                Return mPessoaSingular
            End Get
            Set(ByVal value As Boolean)
                mPessoaSingular = value
            End Set
        End Property

        Public Property DataCriacao() As Date
            Get
                If mEmModoEdicao = False Then 'Ou seja se se trata de uma insercao
                    mDataCriacao = Now.Date
                End If
                Return mDataCriacao
            End Get
            Set(ByVal value As Date)
                mDataCriacao = value
            End Set
        End Property

        Public Property DataUltimaActualizacao() As Date
            Get
                mDataUltimaActualizacao = Now().Date
                Return mDataUltimaActualizacao
            End Get
            Set(ByVal value As Date)
                mDataUltimaActualizacao = value
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

