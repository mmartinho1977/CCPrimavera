Imports Microsoft.VisualBasic


Namespace BE

    Public Class CargaDescarga
        Private mCARGA_Morada As String
        Private mCARGA_Morada2 As String
        Private mCARGA_Localidade As String
        Private mCARGA_CodigoPostal As String
        Private mCARGA_LocalidadePostal As String
        Private mCARGA_Distrito As String
        Private mCARGA_Pais As String
        Private mCARGA_Data As String
        Private mCARGA_Hora As String
        Private mDESCARGA_TipoEntidade As String
        Private mDESCARGA_Entidade As String
        Private mDESCARGA_Nome As String
        Private mDESCARGA_Morada As String
        Private mDESCARGA_Morada2 As String
        Private mDESCARGA_Localidade As String
        Private mDESCARGA_CodigoPostal As String
        Private mDESCARGA_LocalidadePostal As String
        Private mDESCARGA_Distrito As String
        Private mDESCARGA_Pais As String
        Private mDESCARGA_Data As String
        Private mDESCARGA_Hora As String
        ' Private mEmModoEdicao As Boolean ' este objeto não é persistido na base de dados, é gerado em runtime quando solicitado


        Sub New()

            mCARGA_Morada = ""
            mCARGA_Morada2 = ""
            mCARGA_Localidade = ""
            mCARGA_CodigoPostal = ""
            mCARGA_LocalidadePostal = ""
            mCARGA_Distrito = ""
            mCARGA_Pais = ""
            mCARGA_Data = ""
            mCARGA_Hora = ""

            mDESCARGA_TipoEntidade = ""
            mDESCARGA_Entidade = ""
            mDESCARGA_Nome = ""
            mDESCARGA_Morada = ""
            mDESCARGA_Morada2 = ""
            mDESCARGA_Localidade = ""
            mDESCARGA_CodigoPostal = ""
            mDESCARGA_LocalidadePostal = ""
            mDESCARGA_Distrito = ""
            mDESCARGA_Pais = ""
            mDESCARGA_Data = ""
            mDESCARGA_Hora = ""

            ' mEmModoEdicao = False
        End Sub



        Public Property CARGA_Morada() As String
            Get
                Return mCARGA_Morada
            End Get
            Set(ByVal value As String)
                mCARGA_Morada = value
            End Set
        End Property



        Public Property CARGA_Morada2() As String
            Get
                Return mCARGA_Morada2
            End Get
            Set(ByVal value As String)
                mCARGA_Morada2 = value
            End Set
        End Property



        Public Property CARGA_Localidade() As String
            Get
                Return mCARGA_Localidade
            End Get
            Set(ByVal value As String)
                mCARGA_Localidade = value
            End Set
        End Property



        Public Property CARGA_CodigoPostal() As String
            Get
                Return mCARGA_CodigoPostal
            End Get
            Set(ByVal value As String)
                mCARGA_CodigoPostal = value
            End Set
        End Property



        Public Property CARGA_LocalidadePostal() As String
            Get
                Return mCARGA_LocalidadePostal
            End Get
            Set(ByVal value As String)
                mCARGA_LocalidadePostal = value
            End Set
        End Property



        Public Property CARGA_Distrito() As String
            Get
                Return mCARGA_Distrito
            End Get
            Set(ByVal value As String)
                mCARGA_Distrito = value
            End Set
        End Property



        Public Property CARGA_Pais() As String
            Get
                Return mCARGA_Pais
            End Get
            Set(ByVal value As String)
                mCARGA_Pais = value
            End Set
        End Property



        Public Property CARGA_Data() As String
            Get
                Return mCARGA_Data
            End Get
            Set(ByVal value As String)
                mCARGA_Data = value
            End Set
        End Property



        Public Property CARGA_Hora() As String
            Get
                Return mCARGA_Hora
            End Get
            Set(ByVal value As String)
                mCARGA_Hora = value
            End Set
        End Property



        Public Property DESCARGA_TipoEntidade() As String
            Get
                Return mDESCARGA_TipoEntidade
            End Get
            Set(ByVal value As String)
                mDESCARGA_TipoEntidade = value
            End Set
        End Property



        Public Property DESCARGA_Entidade() As String
            Get
                Return mDESCARGA_Entidade
            End Get
            Set(ByVal value As String)
                mDESCARGA_Entidade = value
            End Set
        End Property



        Public Property DESCARGA_Nome() As String
            Get
                Return mDESCARGA_Nome
            End Get
            Set(ByVal value As String)
                mDESCARGA_Nome = value
            End Set
        End Property



        Public Property DESCARGA_Morada() As String
            Get
                Return mDESCARGA_Morada
            End Get
            Set(ByVal value As String)
                mDESCARGA_Morada = value
            End Set
        End Property



        Public Property DESCARGA_Morada2() As String
            Get
                Return mDESCARGA_Morada2
            End Get
            Set(ByVal value As String)
                mDESCARGA_Morada2 = value
            End Set
        End Property



        Public Property DESCARGA_Localidade() As String
            Get
                Return mDESCARGA_Localidade
            End Get
            Set(ByVal value As String)
                mDESCARGA_Localidade = value
            End Set
        End Property



        Public Property DESCARGA_CodigoPostal() As String
            Get
                Return mDESCARGA_CodigoPostal
            End Get
            Set(ByVal value As String)
                mDESCARGA_CodigoPostal = value
            End Set
        End Property



        Public Property DESCARGA_LocalidadePostal() As String
            Get
                Return mDESCARGA_LocalidadePostal
            End Get
            Set(ByVal value As String)
                mDESCARGA_LocalidadePostal = value
            End Set
        End Property



        Public Property DESCARGA_Distrito() As String
            Get
                Return mDESCARGA_Distrito
            End Get
            Set(ByVal value As String)
                mDESCARGA_Distrito = value
            End Set
        End Property



        Public Property DESCARGA_Pais() As String
            Get
                Return mDESCARGA_Pais
            End Get
            Set(ByVal value As String)
                mDESCARGA_Pais = value
            End Set
        End Property



        Public Property DESCARGA_Data() As String
            Get
                Return mDESCARGA_Data
            End Get
            Set(ByVal value As String)
                mDESCARGA_Data = value
            End Set
        End Property



        Public Property DESCARGA_Hora() As String
            Get
                Return mDESCARGA_Hora
            End Get
            Set(ByVal value As String)
                mDESCARGA_Hora = value
            End Set
        End Property



        'Public Property EmModoEdicao() As Boolean
        '    Get
        '        Return mEmModoEdicao
        '    End Get
        '    Set(ByVal value As Boolean)
        '        mEmModoEdicao = value
        '    End Set
        'End Property



        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub

    End Class

End Namespace

