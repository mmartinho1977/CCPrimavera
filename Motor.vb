Imports Microsoft.VisualBasic
Imports CCPrimavera.Others
Imports CCPrimavera.BS
Imports Interop



Public Class Motor
    Private Shared primavera_CommonFilesFolder As String
    Private connectionString As String
    Private empresaPrimavera As String
    Private utilizadorPrimavera As String
    Private passwordPrimavera As String
    Private objMotor As ErpBS900.ErpBS
    Private objArtigos As Artigos
    Private objArtigosPrecos As ArtigosPrecos
    Private objArtigosArmazens As ArtigosArmazens
    Private objTaxasIva As TaxasIva
    Private objVendas As DocumentosVenda
    Private objCargasDescargas As CargasDescargas
    Private objInternos As DocumentosInternos
    Private objPendentes As Pendentes
    Private objExtratosCCT As ExtratosCCT
    Private objContadores As Contadores
    Private objClientes As Clientes
    Private objEntidadesAssociadas As EntidadesAssociadas
    Private objOutrosTerceiros As OutrosTerceiros
    Private objTiposEntidade As TiposEntidade
    Private objMoedas As Moedas
    Private objNacionalidades As Nacionalidades



    Sub New(ByVal connectionString As String, ByVal empresaPrimavera As String, ByVal utilizadorPrimavera As String, ByVal passwordPrimavera As String)
        Me.connectionString = connectionString
        Me.empresaPrimavera = empresaPrimavera
        Me.utilizadorPrimavera = utilizadorPrimavera
        Me.passwordPrimavera = passwordPrimavera


        ' Isto terá que ser recebido por parametro no futuro //ATUALIZAR NO FUTURO (receber por parametro)
        primavera_CommonFilesFolder = "C:\Program Files (x86)\PRIMAVERA\SG900\Apl" 'pasta dos ficheiros comuns especifica da versão do ERP PRIMAVERA utilizada.

        'PRIMAVERA (preparar ambiente para integração)
        'Este handler tem que ser adicionado antes de existir qualquer referência para classes existentes nos Interop's,
        'isto é, no método Main() da aplicação NÃO PODERÁ EXISTIR DECLARAÇÕES DE VARIÁVEIS DE TIPOS EXISTENTES NOS INTEROPS.
        'Com este método, na pasta da aplicação não deverão existir os Interops e as referências para os mesmos deverão ser
        'adicionadas com Copy Local = False e Embed Interop Types = false. (Em C# é Specific version = false)

        AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf CurrentDomain_AssemblyResolve

        ' A abertura da empresa tem que ser feita num procedimento diferente daquele em que está a ser carregado o redirecionamento AssemblyResolve
        AbreEmpresa()

    End Sub




    Private Shared Function CurrentDomain_AssemblyResolve(sender As Object, args As ResolveEventArgs) As System.Reflection.Assembly
        ' variaveis e constantes necessárias
        Dim assemblyFullName As String
        Dim assemblyName As System.Reflection.AssemblyName

        ' identificar a assembly a resolver (recebida como argumento e enviada pelo handler que é disparado por não ter encontrado a dll na pasta deste programa)
        assemblyName = New System.Reflection.AssemblyName(args.Name)
        assemblyFullName = System.IO.Path.Combine(System.IO.Path.Combine(primavera_CommonFilesFolder), assemblyName.Name + ".dll")
        ' se encontrar a assembly dentro da pasta common files, refleti-la nesta aplicação como uma referência válida à DLL
        If (System.IO.File.Exists(assemblyFullName)) Then

            '==================== ativar para ver os assemblys que são carregados através deste resolver ===============
            'Console.WriteLine(assemblyFullName)

            Return System.Reflection.Assembly.LoadFile(assemblyFullName)

        Else
            Return Nothing
        End If

    End Function



    Private Sub AbreEmpresa()

        If Not IsNothing(objMotor) Then
            objMotor = Nothing
        End If

        objMotor = New ErpBS900.ErpBS

        objMotor.AbreEmpresaTrabalho(StdBE900.EnumTipoPlataforma.tpProfissional, empresaPrimavera, utilizadorPrimavera, passwordPrimavera)
    End Sub



    Public Sub Executa(ByVal sqlString As String)
        objMotor.DSO.BDAPL.Execute(sqlString)
    End Sub


    Public Sub Executa(ByVal sqlString As String, ByRef returnValue As System.Int16)
        returnValue = CType(objMotor.DSO.BDAPL.Execute(sqlString).Fields(0).Value, System.Int16)
    End Sub


    Public Sub IniciaTransaccao()
        objMotor.IniciaTransaccao()
    End Sub


    Public Sub TerminaTransaccao()
        objMotor.TerminaTransaccao()
    End Sub


    Public Sub DesfazTransaccao()
        objMotor.DesfazTransaccao()
    End Sub


    Public ReadOnly Property Artigos() As Artigos
        Get
            If IsNothing(objArtigos) Then
                objArtigos = New Artigos(objMotor)
            End If
            Return objArtigos
        End Get
    End Property


    Public ReadOnly Property ArtigosPrecos() As ArtigosPrecos
        Get
            If IsNothing(objArtigosPrecos) Then
                objArtigosPrecos = New ArtigosPrecos(objMotor)
            End If
            Return objArtigosPrecos
        End Get
    End Property


    Public ReadOnly Property ArtigosArmazens() As ArtigosArmazens
        Get
            If IsNothing(objArtigosArmazens) Then
                objArtigosArmazens = New ArtigosArmazens(objMotor)
            End If
            Return objArtigosArmazens
        End Get
    End Property


    Public ReadOnly Property TaxasIva() As TaxasIva
        Get
            If IsNothing(objTaxasIva) Then
                objTaxasIva = New TaxasIva(objMotor)
            End If
            Return objTaxasIva
        End Get
    End Property


    Public ReadOnly Property Vendas() As DocumentosVenda
        Get
            If IsNothing(objVendas) Then
                objVendas = New DocumentosVenda(objMotor)
            End If
            Return objVendas
        End Get
    End Property


    Public ReadOnly Property CargasDescargas() As CargasDescargas
        Get
            If IsNothing(objCargasDescargas) Then
                objCargasDescargas = New CargasDescargas(objMotor)
            End If
            Return objCargasDescargas
        End Get
    End Property


    Public ReadOnly Property Internos() As DocumentosInternos
        Get
            If IsNothing(objInternos) Then
                objInternos = New DocumentosInternos(objMotor)
            End If
            Return objInternos
        End Get
    End Property


    Public ReadOnly Property Pendentes() As Pendentes
        Get
            If IsNothing(objPendentes) Then
                objPendentes = New Pendentes(objMotor, connectionString)
            End If
            Return objPendentes
        End Get
    End Property


    Public ReadOnly Property ExtratosCCT() As ExtratosCCT
        Get
            If IsNothing(objExtratosCCT) Then
                objExtratosCCT = New ExtratosCCT(objMotor, connectionString)
            End If
            Return objExtratosCCT
        End Get
    End Property


    Public ReadOnly Property Contadores() As Contadores
        Get
            If IsNothing(objContadores) Then
                objContadores = New Contadores(objMotor)
            End If
            Return objContadores
        End Get
    End Property


    Public ReadOnly Property Clientes() As Clientes
        Get
            If IsNothing(objClientes) Then
                objClientes = New Clientes(objMotor, connectionString)
            End If
            Return objClientes
        End Get
    End Property


    Public ReadOnly Property EntidadesAssociadas() As EntidadesAssociadas
        Get
            If IsNothing(objEntidadesAssociadas) Then
                objEntidadesAssociadas = New EntidadesAssociadas(objMotor)
            End If
            Return objEntidadesAssociadas
        End Get
    End Property


    Public ReadOnly Property OutrosTerceiros() As OutrosTerceiros
        Get
            If IsNothing(objOutrosTerceiros) Then
                objOutrosTerceiros = New OutrosTerceiros(objMotor, connectionString)
            End If
            Return objOutrosTerceiros
        End Get
    End Property


    Public ReadOnly Property TiposEntidade() As TiposEntidade
        Get
            If IsNothing(objTiposEntidade) Then
                objTiposEntidade = New TiposEntidade
            End If
            Return objTiposEntidade
        End Get
    End Property


    Public ReadOnly Property Moedas() As Moedas
        Get
            If IsNothing(objMoedas) Then
                objMoedas = New Moedas(objMotor)
            End If
            Return objMoedas
        End Get
    End Property


    Public ReadOnly Property Nacionalidades() As Nacionalidades
        Get
            If IsNothing(objNacionalidades) Then
                objNacionalidades = New Nacionalidades(objMotor)
            End If
            Return objNacionalidades
        End Get
    End Property



    Protected Overrides Sub Finalize()
        If Not IsNothing(objArtigos) Then
            objArtigos = Nothing
        End If
        If Not IsNothing(objArtigosPrecos) Then
            objArtigosPrecos = Nothing
        End If
        If Not IsNothing(objArtigosArmazens) Then
            objArtigosArmazens = Nothing
        End If
        If Not IsNothing(objTaxasIva) Then
            objTaxasIva = Nothing
        End If
        If Not IsNothing(objVendas) Then
            objVendas = Nothing
        End If
        If Not IsNothing(objCargasDescargas) Then
            objCargasDescargas = Nothing
        End If
        If Not IsNothing(objInternos) Then
            objInternos = Nothing
        End If
        If Not IsNothing(objPendentes) Then
            objPendentes = Nothing
        End If
        If Not IsNothing(objOutrosTerceiros) Then
            objOutrosTerceiros = Nothing
        End If
        If Not IsNothing(objEntidadesAssociadas) Then
            objEntidadesAssociadas = Nothing
        End If
        If Not IsNothing(objClientes) Then
            objClientes = Nothing
        End If
        If Not IsNothing(objTiposEntidade) Then
            objTiposEntidade = Nothing
        End If
        If Not IsNothing(objContadores) Then
            objContadores = Nothing
        End If
        If Not IsNothing(objMoedas) Then
            objMoedas = Nothing
        End If
        If Not IsNothing(objMotor) Then
            If objMotor.Contexto.EmpresaAberta = True Then
                objMotor.FechaEmpresaTrabalho()
            End If
            objMotor = Nothing
        End If
        MyBase.Finalize()
    End Sub

End Class



