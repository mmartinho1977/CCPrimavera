Imports Microsoft.VisualBasic
Imports CCPrimavera.Specifications.LinhaDocumentoVenda_Specs


Namespace BE

    Public Class LinhasDocumentosVendaCollection

        Inherits System.Collections.Specialized.NameObjectCollectionBase


        Public Overridable Function Add(ByVal LinhaDocumentoVenda As BE.LinhaDocumentoVenda, ByVal key As String) As Integer
            MyBase.BaseSet(key, LinhaDocumentoVenda)
            Me(MyBase.Count - 1).NumeroLinha = (MyBase.Count - 1 + 1).ToString.TrimEnd
        End Function


        Public Overridable Function AddAndGetNewCollection(ByVal LinhaDocumentoVenda As LinhaDocumentoVenda, ByVal index As System.Int32) As LinhasDocumentosVendaCollection
            Dim i As System.Int32
            Dim mReorganizedCollection As LinhasDocumentosVendaCollection

            mReorganizedCollection = New LinhasDocumentosVendaCollection

            If index > Me.Count - 1 Then
                Me.Add(LinhaDocumentoVenda, index.ToString.TrimEnd)
                AddAndGetNewCollection = Me
                Exit Function
            End If

            For i = 0 To MyBase.Count - 1
                Select Case i
                    Case Is < index
                        mReorganizedCollection.Add(Me(i), i.ToString.TrimEnd)
                        mReorganizedCollection(i).NumeroLinha = (i + 1) 'A base do Numerador de Linha da Primavera é 1
                    Case Is = index
                        mReorganizedCollection.Add(LinhaDocumentoVenda, i.ToString.TrimEnd)
                        mReorganizedCollection(i).NumeroLinha = (i + 1)
                        mReorganizedCollection.Add(Me(i), (i + 1).ToString.TrimEnd)
                        mReorganizedCollection(i + 1).NumeroLinha = (i + 1 + 1)
                    Case Else
                        mReorganizedCollection.Add(Me(i), (i + 1).ToString.TrimEnd)
                        mReorganizedCollection(i + 1).NumeroLinha = (i + 1 + 1)
                End Select
            Next

            AddAndGetNewCollection = mReorganizedCollection

            mReorganizedCollection = Nothing
        End Function


        Default Public Overridable Property Item(ByVal index As Integer) As BE.LinhaDocumentoVenda
            Get
                Return DirectCast(MyBase.BaseGet(index), BE.LinhaDocumentoVenda)
            End Get
            Set(ByVal value As BE.LinhaDocumentoVenda)
                MyBase.BaseSet(index, value)
            End Set
        End Property


        Default Public Overridable Property Item(ByVal key As String) As BE.LinhaDocumentoVenda
            Get
                Return DirectCast(MyBase.BaseGet(key), BE.LinhaDocumentoVenda)
            End Get
            Set(ByVal value As BE.LinhaDocumentoVenda)
                MyBase.BaseSet(key, value)
            End Set
        End Property


        Public Overridable Sub Remove(ByVal key As String)
            MyBase.BaseRemove(key)
        End Sub


        Public Overridable Sub RemoveAt(ByVal index As Integer)
            MyBase.BaseRemoveAt(index)
        End Sub


        Public Overridable Function RemoveAtAndGetNewCollection(ByVal index As System.Int32) As LinhasDocumentosVendaCollection
            Dim i As System.Int32
            Dim mReorganizedCollection As LinhasDocumentosVendaCollection

            mReorganizedCollection = New LinhasDocumentosVendaCollection

            For i = 0 To MyBase.Count - 1
                Select Case i
                    Case Is < index
                        mReorganizedCollection.Add(Me(i), i.ToString.TrimEnd)
                        mReorganizedCollection(i).NumeroLinha = (i + 1)
                    Case Is = index
                        ' Não adicionar nada porque se trata do index a elimintar
                    Case Else
                        mReorganizedCollection.Add(Me(i), (i - 1).ToString.TrimEnd)
                        mReorganizedCollection(i - 1).NumeroLinha = (i + 1 - 1)
                End Select
            Next

            RemoveAtAndGetNewCollection = mReorganizedCollection

            mReorganizedCollection = Nothing
        End Function



        Public Function List() As System.Collections.Generic.List(Of BE.LinhaDocumentoVenda)
            Dim mList As New System.Collections.Generic.List(Of BE.LinhaDocumentoVenda)
            Dim i As System.Int32

            For i = 0 To MyBase.Count - 1
                mList.Add(Me.Item(i))
            Next

            List = mList

            mList = Nothing
        End Function



        Public Sub NormalizaLinhas(ByVal QtdMinimaLinhas As System.Int16)
            Dim objLinhaDocumentoVenda As BE.LinhaDocumentoVenda
            Dim i As System.Int16

            If QtdMinimaLinhas > Me.Count Then
                For i = Me.Count To QtdMinimaLinhas
                    objLinhaDocumentoVenda = New BE.LinhaDocumentoVenda
                    With objLinhaDocumentoVenda
                        .TipoLinha = TiposLinhas.Comentario_60
                    End With
                    Me.Add(objLinhaDocumentoVenda, objLinhaDocumentoVenda.Id)
                Next
            End If

            objLinhaDocumentoVenda = Nothing
        End Sub


    End Class

End Namespace

