Imports Microsoft.VisualBasic
Imports CCPrimavera.Specifications.LinhaDocumentoInterno_Specs


Namespace BE

    Public Class LinhasDocumentosInternosCollection

        Inherits System.Collections.Specialized.NameObjectCollectionBase


        Public Overridable Function Add(ByVal LinhaDocumentoInterno As BE.LinhaDocumentoInterno, ByVal key As String) As Integer
            MyBase.BaseSet(key, LinhaDocumentoInterno)
            Me(MyBase.Count - 1).NumeroLinha = (MyBase.Count - 1 + 1).ToString.TrimEnd
        End Function


        Public Overridable Function AddAndGetNewCollection(ByVal LinhaDocumentoInterno As LinhaDocumentoInterno, ByVal index As System.Int32) As LinhasDocumentosInternosCollection
            Dim i As System.Int32
            Dim mReorganizedCollection As LinhasDocumentosInternosCollection

            mReorganizedCollection = New LinhasDocumentosInternosCollection

            If index > Me.Count - 1 Then
                Me.Add(LinhaDocumentoInterno, index.ToString.TrimEnd)
                AddAndGetNewCollection = Me
                Exit Function
            End If

            For i = 0 To MyBase.Count - 1
                Select Case i
                    Case Is < index
                        mReorganizedCollection.Add(Me(i), i.ToString.TrimEnd)
                        mReorganizedCollection(i).NumeroLinha = (i + 1) 'A base do Numerador de Linha da Primavera é 1
                    Case Is = index
                        mReorganizedCollection.Add(LinhaDocumentoInterno, i.ToString.TrimEnd)
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


        Default Public Overridable Property Item(ByVal index As Integer) As BE.LinhaDocumentoInterno
            Get
                Return DirectCast(MyBase.BaseGet(index), BE.LinhaDocumentoInterno)
            End Get
            Set(ByVal value As BE.LinhaDocumentoInterno)
                MyBase.BaseSet(index, value)
            End Set
        End Property


        Default Public Overridable Property Item(ByVal key As String) As BE.LinhaDocumentoInterno
            Get
                Return DirectCast(MyBase.BaseGet(key), BE.LinhaDocumentoInterno)
            End Get
            Set(ByVal value As BE.LinhaDocumentoInterno)
                MyBase.BaseSet(key, value)
            End Set
        End Property


        Public Overridable Sub Remove(ByVal key As String)
            MyBase.BaseRemove(key)
        End Sub


        Public Overridable Sub RemoveAt(ByVal index As Integer)
            MyBase.BaseRemoveAt(index)
        End Sub


        Public Overridable Function RemoveAtAndGetNewCollection(ByVal index As System.Int32) As LinhasDocumentosInternosCollection
            Dim i As System.Int32
            Dim mReorganizedCollection As LinhasDocumentosInternosCollection

            mReorganizedCollection = New LinhasDocumentosInternosCollection

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



        Public Function List() As System.Collections.Generic.List(Of BE.LinhaDocumentoInterno)
            Dim mList As New System.Collections.Generic.List(Of BE.LinhaDocumentoInterno)
            Dim i As System.Int32

            For i = 0 To MyBase.Count - 1
                mList.Add(Me.Item(i))
            Next

            List = mList

            mList = Nothing
        End Function



        Public Sub NormalizaLinhas(ByVal QtdMinimaLinhas As System.Int16)
            Dim objLinhaDocumentoInterno As BE.LinhaDocumentoInterno
            Dim i As System.Int16

            If QtdMinimaLinhas > Me.Count Then
                For i = Me.Count To QtdMinimaLinhas
                    objLinhaDocumentoInterno = New BE.LinhaDocumentoInterno
                    With objLinhaDocumentoInterno
                        .TipoLinha = TiposLinhas.Comentario_60
                    End With
                    Me.Add(objLinhaDocumentoInterno, objLinhaDocumentoInterno.Id)
                Next
            End If

            objLinhaDocumentoInterno = Nothing
        End Sub


    End Class

End Namespace

