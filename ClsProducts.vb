﻿Imports System.IO
Imports System.Net
Imports System.Xml
Imports System.Xml.Linq
Public Class ClsProducts

    Const DirApi As String = "https://www.topkit.es/api"
    Const key As String = "?ws_key=C23D8TF675W562UCJDL27J3UU4AMR4V8"


    Public Product As New Hashtable


    Public Function DameProcuctos() As List(Of String)
        ' ****  Obtenemos la lista de productos para irla recorriendo ****

        Dim proc As New Product
        Dim Lista As New List(Of String)

        Try


            Dim nodoraiz As XElement
            'Dim Xml As String = DameXml(DirApi & "/products" & key)
            Dim Xml As String = DameXml(DirApi & "/products" & key)
            Dim doc As XElement = XElement.Parse(Xml)
            nodoraiz = doc.Element("products")
            Dim rfc As String
            For Each v In nodoraiz.Elements("product")
                rfc = v.Attribute("id").Value
                Lista.Add(rfc)
            Next
        Catch ex As Exception
            Dim ola As String = ""
        End Try

        Return Lista
    End Function

    Public Sub Construye()

        Try

            Dim settings As XmlWriterSettings = New XmlWriterSettings()
            Using writer As XmlWriter = XmlWriter.Create(Application.StartupPath & "\productsE.xml", settings)

                settings.Indent = True ' Begin writing.
                settings.IndentChars = vbTab
                writer.WriteStartDocument()
                writer.WriteStartElement("products") ' Root.


                '*************  OBTENEMOS LA LISTA DE ID DE LOS PRODUCTOS *********
                Dim Productos As List(Of String) = DameProcuctos()

                Dim Cuantos As Integer = Productos.Count
                '    FRM.IniciarProceso()

                '*** Recorremos la lista de productos y vamos rellenando
                ' writer.WriteStartElement("products")
                Dim n As Integer = 0
                For Each Pro In Productos
                    '  FRM.proc = n
                    Dim Producto As New Product
                    Dim pasa As Boolean = Producto.DameProducto(Pro)
                    If pasa Then 'Si el producto ha pasado y no ha sido filtrado

                        'Primero el nodo proudctos
                        With writer

                            .WriteStartElement("product") ' iniciamos un nuevo producto


                            ' Los datos del producto
                            .WriteElementString("ean13", Producto.ean)

                            'nombres de los productos
                            .WriteStartElement("name")
                            Dim x As Integer = 1
                            For Each na In Producto.name
                                Select Case x
                                    Case 1
                                        .WriteElementString("español", na)
                                    Case 2
                                        .WriteElementString("ingles", na)
                                    Case 3
                                        .WriteElementString("italiano", na)
                                    Case 4
                                        .WriteElementString("frances", na)
                                End Select
                                x = x + 1
                            Next

                            .WriteEndElement()


                            .WriteElementString("quantity", Producto.quantity)
                            .WriteElementString("price", Producto.price)
                            .WriteElementString("rate", Producto.rate)
                            .WriteElementString("reference", Producto.reference)


                            'descripciones
                            .WriteStartElement("description")
                            .WriteStartElement("español")
                            .WriteCData(Producto.description.Item(0))
                            .WriteEndElement()
                            .WriteStartElement("ingles")
                            .WriteCData(Producto.description.Item(1))
                            .WriteEndElement()
                            .WriteStartElement("italiano")
                            .WriteCData(Producto.description.Item(2))
                            .WriteEndElement()
                            .WriteStartElement("frances")
                            .WriteCData(Producto.description.Item(3))
                            .WriteEndElement()
                            .WriteEndElement()



                            .WriteElementString("brand", Producto.brand)

                            'imagenes
                            .WriteStartElement("images")
                            For Each ima In Producto.Imagenes
                                .WriteElementString("image", ima)
                            Next
                            writer.WriteEndElement()

                            'Categorias
                            .WriteStartElement("categories")
                            For Each cat In Producto.Categorias
                                .WriteElementString("category", cat)
                            Next
                            writer.WriteEndElement()

                            'Dimensiones de producto montado
                            .WriteStartElement("packaging")
                            If Producto.Logis.Count = 1 Then
                                For Each t In Producto.Logis
                                    .WriteElementString("height", t.height)
                                    .WriteElementString("width", t.width)
                                    .WriteElementString("depth", t.depth)
                                    .WriteElementString("weight", t.weight)
                                Next
                            Else
                                .WriteElementString("height", 0)
                                .WriteElementString("width", 0)
                                .WriteElementString("depth", 0)
                                .WriteElementString("weight", 0)
                            End If
                            .WriteEndElement()

                            'Packaging
                            .WriteStartElement("dimension")
                            .WriteElementString("height", Producto.heightM)
                            .WriteElementString("width", Producto.widthM)
                            .WriteElementString("depth", Producto.depthM)
                            .WriteElementString("weight", Producto.weightM)
                            .WriteEndElement()

                            'features
                            .WriteStartElement("features")
                            For Each elemento As DictionaryEntry In Producto.Features
                                .WriteStartElement("feature")
                                .WriteElementString("name", elemento.Key)
                                .WriteElementString("value", elemento.Value)
                                .WriteEndElement()
                            Next
                            .WriteEndElement()

                            'variations


                            If Producto.variantes.Count > 0 Then
                                .WriteStartElement("variations")
                                For Each oRow In Producto.variantes
                                    .WriteStartElement("variation")
                                    .WriteElementString("ean13", oRow.ean13)
                                    .WriteElementString("reference", oRow.reference)
                                    .WriteElementString("quantity", oRow.quantity)
                                    .WriteStartElement("attributes")
                                    .WriteStartElement("attritute")
                                    .WriteElementString("name", oRow.name)
                                    .WriteElementString("value", oRow.value)
                                    .WriteEndElement()
                                    .WriteEndElement()
                                    .WriteEndElement()
                                Next
                                writer.WriteEndElement()
                            End If

                        End With
                        writer.WriteEndElement()
                    End If
                    n += 1
                    '  FRM.o.Text = n
                Next Pro
                writer.WriteEndElement()
                writer.WriteEndDocument()
                writer.Flush()

            End Using
        Catch ex As Exception
            Dim ola As String = ""
        End Try


    End Sub

    Public Function DameXml(ByVal Dir As String) As String

        Try


            Dim request As HttpWebRequest
            Dim response As HttpWebResponse = Nothing
            Dim reader As StreamReader


            request = DirectCast(WebRequest.Create(Dir), HttpWebRequest)
            request.Method = "GET"
            request.KeepAlive = False
            request.Timeout = 15 * 1000

            response = DirectCast(request.GetResponse(), HttpWebResponse)
            reader = New StreamReader(response.GetResponseStream())

            Dim rawresp As String
            rawresp = reader.ReadToEnd()
            Return rawresp
            Dim ola As String = ""
        Catch ex As Exception
            Dim ola As String = ex.Message
        End Try


    End Function




End Class


