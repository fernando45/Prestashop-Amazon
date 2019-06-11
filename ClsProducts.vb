Imports System.IO
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

            '*** Recorremos la lista de productos y vamos rellenando
            ' writer.WriteStartElement("products")

            For Each Pro In Productos
                Dim Producto As New Product
                Dim pasa As Boolean = Producto.DameProducto(Pro)
                If pasa Then 'Si el producto ha pasado y no ha sido filtrado

                        'Primero el nodo proudctos
                        With writer

                            .WriteStartElement("product") ' iniciamos un nuevo producto


                            ' Los datos del producto
                            .WriteElementString("ean13", Producto.ean)
                            .WriteElementString("name", Producto.name)
                            .WriteElementString("quantity", Producto.quantity)
                            .WriteElementString("price", Producto.price)
                            .WriteElementString("rate", Producto.rate)
                            .WriteElementString("reference", Producto.reference)

                            writer.WriteStartElement("description")
                            writer.WriteCData(Producto.description)
                            writer.WriteEndElement()

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
                            .WriteStartElement("dimension")
                            .WriteElementString("height", Producto.heightM)
                            .WriteElementString("width", Producto.widthM)
                            .WriteElementString("depth", Producto.depthM)
                            .WriteElementString("weight", Producto.weightM)
                            .WriteEndElement()

                            'Packaging
                            .WriteStartElement("packaging")
                            .WriteElementString("height", 0)
                            .WriteElementString("width", 0)
                            .WriteElementString("depth", 0)
                            .WriteElementString("weight", 0)
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


