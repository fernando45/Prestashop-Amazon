Imports System.IO
Imports System.Net
Imports System.Linq
Imports System.Xml



Public Class Product
    Const DirApi As String = "https://www.topkit.es/api"
    Const key As String = "?ws_key=C23D8TF675W562UCJDL27J3UU4AMR4V8"
    Const keyI As String = "?ws_key=28849WPK23C9KL8FESUYDF5H9YTQJL4I"
    Private _doc As String

    'DATOS DEL PRODUCTO
    Public ean As String
    Public name As String
    Public quantity As String 'stock
    Public price As String
    Public rate As String
    Public reference As String 'SKU
    Public description As String
    Public brand As String

    'VARIANTES (COMBINACIONES DE COLORES).
    Public variantes As New List(Of Variante)


    'DIMENSIONES Y EMLALAJE LOGISTICO
    Public Logis As New List(Of packanging)


    'DIMENSIONES DEL PRODUCTO MONTADO
    Public heightM As String, widthM As String, depthM As String, weightM As String

    'FEATURES, PROPIEDADES
    Public Features As Hashtable

    'Imagenes, url de las imagenes
    Public Imagenes As New List(Of String)

    'Categorias
    Public Categorias As New List(Of String)

    Private Function DameValor(ByVal xml As String, ByVal filtro As String) As String
        Dim valor As String = ""
        Try
            Dim m_xmld As XmlDocument
            Dim m_nodelist As XmlNodeList
            Dim m_node As XmlNode


            'Creamos el "Document"
            m_xmld = New XmlDocument()

            'Cargamos el archivo
            m_xmld.LoadXml(xml)



            'Obtenemos la lista de los nodos "name"
            m_nodelist = m_xmld.SelectNodes(filtro)

            'Iniciamos el ciclo de lectura
            For Each m_node In m_nodelist

                valor = m_node.ChildNodes.Item(0).InnerText

            Next
        Catch ex As Exception
            'Error trapping
            Console.Write(ex.ToString())
        End Try
        Return valor
    End Function
    Private Function DameVDimension(ByVal id As String) As String
        Dim valor As String = ""
        Try
            Dim Xml As String = DameXml("https://www.topkit.es/api/product_feature_values/" & id & key)
            valor = DameValor(Xml, "prestashop/product_feature_value/value")

        Catch ex As Exception
            Dim ola As String = ex.Message
        End Try
        Return valor
    End Function
    Private Sub DameDimensiones(ByVal doc As XElement)
        Try

            Dim tests As IEnumerable(Of XElement) =
            From el In doc...<id_feature_value> Select el
            Dim c As Integer = 1
            For Each ele As XElement In tests
                Dim valor As String = CStr(ele).ToString
                Dim aver As String = DameVDimension(valor)

                If aver.Contains(";") Then
                    Dim w As String() = aver.Split(New Char() {";"c})
                    Dim word As String
                    Dim p As New packanging
                    Dim a As Integer = 1
                    For Each word In w

                        Select Case a
                            Case 1
                                p.height = word
                            Case 2
                                p.width = word
                            Case 3
                                p.depth = word
                            Case 4
                                p.weight = word
                        End Select
                        a += 1
                    Next
                    Logis.Add(p)
                End If
            Next
        Catch ex As Exception
            'Error trapping
            Console.Write(ex.ToString())
        End Try

    End Sub
    Private Function DameValorL(ByRef doc As XElement, ByVal filtro As String) As String
        Dim valor As String = ""

        Try

            Dim employees As IEnumerable(Of XElement) = doc.Elements()

            For Each employee In employees
                valor = employee.Element(filtro).Value
            Next employee


        Catch ex As Exception
            Dim ola As String = ex.Message
        End Try
        Return valor
    End Function
    Private Sub EliminaArchivo(ByVal dir As String)
        If File.Exists(dir) Then
            File.Delete(dir)
        End If
    End Sub
    Private Sub DameFeatures()
        Features = New Hashtable
        Features.Add("Origen", "España")
        Features.Add("Fabricante", "TopKit")
    End Sub

    Private Function DameCategoria(ByVal id As String) As String
        Dim valor As String = ""
        Try
            Dim Xml As String = DameXml(DirApi & "/categories/" & id & key)
            valor = DameValor(Xml, "prestashop/category/name")

        Catch ex As Exception
            Dim ola As String = ex.Message
        End Try
        Return valor
    End Function
    Private Sub DameCategorias(ByRef doc As XElement)
        Dim valor As String
        Try
            Dim cate As IEnumerable(Of XElement) = doc.Elements()
            For Each categoria In cate.Descendants("category")
                valor = CStr(categoria)
                Dim cat As String = DameCategoria(valor)
                Categorias.Add(cat)
                Dim ola As String = ""
            Next categoria
        Catch ex As Exception
            Dim ola As String = ex.Message
        End Try

    End Sub

    Private Sub DameImagenes(ByVal doc As XElement)
        Dim valor As String
        Try
            Dim ima As IEnumerable(Of XElement) = doc.Elements()
            For Each imagen In ima.Descendants("image")
                valor = CStr(imagen)
                Dim url As String = DirApi & "/images/products/4/" & valor & keyI
                Imagenes.Add(url)
            Next imagen
        Catch ex As Exception
            Dim ola As String = ex.Message
        End Try
    End Sub

    Private Function DameListaVariantes(ByVal doc As XElement) As List(Of String)
        Dim lista As New List(Of String)
        Try
            ' Dim root As XElement = XElement.Load("TestConfig.xml")
            Dim tests As IEnumerable(Of XElement) =
            From el In doc...<combination> Select el

            For Each ele As XElement In tests
                Dim valor As String = CStr(ele).ToString
                lista.Add(valor)
            Next
        Catch ex As Exception
            'Error trapping
            Console.Write(ex.ToString())
        End Try

        Return lista

    End Function
    Private Function DameColor(ByVal id As String) As String
        Dim valor As String = ""

        Dim Xml As String = DameXml(DirApi & "/product_option_values/" & id & key)

        valor = DameValor(Xml, "prestashop/product_option_value/name")

        Return valor
    End Function
    Private Sub DameVairante(ByVal id As String)

        Dim Xml As String = DameXml(DirApi & "/combinations/" & id & key)

        Dim doc As XElement = XElement.Parse(Xml)
        Dim p As New Variante


        p.ean13 = DameValorL(doc, "ean13") 'El ean
        p.reference = DameValorL(doc, "reference") 'El ean
        p.quantity = DameValorL(doc, "quantity")
        p.name = "Color"
        Dim valor As String = ""
            Dim tests As IEnumerable(Of XElement) =
            From el In doc...<product_option_value> Select el

            For Each ele As XElement In tests
                valor = CStr(ele).ToString
            Next
        p.value = DameColor(valor)

        variantes.Add(p)

    End Sub

    Private Sub DameVariantes(ByVal doc As XElement)

        Dim lista As List(Of String) = DameListaVariantes(doc)



        Try
            If lista.Count > 0 Then
                For Each pro In lista
                    Dim val As String = pro

                    DameVairante(val)
                Next
            End If
        Catch ex As Exception
            'Error trapping
            Console.Write(ex.ToString())
        End Try


    End Sub
    Private Function DameStock(ByVal doc As XElement) As String
        Dim valor As String = ""
        Try
            Dim cate As IEnumerable(Of XElement) = doc.Elements()
            For Each stock In cate.Descendants("stock_available")
                Dim id As String = stock.Element("id")

                Dim x As String = DameXml(DirApi & "/stock_availables/" & id & key)
                Dim d As XElement = XElement.Parse(x)
                valor = DameValorL(d, "quantity")
                Exit For
            Next stock
        Catch ex As Exception
            Dim ola As String = ex.Message
        End Try
        Return valor
    End Function

    Public Function DameProducto(ByVal id As String) As Boolean

        Dim Pasa As Boolean = False
        If id = "1" Then
            Dim olas As String = ""

        End If

        Try
            ' Esto es para metodo de busqueda con link
            Dim Xml As String = DameXml(DirApi & "/products/" & id & key)

            Dim doc As XElement = XElement.Parse(Xml)
            _doc = doc

            Dim active As String, categoria As String = "", id_combi As String = ""
            ' DameDimensiones(doc)

            '***** COGEMOS LOS DATOS DEL PRODUCTO ************************

            'AQUI FILTRAMOS LOS QUE NO TIENEN QUE IR *********************

            'El ean temos que obtenerlo a traves de la variante por defecto
            id_combi = DameValorL(doc, "id_default_combination")
            If Not IsNumeric(id_combi) Then
                Return False
            End If
            Dim X As String = DameXml(DirApi & "/combinations/" & id_combi & key)
            Dim d As XElement = XElement.Parse(X)
            ean = DameValorL(d, "ean13") 'El ean


            active = DameValorL(doc, "active") ' Si el producto es activo
            If active = 0 Then Return False
            categoria = DameValorL(doc, "id_category_default") ' Si el id de categoria es 17 hay que salirse
            If categoria = "17" Then Return False
            'id_combi = DameValorL(doc, "id_combination_default")

            If Not IsNumeric(ean) Then Return False
            If ean = 0 Then Return False
            brand = DameValorL(doc, "manufacturer_name") 'El ean
            If brand <> "TopKit" Then Return False
            'FIN DEL FILTO**************************************************


            ' name = DameValor(Application.StartupPath & "\producto.xml", "prestashop/product/meta_title")
            name = DameValor(Xml, "prestashop/product/meta_title")
            quantity = DameStock(doc) 'El ean
            price = DameValorL(doc, "price") 'El ean
            rate = "21"
            reference = DameValorL(doc, "reference") 'El ean
            description = DameValorL(doc, "description") 'El ean


            '***** COGEMOS LOS DATOS DIMENSIONES DEL PRODUCTO MONTADO   ***********************
            heightM = DameValorL(doc, "height") 'El ean
            widthM = DameValorL(doc, "width") 'El ean
            depthM = DameValorL(doc, "depth") 'El ean
            weightM = DameValorL(doc, "weight") 'El ean

            '***** COGEMOS LOS DATOS DIMENSIONES LOGISTICAS ***********************
            DameDimensiones(doc)

            '***** RELLENAMOS UN HASTABLE CON LAS PROPIEDADES Y LO DEJAMOS DISPONIBLE.
            DameFeatures()

            '***** RELLENAMOS UNA LISTA CON LAS URL DE LAS IMAGENES Y LO DEJAMOS DISPONIBLE.
            DameImagenes(doc)

            '***** RELLENAMOS UNA LISTA CON LAS CATEGORIAS Y LO DEJAMOS DISPONIBLE.
            DameCategorias(doc)

            '************  VARIANTES lo hacemos con un datatable. ******************
            DameVariantes(doc)

            Pasa = True

        Catch ex As Exception
            Dim ola As String = ""

        End Try

        Return Pasa

    End Function
    Private Function DameXml(ByVal Dir As String) As String

        Try


            Dim request As Net.HttpWebRequest
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
    Private Sub Carga()

    End Sub

End Class
