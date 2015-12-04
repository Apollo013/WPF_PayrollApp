Imports System.Xml

Public Class Window1

    Private Sub Window_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        ComboBox1.Items.Add("Billy")
        ComboBox1.Items.Add("Mary")
        ComboBox1.Items.Add("Martha")
        ComboBox1.Items.Add("Bill")
        ComboBox1.Items.Add("Fred")
        ComboBox1.Items.Add("Jim")
        ComboBox1.Items.Add("James")
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles Button1.Click

        Dim writer As New XmlTextWriter("Employees.xml", System.Text.Encoding.UTF8)
        writer.WriteStartDocument(True)

        writer.Formatting = Formatting.Indented
        writer.Indentation = 4

        writer.WriteStartElement("Employees")

        createnode(1, "Billy Meyers", writer)
        createnode(2, "Jane Summers", writer)
        createnode(3, "Phillip Bishop", writer)
        createnode(4, "Marty Maguire", writer)
        createnode(5, "Henry Hill", writer)
        createnode(6, "Mary Jane", writer)

        writer.WriteEndElement()
        writer.WriteEndDocument()
        writer.Close()

    End Sub

    Private Sub createnode(ByVal id As Integer, ByVal aname As String, ByVal writer As XmlTextWriter)

        writer.WriteStartElement("Employee")
        writer.WriteAttributeString("ID", CStr(id))
        writer.WriteAttributeString("Name", aname)
        writer.WriteEndElement()

    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles Button2.Click
        ' Change Marty To Georgina !
        Dim xd As New XmlDocument
        xd.Load("Employees.xml")

        Dim node As XmlNode = xd.SelectSingleNode("/Employees/Employee[ID='4']")


        xd.Save("Employees.xml")
    End Sub
End Class
