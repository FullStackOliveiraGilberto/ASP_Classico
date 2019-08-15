

<%
    Class Category
        Private NameVar

        Public Property Get Name()
            Name = NameVar
        End Property

        Public Property Let Name(nameParam)
            NameVar = nameParam
        End Property
    End Class

    Class Item
        Private NameVar
        Private ItemVar

        Public Property Get Name()
            Name = NameVar
        End Property

        Public Property Let Name(nameParam)
            NameVar = nameParam
        End Property

        Public Property Get Item()
            Item = ItemVar
        End Property

        Public Property Let Item(itemParam)
            ItemVar = itemParam
        End Property
    End Class

    Dim CategoryVar
    Set CategoryVar = New Category

    CategoryVar.Name = "Weight"

    Dim ItemVar
    Set ItemVar = New Item

    ItemVar.Name = "kg"
	ItemVar.Item = "motor de passe para carro de corrida"
    'ItemVar.Category = CategoryVar ' There is no 'Category' property in your class
%>

<html>
    <head>
        <title>UoM Componet Testing</title>
    </head>
    <body>
        <%= ItemVar.Name %><br/>
		<%= ItemVar.Item %><br/>
		<%= CategoryVar.Name %><br/>
		
    </body>
</html>