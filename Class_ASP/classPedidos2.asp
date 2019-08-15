<%

'   Class Category
'        Private NameVar
'        Public Property Get Name()
'            Name = NameVar
'        End Property

        'Public Property Let Name(nameParam)
         '   NameVar = nameParam
        'End Property
'End Class


Class Pedido
'  nPedido
'  TemSuspensao
'  DataInicioContagem
'  DataSuspensao
'  DataReiniciocontAposSuspesao
'  StatusInicioContagem
'  DataPrevista  
    Private nPedidoVar
	Private TemSuspensaoVar
	Private DataInicioContagemVar
	Private DataSuspensaoVar
	Private DataReiniciocontAposSuspesaoVar
	Private StatusInicioContagemVar
	Private DataPrevistaVar
	
    Public Property Get nPedido()             nPedido    	  	 = nPedidoVar    		 End Property
	Public Property Get TemSuspensao()        TemSuspensao    	 = TemSuspensaoVar    	 End Property
	Public Property Get DataInicioContagem()  DataInicioContagem = DataInicioContagemVar End Property
 	Public Property Get DataSuspensao()       DataSuspensao    	 = DataSuspensaoVar      End Property

    Public Property let nPedido(nPedidoParam) 						nPedidoVar 				= nPedidoParam  			End Property
	Public Property let TemSuspensao(TemSuspensaoParam) 			TemSuspensaoVar 		= TemSuspensaoParam  		End Property
	Public Property let DataInicioContagem(DataInicioContagemParam) DataInicioContagemVar 	= DataInicioContagemParam   End Property
	Public Property let DataSuspensao(DataSuspensaoParam) 			DataSuspensaoVar 		= DataSuspensaoParam  		End Property
	
	public function classipedido()
		Dim nPedidoVar2
		Set nPedidoVar2 = New Pedido
		nPedidoVar2.nPedido = "7788994"
		nPedidoVar2.TemSuspensao = "nao"
		nPedidoVar2.DataInicioContagem = "25/11/2014"
		classipedido = nPedidoVar2
    return classipedido 
	
end function 
	
	
	
End Class


Dim nPedidoVar
Set nPedidoVar = New Pedido
nPedidoVar.nPedido = "78254"
nPedidoVar.TemSuspensao = "Sim"
nPedidoVar.DataInicioContagem = "01/12/2014"


public function classipedido2()
	Dim nPedidoVar2
	Set nPedidoVar2 = New Pedido
	nPedidoVar2.nPedido = "7788994"
	nPedidoVar2.TemSuspensao = "nao"
	nPedidoVar2.DataInicioContagem = "25/11/2014"
	classipedido = nPedidoVar2
return classipedido 
end function 
	
	
%>
	
<html>
    <head>
        <title>UoM Componet Testing</title>
    </head>
    <body>
        <%= nPedidoVar.nPedido %><br/>
		<%= nPedidoVar.TemSuspensao %><br/>
		<%= nPedidoVar.DataInicioContagem %><br/>
		
		
		<% ret = classipedido2() %>
		
		<%= ret.nPedido %><br/>
		<%= ret.TemSuspensao %><br/>
		<%= ret.DataInicioContagem %><br/>
		
		
    </body>
</html>

	
	
	
	
