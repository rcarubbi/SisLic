<%
Function addItemPedido(descricao, qt, vlUnitario, idPropostaSubItem)
	dim itens
	itens = session("pedidoItens")
	
	'retorna proximo indice
	redim preserve itens(5, ubound(itens,2) + 1)
	indice = ubound(itens,2)
	
	'salva os dados
	itens(0,indice) = indice
	itens(1,indice) = descricao
	itens(2,indice) = cDbl(vlUnitario)	
	itens(3,indice) = cDbl(qt)
	itens(4,indice) = qt * vlUnitario
	itens(5,indice) = idPropostaSubItem
	
	session("pedidoItens") = itens
	addItemPedido = indice
end function

Function EditItemPedido(indice, descricao, qt, VlUnitario, idPropostaSubItem)
	dim itens
	itens = session("pedidoItens")
	
	'salva os dados
	itens(0,indice) = indice
	itens(1,indice) = descricao
	itens(2,indice) = cDbl(vlUnitario)	
	itens(3,indice) = cDbl(fix(qt))
	itens(4,indice) = qt * vlUnitario
	itens(5,indice) = idPropostaSubItem
	
	session("pedidoItens") = itens
	
	EditItemPedido = indice
end function

Function DelItemPedido(indice)
	dim itens, i
	itens = session("pedidoItens")
	
	'se deletar um item no meio da matriz
	'passar todos os itens para frente
	'e depois excluir o ultimo item
	for i = indice to ubound(itens,2) -1
		
		itens(0,i) = i
		itens(1,i) = itens(1,i+1)
		itens(2,i) = itens(2,i+1)
		itens(3,i) = itens(3,i+1)
		itens(4,i) = itens(4,i+1)
		itens(5,i) = itens(5,i+1)

	next
	
	redim preserve itens(5, ubound(itens,2)-1)

	session("pedidoItens") = itens
		
	'retorna o novo tamanho da matriz
	DelItemPedido = ubound(itens,2)
end function

sub SalvaNovoPedido(idProcesso, Obs, Fornecedor, dtPedido, comprador, formaPagto, prazoEntrega)
	dim sql
			
	SQL = "INSERT INTO Pedido " & _
			   "(DtPedido" & _
			   ",vlTotal" & _
			   ",idFornecedor" & _
			   ",idProcesso" & _
			   ",obs" & _
			   ",comprador" & _
			   ",formaPagto" & _
			   ",prazoEntrega)" & _
           "VALUES" & _
			   "('" & dtPedido & " " & time & "' " & _
			   "," & replace(rtnVlTotal,",",".") & " " & _
			   "," & Fornecedor & " " & _
			   "," & iif(idProcesso = "", "NULL", idProcesso) & _
			   ",'" & obs & "' " & _
			   ",'" & comprador & "' " & _
			   ",'" & formaPagto & "' " & _
			   ",'" & prazoEntrega & "')"
			
	conn.execute(SQL)
	
	SalvaPedidoItens(maxID("Pedido", "NumPedido"))
	

end sub

sub SalvaPedidoItens(NumPedido)
	dim sql
	dim itens, i
	
	itens = session("pedidoItens")
	
	for i = 0 to ubound(itens,2)
		sql = "INSERT INTO PedidoItem " & _
			   "(NumPedido " & _
			   ",idPropostaSubItem " & _			   
			   ",descricao " & _
			   ",quantidade " & _
			   ",vlUnitario) " & _
			   "VALUES " & _
			   "(" & NumPedido & " " & _
			   "," & itens(5,i) & " " & _			   
			   ",'" & itens(1,i) & "' " & _
			   "," & replace(itens(3,i), ",",".") & " " & _
			   "," & replace(itens(2,i), ",",".") & ")"

		conn.execute(SQL)
		
		'quando um item é salvo
		'o estoque é automaticamente atualizado pela trigger T_PedidoAtualizaEstoque
		
	next   
		   
end sub

function rtnVlTotal
	dim i, tot, itens
	
	itens = session("pedidoItens")
	for i = 0 to ubound(itens,2)
		tot = tot + (itens(3,i) * itens(2,i))
	next
	rtnVlTotal = tot
end function

sub pedidoProcessa(NumPedido, nf, vlImposto, dtRecebimento, hrRecebimento, obs, codPlanoContaImposto, codPlanoContaNota)
	dim SQL
	
	sql = "UPDATE pedido set " & _
				"NumNFCompra=" & nf & ", " & _
				"vlImpostos=" & replace(replace(vlImposto,".",""),",",".") & ", " & _
				"dtRecebimento='" & dtRecebimento & " " & hrRecebimento & "', " & _
				"obs='" & replace(obs, "'", "''") & "', " & _		
				"recebido=1, " & _					
				"codPlanoContaImposto=" & codPlanoContaImposto & ", " & _
				"codPlanoContaNota=" & codPlanoContaNota & " " & _
				"WHERE numPedido=" & numPedido
	'RESPONSE.WRITE SQL
	'RESPONSE.END
	conn.execute(SQL)							

end sub

sub PedidoCancela(NumPedido)
	dim sql
	sql = "update pedido set cancelado=1 where numPedido=" & numPedido
	conn.execute(SQL)
end sub

sub PedidoCancelaRecebido(NumPedido, motivoCancelamento)
	dim sql
	' cancelando os lançamentos do fluxo de caixa referentes a este pedido
	sql = "UPDATE fluxoCaixa SET dataCancelamento = getdate(), cancelado = 1, motivoCancelamento = '" & motivoCancelamento & "' where idRelacional = " & numPedido & " and origem =2"
	
	conn.execute(sql)
	' cancelando o pedido
	sql = "Update pedido set cancelado = 1 where numPedido = " & numPedido
	conn.execute(sql)
end sub
%>