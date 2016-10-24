<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="../comercio/funPedido.asp"-->
<!--#include file="funGeraEmpenho.asp"-->
<%call abreBanco%>
<%
dim idProcesso, erro, prazo, validade
dim dtPedido, obs, sql, rs
dim fornecedor, aux
dim dataEmpenho, Observacao, valorTotal
dim pedidoComprador, pedidoPrazoEntrega, pedidoFormaPagto
'verifica se tem algum item ganho selecionado
'pois se a sessão expirou, então tem que redirecionar para a tela inicial
set rs = tempGetItensGanhos
if rs.eof then
	rs.close
	set rs = nothing
	response.Redirect("./")
end if

'redimensiona as matrizes que irao conter as informações dos pedidos
'------------------------------------------------------------------->
sql = "SELECT max(idFornecedor) as maxID from tempEmpItemPedido"
set rs = conn.execute(SQL)
if isNUll(rs("maxID")) then
	redim dtPedido(-1)
	redim obs(-1)
	redim pedidoComprador(-1)
	redim pedidoPrazoEntrega(-1)
	redim pedidoFormaPagto(-1)
else
	redim dtPedido(rs("maxID"))
	redim obs(rs("maxID"))
	redim pedidoComprador(rs("maxID"))
	redim pedidoPrazoEntrega(rs("maxID"))
	redim pedidoFormaPagto(rs("maxID"))
	'colocando um valor inicial nas informações dos pedidos
	for i = 1 to rs("maxID") 
		obs(i) = application("par_24")
		pedidoComprador(i) = application("par_21")
		pedidoPrazoEntrega(i) = application("par_22")
		pedidoFormaPagto(i) = application("par_23")
	next
	
end if
rs.close
set rs = nothing
'<---------------------------------------------------------------------




	
prazo = request("prazo")
validade = request("validade")
dataEmpenho = request("dataEmpenho")
observacao = request("observacao")
valorTotal = request("valorTotal")
idProcesso = request("idProcesso")
'pedidoComprador = request("pedidoComprador")' application("par_21")
'pedidoPrazoEntrega = request("pedidoPrazoEntrega")' application("par_22")
'pedidoFormaPagto = request("pedidoFormaPagto") 'application("par_23")

if request.Form("isPostBack") = "1" then


	SQL = "Select idFornecedor, nome from tempEmpItemPedido Inner Join parceiro on parceiro.id=tempEmpItemPedido.idFornecedor where sessionID='" & session.SessionID & "' group by idFornecedor, nome"
	set rs = conn.execute(SQL)
	while not rs.eof
		dtPedido(rs("idFornecedor")) = recebe_datador("dtPedido_" & rs("idFornecedor"), 0)
		obs(rs("idFornecedor")) = request.Form("obs_" & rs("idFornecedor"))
		pedidoComprador(rs("idFornecedor")) = request.Form("pedidoComprador_" & rs("idFornecedor"))
		pedidoPrazoEntrega(rs("idFornecedor")) = request.Form("pedidoPrazoEntrega_" & rs("idFornecedor"))
		pedidoFormaPagto(rs("idFornecedor")) = request.Form("pedidoFormaPagto_" & rs("idFornecedor"))
		if not isDate(dtPedido(rs("idFornecedor"))) then erro = erro & "Data do pedido inválida para o fornecedor: " & rs("nome") & ". "
		rs.movenext
	wend
	
	
	rs.close
	set rs = nothing
	

	'verifica se houve alteração no estoque
	'----------------------------------------------------------->
	erro = erro & EstoqueConferencia
	'<----------------------------------------------------------	
	
	if erro = "" then
	
		'altera o estoque
		'--------------------------------------------------------------------->
		EstoqueAtualizar
		'<--------------------------------------------------------------------
		
	

		
		'fazendo os pedidos para cada fornecedor
		'--------------------------------------------------------->
		
		redim itens(5,-1)
		
		session("pedidoItens") = itens	
			
		sql = "Select * from tempEmpItemPedido where sessionID='" & session.SessionID & "' order by idFornecedor"
		set rs = server.CreateObject("ADODB.Recordset")
		rs.open sql, conn, 1
		if not rs.eof then
			'pegando o primeiro fornecedor
			aux = rs("idFornecedor")
			while not rs.eof
				fornecedor = rs("idFornecedor")
				
				'os itens tem que ser gravado em pedidos diferentes agrupados pelo fornecedor
				'então se o fornecedor atual é diferente do atenrior, então o pedido já é outro, então ja grava 
				'os atenriores				
				if aux <> fornecedor then
					
					SalvaNovoPedido idProcesso, obs(aux), aux, dtPedido(aux), pedidoComprador(aux),pedidoFormaPagto(aux),pedidoPrazoEntrega(aux)

					aux = rs("idFornecedor")
					
					'limpando os itens dos pedidos anteriores
					redim itens(5,-1)
					session("pedidoItens") = itens
				end if
				'grava o item
				addItemPedido rs("descricao"), rs("qtPedido"), rs("valorunitario"), PropostaSubItem(rs("idPropostaItem"), rs("qtProduto"), rs("descricao"))
				'insere o item como um subItem da proposta

				 
				rs.movenext
			wend
			
			rs.movelast
			SalvaNovoPedido idProcesso, obs(rs("idFornecedor")), rs("idFornecedor"), dtPedido(rs("idFornecedor")), pedidoComprador(rs("idFornecedor")),pedidoFormaPagto(rs("idFornecedor")),pedidoPrazoEntrega(rs("idFornecedor"))
								
			rs.close
			set rs = nothing
			'<--------------------------------------------------------------------

			
				
		end if	
		'salva os dados do empenho
		'-------------------------------------------------------------------->					
		GerarEmpenho idProcesso, dataEmpenho, valorTotal, prazo, validade, observacao
		'<-------------------------------------------------------------------
		
		
		
		'deletando as informações das tabelas temporarias para este usuario
		'-------------------------------------------------------------------->
		tempInicializa
		'<-------------------------------------------------------------------
		
		session("emp_observacao") = ""
		session("pedidoItens") = ""
		'pagina finalizado
		response.Redirect("GeraEmpFim.asp")
	end if

end if

sub Principal
%>

	
<%
	if tempFaltaItens(idProcesso) then
		%>
		<label class="erro">Ainda há itens não compostos (sem + Compra ou + Estoque).</label>
		<br><br>
		<input type="button" name="Voltar" value="Voltar" class="botao" onClick="window.location.href='GeraEmpPrazo.asp?idProcesso=<%=idProcesso%>&ori=Confirm'">
		<%
		exit sub
	end if
	
	
	
	sql = "select parceiro.nome, tempEmpItemPedido.* " & _
			"from tempEmpItemPedido inner join parceiro on tempEmpItemPedido.idFornecedor = parceiro.id " & _
			"WHERE sessionid='" & session.SessionID & "' " & _
			"order by nome, valorUnitario"
	
	set rs = conn.execute(SQL)
	
	%>
	<p>
		<label class="titulo">Empenho</label>
		<br>
		<label class="subtitulo">Para compor o empenho será necessário os seguintes pedidos de compra e as baixas no estoque, confirma?</label>
	</p>
<form action="GeraEmpConfirm.asp" method="post" enctype="application/x-www-form-urlencoded" name="form1" id="form1">

	<p>
	<span class="erro"><%=erro%></span>
	</p>
	<label class="subtitulo">Pedidos</label>
	
	<%
	if rs.eof then
		%><br><label class="texto">:: Nenhum produto para compra foi informado. ::</label><%
	else
		%>
		<table cellpadding="2" cellspacing="0" border="0">	
		<%
		aux = ""
		while not rs.eof 
			fornecedor = rs("nome")
			if aux <> fornecedor then
				%>
				<tr>
					<td colspan="4"><hr noshade color="f5f5f5" size="1"></td>
				</tr>
				<tr>
					<td>
						<span class="cabecalho">Fornecedor: </span>
					</td>
					<td colspan="3">
						<span class="texto"><%=rs("nome")%></span>
					</td>
				</tr>
				<tr>
					<td>
						<span class="cabecalho">Data do Pedido:</span>
					</td>
					<td colspan="3">
						<%datador "dtPedido_" & rs("idFornecedor"), 0, year(now) - 10, year(now) + 1, dtPedido(rs("idFornecedor"))%>
					</td>
				</tr>
					
				<tr>
					<td class="cabecalho">Comprador (Responsável):</td>
					<td colspan="3"><input name="pedidoComprador_<%=rs("idFornecedor")%>" type="text" class="caixa" value="<%=pedidoComprador(rs("idFornecedor"))%>" size="30" maxlength="100"></td>
				</tr>
				<tr>
					<td class="cabecalho">Prazo de Entrega:</td>
					<td colspan="3"><input name="pedidoPrazoEntrega_<%=rs("idFornecedor")%>" type="text" class="caixa" value="<%=pedidoPrazoEntrega(rs("idFornecedor"))%>" size="30" maxlength="100"></td>
				</tr>
				<tr>
					<td class="cabecalho">Forma de Pagamento:</td>
					<td colspan="3"><input name="pedidoFormaPagto_<%=rs("idFornecedor")%>" type="text" class="caixa" value="<%=pedidoFormaPagto(rs("idFornecedor"))%>" size="30" maxlength="100"></td>
				</tr>
				<tr>
					<td>
						<span class="cabecalho">Observações: </span>
					</td>
					<td colspan="3" colspan="3">
						<textarea  class="texto"name="obs_<%=rs("idFornecedor")%>" cols="30" rows="3" id="obs_<%=rs("idFornecedor")%>"><%=obs(rs("idFornecedor"))%></textarea>
					</td>
				</tr>
				<tr>
					<td colspan="4"><span class="cabecalho">Lista de Produtos: </span></td>
				</tr>
				<%
				aux = rs("nome")
			end if
			%>
				<tr>
				  <td>&nbsp;</td>
				  <td align="right"><span class="texto"><%=formata_campo(rs("qtProduto").value)%></span></td>			
				  <td><span class="texto"><%=rs("descricao")%></span></td>
				  <td align="right"><span class="texto"><%=formata_campo(rs("qtProduto") * rs("ValorUnitario"))%></span></td>
		  		</tr>
				
			
			<%
			rs.movenext
		wend
		%>
		</table>
			<%
			
	end if
	rs.close
	set rs = nothing
	%>

	<br><br>
	<label class="subtitulo">Baixa no estoque</label>
	<%
	sql = "SELECT * from tempEmpItemEstoque where sessionID = '" & session.SessionID & "'"
	set rs = conn.execute(SQL)
	
	if rs.eof then
		%><br><label class="texto">:: Nenhum produto do estoque foi informado. ::</label><%
	else
		%>
		<table cellpadding="2" cellspacing="0" border="0">	
		<%
		while not rs.eof
			%>
			<tr>
			  <td align="right"><span class="texto"><%=formata_campo(rs("qt").value)%></span></td>			
			  <td><span class="texto"><%=rs("descricao")%></span></td>
			  <td align="right"><span class="texto"><%=formata_campo(rs("qt") * rs("ValorUnitario"))%></span></td>								
			  <td><span class="texto">(Estoque)</span></td>
			</tr>
			
			<%
			rs.movenext
		wend
		%>
		</table>
		<%
	end if
	
	
	%>
	<p>
		<input type="hidden" name="prazo" value="<%=prazo%>">
		<input type="hidden" name="validade" value="<%=validade%>">	
		<input type="hidden" name="idProcesso" value="<%=idProcesso%>">
		<input type="hidden" name="dataEmpenho" value="<%=dataEmpenho%>">
		<input type="hidden" name="valorTotal" value="<%=valorTotal%>">
				
		<input type="hidden" name="isPostBack" value="1">				
		<input type="button" name="Voltar" value="Voltar" class="botao" onClick="window.location.href='GeraEmpPrazo.asp?idProcesso=<%=idProcesso%>&validade=<%=validade%>&prazo=<%=prazo%>&dataEmpenho=<%=dataEmpenho%>&valorTotal=<%=valorTotal%>&ori=Confirm'">
		<input name="concluir" type="submit" class="botao" id="concluir" value="Concluir">

	</p>
</form>
	<%
end sub

sub Impressao
end sub

sub auxiliar
	
end sub

%>
    
<!--#include file="../layout/layout.asp"-->

