<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->

<%call abreBanco%>
<%

dim NumPedido, erro, acao

NumPedido = request("NumPedido")
if NumPedido = "" or not isNumeric(NumPedido) then response.Redirect("../")
sub principal
	
	'###################################################
	' exibindo os dados do pedido
	
	SQL = "SELECT   Edital.NumProcesso, Parceiro.Nome AS Fornecedor, Pedido.* " & _
			"FROM   Edital RIGHT JOIN " & _
					"(Parceiro INNER JOIN Pedido ON Parceiro.Id = Pedido.idFornecedor) " & _
					"ON Edital.IDProcesso = Pedido.idProcesso " & _			
			"where NumPedido=" & NumPedido
	set rs = conn.execute(SQL)

	if not rs.eof then
	
	%>
	<label class="titulo">
	
	</label>
	
	<%
		if rs("cancelado") = 0 then
			if rs("recebido") = 0 then
				situacao = 0
			else
				situacao = 1
			end if
		else
			situacao = 2
		end if
		%>
		<label class="titulo">Detalhes do Pedido  n.º:<%=rs("NumPedido")%>
			<%
			select case situacao
			case 0
			'quando o pedido esta pendente, imprime em um formato para o fornecedor
				%>
				(Pendente)
				<%
			case 1
				%>
				(Recebido)
				<%		
			case 2
				%>
				(Cancelado)
				<%		
			end select
			%>
		</label><br>
		<label class="subtitulo">Segue abaixo os dados do Pedido</label>
		<br><br>
		<span class="subtitulo">Pedido</span>
	
			<table width="100%"  border="1" cellspacing="3" cellpadding="0" style="border-collapse:collapse ">

			  <tr>
				<td valign="top"><span class="cabecalho">NumPedido</span><br>
				 <span class="texto"><%=rs("NumPedido")%></span></td>
				<td valign="top"><span class="cabecalho">Num. Processo</span><br>
				  <span class="texto"><%=rs("NumProcesso")%></span></td>			
				<td valign="top"><span class="cabecalho">Data do Pedido</span><br>
				  <span class="texto"><%=formata_campo(rs("dtPedido").value)%></span></td>
			
			<tr>
				<td valign="top"><span class="cabecalho">Data de Recebimento</span><br>
				  <span class="texto"><%=formata_campo(rs("dtRecebimento").value)%></span></td>
			
				<td valign="top"><span class="cabecalho">Recebido</span><br>
				  <span class="texto"><%=formata_campo(rs("recebido").value)%></span></td>
			
				<td valign="top"><span class="cabecalho">Pago</span><br>
				  <span class="texto"><%=formata_campo(rs("pago").value)%></span></td>
			  </tr>					
			
			  <tr>
				<td valign="top"><span class="cabecalho">Fornecedor</span><br>
				  <span class="texto"><%=rs("Fornecedor")%></span></td>			  
				<td valign="top"><span class="cabecalho">NF de Compra</span><br>
				  <span class="texto"><%=rs("NumNFCompra")%></span></td>			  
				<td valign="top"><span class="cabecalho">Vl. Impostos</span><br>
				<%'só exibe o valor dos impostos depois de recebido %>
				  <span class="texto"><%=iif(situacao = 0, "", formata_campo(rs("vlImpostos").value))%></span></td>
			  </tr>					
							
			  <tr>
				<td colspan="3" valign="top"><span class="cabecalho">OBS</span><br>
				  <span class="texto"><%=rs("obs")%></span></td>
			  </tr>					          
			</table>
	<%
	else
		%><span class="erro">Nenhum registro encontrado</span>
			<p><a href="#" class="texto" onClick="histoy.back(-1);">Voltar</a></p>
		<%
	end if
	
	rs.close
	set rs = nothing
	' fim dados pedido
	'##################################################
	%>		
	<br>
	
	<%
	'##################################################
	' endereco do cliente
	
	
	SQL = "Select * from PedidoItem where NumPedido=" & NumPedido
	set rs = conn.execute(SQL)
	%>
	<span class="subtitulo">Itens</span>
	<table width="100%"  border="1" cellspacing="0" cellpadding="3" style="border-collapse:collapse ">
	<tr>
	  <td><label class="cabecalho">Descrição</label></td>
	  <td align="right"><label class="cabecalho">Vl. Unitário</label></td>
	  <td align="right"><label class="cabecalho">Qt.</label></td>
	  <td align="right"><label class="cabecalho">Vl. Total</label></td>								
	</tr>
	<%
	dim cor
	cor =false
	while not rs.eof
	cor = not cor
	%>


							<tr bgcolor="<%if not cor then response.Write("#eeeeee")%>">
							  <td><span class="texto"><%=rs("descricao")%></span></td>
							  <td align="right"><span class="texto"><%=formata_campo(rs("vlUnitario").value)%></span></td>
							  <td align="right"><span class="texto"><%=formata_campo(rs("quantidade").value)%></span></td>
							  <td align="right"><span class="texto"><%=formata_campo(rs("quantidade") * rs("vlUnitario"))%></span></td>								
							</tr>
							
	
	<%	rs.movenext
	wend
	rs.close
	set rs = nothing
	%>	
	</table>
	
	<%
end sub

sub Impressao
	response.write request.ServerVariables("SCRIPT_NAME") & "?imp=1&NumPedido=" & NumPedido
end sub

sub auxiliar
	%>
		<input type="button" onClick="window.print();" value="Imprimir" class="botao">
		<br><br>
		<input type="button" onClick="window.location.href='PedidoLista.asp'" value="Voltar" class="botao">

	<%
end sub



%>
<%if request.QueryString("imp") = "" then%>
	<!--#include file="../layout/layout.asp"-->
<%else%>
	<!--#include file="../layout/layout_print.asp"-->
<%end if%>

