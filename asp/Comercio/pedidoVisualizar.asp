<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->

<%call abreBanco%>
<%

dim NumPedido, erro, acao

NumPedido = request("NumPedido")
if NumPedido = "" or not isNumeric(NumPedido) then response.Redirect("../")
sub principal
	dim tot
	'###################################################
	' exibindo os dados do pedido
	
	SQL = "SELECT   Edital.NumProcesso, Parceiro.Nome AS Fornecedor, Pedido.* " & _
			"FROM   Edital RIGHT JOIN " & _
					"(Parceiro INNER JOIN Pedido ON Parceiro.Id = Pedido.idFornecedor) " & _
					"ON Edital.IDProcesso = Pedido.idProcesso " & _			
			"where NumPedido=" & NumPedido
	set rs = conn.execute(SQL)
	
	
'cabecalho do pedido%>
<p align="center">
	<div align="center" class="titulo" style="width:100%;">
		<%=application("par_20")%>
	</div>
	<br>
	<div align="center" class="cabecalho"  style="width:100%;">
	<%=Application("par_13")%> - Fone/Fax: <%=Application("par_14")%> - CEP <%=Application("par_15")%>.<BR>
		C.N.P.J. <%=application("par_16")%> - Inscrição Estadual <%=application("par_17")%> - <%=application("par_18")%> - <%=application("par_19")%>
	</div>

</p>
<div align="center" style="background-color:#CCCCCC;border-width:1px;border-color:#000000;border-style:solid; ">
	<div class="cabecalho" style="width:100%">PEDIDO DE COMPRA</div>
</div>
	<%
	if not rs.eof then
'		if rs("cancelado") = 0 then
'			if rs("recebido") = 0 then
'				situacao = 0
'			else
'				situacao = 1
'			end if
'		else
'			situacao = 2
'		end if
		%>
			<table width="100%"  border="0" cellspacing="3" cellpadding="0" style="border-collapse:collapse ">

			  <tr>
			  	<td valign="top"><span class="cabecalho">Para: </span>
				  <span class="texto"><%=rs("Fornecedor")%></span></td>	
				<td valign="top" align="right"><span class="cabecalho">Data: </span>
					<span class="texto"><%=formata_campo(rs("dtPedido").value)%></span></td>  
			  </tr>
			  <tr>
			  	<td colspan="2"><span class="cabecalho">Prazo de Entrega: </span>
						<span class="texto"><%=rs("prazoEntrega")%></span></td>
			  </tr>
			  <tr>
			  	<td colspan="2"><span class="cabecalho">Condições de Pagamento: </span>
					<span class="texto"><%=rs("formaPagto")%></td>
			  </tr>
			  <tr>
			  	<td colspan="2" class="cabecalho">
					<%=rs("obs")%>
				</td>
			  </tr>
			  <tr>
			  	<td colspan="2">
					<%tot = ItensDoPedido(NumPedido)%>
				</td>
			  </tr>
			  <tr>
			  	<td colspan="2">
					<table width="100%" cellpadding="3" cellspacing="0" border="1" style="border-collapse:collapse" bgcolor="#CCCCCC" bordercolor="#000000">
						<tr>
							<td><span class="cabecalho">COMPRADOR: </span>
								<span class="texto"><%=rs("comprador")%></span></td>
							<td align="right"><span class="cabecalho">VALOR TOTAL DO PEDIDO: </span>
								<span class="texto"><%=formata_campo(cCur(tot))%></span></td>
						</tr>
						<tr>
							<td colspan="2"><span class="cabecalho">N.º do Pedido: </span>
								<span class="texto"><%=rs("numPedido")%></span>
								<br>
								<span class="cabecalho">N.º do Processo: </span>
								<span class="texto"><%=iif(isNull(rs("numProcesso")), "-", rs("numProcesso"))%></span></td>								
						</tr>
					</table>
				
				</td>
			  
			  </tr>
					<!--
					
				<td valign="top"><span class="cabecalho">NumPedido</span><br>
				 <span class="texto"><%=rs("NumPedido")%></span></td>
				<td valign="top"><span class="cabecalho">Num. Processo</span><br>
				  <span class="texto"><%=rs("NumProcesso")%></span></td>			
				
			
		  
				<td valign="top"><span class="cabecalho">NF de Compra</span><br>
				  <span class="texto"><%=rs("NumNFCompra")%></span></td>			  
				<td valign="top"><span class="cabecalho">Vl. Impostos</span><br>
				<%'só exibe o valor dos impostos depois de recebido %>
				  <span class="texto"><%=iif(situacao = 0, "", formata_campo(rs("vlImpostos").value))%></span></td>
			  </tr>					
							
			  <tr>
				<td colspan="3" valign="top"><span class="cabecalho">OBS</span><br>
				  </td>
			  </tr>				
			  -->	          
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

end sub


function ItensDoPedido(NumPedido)
	dim sql
	dim rs
	dim t
	t = 0
	SQL = "Select * from PedidoItem where NumPedido=" & NumPedido
	set rs = conn.execute(SQL)
	%>
	<table width="100%"  border="1" cellspacing="0" cellpadding="3" style="border-collapse:collapse " bordercolor="#000000">
	<tr bgcolor="#CCCCCC">
	  <td align="right"><label class="cabecalho">Qtde.</label></td>
	  <td><label class="cabecalho">Descrição</label></td>
	  <td align="right"><label class="cabecalho">Vl. Unitário</label></td>
	  <td align="right"><label class="cabecalho">Vl. Total</label></td>								
	</tr>
	<%
	dim cor
	cor =false
	while not rs.eof
	cor = not cor
	t = t + (rs("quantidade") * rs("vlUnitario"))
	%>


							<tr bgcolor="<%if not cor then response.Write("#eeeeee")%>">
							  <td align="right"><span class="texto"><%=formata_campo(rs("quantidade").value)%></span></td>
							  <td><span class="texto"><%=rs("descricao")%></span></td>
							  <td align="right"><span class="texto"><%=formata_campo(rs("vlUnitario").value)%></span></td>
							  <td align="right"><span class="texto"><%=formata_campo(rs("quantidade") * rs("vlUnitario"))%></span></td>								
							</tr>
							
	
	<%	rs.movenext
	wend
	rs.close
	set rs = nothing
	%>	
	</table>
	<%
	ItensDoPedido = t
end function
sub Impressao
	response.write request.ServerVariables("SCRIPT_NAME") & "?imp=1&NumPedido=" & NumPedido & "&data=" & now
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

