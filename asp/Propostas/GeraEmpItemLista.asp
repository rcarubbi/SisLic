<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="funGeraEmpenho.asp"-->
<%call abreBanco%>
<%
dim idProcesso, rtn

idProcesso = request("idProcesso")

sub Principal
	'se a origem veio da primeira tela
	if request.QueryString("ori") = "" then
		if request.Form("isPostBack") = "" then
			'limpando registros antigos deste usuario
			tempInicializa
		end if
	
		rtn = tempGravaItensGanhos(request.Form("chkItens"), idProcesso)
	
		'se o usuario não selecionou nenhum empenho
		if rtn = "0" then
			%>
			<p>
			<label class="titulo">Empenho</label><br>
			<label class="subtitulo">Erro</label>
			</p>
			<p>
			<label class="erro">Nenhum item do empenho foi selecionado! É necessário pelo menos 1 (um) item para continuar!</label>
			</p>
			<input name="Voltar" type="button" class="botao" id="voltar" value="Voltar" onClick="history.back(-1);">
	
			<%
			exit sub
		end if
	end if
		
	dim rs
	
	set rs = tempGetItensGanhos
	
	%>
	<p>
		<label class="titulo">Composição do empenho</label><br>
		<label class="subtitulo">Informe os produtos necessários para compor o empenho.</label>
	</p>
	<table width="100%"  border="0" cellspacing="0" cellpadding="3">
	<!--
	<tr>
	  <td><label class="cabecalho">Especificação</label></td>
	  <td><label class="cabecalho">Unidade</label></td>	  
	  <td align="right"><label class="cabecalho">Vl. Unitário</label></td>
	  <td align="right"><label class="cabecalho">Qt.</label></td>
	  <td align="right"><label class="cabecalho">Vl. Total</label></td>			
	  <td></td>					
	  <td></td>
	</tr>
		-->
	<%
	while not rs.eof
	
		%>
			<tr bgcolor="<%=corFundo%>">
			  <td><span class="cabecalho"><%=rs("especificacao")%></span></td>
			  <td><span class="cabecalho"><%=rs("unidade")%></span></td>							  
			  <td align="right"><span class="cabecalho"><%=formata_campo(rs("ValorUnitario").value)%></span></td>
			  <td align="right"><span class="cabecalho"><%=formata_campo(rs("quantidade").value)%></span></td>
			  <td align="right"><span class="cabecalho"><%=formata_campo(rs("quantidade") * rs("ValorUnitario"))%></span></td>								
			  <td><a href="GeraEmpItemPedido.asp?idPropostaItem=<%=rs("idPropostaItem")%>&idProcesso=<%=idProcesso%>&acao=add" class="cabecalho">+Compra</a></td>
			  <td><a href="GeraEmpItemEstoque.asp?idPropostaItem=<%=rs("idPropostaItem")%>&idProcesso=<%=idProcesso%>&acao=add" class="cabecalho">+Estoque</a></td>
			</tr>
			<tr>
				<td colspan="7">
					<%
					listaSubItens rs("idPropostaItem")%>
				</td>
				
			</tr>
			<tr>
				<td colspan="7">
					<hr size="1" color="#CCCCCC">
				</td>
			</tr>	
		
		<%
	
		rs.movenext
	wend
	%>
	</table>
	<%
	rs.close
	set rs = nothing
	
	sql = "select sum(qt*valorUnitario) as tot " & _
			"from tempEmpItemEstoque where sessionID='" & session.SessionID & "' " 
	set rs = conn.execute(SQL)
	if not isNull(rs("tot")) then
		totalGeral = rs("tot")
	end if
	rs.close
	set rs = nothing
	
	sql = "select sum(qtProduto * ValorUnitario) as tot " & _
				"from tempEmpItemPedido where sessionID='" & session.SessionID & "' " 
	
	set rs = conn.execute(SQL)
	if not isNull(rs("tot")) then
		totalGeral = totalGeral + rs("tot")
	end if
	rs.close
	set rs = nothing
	
		
	
	%>
	<input type="button" name="Voltar" value="Voltar" class="botao" onClick="window.location.href='GeraEmpIni.asp?idProcesso=<%=idProcesso%>'">
	<input type="button" name="Proximo" value="Próximo" class="botao" onClick="window.location.href='GeraEmpPrazo.asp?idProcesso=<%=idProcesso%>&valorTotal=<%=totalGeral%>'">
	
	<%
end sub

sub Impressao
end sub

sub auxiliar
	
end sub


sub ListaSubItens(idPropostaItem)
	dim rs, totQt, totVl
	
	'-----------------------------------------------------------------
	'LISTANDO OS ITENS CUJA ORIGEM É DE NOVOS PEDIDOS
	'-----------------------------------------------------------------
	set rs = tempGetSubItensPedidos(idPropostaItem, "")
	
	%>
	<table border="0" cellspacing="0" cellpadding="5">
	<%
	totQt = 0
	totVl = 0

	
	while not rs.eof
	
		%>
			<tr bgcolor="<%=corFundo%>">
			  <td align="right"><span class="texto"><%=formata_campo(rs("qtProduto").value)%></span></td>			
			  <td><span class="texto"><%=rs("descricao")%></span></td>
<!--			  <td align="right"><span class="texto"><%=formata_campo(rs("ValorUnitario").value)%></span></td>-->
<!--			  <td align="right"><span class="texto"><%=formata_campo(rs("qtPedido").value)%></span></td>-->
			  <td align="right"><span class="texto"><%=formata_campo(rs("qtProduto") * rs("ValorUnitario"))%></span></td>								
			  <td><span class="texto">(Pedido)</span></td>
			  <td><a href="GeraEmpItemPedido.asp?id=<%=rs("ID")%>&acao=edit&idProcesso=<%=idProcesso%>" class="texto">Editar</a></td>
			  <td><a href="GeraEmpItemPedido.asp?id=<%=rs("ID")%>&acao=del&idProcesso=<%=idProcesso%>" class="texto">Excluir</a></td>
			</tr>
		<%
		totQt = totQt + rs("qtProduto")
		totVl = totVl + (rs("qtProduto") * rs("ValorUnitario"))
		rs.movenext
	wend
	
	rs.close
	set rs = nothing		
	
	
	'------------------------------------------------------------------------
	'LISTANDO ITENS CUJA ORIGEM É DO ESTOQUE
	'------------------------------------------------------------------------
	set rs = tempGetSubItensEstoque(idPropostaItem, "")

	while not rs.eof
	
		%>
			<tr bgcolor="<%=corFundo%>">
			  <td align="right"><span class="texto"><%=formata_campo(rs("qt").value)%></span></td>			
			  <td><span class="texto"><%=rs("descricao")%></span></td>
			  <td align="right"><span class="texto"><%=formata_campo(rs("qt") * rs("ValorUnitario"))%></span></td>								
			  <td><span class="texto">(Estoque)</span></td>
			  <td><a href="GeraEmpItemEstoque.asp?id=<%=rs("ID")%>&acao=edit&idProcesso=<%=idProcesso%>" class="texto">Editar</a></td>
			  <td><a href="GeraEmpItemEstoque.asp?id=<%=rs("ID")%>&acao=del&idProcesso=<%=idProcesso%>" class="texto">Excluir</a></td>
			</tr>
		<%
		totQt = totQt + rs("qt")
		totVl = totVl + (rs("qt") * rs("ValorUnitario"))
		rs.movenext
	wend
	
	rs.close
	set rs = nothing		
		
	if totQt > 0 then
		%>
	
		<tr>
			<td style="border-top:1px solid #000000" align="right"><span class="cabecalho"><%=totQt%></span></td>
			<td style="border-top:1px solid #000000">&nbsp;</td>
			<td style="border-top:1px solid #000000" align="right"><span class="cabecalho"><%=formata_campo(cCur(totVl))%></span></td>
			<td style="border-top:1px solid #000000">&nbsp;</td>
			<td style="border-top:1px solid #000000">&nbsp;</td>
			<td style="border-top:1px solid #000000">&nbsp;</td>
		</tr>
	<%
	end if
	%>
	</table>
	<%

end sub


%>
    
<!--#include file="../layout/layout.asp"-->

