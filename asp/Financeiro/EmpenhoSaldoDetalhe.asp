<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<%call abreBanco%>

<%
dim idProcesso
dim creditos
dim icms
dim produtos

idProcesso = request.QueryString("idProcesso")
if idProcesso = "" or not isNumeric(idProcesso) then response.Redirect("../")

sub principal
	dim sql
	dim rs
	
	creditos = 0
	icms = 0
	produtos = 0
	
	sql = "SELECT     Edital.NumProcesso, Empenho.DataEmpenho, Empenho.Observacao " & _
			"FROM         Edital INNER JOIN " & _
                      "Empenho ON Edital.IdProcesso = Empenho.idProcesso " & _
					  "WHERE empenho.idProcesso=" & idProcesso
	
	set rs = conn.execute(SQL)
	
	if rs.eof then
		%><label class="erro">Nenhum registro encontrado</label><%
		rs.close 
		set rs = nothing
		exit sub
	else
		%>
		<p>
		<label class="titulo">Detalhes do saldo de Empenho</label><br>
		<label class="subtitulo">Processo: <%=rs("numProcesso")%></label><br>
		<label class="subtitulo">Data: <%=formata_campo(rs("dataEmpenho").value)%></label><br>
		<label class="texto">Observações: <%=rs("observacao")%></label>
		</p>
		<hr noshade size="1" width="100%">
		<table cellpadding="3" cellspacing="0" border="1" style="border-collapse:collapse" width="100%">
			<tr>
				<td class="cabecalho">Descrição</td>
				<td align="right" class="cabecalho">Créditos</td>
				<td align="right" class="cabecalho">%</td>
				<td align="right" class="cabecalho">Icms</td>
				<td class="cabecalho">Produtos</td>												
			</tr>
			<%exibeNota%>
			<tr bgcolor="#f5f5f5">
				<td colspan="5" class="texto">&nbsp;</td>
			</tr>
			<%
			impostoFederal%>
			<tr bgcolor="#f5f5f5">
				<td colspan="5" class="texto">&nbsp;</td>
			</tr>
			<%
			exibePedidos%>
			<tr bgcolor="#f5f5f5">
				<td colspan="5" class="texto">&nbsp;</td>
			</tr>
			<%
			exibeSaldo%>
		</table>
		<%
	
	end if
	rs.close
	set rs = nothing
	
	
end sub

function exibeNota
	dim sql
	dim rs
	dim i
	dim tot
	
	sql = "select NumNF, descricao, valorTotal, impostos from empenhoNF where idProcesso=" & idProcesso
	set rs = conn.execute(SQL)
	
	while not rs.eof
		%>
		<tr class="texto">
			<td><%="Empenho: " & rs("descricao") & " - NF: " & rs("numNF") %></td>
			<td align="right" style="color:#0000FF "><%=formata_campo(rs("valorTotal").value)%></td>
			<td align="right"><%=formata_campo((rs("impostos") * 100)/(rs("valorTotal")))%>%</td>			
			<td align="right" style="color:#FF0000 ">(<%=formata_campo(rs("impostos").value)%>)</td>
			<td>&nbsp;</td>
		</tr>
		<%
		creditos = creditos + rs("valorTotal")
		icms = icms - rs("impostos")

		rs.movenext
	wend
	
	rs.close
	set rs = nothing 
end function

function impostoFederal
	dim sql
	dim rs
	dim tot
	sql = "SELECT SUM(valorTotal) as tot from empenhoNF where idProcesso=" & idProcesso
	set rs = conn.execute(SQL)
	if isNull(rs("tot")) then
		tot = 0
	else
		tot = ccur((rs("tot") * replace(application("par_25"),".",",")) / 100)
	end if
	%>
	<tr class="texto">
		<td>Imposto Federal</td>
		<td>&nbsp;</td>
		<td align="right">5,90%</td>
		<td align="right" style="color:#FF0000 ">(<%=formata_campo(tot)%>)</td>
		<td>&nbsp;</td>
	</tr>
	<%
	icms = icms - tot
	rs.close
	set rs = nothing
end function

function exibePedidos
	dim sql
	dim rs
	
	SQL = "select pedido.NumNFCompra, pedido.vlTotal, pedido.vlImpostos, parceiro.nome from " & _
			"pedido inner join parceiro on parceiro.id=pedido.idFornecedor " & _
			"where recebido=1 and cancelado=0 and idProcesso=" & idProcesso	
			
	set rs = conn.execute(SQL)
	while not rs.eof
		%>
		<tr class="texto">
			<td><%=replace(rs("nome"),chr(13),"<BR>") & " - NF:" & rs("NumNFCompra")%></td>
			<td>&nbsp;</td>
			<td align="right"><%=formata_campo((rs("vlImpostos") * 100)/rs("vlTotal"))%>%</td>
			<td align="right" style="color:#0000FF "><%=formata_campo(rs("vlImpostos").value)%></td>
			<td align="right" style="color:#FF0000 "><%=formata_campo(rs("vlTotal").value)%></td>
		</tr>
		<%
		produtos = produtos + rs("vlTotal")
		icms = icms + rs("vlImpostos")
	
		rs.movenext
	wend
	rs.close
	set rs = nothing
	
end function

function exibeSaldo
	dim aux
	dim aux2
	dim aux3
	'numero absoluto do total do icms (para exibir no padrao (###) ao inves de -###)
	aux = formata_campo(abs(icms))
	
	aux2 = (creditos - produtos + icms) * 100/creditos
	
	aux3 = creditos - produtos + icms
	%>
		<tr class="texto">
			<td>Saldo da Operacao</td>
			<td>&nbsp;</td>
			<td align="right">&nbsp;</td>
			<td align="right" style="color:<%=iif(icms >= produtos, "#0000FF", "#FF0000")%>">
				<%=iif(icms < 0 , "(" & aux & ")", aux)%></td>
			<td align="right" style="color:#FF0000 ">(<%=formata_campo(produtos)%>)</td>
		</tr>	
		<tr class="cabecalho" bgcolor="<%=iif(aux2 >= 0, "#A6D9FF", "#FFA8A8")%>">
			<td>Saldo lucro líquido</td>
			<td>&nbsp;</td>
			<td align="right">&nbsp;</td>
			<td align="right"><%=iif(aux2 < 0 , "(" & formata_campo(abs(aux2)) & ")", formata_campo(aux2))%></td>
			<td align="right"><%=formata_campo(iif(aux3 < 0, "(" & formata_campo(abs(aux3)) & ")", formata_campo(aux3)))%></td>
		</tr>			
	<%
end function
%>

<%sub Impressao
	response.write request.ServerVariables("SCRIPT_NAME") & "?imp=1&idProcesso=" & idProcesso
end sub

sub auxiliar
%>
		<input type="button" onClick="window.print();" value="Imprimir" class="botao">
		<br><br>
		<input type="button" onClick="window.location.href='EmpenhoSaldo.asp'" value="Voltar" class="botao">

	<%
end sub

if request.QueryString("imp") = "" then%>
<!--#include file="../layout/layout.asp"-->
<%else%>
<!--#include file="../layout/layout_print.asp"-->
<%end if%>

<%call fechaBanco%>
