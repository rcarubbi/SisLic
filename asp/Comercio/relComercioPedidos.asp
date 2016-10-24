<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<%call abreBanco%>

<%
dim sql, rs, vcor, cor
dim pago, cancelado, recebido, fornecedor, dtini, dtfim
pago = trim(replace(request("pago"),"'","''"))
cancelado = trim(replace(request("cancelado"),"'","''"))
recebido = trim(replace(request("recebido"),"'","''"))
fornecedor = trim(replace(request("fornecedor"),"'","''"))
dtIni = recebe_datador("dtIni",0)
dtFim = recebe_datador("dtFim",0)

if not isDate(dtIni) then
	dtini = dateadd("m",-1,date)
	dtfim = date
end if
if cdate(dtFim) < cdate(dtIni) then dtfim = dtini
sub principal
	
	if request("isPostBack") then
		
		
		sql = "SELECT p.dtpedido, p.dtrecebimento, p.vltotal, p.vlImpostos, p.pago, f.nome as fornecedor " & _
		      " FROM pedido p INNER JOIN parceiro f on f.id = p.idfornecedor " & _
			  " WHERE p.dtpedido between " & formata_data_sql(dtIni) & " and " & formata_data_sql(dtFim & " 23:59:59") & " and recebido = " & recebido & " and cancelado = " & cancelado
				
			  if pago > 0 then
			  	sql = sql & " and p.pago= " & pago - 1
			  end if
				
			  if fornecedor > 0 then
			  	sql = sql & " and f.id= " & fornecedor
			  end if
				
				
		set rs = conn.execute(sql)
		%>
			<label class="titulo">Relatório de Pedidos</label><br>
			
			<label class="subtitulo"> <%=iif(recebido=0,"Não Recebidos","Recebidos")%><%if pago>0 then response.write iif(pago=1," a pagar"," pagos")%><%if cancelado=1 then response.write " (Cancelados)"%></label><br>
			<label class="subtitulo">Fornecedor: <%=iif(fornecedor=0,"Todos",getValor(fornecedor,"id","nome","parceiro"))%></label><br>
			<label class="subtitulo">No período de: <%=dtini%> à <%=dtfim%></label><BR>
			<table style="border-collapse:collapse;" border="1" width="100%">
				<thead >
					<tr class="cabecalho" style="background-color:#f5f5f5;">
						<th width="25%">Fornecedor</th>
						<th width="20%"><%=iif(recebido=0,"Data Pedido","Data Recebimento")%></th>
						<th width="15%" align="right">Valor Total</th>
						<th width="10%" align="right">Impostos</th>
						<th width="3%">Pago</th>
					</tr>
				</thead>
				<tbody>
		<%while not rs.eof %>		
		<%if vcor then
			cor = "#ffffff"
		else
			cor= "#ffffcc"
		end if%>
				<tr bgcolor="<%=cor%>">
					<td class="texto"><%=rs("fornecedor")%></td>
					<td class="texto" align="center"><%=formata_data(iif(recebido=0,rs("dtpedido"),rs("dtrecebimento")))%></td>
					<td class="texto" align="right"><%=formatcurrency(rs("vltotal"))%></td>
					<td class="texto" align="right"><%=formatcurrency(rs("vlimpostos"))%></td>
					<td class="texto"><%=formata_corSN(formata_campo(cbool(rs("pago"))))%></td>
				</tr>
			
			
		<%	vcor = not vcor
			rs.movenext
		wend	
		%></tbody>
		</table><%
	end if
end sub
%>

<%sub Impressao
	response.write request.ServerVariables("SCRIPT_NAME") & "?pago=" & pago & "&recebido=" & recebido & "&cancelados=" & cancelados & "&imp=1&isPostBack=1"
end sub

sub auxiliar%>
<form name="filtros" method="post" action="">
<input type="hidden" value="1" name="isPostBack">
<table>
	<tr>
		<td><label class="cabecalho">Recebidos</label><br>
			<%dim itemTexto, itemValor
			redim itemTexto(1), itemValor(1)
			itemTexto(0) = "Não"
			itemValor(0) = 0
			itemTexto(1) = "Sim"
			itemValor(1) = 1
			Combo "recebido", itemTexto, itemValor, request.Form("recebido")%>
		</td>
	</tr>
	<tr>
		<td><label class="cabecalho">Canceladas</label><br>
			<%
			redim itemText(1), itemValor(1)
			itemTexto(0) = "Não"
			itemValor(0) = 0
			itemTexto(1) = "Sim"
			itemValor(1) = 1
			Combo "cancelado", itemTexto, itemValor, request.Form("cancelado")%>
		</td>
		
	</tr>
	
	<tr>
		<td><label class="cabecalho">Pagas</label><br>
			<%redim itemTexto(2), itemValor(2)
			itemTexto(0) = "Indiferente"
			itemValor(0) = 0
			itemTexto(1) = "Não"
			itemValor(1) = 1
			itemTexto(2) = "Sim"
			itemValor(2) = 2
			Combo "Pago", itemTexto, itemValor, request.Form("Pago")%>
		</td>
		
	</tr>
	<tr>
		<td><label class="cabecalho">Fornecedor</label><br>
			<%ComboDB "Fornecedor","Nome","id","Select id, nome from parceiro where tipo = 1 and ativo = 1",Fornecedor,"Todos","" %>
		</td>
	</tr>
	<tr><Td>
	<label class="cabecalho">Data inicial</label>
		<br>
		<%
		Datador "dtIni",0,1980, year(now),dtIni
		%></td></tr>
		<tr><td>
		<label class="cabecalho">Data final</label>
		<br>
		<%
		Datador "dtFim",0,1980, year(now),dtFim	
		%>
		</Td></tr>
	<tr>
		<td><br>
		<input type="submit" value="Ok" name="cmdOk" class="botao">&nbsp;
		<input type="button" value="Imprimir" name="cmdOk"  onClick="window.print();" class="botao">
		</td>
	</tr>
</table>
</form>
<%end sub%>
<%if request.QueryString("imp") = "" then%>
<!--#include file="../layout/layout.asp"-->
<%else%>
<!--#include file="../layout/layout_print.asp"-->
<%end if%>

<%call fechaBanco%>