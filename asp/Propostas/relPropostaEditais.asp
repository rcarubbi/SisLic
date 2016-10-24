<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<%call abreBanco%>

<%dim IDCliente, Situacao, sql, rs, vcor, cor, IDEditalTipo, situacaoDesc, participou, participouDesc, dtini, dtfim
IDEditalTipo = request("IDEditalTipo")
Situacao = request("situacao")
IDCliente = request("IDCliente")
participou = request("participou")
situacaoDesc = formataSituacao(situacao)
participouDesc = formataParticipou(participou)
dtIni = recebe_datador("dtIni",0)
dtFim = recebe_datador("dtFim",0)

if not isDate(dtIni) then
	dtini = dateadd("m",-1,date)
	dtfim = date
end if
if cdate(dtFim) < cdate(dtIni) then dtfim = dtini
sub principal
	
	if request("isPostBack") then
			
		sql = "SELECT e.numprocesso, e.abertura, e.encerramento, e.participou, e.situacao, t.descricao as tipo, replace(p.nome,char(13),'<br>') as nome   FROM (edital e INNER JOIN parceiro p ON p.id=e.idParceiro) INNER JOIN EditalTipo t ON e.IdEditalTipo=t.id Where idprocesso <> 0"
			if idEditalTipo <> 0 then
					sql = sql & " and ideditaltipo = " & ideditaltipo
			end if
			if Situacao <> 4 then
				sql = sql & " and Situacao = " & Situacao
			end if
			if idcliente <> 0  then
				sql = sql & " and idparceiro = " & idcliente
			end if
			if participou <> 0 then
				sql = sql & " and participou = " & participou - 1
			end if
			sql = sql & " and e.dataCadastro between " & formata_data_sql(dtIni) & " and " & formata_data_sql(dtFim & " 23:59:59") 
		set rs = conn.execute(sql)
		%>
			<label class="titulo">Relatório de Editais</label><br>
			<label class="subtitulo">Situação: <%=situacaoDesc%></label><BR>
			<label class="subtitulo">Participou: <%=participouDesc%></label><br>
			<label class="subtitulo">Tipo: <%=iif(idEditalTipo=0,"Todos",getValor(idEditalTipo,"id","descricao","editalTipo"))%></label><br>
			<label class="subtitulo">Cliente: <%=iif(idCliente=0,"Todos",getValor(idCliente,"id","nome","parceiro"))%></label><BR>
			<label class="subtitulo">No período de: <%=dtini%> à <%=dtfim%></label>
			<table style="border-collapse:collapse;" border="1" width="100%">
			<table style="border-collapse:collapse;" border="1" width="100%">
				<thead>
					<tr class="cabecalho" style="background-color:#f5f5f5;">
						<th width="">Tipo</th>
						<th width="">Núm. do Processo</th>
						<th width="">Cliente</th>
						<th width="">Data de Abertura</th>
						<%if situacao = 3 then%>
							<th width="">Data de Encerramento</th>
						<%end if%>
						<th width="">Situação</th>
						<th width="">Participou</th>
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
					<td class="texto"><%=rs("tipo")%></td>
					<td class="texto"><%=rs("numProcesso")%></td>
					<td class="texto"><%=rs("nome")%></td>
					<td class="texto"><%=formata_data(rs("abertura"))%></td>
					<%if situacao = 3 then%>
					<td class="texto"><%=iif(isnull(rs("encerramento")),"Não Encerrado",formata_data(rs("encerramento")))%></td>
					<%end if%>
					<td class="texto"><%=formataSituacao(rs("situacao"))%></td>
					<td class="texto" align="center"><%=formata_campo(cbool(rs("participou")))%></td>
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
	response.write request.ServerVariables("SCRIPT_NAME") & "?imp=1&ispostback=1&situacao=" & situacao & "&ideditaltipo=" & idEditaltipo & "&idcliente=" & idcliente

end sub

sub auxiliar%>
<form name="filtros" method="post" action="">
<input type="hidden" value="1" name="isPostBack">
<table>
	<tr>
		<td><label class="cabecalho">Situação</label><br>
			<%dim itemTexto, itemValor
			redim itemTexto(4)
			redim itemValor(4)
			itemTexto(0) = "Todas"
			itemValor(0) = 4
			itemTexto(1) = "Pendente"
			itemValor(1) = 0
			itemTexto(2) = "Em andamento"
			itemValor(2) = 1
			itemTexto(3) = "Cancelado"
			itemValor(3) = 2
			itemTexto(4) = "Finalizado"
			itemValor(4) = 3
			Combo "Situacao", itemTexto, itemValor, request.Form("situacao")%>
		</td>
	</tr>
	<tr>
		<td>
		<label class="cabecalho">Participou</label><br>
		<%
			redim itemTexto(2)
			redim itemValor(2)
			itemTexto(0) = "Indiferente"
			itemValor(0) = 0
			itemTexto(1) = "Não"
			itemValor(1) = 1
			itemTexto(2) = "Sim"
			itemValor(2) = 2
			Combo "Participou", itemTexto, itemValor, request.Form("Participou")%>
		</td>
	</tr>
	<tr>
		<td><label class="cabecalho">Tipo</label><br>
			  <%ComboDB "IDEditalTipo","Descricao","id","Select id, descricao from EditalTipo",IdEditalTipo,"Todos",""%>
			
		</td>
		
	</tr>
	<tr>
		<td><label class="cabecalho">Cliente</label><br>
			    <%ComboDB "IDCliente","Nome","id","Select id, replace(nome,char(13),'. ') as nome from parceiro where tipo = 0 and ativo = 1",Idcliente,"Todos",235 %>
		</td>
		
	</tr>
	<tr><Td>
	<label class="cabecalho">Data inicial</label>
		<br>
		<%
		Datador "dtIni",0,1980, year(now),dtIni
		%>
		<br>
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
<% function formataSituacao(CodSituacao)
dim retorno
	if codsituacao <> "" then
		select case codsituacao
			case 0
				retorno = "Pendente"
			case 1
				retorno = "Em andamento"	
			case 2
				retorno = "Cancelado"
			case 3
				retorno = "Finalizado"
			case 4	
				retorno = "Todas"
			end select 
	end if

	formataSituacao = retorno
end function

function formataParticipou(codParticipou)
	dim retorno
	 select case codParticipou
	 	case 0
			retorno = "Indiferente"
		case 1
			retorno = "Não"
		case 2
			retorno = "Sim"
	end select
	formataParticipou = retorno
end function
%>

<%call fechaBanco%>