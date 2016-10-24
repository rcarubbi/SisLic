<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<%call abreBanco%>

<%dim cpfUsuario, Situacao, clienteid, sql, rs, vcor, cor, dtini, dtfim
cpfUsuario = request("cpfUsuario")
Situacao = request("situacao")
clienteid = request("clienteid")
situacaoDesc = formataSituacaoProp(situacao)
dtIni = recebe_datador("dtIni",0)
dtFim = recebe_datador("dtFim",0)
if not isDate(dtIni) then
	dtini = dateadd("m",-1,date)
	dtfim = date
end if
if cdate(dtFim) < cdate(dtIni) then dtfim = dtini

sub principal
	
	if request("isPostBack") then
			
		sql = "SELECT edi.numProcesso, pro.ValorTotal, pro.Situacao, usu.Nome as usuario, replace(par.nome,char(13),'<br>') as Cliente from proposta pro " & _
		      "inner join edital edi on edi.idProcesso=pro.idprocesso " & _ 
			  "inner join usuarios usu on usu.cpf=pro.cpfUsuario " & _ 
			  "inner join parceiro par on par.id=edi.idparceiro " & _
			  "where pro.idprocesso <> 0 and pro.dataCadastro between " & formata_data_sql(dtIni) & " and " & formata_data_sql(dtFim & " 23:59:59") 
			if cpfUsuario <> 0 then
					sql = sql & " and Usu.cpf = " & cpfUsuario
			end if
			if Situacao <> 4 then
				sql = sql & " and pro.Situacao = " & Situacao
			end if
			if clienteid <> 0 then
				sql = sql & " and par.id = " & clienteid
			end if
			
		set rs = conn.execute(sql)
		%>
			<label class="titulo">Relatório de Propostas</label><br>
			<label class="subtitulo">Situação: <%=situacaoDesc%></label><BR>
			<label class="subtitulo">Responsável: <%=iif(cpfusuario=0,"Todos",getValor(cpfUsuario, "cpf", "nome", "usuarios"))%></label><BR>
			<label class="subtitulo">Cliente: <%=iif(Clienteid=0,"Todos",getValor(clienteid, "id", "nome", "parceiro"))%></label><BR>
			<label class="subtitulo">No período de: <%=dtini%> à <%=dtfim%></label>
			
			
			<table style="border-collapse:collapse;" border="1" width="100%">
				<thead>
					<tr class="cabecalho" style="background-color:#f5f5f5;">
						<th width="">Núm. do Processo</th>
						<th width="">Cliente</th>
						<th width="">Valor da Proposta</th>
						<th width="">Situação</th>
						<th width="">Responsável</th>
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
					<td class="texto"><%=rs("numProcesso")%></td>
					<td class="texto"><%=rs("cliente")%></td>
					<td class="texto"><%=formata_Campo(ccur(rs("Valortotal")))%></td>
					<td class="texto"><%=formataSituacaoProp(rs("situacao"))%></td>
					<td class="texto"><%=rs("usuario")%></td>
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
	response.write request.ServerVariables("SCRIPT_NAME") & "?imp=1&ispostback=1&situacao=" & situacao & "&cpfUsuario=" & cpfUsuario & "&clienteid=" & clienteid
end sub

sub auxiliar%>
<form name="filtros" method="post" action="">
<input type="hidden" value="1" name="isPostBack">
<table>
	<tr>
		<td><label class="cabecalho">Situação</label><br>
			<%dim itemTexto, itemValor
			redim itemTexto(3)
			redim itemValor(3)
			itemTexto(0) = "Todas"
			itemValor(0) = 4
			itemTexto(1) = "Pendente"
			itemValor(1) = 0
			itemTexto(2) = "Empenhada"
			itemValor(2) = 1
			itemTexto(3) = "Arquivada"
			itemValor(3) = 2
			Combo "Situacao", itemTexto, itemValor, request.Form("situacao")%>
		</td>
	</tr>
	<tr>
		<td><label class="cabecalho">Responsável</label><br>
			  <%ComboDB "cpfUsuario","nome","cpf","Select nome, cpf from Usuarios",cpfUsuario,"Todos",""%>
			
		</td>
		
	</tr>
	<tr>
		<td><label class="cabecalho">Cliente</label><br>
			    <%ComboDB "clienteid","Nome","id","Select id, replace(nome,char(13),'. ') as nome from parceiro where tipo = 0 and ativo = 1",clienteid,"Todos","230" %>
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
<% function formataSituacaoProp(CodSituacao)
dim retorno
	if codsituacao <> "" then
		select case codsituacao
			case 4
				retorno = "Todas"
			case 0
				retorno = "Pendente"	
			case 1
				retorno = "Empenhada"
			case 2
				retorno = "Arquivada"
			
			end select 
	end if

	formataSituacaoProp = retorno
end function


%>

<%call fechaBanco%>