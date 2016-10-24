<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="../funcoes/gridgeral.asp"-->
<%call abreBanco%>

<%
dim acao, idProcesso
dim dtini, dtfim


acao = request.Form("acao")
idProcesso = request.Form("optLinha")

if acao = "criarNF" then
	response.Redirect("empenhoNFCriar.asp?acao=add&idProcesso=" & idProcesso)
elseif acao = "verNF" then
	response.Redirect("empenhoNFVer.asp?idProcesso=" & idProcesso)
elseif acao = "Del" then
 	response.Redirect("EmpenhoDeletar.asp?idProcesso=" & idProcesso)
end if

dtIni = recebe_datador("dtIni",0)
dtFim = recebe_datador("dtFim",0)


if not isDate(dtIni) then
	dtini = dateadd("m",-1,date)
	dtfim = date
end if
if cdate(dtFim) < cdate(dtIni) then dtfim = dtini
sub principal
	formLista
end sub

sub Impressao
end sub

sub auxiliar
	%>
	<form name="form1" action="empenholista.asp" method="post" enctype="application/x-www-form-urlencoded">
		<input type="hidden" name="filtros" value="1">
		<%
		dim itemTexto(1), itemValor(1)
				
		itemTexto(0) = "Sem nota fiscal"
		itemValor(0) = 0		
		
		itemTexto(1) = "Com nota fiscal"
		itemValor(1) = 1
		
		%><label class="cabecalho">Situação</label><br><%
		Combo "situacao", itemTexto, itemValor, request("situacao")
		
		%>
		<br>
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
		<p>
		<input type="submit" name="OK" value="OK" class="botao">
		</p>
	</form>
	<%
end sub

sub FormLista
	dim sql, rs
	dim situacao
	
	situacao = request("situacao")
	'por padrao situação é sempre pendente
	if situacao = "" then situacao = "0"
			
	sql = "SELECT distinct edital.numProcesso, replace(parceiro.nome,char(13),'<br>') as nome, empenho.DataEmpenho, empenho.ValorTotal, prazo, validade, observacao, empenho.idProcesso " & _
			"FROM empenho inner join edital on empenho.idProcesso = edital.idProcesso inner join parceiro on parceiro.id = edital.idParceiro"
			
	if situacao = "0" then 'sem nota fiscal
		sql = sql & " left JOIN EmpenhoNF ON Empenho.idProcesso = empenhonf.idprocesso"
		sql = sql & " WHERE empenhonf.idprocesso is null and empenho.dataEmpenho between " & formata_data_sql(dtIni) & " and " & formata_data_sql(dtFim & " 23:59:59")
	else 'com nota fiscal
		sql = sql & " INNER JOIN EmpenhoNF ON Empenho.idProcesso = EmpenhoNF.idProcesso"
		sql = sql & " WHERE empenho.dataEmpenho between " & formata_data_sql(dtIni) & " and " & formata_data_sql(dtFim & " 23:59:59")
	end if
	
	
	
	set rs = conn.execute(sql)
	
	' configurando o grid
	dim colunas(5,5), botoes(3, 6)
	
	colunas(0,0) = "Núm. do Processo"
	colunas(1,0) = 0
	colunas(2,0) = "10%"
	
	colunas(0,1) = "Cliente"
	colunas(1,1) = 1
	colunas(2,1) = "25%"
	
	colunas(0,2) = "Data"
	colunas(1,2) = 2
	colunas(2,2) = "15%"
	
	colunas(0,3) = "Valor"
	colunas(1,3) = 3
	colunas(2,3) = "20%"
	
	colunas(0,4) = "Prazo (dias)"
	colunas(1,4) = 4
	colunas(2,4) = "10%"

	colunas(0,5) = "Validade (dias)"
	colunas(1,5) = 5
	colunas(2,5) = "10%"	

	botoes(0,0) = false
	botoes(0,1) = false
	botoes(0,2) = false
	botoes(0,3) = false		
	
	botoes(0,4) = true
	botoes(1,4) = "Criar Nota Fiscal"
	botoes(2,4) = "criarNF"
	botoes(3,4) = true	
	
	if situacao = 1 then
		botoes(0,5) = true
		botoes(1,5) = "Ver notas fiscais"
		botoes(2,5) = "verNF"
		botoes(3,5) = true		
	end if
	
	botoes(0,6) = true
	botoes(1,6) = "Deletar Empenho"
	botoes(2,6) = "Del"
	botoes(3,6) = true	
	
	' fim da configuração do grid
	%>
	<form action="EmpenhoLista.asp" method="post">
	<label class="titulo">Empenhos</label>
	<br>
	<label class="subtitulo">Lista dos Empenhos - 
		<%
		select case situacao
		case 0
		%>
		(Sem nota fiscal)
		<%
		case 1
		%>
		(Com nota fiscal)
		<%		
		end select
	%>
	</label><br>	
	<%
	
		monta_grid rs, "", "", "", 0, 7, Colunas, Botoes, ""
		rs.close : set rs = nothing
	%></form><%
end sub

'arquiva a proposta atraves do id dela
sub propostaArquivar(id)
	dim sql
	
	sql = "update proposta set situacao = 2 where idProcesso = " & id
	conn.execute(SQL)

end sub

'recupera a proposta arquivada atraves do id dela
sub propostaRecuperar(id)
	dim sql
	
	sql = "update proposta set situacao = 0 where idProcesso = " & id
	conn.execute(SQL)

end sub
%>
<!--#include file="../layout/layout.asp"-->
 