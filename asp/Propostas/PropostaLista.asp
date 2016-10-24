<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="../funcoes/gridgeral.asp"-->
<%call abreBanco%>

<%
dim acao, idProcesso
dim dtIni
dim dtFim

acao = request.Form("acao")
idProcesso = request.Form("optLinha")
	
if acao = "edit" then
	session.Contents.Remove("PropostaItens")
	response.Redirect("CadProposta.asp?redir=PropostaLista.asp&idProcesso=" & idProcesso & "&acao=" & acao)
elseif acao = "del" then
	if isNumeric(idProcesso) and idProcesso<> "" then
		conn.execute("DELETE FROM Proposta WHERE idProcesso=" & idProcesso)
	end if
elseif acao = "Proposta Perdida" then
	if isNumeric(idProcesso) and idProcesso<> "" then
		conn.execute("sp_propostaSituacao(2," & idProcesso & ")")
		'PropostaSituacao idProcesso, 2
	end if	
elseif acao = "rec" then
	if isNumeric(idProcesso) and idProcesso<> "" then
		conn.execute("sp_propostaSituacao(0," & idProcesso & ")")
		'PropostaSituacao idProcesso, 0
	end if
elseif acao = "gerarEmp" then
	response.Redirect("GeraEmpIni.asp?idProcesso=" & idProcesso)
elseif acao = "vizualizarProp" then
	response.Redirect("VizualizaProposta.asp?idProcesso=" & idProcesso)
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
	response.write request.ServerVariables("SCRIPT_NAME") & "?imp=1&idprocesso=" & idprocesso
end sub


sub auxiliar
	%>
	<form name="form1" action="" method="post" enctype="application/x-www-form-urlencoded">
		<input type="hidden" name="filtros" value="1">
		<%
		dim itemTexto(2), itemValor(2)
		
		itemTexto(0) = "Pendentes"
		itemValor(0) = 0
		
		itemTexto(1) = "Empenhadas"
		itemValor(1) = 1
		
		itemTexto(2) = "Perdidas"
		itemValor(2) = 2
				
		
		
		%><label class="cabecalho">Situação</label><br><%
		Combo "situacao", itemTexto, itemValor, request.Form("situacao")
		
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
	
	situacao = request.Form("situacao")
	'por padrao situação é sempre pendente
	if situacao = "" then situacao = "0"
			
	sql = "Select e.numProcesso, replace(pa.nome,char(13),'<br>') as cliente, " & _
		"pr.valorTotal, pr.dataCadastro, u.Nome, pr.IdProcesso as usuario from Proposta pr INNER JOIN " & _
		"edital e on pr.idProcesso = e.idProcesso INNER JOIN usuarios u " & _
		"on u.cpf=pr.cpfUsuario INNER JOIN parceiro pa on e.idparceiro = pa.id "& _
		"WHERE pr.situacao=" & situacao & " AND (pr.dataCadastro BETWEEN " & _
		formata_data_sql(dtIni) & " AND " & formata_data_sql(dtFim & " 23:59:59") & _
		") order by pr.DataCadastro Desc "
		 
	set rs = conn.execute(sql)
	
	' configurando o grid
	dim colunas(5,4), botoes(3, 7)
	
	colunas(0,0) = "Núm. do Processo"
	colunas(1,0) = 0
	colunas(2,0) = "10%"
	
	colunas(0,1) = "Cliente"
	colunas(1,1) = 1
	colunas(2,1) = "25%"
	
	colunas(0,2) = "Valor Total"
	colunas(1,2) = 2
	colunas(2,2) = "15%"
	
	colunas(0,3) = "Criada por"
	colunas(1,3) = 4
	colunas(2,3) = "20%"
	
	colunas(0,4) = "Data de Cadastro"
	colunas(1,4) = 3
	colunas(2,4) = "10%"

	
	botoes(0,7) = true
	botoes(1,7) = "Visualizar"
	botoes(2,7) = "vizualizarProp"
	botoes(3,7) = true

	if situacao = 0 then
	
		botoes(0,0) = false
		botoes(0,1) = true
		botoes(0,2) = true	
		botoes(0,3) = false		
		
		botoes(0,4) = true
		botoes(1,4) = "Gerar Empenho"
		botoes(2,4) = "gerarEmp"
		botoes(3,4) = true
		
		botoes(0,5) = true
		botoes(1,5) = "Proposta Perdida"
		botoes(2,5) = "Proposta Perdida"
		botoes(3,5) = true	
				
				
	elseif situacao = 2 then
		botoes(0,0) = false
		botoes(0,1) = false
		botoes(0,2) = false	
		botoes(0,3) = true
		
		
	elseif situacao = 3 then
		botoes(0,0) = false
		botoes(0,1) = false
		botoes(0,2) = false	
		botoes(0,3) = true	
		
		botoes(0,4) = false
		botoes(0,5) = false
		
		
	elseif situacao = 1 then
		botoes(0,0) = false
		botoes(0,1) = false
		botoes(0,2) = false	
		botoes(0,3) = false			
		
		botoes(0,4) = false
		botoes(0,5) = false	
		
		botoes(0,7) = true
	end if				

	
	
	' fim da configuração do grid
	%>
	<form action="<%=request.servervariables("SCRIPT_NAME")%>" method="post">
	<label class="titulo">Proposta</label>
	<br>
	<label class="subtitulo">Lista das Propostas - 
		<%
		select case situacao
		case 0
		%>
		(Pendentes)
		<%
		case 1
		%>
		(Empenhadas)
		<%		
		case 2
		%>
		(Perdidas)
		<%		
		end select
	%>
	</label><br>	
	<%
	
		monta_grid rs, "", "", "", 0, 5, Colunas, Botoes, ""
		rs.close : set rs = nothing
	%></form><%
end sub


%>
<!--#include file="../layout/layout.asp"-->
 