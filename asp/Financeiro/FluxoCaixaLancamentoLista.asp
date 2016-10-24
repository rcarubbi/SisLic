<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="../funcoes/gridGeral.asp"-->
<!--#include file="funcoes.asp"-->
<%call abreBanco%>

<%
dim acao, msg, situacao, sqlCancelado, rsCancelado, CancelamentoAprovado
dim dtini, dtfim

situacao = request("situacao")
acao = request.Form("acao")

if acao = "add" then
	response.Redirect("cadFluxoCaixaLancamento.asp?redir=FluxoCaixaLancamentoLista.asp&acao=add")
elseif acao = "delesp" then
	sqlCancelado = "SELECT idRelacional from fluxocaixa where id = " & request.form("optLinha")
	set rsCancelado = conn.execute(sqlCancelado)
	if isnull(rsCancelado("idRelacional"))  then
		CancelamentoAprovado = true
	else
		CancelamentoAprovado = false
	end if
	rsCancelado.close : set rsCancelado = nothing
	
	if CancelamentoAprovado then
		response.Redirect("cadFluxoCaixaLancamento.asp?redir=FluxoCaixaLancamentoLista.asp&acao=del&id="&request.form("optlinha"))
	else
		msg = "Só é possível cancelar lançamentos manuais. "
	end if
elseif acao = "vermot" then
	response.Redirect("cadFluxoCaixaLancamento.asp?redir=FluxoCaixaLancamentoLista.asp&acao=vermot&id="&request.form("optlinha"))
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
	<form name="form1" action="" method="post" enctype="application/x-www-form-urlencoded">
		<input type="hidden" name="filtros" value="1">
		<%
		dim itemTexto(1), itemValor(1)
		
		itemTexto(0) = "Ativos"
		itemValor(0) = 0
		
		itemTexto(1) = "Cancelados"
		itemValor(1) = 1
				
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
	dim colunas(4,3), botoes(3, 5), sql, rs
	
	situacao = request.Form("situacao")
	
	if situacao =  "" then situacao = 0
		
	colunas(0,0) = "Plano Conta"
	colunas(1,0) = 4
	colunas(2,0) = "30%"

	colunas(0,1) = "Ref."
	colunas(1,1) = 1
	colunas(2,1) = "30%"
	
	colunas(0,2) = "Data de Lançamento"
	colunas(1,2) = 3
	colunas(2,2) = "30%"	
	
	colunas(0,3) = "Valor"
	colunas(1,3) = 2
	colunas(2,3) = "25%"

	
	'botoes
	botoes(0,0) = false 'novo
	botoes(0,1) = false 'editar
	botoes(0,2) = false  'deletar	
	botoes(0,3) = false  'recuperar
	
	if situacao = 0 then
		botoes(0,0) = true 'novo
		botoes(0,4) = true
		botoes(1,4) = "Deletar"
		botoes(2,4) = "delesp"
		botoes(3,4) = true
	else
		botoes(0,5) = true
		botoes(1,5) = "Ver Motivo"
		botoes(2,5) = "vermot"
		botoes(3,5) = true
	end if
	
	
	
	sql = "SELECT     FluxoCaixa.Id, FluxoCaixa.Referente, FluxoCaixa.Valor, FluxoCaixa.DataLancamento, PlanoContas.Descricao AS ContaDescricao " & _
			" FROM    FluxoCaixa INNER JOIN " & _
            "		  PlanoContas ON FluxoCaixa.CodPlanoConta = PlanoContas.CodConta" &_
			" WHERE   cancelado = " & situacao  & " and fluxoCaixa.dataLancamento  between " & formata_data_sql(dtIni) & " and " & formata_data_sql(dtFim & " 23:59:59")    
			
				   
			
		   
	set rs = conn.execute(sql)

	
	' fim da configuração do grid
%>

<form action="<%=request.ServerVariables("SCRIPT_NAME")%>" method="post">

	<label class="titulo">Lançamentos</label>
	<br>
	<label class="subtitulo">Fluxo de Caixa</label>
	<p class="erro">
		&nbsp;<%=msg%>
	</p>
	<%
	monta_grid rs, "", "", "", 0, 0, Colunas, Botoes, ""
	rs.close : set rs = nothing



%></form>
 
<%
end sub



%>
<!--#include file="../layout/layout.asp"-->
<%call FechaBanco%>
 