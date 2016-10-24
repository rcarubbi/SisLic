<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="../funcoes/gridGeral.asp"-->
<!--#include file="funcoes.asp"-->
<%call abreBanco%>

<%
dim acao, msg, filtros, PrevPagto


PrevPagto = request.form("PrevPagto")
filtros = request.form("filtros")
acao = request.Form("acao")

if acao = "add" then
	response.Redirect("cadFluxoCaixaPrevisao.asp?acao=add")
elseif acao = "edit" then
	response.Redirect("cadFluxoCaixaPrevisao.asp?acao=edit&id=" & request.Form("optLinha"))
elseif acao = "del" then
	if isNumeric(request.Form("optLinha")) and request.Form("optLinha") <> "" then
		FCPDeleta request.Form("optLinha")
	end if
elseif acao = "lan" then
	if isNumeric(request.Form("optLinha")) and request.Form("optLinha") <> "" then
		sql = "SELECT * from FluxoCaixaPrevisao where id=" & request.Form("optlinha")
		set rs = conn.execute(SQL)
		if not rs.eof then
			FluxoCaixaLancar "", rs("CodPlanoConta"), rs("Referente"), rs("Valor")
			msg = "Lançamento realizado com sucesso!"
		end if
		
		rs.close
		set rs = nothing
		
	end if
end if

sub principal

	formLista
	
end sub

sub Impressao
end sub

sub auxiliar%>

	<form name="form1" action="" method="post" enctype="application/x-www-form-urlencoded">
		<input type="hidden" name="filtros" value="1">
		<%
		dim itemTexto(3), itemValor(3)
		
		itemTexto(0) = "Todas"
		itemValor(0) = 0
		
		itemTexto(1) = "Para hoje"
		itemValor(1) = 1
		
		itemTexto(2) = "Próxima Semana"
		itemValor(2) = 2
		
		itemTexto(3) = "Para o mês"
		itemValor(3) = 3
		
			
		%><label class="cabecalho">Prev. Pagto.</label><br><%
		Combo "PrevPagto", itemTexto, itemValor, request.Form("PrevPagto")
		
		%>
		<p>
		<input type="submit" name="OK" value="OK" class="botao">
		</p>
	</form>
	<%
end sub

sub FormLista
	dim colunas(4,3), botoes(3, 4), sql, rs
	
	situacao = request.Form("situacao")
	
	if situacao =  "" then situacao = 0
		
	colunas(0,0) = "Plano Conta"
	colunas(1,0) = 4
	colunas(2,0) = "30%"

	colunas(0,1) = "Valor"
	colunas(1,1) = 2
	colunas(2,1) = "25%"
	
	colunas(0,2) = "Vencimento Dia"
	colunas(1,2) = 3
	colunas(2,2) = "15%"

	colunas(0,3) = "Ref."
	colunas(1,3) = 1
	colunas(2,3) = "30%"
	
	'botoes
	botoes(0,0) = true 'novo
	botoes(0,1) = true 'editar
	botoes(0,2) = true  'deletar
	botoes(0,3) = false	 'recuperar
	
	botoes(0,4) = true
	botoes(1,4) = "Lançar Agora"
	botoes(2,4) = "lan"
	botoes(3,4) = true

	
	sql = "SELECT     FluxoCaixaPrevisao.Id, FluxoCaixaPrevisao.Referente, FluxoCaixaPrevisao.Valor, " & _
                      "FluxoCaixaPrevisao.VencimentoDia, PlanoContas.Descricao AS PlanoContaDescricao " & _
		  "FROM        PlanoContas INNER JOIN " & _
                      "FluxoCaixaPrevisao ON PlanoContas.CodConta = FluxoCaixaPrevisao.CodPlanoConta " 
					  
		select case prevpagto  
		 	
			case "1"
				sql = sql & " WHERE FluxoCaixaPrevisao.VencimentoDia = " & day(now)
			case "2"
				sql = sql & " WHERE FluxoCaixaPrevisao.VencimentoDia between " & day(now) & " and " & day(now) + 6
			case "3"
				sql = sql & " WHERE FluxoCaixaPrevisao.VencimentoDia >= " & day(now)
		end select 
		
		
		
	
			  
	sql = sql & "ORDER BY PlanoContas.Descricao"
	set rs = conn.execute(sql)

	
	' fim da configuração do grid
%>
<form action="FluxoCaixaPrevisaoLista.asp" method="post">

	<label class="titulo">Fluxo caixa automático</label>
	<br>
	<label class="subtitulo">Lista de lançamentos pré-cadastrados</label>
	<p class="erro">
		&nbsp;<%=msg%>
	</p>
	<%
	monta_grid rs, "", "", "", 0, 0, Colunas, Botoes, ""
	rs.close : set rs = nothing



%></form><%
end sub



%>
<!--#include file="../layout/layout.asp"-->
<%call FechaBanco%>
 