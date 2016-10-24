<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="../funcoes/gridgeral.asp"-->
<%call abreBanco%>

<%
dim acao
dim erro
dim dtIni
dim dtFim

acao = request.Form("acao")

if acao = "edit" or acao = "add" then
	response.Redirect("CadEdital.asp?redir=EditalLista.asp&idProcesso=" & request.Form("optLinha") & "&acao=" & acao)
elseif acao = "del" then
	idProcesso = request.Form("optLinha")
	if isNumeric(idProcesso) and idProcesso<> "" then
		if not temPropostas(idProcesso) then	
			sql = "exec sp_editalDeletaRecupera 2," & idProcesso
			conn.execute(SQL)
		else
			erro = "Esse Edital tem propostas pendentes, cancele as propostas antes de cancelar o edital!"
			
		end if
	end if
elseif acao = "rec" then
	idProcesso = request.Form("optLinha")
	if isNumeric(idProcesso) and idProcesso<> "" then
		sql = "exec sp_editalDeletaRecupera 1," & idProcesso
		conn.execute(SQL)
	end if
elseif acao = "AddProp" then
	session.Contents.Remove("PropostaItens")
	response.Redirect("CadProposta.asp?redir=EditalLista.asp&idProcesso=" & request.Form("optLinha") & "&acao=add")
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
		dim itemTexto(3), itemValor(3)
		
		itemTexto(0) = "Pendentes"
		itemValor(0) = 0
		
		itemTexto(1) = "Em andamento"
		itemValor(1) = 1
		
		itemTexto(2) = "Cancelados"
		itemValor(2) = 2
		
		itemTexto(3) = "Finalizados"
		itemValor(3) = 3
		
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
	'por padrao situação é sempre ativo
	if situacao = "" then situacao = "0"
			
	sql = "Select e.idProcesso, e.NumProcesso, e.Descricao, t.Descricao as Tipo, replace(p.Nome,char(13),'<br>') as Cliente, e.DataCadastro from Edital e INNER JOIN" & _ 
		  " EditalTipo t ON t.id = e.IDEditalTipo INNER JOIN Parceiro p ON p.id = e.IdParceiro" & _ 
		  " WHERE e.situacao = " & situacao & " and e.datacadastro between " & formata_data_sql(dtIni) & " and " & formata_data_sql(dtFim & " 23:59:59") & " order by DataCadastro desc"
		  'response.write sql
		  'response.End()	
	set rs = conn.execute(sql)
	
	' configurando o grid
	dim colunas(4,4), botoes(3, 4)
	
	colunas(0,0) = "Núm. do Processo"
	colunas(1,0) = 1
	colunas(2,0) = "10%"
		
	colunas(0,1) = "Edital"
	colunas(1,1) = 2
	colunas(2,1) = "70%"
		
	colunas(0,2) = "Tipo"
	colunas(1,2) = 3
	colunas(2,2) = "25%"
	
	colunas(0,3) = "Cliente"
	colunas(1,3) = 4
	colunas(2,3) = "25%"
	
	colunas(0,4) = "Data de Cadastro"
	colunas(1,4) = 5
	colunas(2,4) = "25%"
	

	select case situacao 
		case 0 ' pendentes
			botoes(0,0) = true
			botoes(0,1) = true
			botoes(0,2) = true		
			botoes(0,3) = false	
			' verifica se o usuario tem permissao para manipular as propostas
			
			if not TemRestricoes(17) then
				botoes(0,4) = true		
				botoes(1,4) = "Nova Proposta" 			
				botoes(2,4) = "AddProp"
				botoes(3,4) = true		
			end if
			
		case 1 ' em andamento
			botoes(0,0) = false
			botoes(0,1) = true
			botoes(0,2) = true		
			botoes(0,3) = false			
		case 2 ' cancelados
			botoes(0,0) = false
			botoes(0,1) = false
			botoes(0,2) = false		
			botoes(0,3) = true	
		case 3 ' finalizados
			botoes(0,0) = false
			botoes(0,1) = false
			botoes(0,2) = false		
			botoes(0,3) = false	
	end select
	
	
	' fim da configuração do grid
	%>
	<form action="EditalLista.asp" method="post">
	<input type="hidden" name="situacao" value="<%=situacao%>">
	<label class="titulo">Edital</label>
	<br>
	<label class="subtitulo">Lista dos Editais</label>
	<p class="erro">
		<%=erro%>
	</p>
	<%
	
		monta_grid rs, "", "", "", 0, 0, Colunas, Botoes, ""
		rs.close : set rs = nothing
	%></form><%
end sub


function temPropostas(idProcesso)
	dim sql
	dim rs
	
	sql = "Select * from proposta where idProcesso=" & idProcesso
	
	set rs = conn.execute(SQL)
	
	if not rs.eof then
		temPropostas = true
	else
		temPropostas = false
	end if
	
	rs.close
	set rs = nothing
end function
'dim x 

'set rs = conn.execute("select codConta, codContaSuperior, descricao from planocontas where ativo=1 order by codContaSuperior")
'x = rs.getrows

'getContas 0, 0
'rs.close
'set rs = nothing

'Function getContas(id_superior, nivel)
'	dim contaSuperior
'	dim  i
	
'	for i = 0 to ubound(x,2)
		'verifica se é a conta raiz
		'se é raiz, conver o nulo para zero
'		if isNull(x(1,i)) then
'			contaSuperior = 0
'		else
'			contaSuperior = x(1,i)
'		end if
		'verifica se o item a ser exibido pertence a conta pai passada como parametro
'		if cInt(contaSuperior) = cInt(id_superior) then
			'exibe os itens tabulados de acordo com o nivel
'			response.Write(string(nivel, "__") & x(2,i) & "<br>")
			
'			getContas x(0,i), nivel + 1 
			
'		end if	
'	next
'end function
%>
<!--#include file="../layout/layout.asp"-->
 