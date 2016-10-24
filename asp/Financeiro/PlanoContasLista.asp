<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="../funcoes/gridgeral.asp"-->
<%call abreBanco%>

<%
dim acao, tipo, situacao

if request.Form("filtros") = "1" then
	situacao = request.Form("situacao")
else
	situacao = request.querystring("situacao")
end if

'por padrao situação é sempre ativo
if situacao = "" or not isNumeric(situacao) then 
	situacao = 1
else
	situacao = CInt(situacao)
end if

tipo = request("tipo")

if tipo = "" or not isNumeric(tipo) then tipo = 0

acao = request.Form("acao")
select case acao
case "add"
	response.Redirect("cadPlanoContas.asp?acao=add&tipo=" & tipo)
case "edit"
	response.Redirect("cadPlanoContas.asp?acao=edit&id=" & request.Form("optlinha") & "&tipo=" & tipo)	
case "del"
	if request.Form("optLinha") <> "" and isNumeric(request.Form("optLinha")) then
		
		sql = "select * from planocontas where ativo = 1 AND codContaSuperior=" & request.Form("optLinha")
		
		set rs = conn.execute(SQL)
		
		if not rs.eof then
			erro = "Essa conta não pode ser deletada pois ela possui contas filhas. Delete antes as contas filhas para depois deletar essa conta."
		else
			sql = "update planocontas set ativo = 0 where codConta=" & request.Form("optLinha")
			conn.execute(SQL)
		end if
		
		rs.close
		set rs = nothing
	end if
case "rec"
	if request.Form("optLinha") <> "" and isNumeric(request.Form("optLinha")) then
		
		sql = "SELECT     PlanoContas_pai.ativo " & _
				"FROM     PlanoContas LEFT OUTER JOIN " & _
                    		"PlanoContas AS PlanoContas_pai ON PlanoContas.CodContaSuperior = PlanoContas_pai.CodConta " & _
				"WHERE 	  PlanoContas.codConta=" & request.Form("optLinha")
		
		set rs = conn.execute(SQL)
		
		
		if not rs.eof then

			if isNull(rs("ativo")) then
				ok = true
			elseif rs("ativo") = false then
				ok = false
			elseif rs("ativo") = true then
				ok = true
			end if

		else
			ok = true
		end if

		if ok then
			sql = "update planocontas set ativo = 1 where codConta=" & request.Form("optLinha")
			conn.execute(SQL)
		else
			erro = "Essa conta não pode ser recuperada pois ela é filha de uma conta que esta deletada. Recupere a conta pai antes de recuperar essa conta."
		end if		
		
		rs.close
		set rs = nothing

	end if

end select


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
		itemValor(0) = 1
		
		itemTexto(1) = "Deletados"
		itemValor(1) = 0
		

		%><label class="cabecalho">Situação</label><br><%
		Combo "situacao", itemTexto, itemValor, situacao
		
		%>
		<p>
		<input type="submit" name="OK" value="OK" class="botao">
		</p>
	</form>
	<%
end sub

sub FormLista
	
	dim  mConta
	

			
'	sql = "select CodConta, Descricao, TipoDesc = " & _
'			"case when tipo = 0 then 'Crédito' " & _
'			    "when tipo = 1 then 'Débito' " & _
'			"end " & _
'			"from PlanoContas WHERE tipo=" & tipo & " AND ativo=" & situacao

	redim mConta(2,-1)

	if situacao = 1 then
		montaContaAtiva mConta, "NULL", 0
	else
		montaContaDeletada mConta
	end if
	
	' configurando o grid
	dim t
	
	if situacao = 1 then
		t = 0
	else
		t = 1
	end if
	
	dim colunas()
	
	redim preserve colunas(4, t)
	
	colunas(0,0) = "Descrição"
	colunas(1,0) = 1
	colunas(2,0) = "100%"	

	'só exibe o nome da tabela pai se esta deletado
	if situacao = 0 then
		colunas(0,1) = "Conta Pai"
		colunas(1,1) = 2
		colunas(2,1) = "50%"		
	end if

	dim botoes(3, 3)

		
	'botoes
	if situacao = 1 then
		botoes(0,0) = true
		botoes(0,1) = true
		botoes(0,2) = true		
		botoes(0,3) = false	
	else
		botoes(0,0) = false
		botoes(0,1) = false
		botoes(0,2) = false			
		botoes(0,3) = true		
	end if

	' fim da configuração do grid
	%>
	<form action="PlanoContasLista.asp?tipo=<%=tipo%>&situacao=<%=situacao%>" method="post">
	<label class="titulo">Plano de Contas
	<%if tipo = 0 then
		%>(Crédito<%
	else
		%>(Débito<%
	end if
	
	if situacao = 0 then
		%> - Deletadas )<%
	else
		%>)<%
	end if
	%>
	</label>
	<br>
	<label class="subtitulo">Lista dos Plano de Contas</label>
	<p class="erro">
		&nbsp;<%=erro%>
	</p>
	<%
	
		monta_grid mConta, "", "", "", 0, 0, Colunas, Botoes, ""
	
	
	%></form><%
end sub

'as contas deletadas não tem nivel
sub MontaContaDeletada(ByRef mConta)
	dim rs
	
	'sql = "select codConta, descricao from planoContas where tipo=" & tipo & " and ativo=0"
	sql = "SELECT     PlanoContas.CodConta, PlanoContas.Descricao, PlanoContas_1.Descricao AS Pai " & _
			"FROM     PlanoContas left JOIN " & _
					 "PlanoContas AS PlanoContas_1 ON PlanoContas.CodContaSuperior = PlanoContas_1.CodConta " & _
					 "WHERE     PlanoContas.ativo=0 and PlanoContas.tipo = " & tipo
	'response.Write(sql)
	set rs = conn.execute(SQL)
	
	while not rs.eof
		addConta mConta, rs("codConta").value, rs("descricao").value, rs("pai")
		rs.movenext
	wend
	
	rs.close
	set rs = nothing
	
end sub

'as contas ativas tem nivel
sub montaContaAtiva(byRef mConta, codContaSuperior, nivel)
	dim rs
	
	sql = "select codConta, descricao from planoContas where tipo=" & tipo & " AND ativo=1"
	if codContaSuperior = "NULL" then
		sql = sql & " AND isNull(codContaSuperior,0) = 0"
	else
		sql = sql & " AND codContaSuperior=" & codContaSuperior
	end if
	
	sql = sql & " ORDER by codContaSuperior"
	'response.Write(sql & "<hr>")
	set rs = conn.execute(SQL)

	while not rs.eof
	
		addConta mConta, rs("codConta").value, (string2(nivel, "&nbsp;&nbsp;") & "" & rs("descricao").value), ""
		
		montaContaAtiva mConta, rs("codConta").value, nivel + 1
		
		rs.movenext		
	wend
	
	rs.close
	set rs = nothing
	
end sub

'função igual a STRING, mas a STRING ao passar um codigo html, ela formata para asp, da errado...!!!!
function String2(qt, str)
	dim rtn
	for i = 1 to qt
		rtn = rtn & str
	next
	String2 = rtn
end function

sub addConta(byRef mConta, codConta, descricao, pai)
	redim preserve mConta(2, ubound(mConta,2) + 1)
	
	mConta(0, ubound(mConta,2)) = codConta
	mConta(1, ubound(mConta,2)) = descricao
	mConta(2, ubound(mConta,2)) = pai	

end sub
%>

<!--#include file="../layout/layout.asp"-->
 