<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<%
call abreBanco

dim acao, codConta, erro, tipo

acao = request("acao")
codConta = request("id")
tipo = request("tipo")

if tipo = "" then 
	tipo = 0
end if

sub Impressao
end sub

sub auxiliar
end sub


sub principal
	dim codContaSuperior, descricao
	
	if request.Form("isPostBack") = "" and acao = "edit" then
		sql = "select * from planocontas where codConta = " & codConta
		set rs = conn.execute(SQL)
		if not rs.eof then
			descricao = rs("descricao")
			codContaSuperior = rs("codContaSuperior")
		end if
		rs.close
		set rs = nothing
		
	elseif request.Form("isPostBack") = "1" then
		codContaSuperior = request.Form("CodContaSuperior")
		descricao = request.Form("descricao")	
		
		if codConta = "" or codConta = "0" AND not IsNumeric(codConta) then 
			erro = erro & "Código inválido. "
		elseif acao = "add" then
			if verifica_existe(codConta, "codConta", "PlanoContas") then 
				erro = erro & "Código já usado por outra conta. "
			end if
		end if
		
		if codContaSuperior = "0" or not IsNumeric(codContaSuperior) then codContaSuperior = "NULL"
		if descricao = "" then erro = erro & "Descrição inválida. "
	
		if erro = "" then
		
			select case acao
				case "add"
		
					SQL = "INSERT INTO PlanoContas (" & _
								"codConta, " & _
								"codContaSuperior, " & _
								"descricao, " & _
								"tipo " & _
								") Values ( " & _
								codConta & ", " & _
								codContaSuperior & ", " & _
								"'" & descricao & "', " & _
								"'" & tipo & "' " & _
								")"
					conn.execute(SQL)
				
				case "edit"
					SQL = "UPDATE PlanoContas set " & _
								"codContaSuperior=" & codContaSuperior & ", " & _
								"descricao='" & descricao & "' " & _
								"WHERE codConta=" & codConta
					conn.execute(SQL)
				
			end select
	
			response.Redirect("PlanoContasLista.asp?tipo=" & tipo)
		end if
	end if
	
	
%>
	<form action="cadPlanoContas.asp" method="post" enctype="application/x-www-form-urlencoded" name="form1" target="_self">
		<input type="hidden" name="acao" value="<%=acao%>">
		<input type="hidden" name="isPostBack" value="1">
		<input type="hidden" name="tipo" value="<%=tipo%>">		
	 
		<label class="titulo">Cadastro de Plano de Contas</label><br>
		<label class="subtitulo">Informe os dados do Plano de Conta</label><br>
		<br>
		
		<label class="cabecalho">Código:</label>
		<br>
		<input name="id" type="text" class="caixa" id="id" size="10" maxlength="7" value="<%=CodConta%>" <%
			if acao = "edit" then response.Write("readonly")
		%>>
		<br>		
		<label class="cabecalho">Descrição:</label>
		<br>
		<input name="descricao" type="text" class="caixa" id="descricao" size="40" maxlength="255" value="<%=descricao%>">
		<br>
		<label class="cabecalho">Conta Mãe</label>
		<br>
		<%
		if acao = "edit" then 'retorna todas as contas menos ela mesma (caso ela já esteja cadastrada)
			ComboDB "codContaSuperior","Descricao","codConta","Select codConta, descricao from PlanoContas where codConta <> " & codConta & " AND tipo = " & tipo & " and ativo = 1", codContaSuperior , "",""
		else
			ComboDB "codContaSuperior","Descricao","codConta","Select codConta, descricao from PlanoContas where tipo = " & tipo & " and ativo = 1", codContaSuperior , "",""
		end if
		%>
		<br>
		<br>

		<input name="btnOK" type="submit" class="botao" id="btnOK" value="OK">
		<input name="btnCancelar" type="button" class="botao" id="btnCancelar" value="Cancelar" onClick="window.location.href='PlanoContasLista.asp?tipo=<%=tipo%>'">
		<P class="erro">
		&nbsp;<%=erro%>
		</P>
	</form>
<%end sub


%>

<!--#include file="../layout/layout.asp"-->
<%call FechaBanco%>
