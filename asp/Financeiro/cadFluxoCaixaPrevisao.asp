<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<%
call abreBanco

dim acao, id

acao = request("acao")
id = request("id")

sub Impressao
end sub

sub auxiliar
end sub


sub principal

	dim codPlanoConta, referente, Valor, VencimentoDia
	
	if request.Form("isPostBack") = "" and acao = "edit" then
		sql = "select * from FluxoCaixaPrevisao where id=" & id
		set rs = conn.execute(SQL)
		
		if not rs.eof then
			codPlanoConta = rs("codPlanoConta")
			referente = rs("referente")
			valor = rs("valor")
			vencimentoDia = rs("VencimentoDia")
		end if
		rs.close
		set rs = nothing
	elseif request.Form("isPostBack") = "1" then
		codPlanoConta = request.Form("codPlanoConta")
		referente = request.Form("referente")
		valor = request.Form("valor")
		vencimentoDia = request.Form("VencimentoDia")	
		
		
		if not isNumeric(codPlanoConta) or codPlanoConta = "" or codPlanoConta = "0" then 
			erro = erro & "Plano de Conta inválido. "
			codPlanoConta = 0
		end if
		
		if not isNumeric(valor) or valor = "" or valor = "0" then erro = erro & "Valor inválido. "
		
		if not isNumeric(vencimentoDia) or vencimentoDia = "" or vencimentoDia = "0" then 
			erro = erro & "Dia de vencimento inválido."
			vencimentoDia = 0
		end if
		
		if erro = "" then

			select case acao
			case "add"
				sql = "INSERT INTO FluxoCaixaPrevisao " & _
							"(CodPlanoConta" & _
							",Referente" & _
							",Valor" & _
							",VencimentoDia) " & _
						 "VALUES " & _
							 "(" & codPlanoConta &_
							 ", '" & referente & "'" & _
							 ", " & valor & _
							 ", " & vencimentoDia & ")"
							   
				conn.execute(SQL)
			
			case "edit"
				sql ="UPDATE FluxoCaixaPrevisao " & _
						   "SET CodPlanoConta = " & codPlanoConta & _
							  ",Referente = '" & referente & "'" & _
							  ",Valor = " & valor & _
							  ",VencimentoDia = " & vencimentoDia & _
						 "WHERE id=" & id
				conn.execute(SQL)
			end select
			
			response.Redirect("FluxoCaixaPrevisaoLista.asp")
			
		end if
	
	end if
	
%>
	<form action="" method="post" enctype="application/x-www-form-urlencoded" name="form1" target="_self">
		<input type="hidden" name="acao" value="<%=acao%>">
		<input type="hidden" name="id" value="<%=id%>">
		<input type="hidden" name="isPostBack" value="1">
	 
		  <label class="titulo">Cadastro de Lançamentos automáticos</label><br>
		  <label class="subtitulo">Informe os dados do lançamento</label><br>
		  <br>
	 	
	  <label class="cabecalho">Plano de Conta (C=Crédito, D=Débito):</label>
	  <br>
	   <%
	   sql = "Select codConta, " & _				
				"descr = case " & _
					"when tipo=0 then '[C] ' + descricao " & _
					"else '[D] ' + descricao " & _
				"end  " & _				
				"from PlanoContas where ativo=1 order by descr"

	   ComboDB "codPlanoConta","descr","codConta",sql,codPlanoConta, "",""%>
	  <br>
	  <label class="cabecalho">Ref.:</label>
	  <br>
	  <input name="referente" type="text" class="caixa" id="referente" size="40" maxlength="100" value="<%=referente%>">
	  <br>
	  <label class="cabecalho">Valor:</label>
	  <br>
	  <input name="valor" type="text" class="caixa" id="valor" size="40" maxlength="100" value="<%=valor%>">
	  <br>  
  	  <label class="cabecalho">Dia de vencimento:</label>
	  <br> 
   	  <%
		dim itemTexto(31), itemValor(31), i
		
		for i = 1 to 31
			itemTexto(i) = i
			itemValor(i) = i		
		next
		Combo "VencimentoDia", itemTexto, itemValor, vencimentoDia	  
	  %>
	  <br>
      <br>

	  <input name="btnOK" type="submit" class="botao" id="btnOK" value="OK">
	  <input name="btnCancelar" type="button" class="botao" id="btnCancelar" value="Cancelar" onClick="window.location.href='<%=redir%>'">
	  <P class="erro">
		&nbsp;<%=erro%>
	  </P>
	</form>
<%end sub


%>

<!--#include file="../layout/layout.asp"-->
<%call FechaBanco%>
