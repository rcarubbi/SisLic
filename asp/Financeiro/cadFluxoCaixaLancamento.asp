<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="./funcoes.asp"-->
<%
call abreBanco

dim acao, id

acao = request("acao")
id = request("id")
redir = request("redir")
sub Impressao
end sub

sub auxiliar
end sub


sub principal

	dim codPlanoConta, referente, Valor, VencimentoDia, motivoCancelamento
	
	
	if request.Form("isPostBack") = "1" and acao  = "add" then
		codPlanoConta = request.Form("codPlanoConta")
		referente = request.Form("referente")
		valor = request.Form("valor")
				
		if not isNumeric(codPlanoConta) or codPlanoConta = "" or codPlanoConta = "0" then 
			erro = erro & "Conta inválida. "
			codPlanoConta = 0
		end if
		
		if not isNumeric(valor) or valor = "" or valor = "0" then erro = erro & "Valor inválido. "
		
			
		if erro = "" then

			select case acao
			case "add"
				
				FluxoCaixaLancar "", CodPlanoConta, Referente, Valor
				
			
			case "edit"
				sql = "UPDATE FluxoCaixa " & _
				"SET CodPlanoConta = " & CodPlanoConta & _
				  ",referente = '"&referente&"' " & _
				  ",valor = "&valor& _
				  ",DataCancelamento = getDate() " & _
				  "WHERE id = " & id
				'conn.execute(SQL)
			end select
			
			response.Redirect("FluxoCaixaLancamentoLista.asp")
			
		end if
	elseif  request.Form("isPostBack") = "1" and acao  = "del"  then 
		motivoCancelamento = request("motivoCancelamento")
		sql = "Update fluxocaixa set cancelado = 1, dataCancelamento = getDate(),  motivoCancelamento = '" & motivoCancelamento & "' where id=" & id
		
		
		
		if trim(motivoCancelamento) = "" then
			erro = "Preencha o motivo do cancelamento. "
		end if
		
		if erro = "" then
		conn.execute(sql)
		response.Redirect("FluxoCaixaLancamentoLista.asp")	
		end if
	elseif  acao  = "vermot"  then 
		sql = "Select motivoCancelamento, referente from fluxocaixa where id = " & id
		set rs = conn.execute(sql)
		if not rs.eof then
			motivoCancelamento = rs("motivoCancelamento")
			referente = rs("referente")
		end if
		rs.close : set rs = nothing
	end if
	
%>
	<script language="javascript" src="../Funcoes/funcoes.js" type="text/javascript"></script>
	<form action="" method="post" enctype="application/x-www-form-urlencoded" name="form1" target="_self">
		<input type="hidden" name="acao" value="<%=acao%>">
		<input type="hidden" name="id" value="<%=id%>">
		<input type="hidden" name="isPostBack" value="1">
		 <%if acao = "add" then%>
		  <label class="titulo">Cadastro de Lançamentos</label><br>
		  <label class="subtitulo">Informe os dados do lançamento</label><br>
		  <br>
			
		  <label class="cabecalho">Conta (C=Crédito, D=Débito):</label>
		  <br>
		   <%
		   sql = "Select codConta, " & _				
					"descr = case " & _
						"when tipo=0 then '[C] ' + descricao " & _
						"else '[D] ' + descricao " & _
					"end  " & _				
					"from PlanoContas where ativo=1 and codConta not in(select codcontasuperior from planocontas where codcontasuperior is not null	and ativo = 1) order by descr"
			
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
		  <br>
		  <input name="btnOK" type="submit" class="botao" id="btnOK" value="OK">
		  <input name="btnCancelar" type="button" class="botao" id="btnCancelar" value="Cancelar" onClick="window.location.href='<%=redir%>'">
	  <%elseif acao = "del" then%>
		  <label class="titulo">Cancelamento de Lançamentos</label><br>
		  <label class="subtitulo">Informe o motivo do cancelamento</label><br>
		  <br>
		  <label class="cabecalho">Motivo:</label>
		  <br>
  	      <textarea class="caixa" cols="40" rows="3" name="motivoCancelamento" onKeyUp="maxLengthTextArea(this, event, 255);"><%=motivoCancelamento%></textarea>
		   <br>
		   <BR>
		   <input name="btnOK" type="submit" class="botao" id="btnOK" value="OK">
		   <input name="btnCancelar" type="button" class="botao" id="btnCancelar" value="Cancelar" onClick="window.location.href='<%=redir%>'">
	  <%elseif acao = "vermot" then%>
	  	 <label class="titulo">Motivo de Cancelamento</label><br>
		 <label class="subtitulo">Lançamento ref. à <%=referente%></label><br>
		 <br>
		 <label class="cabecalho">Motivo:</label>
		 <br>
  	     <textarea class="caixa" cols="40" rows="3" name="motivoCancelamento" onKeyUp="maxLengthTextArea(this, event, 255);"><%=motivoCancelamento%></textarea>
		 <br>
		 <BR>
		  <input name="btnCancelar" type="button" class="botao" id="btnOk" value="OK" onClick="window.location.href='<%=redir%>'">
      <%end if%>
		<P class="erro">
		&nbsp;<%=erro%>
		</P>
	</form>
	
<%end sub


%>

<!--#include file="../layout/layout.asp"-->
<%call FechaBanco%>
