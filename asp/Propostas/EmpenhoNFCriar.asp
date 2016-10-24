<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="../funcoes/gridGeral.asp"-->

<script language="javascript" src="../Funcoes/funcoes.js" type="text/javascript"></script>

<%call abreBanco%>
<%
dim acao, erro
dim valorTotal, percImposto, valorImposto, valorTotalImposto, descricao, observacao, numNF
dim codPlanoContaImposto, codPlanoContaNota
dim data
dim idProcesso

acao = request("acao")
idProcesso = request("idProcesso")

if request.Form("isPostBack") = "1" then
	numNF = request.Form("numNF")
	valorTotal = request.Form("valorTotal")
	percImposto = request.Form("percImposto")
	valorImposto = request.Form("valorImposto")	
	valorTotalImposto = request.Form("valorTotalImposto")
	descricao = request.Form("descricao")
	observacao = request.Form("observacao")
	data = recebe_datador("data", 0)
	codPlanoContaImposto = request.Form("codPlanoContaImposto")
	codPlanoContaNota = request.Form("codPlanoContaNota")
	
	if trim(numNF) = "" or not isNumeric(numNF) then erro = "Número da Nota Fiscal inválido. "
	if not isNumeric(valorTotal) or valorTotal = "" then erro = erro & "Valor total inválido. "
	if not isNumeric(percImposto) or percImposto = "" then erro = erro & "Percentual de imposto inválido. "	
	if not isNumeric(valorImposto) or valorImposto = "" then erro = erro & "Valor do imposto inválido. "
	if valorTotalImposto = "" or Not IsNumeric(valorTotalImposto) then erro = erro & "Valor total do imposto inválido. "	
	if descricao = "" then erro = erro & "Ref. inválida. "
	if not isdate(data) then erro = erro & "Data inválida. "
	
	if not isNumeric(codPlanoContaImposto) or codPlanoContaImposto = "" then 
		erro = erro & "Plano de contas para os Impostos inválida. "
	
	end if
	if not isNumeric(codPlanoContaNota) or codPlanoContaNota = "" then 
		erro = erro & "Plano de contas para as Notas inválida. "
	end if
	
	if erro = "" then
		'response.Write("a" & acao)
		if acao = "add" then 
			sql = "select numNF from empenhoNF where numNF = " & numNF
			set rs = conn.execute(SQL)
			if not rs.eof then
				erro = "Número da nota fiscal já cadastrada. "
			end if
			rs.close
			set rs = nothing
			
			if erro = "" then
				sql = "INSERT INTO EmpenhoNF " & _
							   "(NumNF " & _
							   ",Descricao" & _
							   ",Observacoes" & _
							   ",ValorTotal" & _
							   ",idProcesso" & _
							   ",data" & _
							   ",codPlanoContaImposto" & _
							   ",codPlanoContaNota" & _														   
							   ",impostos)" & _
							   "VALUES" & _
							   "(" & NumNF & _
							   ",'" & Descricao & "'" & _
							   ",'" & Observacoes & "'" & _
							   "," & ValorTotal & _
							   "," & idProcesso & _
							   "," & formata_data_sql(data) & _
							   "," & codPlanoContaImposto & _
							   "," & codPlanoContaNota & _
							   "," & valorImposto & ")"
							'  response.Write(sql)
							'response.write(ValorTotalImposto & " " & valorImposto)
							 ' response.End()
				conn.execute(SQL)
				
				response.Redirect("EmpenhoLista.asp")
			end if

		elseif acao = "edit" then 



		end if

		
	end if
elseif acao = "edit" then
	
elseif acao = "del" then

else

	dim rAux
	
	if application("par_26") <> "" and isNumeric(application("par_26")) then
		sql = "Select codConta from PlanoContas where ativo=1 and codConta=" & application("par_26")
		set rAux = conn.execute(SQL)
		if not rAux.eof then
			codPlanoContaNota = rAux("codConta")
		end if
		rAux.close
		set rAux = nothing
	end if
	if application("par_27") <> "" and isNumeric(application("par_27")) then
		sql = "Select codConta from PlanoContas where ativo=1 and codConta=" & application("par_27")
		set rAux = conn.execute(SQL)
		if not rAux.eof then
			codPlanoContaImposto = rAux("codConta")
		end if
		rAux.close
		set rAux = nothing
	end if
end if


sub principal
%>

<script language="javascript">

	function calculaVlImposto(){

		var x = form1.valorTotal.value
		x = x.replace(".", "");
		x = x.replace(",",".");

		if((isNaN(x) || isNaN(form1.percImposto.value)) == false) {
			form1.valorImposto.value = (parseFloat(x) * (parseFloat(form1.percImposto.value/100)));
			form1.valorImposto.value = parseFloat(form1.valorImposto.value).toFixed(2);
			calculaVlTotalNota();
		}
	}
	function calculaVlTotalNota(){
		var x = form1.valorTotal.value
		x = x.replace(".", "");
		x = x.replace(",",".");
		form1.valorTotalImposto.value = parseFloat(form1.valorImposto.value) + parseFloat(x);
		form1.valorTotalImposto.value = parseFloat(form1.valorTotalImposto.value).toFixed(2);
	}
</script>
<form name="form1" id="form1" method="post" action="empenhoNFCriar.asp">
	<input type="hidden" name="isPostBack" value="1">
	<input type="hidden" name="acao" value="<%=acao%>">
	<input type="hidden" name="idProcesso" value="<%=idProcesso%>">
	
  <p><font class="titulo">Nota fiscal</font><br>
    <font class="subtitulo">Nova nota fiscal</font><br>
    <br>

	<label class="cabecalho">Número da Nota Fiscal:</label><br>
	<input name="numNF" type="text" class="caixa" id="numNF" size="10" maxlength="20" value="<%=NumNF%>">   <br>

	<label class="cabecalho">Data:</label><br>
	<%datador "data", 0, year(now) - 10, year(now) + 1, data%>
	
	
<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td width="120" valign="bottom">
			<label class="cabecalho">Valor da nota (sem os impostos):</label><br>
			<input name="valorTotal" onChange="javascript: calculaVlImposto();" type="text" class="caixa" id="valorTotal" size="10" maxlength="15" value="<%=valorTotal%>">   
		</td>
		<td valign="bottom">
			<font class="cabecalho">Plano de Contas p/ a Nota:</font><br>
			<%
			
			sql = "select * from planoContas where tipo = 0 and ativo = 1 and codConta not in" & _
					"(select isNull(codContaSuperior, 0) from planoContas)"
			
			ComboDB "codPlanoContaNota", "descricao", "codConta", sql,  codPlanoContaNota, "",""
			%>
		</td>
	</tr>	
</table>
		
	<label class="cabecalho">(%)Imposto:</label><br>
	<input name="percImposto" onChange="javascript: calculaVlImposto();" type="text" class="caixa" id="percImposto" size="10" maxlength="15" value="<%=percImposto%>">   

<table border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td width="120">
			<label class="cabecalho">Valor Imposto:</label><br>
			<input name="valorImposto" type="text" class="caixa" id="valorImposto" style="background-color:#c9c9c9"  readonly size="10" maxlength="15" value="<%=valorImposto%>">   
		</td>
		<td>
			<font class="cabecalho">Plano de contas p/ o Imposto:</font><br>
		<%
		
		sql = "select * from planoContas where tipo = 1 and ativo = 1 and codConta not in" & _
				"(select isNull(codContaSuperior, 0) from planoContas)"
				
		ComboDB "codPlanoContaImposto", "descricao", "codConta", sql,  codPlanoContaImposto, "",""
		%>
		</td>
	</tr>	
</table>
	
	<label class="cabecalho">Valor total da nota (com os impostos):</label><br>
	<input name="valorTotalImposto" readonly style="background-color:#c9c9c9 " type="text" class="caixa" id="valorTotalImposto" size="10" maxlength="8" value="<%=valorTotalImposto%>">   
	<br>
	<label class="cabecalho">Ref:</label><br>
	<input name="descricao" type="text" class="caixa" value="<%=descricao%>" size="40">
	<br>	
	<label class="cabecalho">Observações</label><br>
    <textarea name="observacao" cols="40" rows="5" class="caixa" id="descricao" onKeyUp="maxLengthTextArea(this, event, 200);"><%=observacao%></textarea>

        <br><br>
    <input name="btnOK" type="submit" id="btnOK" value="OK" class="botao">
    <input name="btnCancelar" type="button" id="btnCancelar" value="Cancelar" class="botao" onClick="window.location.href='empenhoLista.asp'">
  <p class="erro">
		&nbsp;<%=erro%>
	</p>
</p>
  </form>

<%
end sub


sub Impressao
end sub

sub auxiliar

end sub

%>

<!--#include file="../layout/layout.asp"-->
<%call FechaBanco%>