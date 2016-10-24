<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="funGeraEmpenho.asp"-->
<%call abreBanco%>
<%
dim idProcesso, erro
dim prazo, validade
dim dataEmpenho, Observacao, valorTotal


idProcesso = request("idProcesso")
prazo = request("prazo")
validade = request("validade")
dataEmpenho = request("dataEmpenho")
valorTotal = request("valorTotal")

observacao = session("emp_observacao")

if request.Form("isPostBack") = "1" then
	dataEmpenho = recebe_Datador("dataEmpenho", 0)
	if not isNumeric(valorTotal) or valorTotal = "" then erro = erro & "Valor total inválido. "
	if not isDate(dataEmpenho) then erro = erro & "Data do empenho inválido. "
	if not isNUmeric(prazo) or prazo= "" then erro = erro & "Prazo inválido. "
	if not isNumeric(validade) or validade = "" then erro = erro & "Validade inválida. "
	observacao = request.Form("observacao")

	session("emp_observacao") = observacao
	
	if erro = "" then
		response.Redirect("GeraEmpConfirm.asp?idProcesso=" & idProcesso & "&prazo=" & prazo & "&validade=" & validade & "&dataEmpenho=" & dataEmpenho & "&valorTotal=" & valorTotal)
	end if
	
end if

sub Principal
	if tempFaltaItens(idProcesso) then
		%>
		<label class="erro">Ainda há itens não compostos (sem + Compra ou + Estoque).</label>
		<br><br>
		<input type="button" name="Voltar" value="Voltar" class="botao" onClick="window.location.href='GeraEmpItemLista.asp?idProcesso=<%=idProcesso%>&ori=EmpPrazo'">
		<%
		exit sub
	end if
	
	%>
	<form method="post" action="GeraEmpPrazo.asp" enctype="application/x-www-form-urlencoded" name="form1" id="form1">
        <input name="idProcesso" type="hidden" id="idProcesso" value="<%=idProcesso%>">
		<input type="hidden" name="isPostBack" value="1">
		<p>
		<label class="titulo">Empenho</label><br>
		<label class="subtitulo">Informações do Empenho</label>
		</p>

		<label class="cabecalho">Data do Empenho:</label><br>
		<%datador "dataEmpenho", 0, year(now) - 10, year(now) + 1, dataEmpenho%>
		<br>
		<label class="cabecalho">Valor total:</label><br>
		<input name="valorTotal" type="text" value="<%=formatnumber(valorTotal,2)%>" size="20" maxlength="10" class="texto">
		<br>
		<label class="cabecalho">Prazo (dias):</label><br>
		<input name="prazo" type="text" value="<%=prazo%>" size="10" maxlength="8" class="texto">
		<br>
		<label class="cabecalho">Validade (dias):</label><br>
		<input name="validade" type="text" value="<%=validade%>" size="10" maxlength="8" class="texto">
		<br>
		<label class="cabecalho">Observações:</label><br>
		<textarea name="observacao" cols="30" rows="3" class="texto" id="observacao"><%=observacao%></textarea>
		<br>
		

		<label class="erro"><%=erro%>&nbsp;</label>
		<p>
		<input type="button" name="Voltar" value="Voltar" class="botao" onClick="window.location.href='GeraEmpItemLista.asp?idProcesso=<%=idProcesso%>&ori=Prazo'">
		<input type="submit" name="Proximo" value="Próximo" class="botao">
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

