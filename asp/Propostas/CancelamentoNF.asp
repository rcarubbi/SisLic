<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<%call abreBanco%>
<%
dim numNF, motivo, idProcesso, numProcesso
numNF = request("numNF")
idProcesso = request("idprocesso")
numProcesso = getNumprocesso(idprocesso)
if request.Form("isPostBack") <> "" then
	motivo = request.form("motivoCancelamento")
	if trim(motivo) <> "" then
		CancelaNota numNF, motivo
		response.Redirect("EmpenhoNFver.asp?idprocesso="&idprocesso)
	else
		erro = " O motivo do cancelamento � obrigat�rio. "
	end if
end if

sub Impressao
end sub

sub auxiliar

end sub

sub principal

%>
	<form name="form1" method="post" action="CancelamentoNF.asp">
		<input type="hidden" name="isPostBack" value="1">
		<input type="hidden" name="idProcesso" value="<%=idProcesso%>">	
		<input type="hidden" name="numnf" value="<%=numnf%>">	
			<font class="titulo">Cancelamento da Nota Fiscal n�: <%=numNF%><BR>Empenho ref. ao processo n�: <%=numProcesso%></font><br><BR>
			
			<label class="subtitulo">Informe o motivo do cancelamento</label>
		 	<br>
		  	<label class="cabecalho">Motivo:</label>
		  	<br>
  	      	<textarea class="caixa" cols="40" rows="3" name="motivoCancelamento" onKeyUp="maxLengthTextArea(this, event, 255);"><%=motivoCancelamento%></textarea>
		  	<br>
		   	<BR>
			<font class="subtitulo">Comfirma o cancelamento da Nota Fiscal n�: <%=numnf%>?</font><br>
			<font class="erro">Aten��o: Ap�s o cancelamento desta nota fiscal n�o � poss�vel recuper�-la!</font>
			<p>
				<input type="submit" name="btnSim" value="Sim" class="botao">&nbsp;
				<input type="button" name="btnNao" value="N�o" class="botao" onClick="window.location.href='empenhonfver.asp?idprocesso=<%=idprocesso%>'">			
			</p>
			<label class="erro"><%=erro%></label>
	</form>

<%end sub

sub CancelaNota(numnf, motivoCancelamento)
	dim sql, x
	
	
	' cancelando Lancamentos Referentes a nota fiscal
	sql = "update fluxoCaixa set cancelado = 1, dataCancelamento = getDate(), motivoCancelamento= '" & motivoCancelamento & "' where origem = 1 and idRelacional=" & numnf
	conn.execute(sql)
		
	' deletando a nota fiscal 
	sql = "Delete from empenhoNF where numnf =" & numnf
	conn.execute(sql)
	

end sub

%>
<!--#include file="../layout/layout.asp"-->
<%call FechaBanco%>