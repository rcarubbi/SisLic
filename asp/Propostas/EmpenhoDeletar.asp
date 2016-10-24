<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<%call abreBanco%>
<%
dim idProcesso, motivo, numProcesso
idProcesso = request("idProcesso")
numProcesso = getNumProcesso(idProcesso)
if request.Form("isPostBack") <> "" then
	motivo = request.form("motivoCancelamento")
	if trim(motivo) <> "" then
		DeletaEmpenho idProcesso, motivo
		response.Redirect("empenhoLista.asp")
	else
		erro = "O motivo do cancelamento é obrigatório. "
	end if
end if

sub Impressao
end sub

sub auxiliar

end sub

sub principal

%>
	<form name="form1" method="post" action="EmpenhoDeletar.asp">
		<input type="hidden" name="isPostBack" value="1">
		<input type="hidden" name="idProcesso" value="<%=idProcesso%>">	
				
			<font class="titulo">Cancelamento do Empenho ref. ao Processo nº: <%=numProcesso%></font><br><BR>
			<label class="subtitulo">Informe o motivo do cancelamento</label>
		 	<br>
		  	<label class="cabecalho">Motivo:</label>
		  	<br>
  	      	<textarea class="caixa" cols="40" rows="3" name="motivoCancelamento" onKeyUp="maxLengthTextArea(this, event, 255);"><%=motivoCancelamento%></textarea>
		  	<br>
		   	<BR>
			<font class="subtitulo">Comfirma o cancelamento do Empenho ref. ao Processo nº: <%=numProcesso%>?</font><br>
			<font class="erro">Atenção: Após o cancelamento do empenho não é possível recuperá-lo!</font>
			<p>
				<input type="submit" name="btnSim" value="Sim" class="botao">&nbsp;
				<input type="button" name="btnNao" value="Não" class="botao" onClick="window.location.href='Empenholista.asp'">			
			</p>
			<label class="erro"><%=erro%></label>
	</form>

<%end sub

sub DeletaEmpenho(idProcesso, motivoCancelamento)
	dim sql, x, rsnotas
	' recuperando as notas referentes a este empenho
	sql = "Select numnf from empenhonf where idprocesso=  " & idprocesso
	set rsnotas = conn.execute(sql)
	while not rsnotas.eof
		' cancelando Lancamentos Referentes as notas fiscais do empenho
		sql = "update fluxoCaixa set cancelado = 1, dataCancelamento = getDate(), motivoCancelamento= '" & motivoCancelamento & "' where origem = 1 and idRelacional=" & rsnotas("numnf")
		conn.execute(sql)
		rsnotas.movenext
	wend
	rsnotas.close : set rsnotas = nothing
	' deletando notas fiscais referentes ao Empenho
	sql = "Delete from empenhoNF where idProcesso =" & idProcesso
	conn.execute(sql)
	' deletando Empenho
	sql = "Delete From empenho where idProcesso = " & idProcesso
	conn.execute(sql)
	' setando a proposta como pendente novamente
	sql = "update proposta set situacao = 0 where idprocesso=" & idprocesso
	conn.execute(sql)
end sub

%>
<!--#include file="../layout/layout.asp"-->
<%call FechaBanco%>