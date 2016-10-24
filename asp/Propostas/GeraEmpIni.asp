<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->

<script language="javascript" type="text/javascript" src="../Funcoes/funcoes.js"></script>

<%call abreBanco%>
<%

dim idProcesso, erro, acao

idProcesso = request("idProcesso")

session("emp_observacao") = ""

sub principal
	%>
<form action="GeraEmpItemLista.asp" method="post" enctype="application/x-www-form-urlencoded" name="form1" id="form1">
	<input type="hidden" name="idProcesso" value="<%=idProcesso%>">
	<%
	'###################################################
	' exibindo os dados Da proposta
	
	SQL = "SELECT     Edital.NumProcesso, Proposta.*, Usuarios.Nome " & _
		 	"FROM         Edital INNER JOIN " & _
            	          "Proposta ON Edital.IdProcesso = Proposta.IdProcesso INNER JOIN " & _
                	      "Usuarios ON Proposta.CpfUsuario = Usuarios.CPF " & _
			"WHERE Proposta.idProcesso = " & idProcesso
			
	
	set rs = conn.execute(SQL)
	
'	response.Write(sql)
'	response.End()
	if not rs.eof then
		%>
		<label class="titulo">Gerar empenho do Processo n.º: <%=rs("NumProcesso") & "?"%></label><br>
		<label class="subtitulo">Selecione os itens ganhos.</label>
		<br><br>
		<span class="subtitulo">Proposta</span>
	
			<table width="99%"  border="1" cellspacing="3" cellpadding="0" style="border-collapse:collapse ">

			  <tr>
				<td valign="top"><span class="cabecalho">N.º Processo</span><br>
				 <span class="texto"><%=rs("NumProcesso")%></span></td>
				<td valign="top"><span class="cabecalho">Responsável</span><br>
				  <span class="texto"><%=rs("nome")%></span></td>			
				<td valign="top"><span class="cabecalho">Data de cadastro</span><br>
				  <span class="texto"><%=formata_campo(rs("DataCadastro").value)%></span></td>
			<!--<tr>
				<td valign="top"><span class="cabecalho">Prazo de entrega</span><br>
				  <span class="texto"><%'=formata_campo(rs("PrazoEntrega").value)%></span></td>-->
			
				<!--<td valign="top"><span class="cabecalho">Válido até</span><br>
				  <span class="texto"><%'=formata_campo(rs("ValidadeProposta").value)%></span></td>
			
				<td valign="top"><span class="cabecalho">Garantia</span><br>
				  <span class="texto"><%'=rs("garantia")%> meses</span></td>
			  </tr>					
			  <tr>
				<td valign="top"><span class="cabecalho">Vl. Total</span><br>
				  <span class="texto"><%'=formata_campo(rs("ValorTotal").value)%></span></td>			  
				<!--<td valign="top"><span class="cabecalho">Impostos</span><br>
				  <span class="texto"><%'=formata_campo(rs("Impostos").value)%></span></td>			
				<td valign="top"><span class="cabecalho">Condição de Pagamento</span><br>
				  <span class="texto"><%'=rs("CondicaoPagto")%></span></td> 
			  </tr>					--> 
							
			  <tr>
				<td colspan="3" valign="top"><span class="cabecalho">Outras Informações</span><br>
				  <span class="texto"><%=rs("OutrasInf")%></span></td>
			  </tr>					          
			</table>
	<%
	else
		%><span class="erro">Nenhum registro encontrado</span>
			<p><a href="#" class="texto" onClick="histoy.back(-1);">Voltar</a></p>
		<%
	end if
	
	rs.close
	set rs = nothing
	' fim dados pedido
	'##################################################
	%>		
	<br>
	
	<%
	'##################################################
	' endereco do cliente
	
	
	SQL = "Select * from PropostaItem where idProcesso=" & idProcesso
	set rs = conn.execute(SQL)
	%>
	<span class="subtitulo">Itens</span>
	<table width="99%"  border="1" cellspacing="0" cellpadding="3" style="border-collapse:collapse ">
	<tr>
  	  <td></td>
	  <td nowrap><label class="cabecalho">Núm. Item</label></td>
	  <td><label class="cabecalho">Especificação</label></td>
	  <td><label class="cabecalho">Unidade</label></td>	  
	  <td align="right"><label class="cabecalho">Vl. Unitário</label></td>
	  <td align="right"><label class="cabecalho">Qt.</label></td>
	  <td align="right"><label class="cabecalho">Vl. Total</label></td>								
	</tr>
	<%
	dim cor
	cor =false
	while not rs.eof
	cor = not cor
	
	if not cor then 
		corFundo = "#eeeeee"
	else
		corFundo = ""
	end if
	%>



							<tr onClick="javascript: mudaCor(this, <%="'" & corFundo & "', '#ffffcc'"%>); chk_<%=rs("id")%>.checked =! chk_<%=rs("id")%>.checked;" style="cursor:hand;" bgcolor="<%=corFundo%>">
							  <td><input type="checkbox" name="chkItens" id="chk_<%=rs("id")%>" value="<%=rs("id")%>" onClick="this.checked =! this.checked"></td>
							  <td><span class="texto"><%=rs("numitem")%></span></td>							  
							  <td><span class="texto"><%=rs("especificacao")%></span></td>
							  <td><span class="texto"><%=rs("unidade")%></span></td>							  
							  <td align="right"><span class="texto"><%=formata_campo(rs("ValorUnitario").value)%></span></td>
							  <td align="right"><span class="texto"><%=formata_campo(rs("quantidade").value)%></span></td>
							  <td align="right"><span class="texto"><%=formata_campo(rs("quantidade") * rs("ValorUnitario"))%></span></td>								
							</tr>


	
	<%	rs.movenext
	wend
	rs.close
	set rs = nothing
	%>	
	</table>
	<p>
	<input type="button" onClick="window.location.href='PropostaLista.asp'" value="Voltar" class="botao">
	<input type="submit" value="Próximo" class="botao">
	</p></form>
	<%
end sub

sub Impressao
end sub

sub auxiliar
	
end sub



%>
    
<!--#include file="../layout/layout.asp"-->

