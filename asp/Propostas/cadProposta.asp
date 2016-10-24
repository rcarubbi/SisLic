<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->	
<!--#include file="../funcoes/gridgeral.asp"-->
<script language="javascript" src="../Funcoes/funcoes.js" type="text/javascript"></script>
<%
call abreBanco

dim acaoItem, idProcesso, erro, acao
acao = request("acao")
acaoItem = request("acaoItem")
idProcesso = request("idProcesso")
redir = request("redir")


sub Impressao
end sub

sub auxiliar
end sub


sub principal
	dim contcad, contItens
	dim numProcesso, ValorTotal 
	dim sql, rs, sqlItens, rsItens, indice, propostaItens
	
	contcad=request.QueryString("contcad")
	
	

	if request.form("ispostback") <> "" or contcad<>"1" then 

		session("OutrasInf") = request.form("OutrasInf")
		session("infcabecalho") = request.form("infcabecalho")
		session("responsavel") = request.form("responsavel")
		session("dataProposta") = recebe_datador("dataProposta",0)
		if acaoItem = "add" or acaoItem = "edit" then ' se vai pra tela de novo item da proposta salvar o que ja foi preenchido até então.
			'response.write(request.Form("optLinha"))
'response.end
			response.redirect("cadPropItem.asp?redir="&server.urlencode("CadProposta.asp?contcad=1&acao="&acao&"&idProcesso="&idProcesso&"&redir="&redir)&"&iditem="&request.Form("optLinha")&"&idProcesso="&idProcesso&"&acaoItem="&acaoItem)
			
		elseif acaoItem = "del" then
		
			indice = getIndice(request.Form("optLinha"),session("propostaItens"))
			DelItemProposta(indice)
			
		end if

		
		if acao = "edit" and request.Form("isPostBack") = "" then
		'caso a primeira vez que passa e ação igual a editar recuperar do banco os dados no formulario
					
			SQL = "Select p.OutrasInf, p.InfCabecalho, p.cpfResponsavel, p.dataProposta from proposta p where idProcesso='" & idProcesso & "'"

			set rs = conn.execute(sql)
			if not rs.eof then
				
				session("OutrasInf") = rs("OutrasInf")	
				session("infCabecalho") = rs("infCabecalho")
				session("responsavel") = rs("cpfResponsavel")
				session("dataProposta") = rs("dataProposta")
				sqlItens = " Select i.numItem, i.especificacao, i.ValorUnitario, i.Quantidade, (i.Quantidade*i.valorunitario) as TotItem, i.unidade, i.marca" & _
					        " from propostaItem i where idProcesso=" & idProcesso 
							
				set rsItens = conn.execute(sqlItens)
				redim PropostaItens(7,-1)  
				contItens = 1
				while not rsItens.eof 
					redim preserve propostaitens(7,ubound(PropostaItens,2)+1)
						propostaItens(0,ubound(propostaItens,2)) = contItens
					for x=1 to ubound(propostaItens)
						propostaItens(x,ubound(propostaItens,2)) = rsItens.fields(x-1)
					next
					contItens = contItens + 1
					rsItens.movenext	
				wend
				session("PropostaItens") = PropostaItens
				rsItens.close : Set rsItens = nothing
			end if
							
			rs.close : set rs = nothing
		elseif request.Form("isPostBack") <> "" and request.Form("btnOKProp") <> "" then ' se é postBack recuperar dados do formulario
				
			' Validações
			erro = ""	
			
			if isarray(Session("propostaItens")) then
				if ubound(session("propostaItens"),2) < 0 then erro = "A proposta deve possuir ao menos um item. "
			else
				erro = "A proposta deve possuir ao menos um item. "
			end if
			if not isdate(session("dataProposta")) then
				erro = erro & "Data da proposta inválida. "
			end if
			if session("responsavel") = "" or not isNumeric(session("responsavel")) or session("responsavel") = "0" then
				erro = erro & "Selecione um responsável para assinar a proposta. "
			end if
			
			' Fim das Validações
		
			
		
		end if
		
							
		if erro = "" and request.form("ispostback") <> "" and request.Form("btnOKProp") <> ""  then
			
			valorTotal = getTotalItens()
		
			select case acao
				case "add"
				
				
						SQL = "INSERT INTO Proposta " & _
							  " (IdProcesso " & _
							  " ,ValorTotal " & _
							
							  " ,OutrasInf " & _
							
							  " ,DataCadastro " & _
							  " ,CpfUsuario " & _ 
							  " ,infCabecalho " & _
							   " ,dataProposta " & _
							  " , cpfResponsavel " & _
							  " )" & _ 
						  " VALUES " & _
							  " ("&IdProcesso& _
							  " ,"&replace(ValorTotal,",",".")& _
							
							  " ,'"&session("OutrasInf")&"'"& _
							
							  " ,'"&now&"'"& _
							  " ,"&session("usuario_cpf")&_
							  " ,'"&session("infCabecalho")&"'"&_
							  " ,'"&session("dataProposta")&"'"&_
							  " ,"&session("responsavel")&_
							  ")"
							  
							  
							
							
							  conn.execute(sql)
							  
							gravaPropostaitens()
							
						
					
			case "edit"
				
		
				SQL = "UPDATE Proposta " & _
					"  SET " &_
					"      ValorTotal = "&replace(ValorTotal,",",".")& _
					"     ,OutrasInf = '"&session("OutrasInf")&"'"& _
					"     ,infCabecalho = '"&session("infcabecalho")&"'"& _
					"     ,dataProposta = '"&session("dataProposta")&"'"& _
					"     ,cpfresponsavel = "&session("responsavel")& _
					"     WHERE idProcesso = " & idProcesso
					
						conn.execute(sql)
						  
						  
					
					editaPropostaItens()
						  
				conn.execute(SQL)
				 
			end select
			
			
			session("OutrasInf") = ""
			session("infCabecalho") = ""
			session("responsavel") = ""
			session("dataProposta") = ""
			
			response.Redirect(redir)
		end if
	
	end if
%>
	
	   <form action="cadProposta.asp" method="post" enctype="application/x-www-form-urlencoded" name="form1" target="_self">
		  <label class="titulo">Cadastro de Propostas</label><br>
		  <label class="subtitulo">Informe os dados da Proposta</label><br>
		  <br>
	 	
	  <label class="cabecalho">Núm. do Processo</label>
	  <br>
	  <input name="NumProcesso" type="text" class="caixa" id="NumProcesso" size="25" maxlength="20" readonly value="<%=getNumProcesso(idProcesso)%>">
	  <br>
	   <label class="cabecalho">Informações adicionais de cabeçalho</label> <br>
	  <textarea name="InfCabecalho" Cols="113" rows="7" class="caixa" id="InfCabecalho" onKeyUp="maxLengthTextArea(this, event, 1000);"><%=session("InfCabecalho")%></textarea><br>
	
	  <%PropostaItensLista%>
	
		<input type="hidden" name="acao" value="<%=acao%>">
		
		<input type="hidden" name="redir" value="<%=redir%>">		
		<input type="hidden" name="isPostBack" value="1">
		<input type="hidden" name="idProcesso" value="<%=idProcesso%>">
	 	
	
	  
	
	    <label class="cabecalho">Outras Informações</label> <br>
	  <textarea name="OutrasInf" Cols="113" rows="7" class="caixa" id="OutrasInf" onKeyUp="maxLengthTextArea(this, event, 1000);"><%=session("OutrasInf")%></textarea><br>
	<BR>
		<Label class="cabecalho">Data da Proposta</Label><BR>
		<%datador "dataProposta", 0, year(now)-1, year(now) + 5, session("dataProposta")%><BR><BR>	
		<label class="cabecalho">Responsável</label> <br>
		<%ComboDB "responsavel", "nome", "cpf", "Select pr.cpf, pr.nome from propostaResponsavel pr where pr.ativo=1", cDbl(session("Responsavel")), "","200"%>
	<BR><BR>
	  <input name="btnOKProp" type="submit" class="botao" id="btnOKProp" value="OK">
	  <input name="btnCancelar" type="button" class="botao" id="btnCancelar" value="Cancelar" onClick="window.location.href='<%=redir%>'">
	  <P class="erro">
		&nbsp;<%=erro%>
	  </P>
	</form>
<%end sub%>

<%
sub PropostaItensLista
	dim sql, rs
	
	if not isarray(session("PropostaItens")) then
		redim PropostaItens(7,-1)
	else
		PropostaItens = session("PropostaItens")
	end if
	
	' configurando o grid
	dim colunas(4,4), botoes(3, 4)
	
	colunas(0,0) = "Núm. do Item"
	colunas(1,0) = 1
	colunas(2,0) = "15%"
		
	colunas(0,1) = "Valor Unitário"
	colunas(1,1) = 3
	colunas(2,1) = "25%"	
		
	colunas(0,2) = "Quantidade"
	colunas(1,2) = 4
	colunas(2,2) = "15%"
		
	colunas(0,3) = "Unidade"
	colunas(1,3) = 6
	colunas(2,3) = "15%"
	
	colunas(0,4) = "Valor Total"
	colunas(1,4) = 5	
	colunas(2,4) = "25%"
	colunas(3,4) = true
		
	botoes(0,0) = true
	botoes(0,1) = true
	botoes(0,2) = true		
	
							
	' fim da configuração do grid
	%>
	<br>
	<label class="subtitulo">Itens da proposta</label>
	
	<%
	
	monta_grid PropostaItens, "acaoitem", "", "", 0, 0, colunas, botoes, ""
		
end sub

Function DelItemProposta(indice)
	dim itens, i
	itens = session("PropostaItens")
	
	'se deletar um item no meio da matriz
	'passar todos os itens para frente
	'e depois excluir o ultimo item
	
	if indice >= 0 then
	for i = indice to ubound(itens,2) -1
		
		itens(0,i) = itens(0,i+1)
		itens(1,i) = itens(1,i+1)
		itens(2,i) = itens(2,i+1)
		itens(3,i) = itens(3,i+1)
		itens(4,i) = itens(4,i+1)
		itens(5,i) = itens(5,i+1)
		itens(6,i) = itens(6,i+1)
		itens(7,i) = itens(7,i+1)
	next
		redim preserve itens(7, ubound(itens,2)-1)
	else
		redim itens(7,-1)
	end if

	session("PropostaItens") = itens
		
	'retorna o novo tamanho da matriz
	DelItemProposta = ubound(itens,2)
end function

function getIndice(idItem,PropostaItens)
dim x, retorno


	for x = 0 to ubound(PropostaItens,2)
	
		if PropostaItens(0,x) = idItem then
			retorno = x
			exit for
		else
			retorno = -1
		end if
	next
	getIndice = retorno
end function

sub gravaPropostaItens()
	dim propostaitens, sql, x
	
	propostaItens = session("propostaItens")

	for x = 0 to ubound(PropostaItens,2)

		SQL = "INSERT INTO PropostaItem"& _
			" (IdProcesso"& _
			" ,NumItem"& _
			" ,especificacao"& _
			" ,unidade"& _
			" ,quantidade"& _
			" ,ValorUnitario"&_
			" ,marca)"& _
			"  VALUES"& _
			" ("&IdProcesso& _
			" ,'"&PropostaItens(1,x)&"'"& _
			" ,'"&PropostaItens(2,x)&"'"& _
			" ,'"&PropostaItens(6,x)&"'"& _
			" ,"&PropostaItens(4,x)& _
			" ,"&replace(PropostaItens(3,x), ",", ".")& _
			" , '"&PropostaItens(7,x)&"')"
			
		conn.execute(SQL)
		
	
	next
	session.Contents.Remove("PropostaItens")
	
end sub

sub editaPropostaItens()

	'Como a atualização de registros que esta na matriz com o banco de dados é MUITO complicado
	'deletei todos os registros que estão na propostaItem referente ao processo, e 
	'gravo td de novo
	sql = "DELETE FROM PropostaItem WHERE idProcesso=" & idProcesso
	
	conn.execute(SQL)
	
	gravaPropostaItens
	
end sub

function getTotalItens()
	dim retorno, x, PropostaItens
	PropostaItens = session("propostaItens")
	for x = 0 to ubound(PropostaItens,2)
		retorno = retorno + cdbl(PropostaItens(5,x))
	next
	getTotalItens = retorno
end function

%>
<!--#include file="../layout/layout.asp"-->
<%call FechaBanco%>

						