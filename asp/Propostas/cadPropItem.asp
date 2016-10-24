<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<!--#include file="../funcoes/gridgeral.asp"-->
<%
call abreBanco

dim acaoItem, idProcesso, erro, redir, numItem, iditem, iditemant, PropostaItens
Dim especificacao, quantidade, valorUnitario, unidade , marca
dim indiceItem
acaoItem = request("acaoItem")
idProcesso = request("idProcesso")

redir = request("redir")



sub Impressao
end sub

sub auxiliar
end sub


sub principal
		
		'caso a primeira vez que passa e ação igual a editar
		if acaoItem = "edit" and request.Form("isPostBack") = "" then
			
			
			CarregaLinha()
			
		elseif request.Form("isPostBack") <> "" then ' se é postBack recuperar dados do formulario
			idItem = request.form("iditem")
			idItemAnt = request.form("idItemAnt")
			NumItem = request.form("numItem")
			especificacao = request.form("especificacao")
			quantidade = request.form("quantidade")
			valorUnitario = request.form("valorUnitario")
			unidade = request.form("unidade")
			indiceItem= request.form("optlinha")
			marca = request.form("marca")
			
			
			' Validações
			
			
			If not isnumeric(NumItem) or numItem = "" then erro = "Número do item de proposta Inválido. "
			'elseif verifica_existe_matriz(NumItem,0,session("PropostaItens")) and acaoitem="add" then
				'erro = erro & "Já existe outro item de proposta com este número. "
			'elseif numItem <> numItemAnt and acaoitem = "edit" then
			
				'if verifica_existe_matriz(NumItem, 0,session("PropostaItens")) then
					'erro = erro & "Já existe outro item de proposta com este número. "
				'end if
			'end if
			If trim(especificacao) = "" then 
				erro = erro & "Especificacao Invalida. "
			end if	
			if not isnumeric(quantidade) then
				erro = erro & "Quantidade inválida. "
			end if
			if not isnumeric(valorUnitario) then
				erro = erro & "Valor Unitário inválido. "
			end if
			if trim(unidade) = "" then
				erro = erro & "Unidade de medida inválida. "
			end if
			' Fim das Validações
									
			if erro = "" and request.form("ispostback") <> "" then
'response.write(ubound(Session("PropostaItens"),2) & "-" & idItem)
'response.end
				select case acaoItem
					case "add"
					
						if isarray(session("PropostaItens")) then
							
							PropostaItens = Session("PropostaItens")
							redim preserve PropostaItens(7,ubound(PropostaItens,2)+1)
							addPropostaItem
						else
							redim PropostaItens(7,0)
							addPropostaItem
						end if
						
						response.Redirect(redir)
					case "edit"
										
						editPropostaItem(idItem)
					 	response.Redirect(redir)
					
				end select
			
			end if
		elseif request("ispostback") = "" and acaoitem = "add" then
			
			if isarray(session("PropostaItens")) then
				idItemAnt = ubound(session("propostaItens"),2)
			else
				idItemant= -1
			end if
			
		end if
	
%>
	<form action="cadPropItem.asp" method="post" enctype="application/x-www-form-urlencoded" name="form1" target="_self">
		<input type="hidden" name="acaoitem" value="<%=acaoitem%>">
		<input type="hidden" name="redir" value="<%=redir%>">	
		<input type="hidden" name="idItemAnt" value="<%=idItemAnt%>">	
		<input type="hidden" name="isPostBack" value="1">
		<input type="hidden" name="idProcesso" value="<%=idProcesso%>">
		<input type="hidden" name="indiceItem" value="<%=indiceItem%>">
		<input type="hidden" name="idItem" value="<%=iif(acaoitem="add",idItemant+1,idItem)%>">
	 	
		  <label class="titulo">Cadastro de Itens de Proposta</label><br>
		  <label class="subtitulo">Informe os dados do Item da Proposta</label><br>
		  <br>
	 	
	  <label class="cabecalho">Núm. do Processo</label>
	  <br>
	  <input name="NumProcesso" type="text" class="caixa" id="NumProcesso" size="25" maxlength="20" readonly="true" value="<%=getNumProcesso(idProcesso)%>">
	  <br>
	  <label class="cabecalho">Núm. do Item</label>
	  <br>
	  <input name="NumItem" type="text" class="caixa" id="NumItem" size="25" maxlength="20" value="<%=NumItem%>">
	  <br>
      <label class="cabecalho">Especificação do Item</label>
	  <br>
	  <textarea cols="40" rows="5" class="caixa" name="especificacao" id="especificacao" onKeyDown="if(this.value.length>=1000){this.value=this.value.slice(0,1000); return false;}"><%=especificacao%></textarea>
	  <br>
	  <label class="cabecalho">Marca</label>
	  <br>
	  <input name="marca" type="text" class="caixa" id="marca" size="25" maxlength="50" value="<%=marca%>">
	   <br>
	  <label class="cabecalho">Quantidade</label>
	  <br>
	  <input name="Quantidade" type="text" class="caixa" id="Quantidade" size="13" maxlength="9" value="<%=Quantidade%>">
	  <br>
	  <label class="cabecalho">Unidade de medida</label>
	  <br>
	  <input name="Unidade" type="text" class="caixa" id="Unidade" size="15" maxlength="10" value="<%=Unidade%>">
	  <br>
	  <label class="cabecalho">Valor Unitário</label>
	  <br>
	  <input name="valorUnitario" type="text" class="caixa" id="valorUnitario" size="25" maxlength="20" value="<%=valorUnitario%>">
	  <br>
	  <br>
	  <input name="btnOK" type="submit" class="botao" id="btnOK" value="OK">
	  <input name="btnCancelar" type="button" class="botao" id="btnCancelar" value="Cancelar" onClick="window.location.href='<%=redir%>'">
	  <P class="erro">
		&nbsp;<%=erro%>
	  </P>
	</form>
<%end sub

sub addPropostaItem()

	PropostaItens(0,ubound(PropostaItens,2)) = idItem
	PropostaItens(1,ubound(PropostaItens,2)) = cStr(NumItem)
	PropostaItens(2,ubound(PropostaItens,2)) = especificacao
	PropostaItens(3,ubound(PropostaItens,2)) = ccur(valorUnitario)
	PropostaItens(4,ubound(PropostaItens,2)) = cSng(Quantidade)
	PropostaItens(5,ubound(PropostaItens,2)) = cCur(ValorUnitario*Quantidade)
	PropostaItens(6,ubound(PropostaItens,2)) = Unidade
	PropostaItens(7,ubound(PropostaItens,2)) = Marca
	session("propostaItens") = PropostaItens
end sub

sub EditPropostaItem(indiceItem)

	PropostaItens = session("propostaItens")



	PropostaItens(0,indiceItem) = idItem
	PropostaItens(1,indiceItem) = cstr(NumItem)		
	PropostaItens(2,indiceItem) = especificacao
	PropostaItens(3,indiceItem) = ccur(valorUnitario)
	PropostaItens(4,indiceItem) = cSng(Quantidade)
	PropostaItens(5,indiceItem) = ccur(ValorUnitario*Quantidade)
	PropostaItens(6,indiceItem) = Unidade
	PropostaItens(7,indiceItem) = Marca
	session("propostaItens") = PropostaItens

end sub


sub CarregaLinha()
	PropostaItens = session("PropostaItens")
	idItem = request("idItem")
	idItemAnt = request("idItem")
	
		
	'for x = 0 to ubound(PropostaItens,2)
	'	if PropostaItens(0,x) = cint(idItem) then
			'indiceItem = x
			'exit for
		'else
		'	indiceItem = -1
		'end if
		
	'next
	
	numItem =  PropostaItens(1,idItem)
	especificacao = PropostaItens(2,idItem)
	valorUnitario = PropostaItens(3,idItem)
	quantidade = PropostaItens(4,idItem)
	unidade = PropostaItens(6,idItem)
	marca = PropostaItens(7,idItem)
end sub




%>

<!--#include file="../layout/layout.asp"-->
<%call FechaBanco%>
