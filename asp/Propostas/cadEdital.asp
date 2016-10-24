<!--#include file="../Dados/DBfuncoes.asp"-->
<!--#include file="../funcoes/funcoes.asp"-->
<%
call abreBanco

dim acao, idProcesso, erro

acao = request("acao")
idProcesso = request("idProcesso")
redir = request("redir")

sub Impressao
end sub

sub auxiliar
end sub


sub principal
	Dim NumProcesso, NumEdital, Descricao, AberturaData, AberturaHora, EncerramentoData, EncerramentoHora, Situacao, Participou, IdParceiro, _ 
		IdEditalTipo, DataCadastro, contcad
	
	contcad=request.QueryString("contcad")
	
	
	
	
	if request.form("ispostback") = "" and contcad="1" then 
	' caso venha da pagina de cadastro de usuarios recuperar os campos que ja tinham sido preenchidos
		IdProcesso=session("IdProcesso")  
		NumProcesso=session("NumProcesso") 
		NumEdital=session("NumEdital")
		Descricao=session("Descricao") 
		AberturaData=session("AberturaData") 
		AberturaHora=session("AberturaHora") 
		IdEditalTipo =session("IdEditalTipo") 
		session("IdProcesso") = ""
		session("NumProcesso") = ""
		session("NumEdital") = ""
		session("Descricao") = ""
		session("AberturaData") = ""
		session("AberturaHora") = ""
		session("IdEditalTipo") = ""
	
	else
		'caso a primeira vez que passa e ação igual a editar
		if acao = "edit" and request.Form("isPostBack") = "" then
			SQL = " Select NumProcesso, NumEdital, Descricao, Abertura, Encerramento, Situacao, Participou, IdParceiro, " & _ 
				  " IdEditalTipo, DataCadastro from edital where idProcesso=" & idProcesso
			
			set rs = conn.execute(SQL)
			
			if not rs.eof then
				NumProcessoAnt = rs("NumProcesso")
				NumProcesso = rs("NumProcesso")
				NumEdital = rs("NumEdital")
				Descricao = rs("Descricao")
				Abertura = rs("Abertura")
				AberturaData = formatdatetime(Abertura,2)
				AberturaHora = formatdatetime(Abertura,3)
				Encerramento = rs("Encerramento")
				Situacao = rs("Situacao")
				Participou = rs("Participou")
				IdParceiro = rs("IdParceiro")
				IdEditalTipo = rs("IdEditalTipo")
				DataCadastro = rs("DataCadastro")
			end if
			rs.close
			set rs = nothing
		elseif request.Form("isPostBack") <> "" then ' se é postBack recuperar dados do formulario
				IdProcesso = request.form("IdProcesso")
				NumProcesso = request.form("NumProcesso")
				NumEdital = request.form("NumEdital")
				Descricao = request.form("Descricao")
				AberturaData = recebe_datador("AberturaData",0)
				AberturaHora = recebe_datador("AberturaHora",1)
				IdParceiro = request.form("IdParceiro")
				IdEditalTipo = request.form("IdEditalTipo")
				NumProcessoAnt = request.form("NumProcessoAnt")
				if request.form("cmdNovoCLiente") <> "" then ' se vai pra tela de novo cliente salvar o que ja foi preenchido até então.
					session("IdProcesso") = IdProcesso
					session("NumProcesso") = NumProcesso
					session("NumEdital") = NumEdital
					session("Descricao") = Descricao
					session("AberturaData") = AberturaData
					session("AberturaHora") = AberturaHora
					session("IdEditalTipo") = IdEditalTipo
					response.redirect("cadCliente.asp?redir="&server.urlencode("CadEdital.asp?contcad=1&acao="&acao&"&idProcesso="&idProcesso&"&redir="&redir)&"&tipo=0&acao=add")
				end if
			

			' Validações
			
			If trim(numProcesso) = "" then 
				erro = "Número do Processo Inválido. "
			elseif verifica_existe(NumProcesso,"NumProcesso","Edital") and acao="add" then
				erro = "Já existe outro processo com este número. "
			elseif acao = "edit" and NumProcesso <> NumProcessoAnt then
				if verifica_existe(NumProcesso, "NumProcesso","Edital") then
					erro = erro &  "Já existe outro processo com este número. "
				end if	
			end if
			If trim(NumEdital) = "" then erro = erro & "Número do Edital Invalido. "
			if ideditaltipo =  0 then erro = erro & "Selecione um tipo de edital. "
			if idparceiro = 0 then erro = erro & "Selecione um Cliente. "
			if AberturaData <> "" then
				If not isdate(AberturaData) then erro = erro & "Data de Abertura Invalida. "
			end if
			if AberturaHora <> "" then
				If not isdate(AberturaHora) then erro = erro & "Hora de Abertura Invalida. "
			end if
		end if
		' Fim das Validações
							
		if erro = "" and request.form("ispostback") <> "" then
			select case acao
				case "add"
				
					SQL = "INSERT INTO Edital " & _
							"(NumProcesso,  " & _
							"NumEdital, " & _
							"Abertura,  " & _
							"Situacao, " & _
						 	"Descricao,  " & _
							"IdParceiro, " & _
							"IdEditalTipo, " & _
							"DataCadastro) " & _
							"VALUES (" & _
							"'" & NumProcesso & "', " & _
							"'" & NumEdital & "', " & _
							formata_data_sql(AberturaData & " " & AberturaHora) & ", " & _
							"0,'" & _
							Descricao & "', " & _
							IdParceiro & ", " & _
							IdEditalTipo & ", " & _
							formata_data_sql(now) & ")"
					
					
						
					conn.execute(SQL)
					response.Redirect(redir)
			case "edit"
			
		
				SQL = "UPDATE edital SET "
				sql = sql & " NumProcesso = '" & NumProcesso & "', "
				sql = sql & " NumEdital = '" & NumEdital & "', " & _
				" Abertura = " & formata_data_sql(AberturaData & " " & AberturaHora) & ", " & _
				" IdParceiro = " & IdParceiro & ", " & _
				" Descricao = '" & Descricao & "', " & _
				" IdEditalTipo = " & IdEditalTipo & _
				" WHERE IdProcesso=" & idProcesso
				' response.write sql
				' response.End()
				conn.execute(SQL)
				 
					response.Redirect(redir)
			end select
			
			
		end if
	
	end if
%>
	<form action="" method="post" enctype="application/x-www-form-urlencoded" name="form1" target="_self">
		<input type="hidden" name="acao" value="<%=acao%>">
		<input type="hidden" name="NumProcessoAnt" value="<%=NumProcessoAnt%>">
		<!--<input type="hidden" name="NumProcesso" value="<%=NumProcesso%>">-->
		<input type="hidden" name="redir" value="<%=redir%>">		
		<input type="hidden" name="isPostBack" value="1">
		<input type="hidden" name="idProcesso" value="<%=idProcesso%>">
	 
		  <label class="titulo">Cadastro de Edital</label><br>
		  <label class="subtitulo">Informe os dados do edital</label><br>
		  <br>
	 	
	  <label class="cabecalho">Núm. do Processo</label>
	  <br>
	  <input name="NumProcesso" type="text" class="caixa" id="NumProcesso" size="25" maxlength="20" value="<%=NumProcesso%>">
	  <br>
	  <label class="cabecalho">Núm. do Edital</label>
	  <br>
	  <input name="NumEdital" type="text" class="caixa" id="NumEdital" size="25" maxlength="20" value="<%=NumEdital%>">
	  <br>
      <label class="cabecalho">Tipo do Edital</label>
	  <br>
	   <%ComboDB "IdEditalTipo","Descricao","id","Select id, descricao from EditalTipo",IdEditalTipo,"",""%>
	  <br>
	  <label class="cabecalho">Cliente</label>
	  <br>
	  <%ComboDB "IdParceiro","nome","id","Select id, replace(nome,char(13),'. ') as nome from parceiro where tipo = 0 and ativo = 1",IdParceiro,"","500" %>
	  &nbsp; &nbsp;<input type="submit" class="botao" value="Novo Cliente" name="CmdNovoCLiente" id="CmdNovoCLiente">
	  <br>
	  <label class="cabecalho">Descrição do Edital</label>
	  <br>
	  <input name="Descricao" type="text" class="caixa" id="Descricao" size="50" maxlength="255" value="<%=Descricao%>">
	  <br>
	  <label class="cabecalho">Data da Abertura</label>
	  <br>
	  <%Datador "AberturaData",0,year(date),year(date)+5,AberturaData%>
	  <br>
	  <label class="cabecalho">Hora da Abertura</label>
	  <br>
	  <%Datador "AberturaHora",1,0,0,AberturaHora%>
	  <br>
      <br>

	  <input name="btnOK" type="submit" class="botao" id="btnOK" value="OK">
	  <input name="btnCancelar" type="button" class="botao" id="btnCancelar" value="Cancelar" onClick="window.location.href='<%=redir%>'">
	  <P class="erro">
		&nbsp;<%=erro%>
	  </P>
	</form>
<%end sub

'function validaNumProcesso(numProcesso)

'	if (instr(numProcesso,"/") <> len(numProcesso) - 4) then
	'	validaNumProcesso = false
	'else
	'	 if not isnumeric(left(numProcesso,instr(numProcesso,"/")-1)) or not isnumeric(right(numprocesso,4)) then
	'	 	validaNumProcesso = false
	'	 else
		'	validaNumProcesso = true
		' end if
	'end if
	
'end function
%>

<!--#include file="../layout/layout.asp"-->
<%call FechaBanco%>
