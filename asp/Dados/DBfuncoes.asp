<%
' #########################################################################################################
' SisLic - Biblioteca de Funções de Banco de Dados
' Data de Criação: 09/04/2006 
' Criador: Raphael Carubbi Neto/ José Carlos da Silva Borges	
' #########################################################################################################

'Rotina para esabelecer conexão com o Banco de Dados através do Objeto Conn
Dim Conn

Sub AbreBanco

	Set Conn = Server.CreateObject("ADODB.Connection")
	'Conn.open "Provider=SQLOLEDB.1;Password=Nota10;Persist Security Info=True;User ID=sa;Initial Catalog=SisLic;Data Source=MICRO_3\SQLEXPRESS"
	conn.open "Provider=SQLNCLI;Server=MICRO7\SQLEXPRESS;Database=SisLic;UID=sa;PWD=!2C9|2R2#SL1;" 

End sub

' Rotina para fechar a conexão com o banco de dados
Sub FechaBanco

	Conn.close : Set Conn = Nothing
	
End Sub

sub ComboDB(nomeCombo, ItemTexto, ItemValor, ComandoSQL, ValorSel, texto, tamanho)
	dim rsItens
	set rsItens = conn.execute(ComandoSQL)
	if ValorSel = "" or not isNumeric(valorSel) then valorSel = 0
	if tamanho = "" then tamanho = 130
	%>
	
		
	<select name="<%=nomeCombo%>" class="caixa" style="width:<%=tamanho%>px;">
	
	<option value="0"><%=texto%></option>
	<%while not rsItens.eof
	%>
		
		<option value="<%=rsItens(ItemValor)%>" <%if cdbl(ValorSel) = cdbl(rsItens(ItemValor)) then response.write "selected"%>><%=rsItens(ItemTexto)%></option>
	<%rsItens.movenext
	wend%>
	</select> 
	<%
end sub

function verifica_existe(valor, campo, tabela)
	dim retorno, sql, rs, tipodado, aux	
		
	sql = "Select " & campo & " from " & tabela
	set rs = conn.execute(sql)
	tipodado = rs.fields(0).type
	rs.close : set rs = nothing
	select case tipoDado
		case 200 , 201 , 202 , 203 'string
			aux = "'" & valor & "'"

		case 133 , 134 , 135 'data hora
			aux  = "'" & day(valor) & "/" & month(valor) & "/" & year(valor) & " " & hour(valor) & ":" & minute(valor) & ":" & second(valor) & "'"
		case else 
			aux  = valor	
	end select
	
	sql = "Select " & campo & " from " & tabela & " Where " & campo & " = " & aux
	
	set rs = conn.execute(sql)
	if not rs.eof then
		retorno = true
	else
		retorno = false
	end if
	rs.close : set rs = nothing
	
	verifica_existe = retorno
end function

function maxID(tabela, campoID)
	dim sql, id, rs
	
	sql = "SELECT max(" & campoID & ") as mid from " & tabela
	set rs = conn.execute(SQL)

	if isNull(rs("mid")) then
		maxID = 0
	else
		maxID = rs("mid")
	end if
	rs.close
	set rs = nothing
end function

function getNumProcesso(idProcesso)
	dim retorno, sql, rs
	sql = "Select numProcesso from edital where idProcesso = " & idProcesso
	set rs = conn.execute(sql)
	if not rs.eof then
		retorno = rs("numProcesso")
	else
		retorno = 0
	end if
	rs.close : set rs = nothing
	getNumProcesso = retorno	
end function

function getValor(id, campoId, CampoRetorno, tabela)
	dim retorno, sql, rs
	
	sql = "Select " & CampoRetorno & " from " & tabela & " where " 
	if typename(id) = "String" then
		sql = sql & campoId & " = '" & id & "'" 
	else
		sql = sql & campoId & " = " & id
	end if
	
	set rs = conn.execute(Sql)
	if not rs.eof then
		retorno = rs(CampoRetorno)
	else
		retorno = ""
	end if
	rs.close : set rs = nothing
	getValor = retorno
	

end function


%>