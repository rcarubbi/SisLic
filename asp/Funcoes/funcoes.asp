<%
' #########################################################################################################
' SisLic - Biblioteca de Funções
' Data de Criação: 09/04/2006 
' Criador: Raphael Carubbi Neto	
' #########################################################################################################

'###################################
'validação de usuário
'as paginas login e logar não precisao de login
if session("usuario_cpf") = "" then
	if inStr(lcase(request.ServerVariables("url")), "login.asp") = 0 and _
		inStr(lcase(request.ServerVariables("url")), "logar.asp") = 0 then
		
		response.Redirect("/sislic/login.asp?redir=" & request.ServerVariables("URL"))
	end if
end if


Function FormataCEP(cep)

	cepleft = left(cep,5)
	
	cepright = right(cep,3)
	
	cepfinal = cepleft & "-" & cepright
	
	FormataCEP = cepfinal

End Function

'##########################################################################################################
'função para carregar o modulo desejado passando o id do módulo
'id:
'1 - Propostas
'2 - Comércio
'3 - Financeiro
'4 - Operadores
sub ModuloAtual(m)

	select case m
	case 1
	case 2
	case 3
	case 4
		session("modulo_id") = 4
		session("modulo_nome") = "Usuarios"
		session("modulo_menu") = "mnu_usuario.asp"		
	end select

end sub
'##########################################################################################################

' Funcao IIF codicional
function IIF(condicao,retornoTrue, retornoFalse)

	if condicao = true then
		IIF = retornoTrue
	else
		IIF = retornoFalse
	end if

end function

'nomeCombo = nome da combo para retornar o valor atraves do request
'itemTexto = Matriz com os textos a serem exibidos para o usuario
'ItemValor = Matriz com os valores a serem guardadas na combo
'a posição X da matriz itemTexto, corresponde a mesma posição na matriz ItemValor
sub Combo(nomeCombo, ItemTexto, ItemValor, ValorSel)
	dim i, sel
	
	
	
	%>
	
	<select name="<%=nomeCombo%>" class="caixa">
	<%
	for i = 0 to ubound(itemValor)
	on error resume next
	response.write cInt(ValorSel) 
	%>
		<option value="<%=ItemValor(i)%>" <%if cInt(ValorSel) = ItemValor(i) AND valorSel <> "" then response.write "selected"%>><%=ItemTexto(i)%></option>	
	<%
	next
	%>

	</select> 
	<%
end sub


Function FormataCPF(cpf)
	if not isnull(cpf) then
		vtam = len(cpf)
		if vtam = 11 then
			 cpf1= left(cpf,3)
			 cpf2 = mid(cpf,4,3)
			 cpf3 = mid(cpf,7,3)
			 cpf4 = right(cpf,2)
			 cpffinal = cpf1 & "." & cpf2 & "." & cpf3 & "-" & cpf4
		elseif vtam = 10 then
			 cpf1= left(cpf,2)
			 cpf2 = mid(cpf,3,3)
			 cpf3 = mid(cpf,6,3)
			 cpf4 = right(cpf,2)
			 cpffinal = cpf1 & "." & cpf2 & "." & cpf3 & "-" & cpf4
		elseif vtam = 9 then
			 cpf1= left(cpf,1)
			 cpf2 = mid(cpf,2,3)
			 cpf3 = mid(cpf,5,3)
			 cpf4 = right(cpf,2)
			 cpffinal = cpf1 & "." & cpf2 & "." & cpf3 & "-" & cpf4
		end if
		FormataCPF = cpffinal
	else
		FormataCPF = "N/D"
	end if
End Function

function valida_cpf(cpf)
	dim retorno, dv, Verif_dv, soma
	dv = cint(right(cpf,2))
	if len(cpf) = 11 then
	for x=1 to 9 
		
		soma = soma + (cint(mid(cpf,x,1))*x)
	next
	else
		retorno = false
		exit function
	end if
	Verif_dv = cstr(soma mod 11 mod 10)
	soma = 0
	for x=0 to 9
		soma = soma + (cint(mid(cpf,x+1,1))*x)
	next
	Verif_dv = Verif_dv & cstr(soma mod 11 mod 10)
	if cint(verif_dv) = dv then
		retorno = true
	else
		retorno = false
	end if
	valida_cpf = retorno
end function

Function Valida_Cnpj(Num_Cnpj)
      Dim dig1, dig2 
      ' Variáveis que irão receber os dígitos calculados
      Dim p_dig1, p_dig2 
      'Variaveis usadas somente p/ auxiliar no
      'calculo do dígito
      Dim str_partes, dig_fim 
      ' Vão receber respectivamente os nove
      'caracteres iniciais e os dois ultimos
      Dim parte1, parte2, parte3, parte4 
      'Recebem os números do cpf sem os
      'pontos
      Dim soma_dig1, soma_dig2, cont 
      'Recebem o valor da soma e
      'multiplicação dos números
      'Inicia a eliminação dos Pontos e traços do número

	  if len(Num_Cnpj) < 14 then
	  		Valida_Cnpj = false
			exit function
	  end if	

      str_partes = left(Num_Cnpj, len(Num_Cnpj) - 2) 'parte1 + parte2 + parte3 + parte4
      dig_fim = Right(Num_Cnpj, 2)

      soma_dig1 = 0
      soma_dig2 = 0
      For cont = 1 To Len(str_partes)
            soma_dig1 = soma_dig1 + cInt(Mid(str_partes, cont, 1)) * IIf(cont < 5, 6 - cont, 14 - cont)
            soma_dig2 = soma_dig2 + cInt(Mid(str_partes, cont, 1)) * IIf(cont < 6, 7 - cont, 15 - cont)
      Next 
      p_dig1 = soma_dig1 Mod 11
      dig1 = IIf(p_dig1 = 0 Or p_dig1 = 1, 0, 11 - p_dig1)
      p_dig2 = soma_dig2 + dig1 * 2
      p_dig2 = p_dig2 Mod 11
      dig2 = IIf(p_dig2 = 0 Or p_dig2 = 1, 0, 11 - p_dig2)
      If dig_fim <> dig1 & dig2 Then
            Valida_Cnpj= False 'Retorna false se o número não for válido
      Else
            Valida_Cnpj= True 'Retorna true se o número for váldo
      End If
End Function

function formata_data_sql(valor)
	
	formata_data_sql = "'" & day(valor) & "/" & month(valor) & "/" & year(valor) & " " & hour(valor) & ":" & minute(valor) & ":" & second(valor) & "'"

end function

function Formata_Data(valor)
	if isdate(valor) then
		mes = "jan,fev,mar,abr,mai,jun,jul,ago,set,out,nov,dez"
		mes = split(mes,",")
		Formata_Data = day(valor) & " " & mes(month(valor)-1) & " " & year(valor) & " às " & hour(valor) & ":" & minute(valor) '& ":" & second(valor)
	else
		Formata_data = ""
	end if
end function

function Formata_Data_abrev(valor)
	if isdate(valor) then
		mes = "jan,fev,mar,abr,mai,jun,jul,ago,set,out,nov,dez"
		mes = split(mes,",")
		Formata_Data_abrev = day(valor) & " " & mes(month(valor)-1) & " " & year(valor) 
	else
		Formata_Data_abrev = ""
	end if
end function

function recebe_datador(nomeDatador,tipoDatador)
	dim vl1, vl2, vl3, retorno
	if tipoDatador = 0 then

	 	vl1 = request.form("dia_"&nomeDatador)
		vl2 = request.form("mes_"&nomeDatador)
		vl3 = request.form("ano_"&nomeDatador)
		retorno = vl1&"/"&vl2&"/"&vl3
	else
		vl1 = request.form("hora_"&nomeDatador)
		vl2 = request.form("minuto_"&nomeDatador)
		vl3 = request.form("segundo_"&nomeDatador)
		retorno = vl1&":"&vl2&":"&vl3
	end if
	recebe_datador = retorno
end function

sub Datador(nomeDatador,tipoDatador,AnoInic, AnoFinal,valorAtual)
	' tipoDatador --> 0 = Data; 1= Hora
	dim dia, mes, ano, mesnome, hora, minuto, segundo
	mesnome = "Jan,Fev,Mar,Abr,Mai,Jun,Jul,Ago,Set,Out,Nov,Dez"
	mesnome = split(mesnome,",")
	if valorAtual = "" or not isdate(valorAtual) then valoratual = now

	
	if tipoDatador = 0 then%>
		<Select name="dia_<%=nomeDatador%>" class="caixa" id="dia_<%=nomeDatador%>">
			<%for dia = 1 to 31 %>		
				<option value="<%=dia%>" <%if day(valorAtual)=dia then response.write "Selected"%>><%=dia%></option>
			<%Next%>
		</Select>
		<label class="cabecalho">/</label>
		<Select name="mes_<%=nomeDatador%>" class="caixa" id="mes_<%=nomeDatador%>">
			<%for mes = 1 to 12 %>		
				<option value="<%=mes%>" <%if month(valorAtual)=mes then response.write "Selected"%>><%=mesnome(mes-1)%></option>
			<%Next%>
		</Select>
		<label class="cabecalho">/</label>
		<Select name="ano_<%=nomeDatador%>" class="caixa" id="ano_<%=nomeDatador%>">
			<%for ano = AnoInic to AnoFinal %>		
				<option value="<%=ano%>" <%if year(valorAtual)=ano then response.write "Selected"%>><%=ano%></option>
			<%Next%>
		</Select>
		
<%	else %>
	<Select name="hora_<%=nomeDatador%>" class="caixa" id="hora_<%=nomeDatador%>">
			<%for hora = 0 to 23 %>		
				<option value="<%=hora%>" <%if Hour(valorAtual)=hora then response.write "Selected"%>><%=hora%></option>
			<%Next%>
		</Select>
		<label class="cabecalho">:</label>
		<Select name="minuto_<%=nomeDatador%>" class="caixa" id="minuto_<%=nomeDatador%>">
			<%for minuto = 0 to 59 %>		
				<option value="<%=minuto%>" <%if minute(valorAtual)=minuto then response.write "Selected"%>><%=minuto%></option>
			<%Next%>
		</Select>
		<label class="cabecalho">:</label>
		<Select name="segundo_<%=nomeDatador%>" class="caixa" id="segundo_<%=nomeDatador%>">
			<%for segundo = 0 to 59 %>		
				<option value="<%=segundo%>" <%if second(valorAtual)=segundo then response.write "Selected"%>><%=segundo%></option>
			<%Next%>
		</Select>
	
<%	end if



end sub

'Function URLDecode(sConvert)
'    Dim aSplit
'    Dim sOutput
'    Dim I
'    If IsNull(sConvert) Then
'       URLDecode = ""
'       Exit Function
'    End If
	
    ' convert all pluses to spaces
'    sOutput = REPLACE(sConvert, "+", " ")
	
    ' next convert %hexdigits to the character
'    aSplit = Split(sOutput, "%")
	
'    If IsArray(aSplit) Then
'		if ubound(aSplit) > -1 then
'			sOutput = aSplit(0)
'			For I = 0 to UBound(aSplit) - 1
'			sOutput = sOutput & _
'			  Chr("&H" & Left(aSplit(i + 1), 2)) &_
'			  Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
'			Next
'		end if
'    End If
	
'    URLDecode = sOutput
'End Function

function TemRestricoes(itemMenu)

	dim sql, rs, retorno

	sql= "SELECT id_menuitem FROM restricoes where id_restricoesGrupo=" & session("usuario_grupo") & " AND id_menuitem = " & itemMenu

	set rs = conn.execute(sql)
	
	if rs.eof then
		retorno = false
	else
		retorno = true
	end if
	
	rs.close : set rs = nothing
	
	TemRestricoes = retorno
end function

function formata_campo(valor)
	dim formato, tipo
	err.number = 0
	
	on error resume next

	
	tipo = typename(valor)
	
	if err.number > 0 then 'se o tipo não foi reconhecido
		tipo = ""
	end if
	
	select case tipo
		case "Integer" 'int
			formato = valor 
		case "String" 'string
			formato = valor
		case "Date"
			mes = "jan,fev,mar,abr,mai,jun,jul,ago,set,out,nov,dez"
			mes = split(mes,",")
			formato = day(valor) & " " & mes(month(valor)-1) & " " & year(valor) 
			if hour(valor) <> 0 or minute(valor) <> 0 then
				formato = formato & " às " & hour(valor) & ":" & string(2-len(minute(valor)), "0") & minute(valor)
			end if
			
		case "Boolean" 'boolean
			if valor = true or valor = "1" or valor="-1"	then 
				formato = "Sim"
			else
				formato = "Não"
			end if
		case "Double"
			formato = formatnumber(valor,2)
		case "Currency"
			formato = "R$ " & formatnumber(valor,2)
		case "Null"
			formato = "-"
		case else
			formato = valor 
	end select
	formata_campo = formato
end function

function verifica_existe_matriz(valor,indCampo,matriz)
	dim	retorno , x
	
	if isarray(matriz) then
		for x=0 to ubound(matriz,2)
			if matriz(indcampo,x) = valor then 
				retorno = true
				exit for
			else
				retorno = false
			end if
		next
	else
		retorno = false
	end if
	verifica_existe_matriz = retorno
	
end function



'###########################################################
'função para exibir numero por extenso
'---------------->
'dim extUnidade(9), dezena(17), centena(9)

'Public Function unidades(numero)
'FAZ PARTE DA FUNCAO QUE ESCREVE UM NUMERO POR EXTENSO
'   unidades = extUnidade(numero)
   
'End Function

'Public Function dezenas(numero)
'FAZ PARTE DA FUNCAO QUE ESCREVE UM NUMERO POR EXTENSO
'   If numero > 9 And numero < 21 Then
'      dezenas = dezena(numero - 10)
'   ElseIf numero >= 21 And numero <= 99 Then
'      If Right(numero, 1) = "0" Then
'         dezenas = dezena(Int(Left(numero, 1) + 8))
'      Else
'         dezenas = dezena(Int(1 & Int(Left(numero, 1) - 2))) & " e " & extUnidade(Int(Right(numero, 1)))
'      End If
'   End If
'End Function

'Public Function centenas(numero)
'FAZ PARTE DA FUNCAO QUE ESCREVE UM NUMERO POR EXTENSO
'   If Mid(numero, 2, 2) = 0 Then
'      If numero = 100 Then
'         centenas = "cem"
'      Else
'         centenas = centena(Left(numero, 1))
'      End If
'   ElseIf Mid(numero, 3, 1) = 0 Then
'      centenas = centena(Left(numero, 1)) & " e " & dezenas(Mid(numero, 2, 2))
'   ElseIf Mid(numero, 2, 1) = 0 Then
'      centenas = centena(Int(Left(numero, 1))) & " e " & extUnidade(Int(Right(numero, 1)))
'   Else
'      centenas = centena(Int(Left(numero, 1))) & " e " & dezenas(Int(Right(numero, 2)))
'   End If
'End Function

'Function ext(numero)
'FUNCAO QUE ESCREVE UM NUMERO POR EXTENSO

'   If numero > 999999.99 Then ext = "Número fora dos padrões válidos !": Exit Function
'      extUnidade(0) = "zero"
'      extUnidade(1) = "um"
'      extUnidade(2) = "dois"
'      extUnidade(3) = "três"
'      extUnidade(4) = "quatro"
'      extUnidade(5) = "cinco"
'      extUnidade(6) = "seis"
'      extUnidade(7) = "sete"
'      extUnidade(8) = "oito"
'      extUnidade(9) = "nove"
'      dezena(0) = "dez"
'      dezena(1) = "onze"
'      dezena(2) = "doze"
'      dezena(3) = "treze"
'      dezena(4) = "quatorze"
'      dezena(5) = "quinze"
'      dezena(6) = "dezesseis"
'      dezena(7) = "dezessete"
'      dezena(8) = "dezoito"
'      dezena(9) = "dezenove"
'      dezena(10) = "vinte"
'      dezena(11) = "trinta"
'      dezena(12) = "quarenta"
'      dezena(13) = "cinquenta"
'      dezena(14) = "sessenta"
'      dezena(15) = "setenta"
'      dezena(16) = "oitenta"
'      dezena(17) = "noventa"
'      centena(1) = "cento"
'      centena(2) = "duzentos"
'      centena(3) = "trezentos"
'      centena(4) = "quatrocentos"
'      centena(5) = "quinhentos"
'      centena(6) = "seiscentos"
'      centena(7) = "setecentos"
'      centena(8) = "oitocentos"
'      centena(9) = "novecentos"
'      inteiro = Int(numero)
'      'aqui eu vou verificar se o inteiro é zero e então não vou tratar os inteiros só os centavos
'      If inteiro <> 0 Then
'        tamanho = Len(inteiro)
'      Else
'        tamanho = 0
'      End If
'      Select Case tamanho
'         Case 1
'            ext = unidades(inteiro)
'         Case 2
'            ext = dezenas(inteiro)
'         Case 3
'            ext = centenas(inteiro)
'         Case 4
'            If Right(inteiro, 3) = 0 Then
'               ext = unidades(Left(inteiro, 1)) & " mil"
'            Else
'               If Int(Right(inteiro, 3)) > 99 Then
'                  ext = unidades(Left(inteiro, 1)) & " mil e " & centenas(Int(Right(inteiro, 3)))
'               ElseIf Int(Right(inteiro, 3)) > 9 And Int(Right(inteiro, 3)) < 100 Then
'                  ext = unidades(Left(inteiro, 1)) & " mil e " & dezenas(Int(Right(inteiro, 3)))
'               ElseIf Int(Right(inteiro, 3)) < 10 Then
'                  ext = unidades(Left(inteiro, 1)) & " mil e " & unidades(Int(Right(inteiro, 3)))
'               End If
'            End If
'         Case 5
'            If Right(inteiro, 3) = 0 Then
'               ext = dezenas(Left(inteiro, 2)) & " mil "
'            Else
'               If Int(Right(inteiro, 3)) > 99 Then
'                  ext = dezenas(Left(inteiro, 2)) & " mil e " & centenas(Right(inteiro, 3))
'               ElseIf Int(Right(inteiro, 3)) > 9 And Int(Right(inteiro, 3)) < 100 Then
'                  ext = dezenas(Left(inteiro, 2)) & " mil e " & dezenas(Int(Right(inteiro, 3)))
'               ElseIf Int(Right(inteiro, 3)) < 10 Then
'                  ext = dezenas(Left(inteiro, 2)) & " mil e " & unidades(Int(Right(inteiro, 3)))
'               End If
'            End If
'         Case 6
'            If Right(inteiro, 3) = 0 Then
'               ext = centenas(Left(inteiro, 3)) & " mil "
'            ElseIf Int(Right(inteiro, 3)) > 99 Then
'               ext = centenas(Left(inteiro, 3)) & " mil e " & centenas(Int(Right(inteiro, 3)))
'            ElseIf Int(Right(inteiro, 3)) > 9 And Int(Right(inteiro, 3)) < 100 Then
'               ext = centenas(Left(inteiro, 3)) & " mil e " & dezenas(Int(Right(inteiro, 3)))
'            ElseIf Int(Right(inteiro, 3)) < 10 Then
'               ext = centenas(Left(inteiro, 3)) & " mil e " & unidades(Int(Right(inteiro, 3)))
'            End If
'      End Select
'      If ext <> "" Then
'        If ext = "um" Then
'          ext = ext & " real"
'        Else
'          ext = ext & " reais"
'        End If
'      End If
'      If numero - Int(numero) <> 0 Then
'         Dim fra
'         fra = Right(numero, 2)
'         'verificando se o inteiro é zero
'         If Int(numero) <> 0 Then
'            If InStr(1, fra, ",") <> 0 Then fra = Right(fra, 1) * 10
'            If fra >= 10 Then
'               ext = ext & " e " & dezenas(fra) & " centavos"
'            ElseIf fra < 10 Then
'               ext = ext & " e " & unidades(fra) & " centavos"
'            End If
'         Else
'            'aqui eu tiro o "e", pq só é para os centavos
'            If InStr(1, fra, ",") <> 0 Then fra = Right(fra, 1) * 10
'            If fra >= 10 Then
'               ext = ext & dezenas(fra) & " centavos"
'            ElseIf fra < 10 Then
'               ext = ext & unidades(fra) & " centavos"
'            End If
'         End If
'      End If
'End Function
'<---------------------------------------------------------
'fim função extenso
'############################################################

function formata_corSN(valor) 
	dim retorno
	
		if valor = "Não" then
			retorno = "<font color=red>"&valor&"</font>"
		else
			retorno =  "<font color=blue>"&valor&"</font>"
		end if
		
		formata_corSN = retorno
end function

function nulo(valor)

	if valor = "" or isnull(valor) or isempty(valor) then
		nulo = 0
	else
		nulo = valor
	end if
end function


Function FormataCNPJ(cnpj)
	if not isnull(cnpj) then
		vtam = len(cnpj)
		if vtam = 14 then
			 cnpj1 = left(cnpj,2)
			 cnpj2 = mid(cnpj,3,3)
			 cnpj3 = mid(cnpj,6,3)
			 cnpj4 = mid(cnpj,9,4)
			 cnpj5 = right(cnpj,2)
			 cnpjfinal = cnpj1 & "." & cnpj2 & "." & cnpj3 & "/" & cnpj4 & "-" & cnpj5
		elseif vtam = 13 then
			 cnpj1 = left(cnpj,1)
			 cnpj2 = mid(cnpj,2,3)
			 cnpj3 = mid(cnpj,5,3)
			 cnpj4 = mid(cnpj,8,4)
			 cnpj5 = right(cnpj,2)
			 cnpjfinal = cnpj1 & "." & cnpj2 & "." & cnpj3 & "/" & cnpj4 & "-" & cnpj5
		end if
		FormataCNPJ = cnpjfinal
	else
		FormataCNPJ = "N/D"
	end if
End Function






%>