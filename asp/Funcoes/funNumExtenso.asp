<%

Dim x_Centavos, x_I, x_J, x_K, x_Numero, x_QtdCentenas, x_TotCentenas, x_TxtExtenso( 900 ), x_TxtMoeda( 6 ), x_ValCentena( 6 ), x_Valor, x_ValSoma

' Matrizes de textos
x_TxtMoeda( 1 ) = "rea"
x_TxtMoeda( 2 ) = "mil"
x_TxtMoeda( 3 ) = "milh"
x_TxtMoeda( 4 ) = "bilh"
x_TxtMoeda( 5 ) = "trilh"

x_TxtExtenso( 1 ) = "um"
x_TxtExtenso( 2 ) = "dois"
x_TxtExtenso( 3 ) = "tres"
x_TxtExtenso( 4 ) = "quatro"
x_TxtExtenso( 5 ) = "cinco"
x_TxtExtenso( 6 ) = "seis"
x_TxtExtenso( 7 ) = "sete"
x_TxtExtenso( 8 ) = "oito"
x_TxtExtenso( 9 ) = "nove"
x_TxtExtenso( 10 ) = "dez"
x_TxtExtenso( 11 ) = "onze"
x_TxtExtenso( 12 ) = "doze"
x_TxtExtenso( 13 ) = "treze"
x_TxtExtenso( 14 ) = "quatorze"
x_TxtExtenso( 15 ) = "quinze"
x_TxtExtenso( 16 ) = "dezesseis"
x_TxtExtenso( 17 ) = "dezessete"
x_TxtExtenso( 18 ) = "dezoito"
x_TxtExtenso( 19 ) = "dezenove"
x_TxtExtenso( 20 ) = "vinte"
x_TxtExtenso( 30 ) = "trinta"
x_TxtExtenso( 40 ) = "quarenta"
x_TxtExtenso( 50 ) = "cinquenta"
x_TxtExtenso( 60 ) = "sessenta"
x_TxtExtenso( 70 ) = "setenta"
x_TxtExtenso( 80 ) = "oitenta"
x_TxtExtenso( 90 ) = "noventa"
x_TxtExtenso( 100 ) = "cento"
x_TxtExtenso( 200 ) = "duzentos"
x_TxtExtenso( 300 ) = "trezentos"
x_TxtExtenso( 400 ) = "quatrocentos"
x_TxtExtenso( 500 ) = "quinhentos"
x_TxtExtenso( 600 ) = "seiscentos"
x_TxtExtenso( 700 ) = "setentos"
x_TxtExtenso( 800 ) = "oitocentos"
x_TxtExtenso( 900 ) = "novecentos"

' Função Principal de Conversão de Valores em Extenso
Function Extenso( x_Numero )
	x_Numero = FormatNumber( x_Numero , 2 )
	x_Centavos = right( x_Numero , 2 )
	x_ValCentena( 0 ) = 0
	x_QtdCentenas = int( ( len( x_Numero ) + 1 ) / 4 )
	
	For x_I = 1 to x_QtdCentenas
		x_ValCentena( x_I ) = "" 
	Next
	'
	x_I = 1
	x_J = 1
	For x_K = len( x_Numero ) - 3 to 1 step -1
		x_ValCentena( x_J ) = mid( x_Numero , x_K , 1 ) & x_ValCentena( x_J )
	if x_I / 3 = int( x_I / 3 ) then
	x_J = x_J + 1
	x_K = x_K - 1
	end if
	x_I = x_I + 1
	next
	x_TotCentenas = 0
	Extenso = "" 
	For x_I = x_QtdCentenas to 1 step -1
		
		x_TotCentenas = x_TotCentenas + int( x_ValCentena( x_I ) )		
		if int( x_ValCentena( x_I ) ) <> 0 or ( int( x_ValCentena( x_I ) ) = 0 and x_I = 1 )then
			if int( x_ValCentena( x_I ) = 0 and int( x_ValCentena( x_I + 1 ) ) = 0 and x_I = 1 )then
				Extenso = Extenso & ExtCentena( x_ValCentena( x_I ) , x_TotCentenas ) & " de " & x_TxtMoeda( x_I )
			else
				Extenso = Extenso & ExtCentena( x_ValCentena( x_I ) , x_TotCentenas ) & " " & x_TxtMoeda( x_I )
			end if
			
			if int( x_ValCentena( x_I ) ) <> 1 or ( x_I = 1 and x_TotCentenas <> 1 ) then
			
				Select Case x_I
				Case 1
					Extenso = Extenso & "is"
				Case 3, 4, 5
					Extenso = Extenso & "ões"
				End Select 
			else
				Select Case x_I
				Case 1
				Extenso = Extenso & "l"
				Case 3, 4, 5
				Extenso = Extenso & "ão"
				End Select 
			end if
		end if
		if int( x_ValCentena( x_I - 1 ) ) = 0 then
			Extenso = Extenso
		else
			if ( int( x_ValCentena( x_I + 1 ) ) = 0 and ( x_I + 1 ) <= x_QtdCentenas ) or ( x_I = 2 ) then
				Extenso = Extenso & " e "
			else
				Extenso = Extenso & ", "
			end if
		end if 
	next
	
	if x_Centavos > 0 then
		if int( x_Centavos ) = 1 then
			Extenso = Extenso & " e " & ExtDezena( x_Centavos ) & " centavo"
		else
			Extenso = Extenso & " e " & ExtDezena( x_Centavos ) & " centavos"
		end if
	end if
	Extenso = UCase( Left( Extenso , 1 ) )&right( Extenso , len( Extenso ) - 1 )
End Function

function ExtDezena( x_Valor )
	' Retorna o Valor por Extenso referente à DEZENA recebida
	ExtDezena = ""
	if int( x_Valor ) > 0 then
		if int( x_Valor ) < 20 then
			ExtDezena = x_TxtExtenso( int( x_Valor ) )
		else
			ExtDezena = x_TxtExtenso( int( int( x_Valor ) / 10 ) * 10 )
			if ( int( x_Valor ) / 10 ) - int( int( x_Valor ) / 10 ) <> 0 then
				ExtDezena = ExtDezena & " e " & x_TxtExtenso( int( right( x_Valor , 1 ) ) )
			end if
		end if
	end if
End Function

function ExtCentena( x_Valor, x_ValSoma )
	ExtCentena = ""
	
	if int( x_Valor ) > 0 then
		
		if int( x_Valor ) = 100 then
			ExtCentena = "cem"
		else
			if int( x_Valor ) < 20 then
				if int( x_Valor ) = 1 then
					If x_ValSoma - int( x_Valor ) = 0 then
						ExtCentena = "um"
					else
						ExtCentena = x_TxtExtenso( int( x_Valor ) )
					end if
				else
					ExtCentena = x_TxtExtenso( int( x_Valor ) )
				end if
			else
				if int( x_Valor ) < 100 then
					ExtCentena = ExtDezena( right( x_Valor , 2 ) )
				else 
					ExtCentena = x_TxtExtenso( int( int( x_Valor ) / 100 )*100 )
					if ( int( x_Valor ) / 100 ) - int( int( x_Valor ) / 100 ) <> 0 then
						ExtCentena = ExtCentena & " e " & ExtDezena( right( x_Valor , 2 ) )
					end if
				end if
			end if
		end if
	end if
End Function

%>
