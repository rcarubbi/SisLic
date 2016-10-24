/*
Arquivo de funções javascript
Data: 15/04/2006
*/
/*
function addControles(controle, tipo) //adiciona um controle a uma coleção de controles
{
	var ControlsCollection = new 	

}


function validacao(controle, tipo) // Função que valida controles de tela
/*
Utilizar o parâmetro tipo para o tipo de validação onde:
1. Inteiros Obrigatorios
2. Inteiros Não Obrigatorios
3. Caracteres Obrigatórios
4. Caracteres Não Obrigatórios
5. Data válida no formato dd/mm/yyyyy

{
	var regra 
	switch tipo
	{
		case 1: regra = /^\d+$/;break;
		case 2: regra = /^\d?+$/;break;
		case 3: regra = /^\w+$/;break;
		case 4: regra = /^\w+?$/;break;
		case 5: regra = /^\d{1,2}\/\d{1,2}\/\d{1,4}$/;break; 
	}
	if (regra.test(controle)==false)
	{
		alert("Valor invalido");
	}
	
		return regra.test(controle);
		
}
*/


function maxLengthTextArea(obj, evento, maxLength) {


	if(obj.value.length>=maxLength) {
		switch (evento.keyCode) {
		case 8:
			return true;			
			break 
		case 46:
			return true;			
			break 
		case 37:
			return true;			
			break
		case 38:
			return true;			
			break 
		case 39:
			return true;			
			break
		case 40:
			return true;			
			break
		default:
		//alert('xxx');
			obj.value=obj.value.slice(0,1000);
			return false;
		}
	
	} else {
		return true;
	}

}

	
function mudaCor(obj, corNormal, corSelecionado) {
	if(obj.style.background == corSelecionado) {
		obj.style.background = corNormal;
	} else {
		obj.style.background = corSelecionado;
	}	
}

	


