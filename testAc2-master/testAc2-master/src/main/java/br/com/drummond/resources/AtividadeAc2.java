package br.com.drummond.resources;

import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class AtividadeAc2 {
	
	@RequestMapping(value = "/api3/{valor1}/{valor2}", method = RequestMethod.GET)
	public Integer somar(@PathVariable Integer valor1, @PathVariable Integer valor2) {
		
		Integer resultado;
		
		resultado = valor1 / valor2;

		return resultado;
	}
}