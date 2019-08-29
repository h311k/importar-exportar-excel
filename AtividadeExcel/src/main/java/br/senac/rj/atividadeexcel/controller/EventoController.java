package br.senac.rj.atividadeexcel.controller;

import java.io.ByteArrayInputStream;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import br.senac.rj.atividadeexcel.service.EventoService;
import br.senac.rj.atividadeexcel.util.ImportaExportaExcel;

@RestController
@RequestMapping("/exportar")
public class EventoController {
	
	@Autowired
	EventoService service;
	
	@GetMapping("/eventos")
	public ResponseEntity<InputStreamResource> downloadExcel() {
		ImportaExportaExcel importaExportaExcel = new ImportaExportaExcel();
		ByteArrayInputStream resultado = importaExportaExcel.exportaExcel(service.findAll());
		
		HttpHeaders headers = new HttpHeaders();
        headers.add("Content-Disposition", "attachment; filename=Eventos.xlsx");
    
     return ResponseEntity
                  .ok()
                  .headers(headers)
                  .body(new InputStreamResource(resultado));   
	}

}
