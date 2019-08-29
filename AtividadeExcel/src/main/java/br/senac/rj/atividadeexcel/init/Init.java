package br.senac.rj.atividadeexcel.init;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.ApplicationListener;
import org.springframework.context.event.ContextRefreshedEvent;
import org.springframework.stereotype.Component;

import br.senac.rj.atividadeexcel.domain.Evento;
import br.senac.rj.atividadeexcel.repository.EventoRepository;
import br.senac.rj.atividadeexcel.util.ImportaExportaExcel;

@Component
public class Init implements ApplicationListener<ContextRefreshedEvent> {
	
	@Autowired
	EventoRepository repository;

	@Override
	public void onApplicationEvent(ContextRefreshedEvent event) {
		ImportaExportaExcel importaExportaExcel = new ImportaExportaExcel();
		List<Evento> eventos = importaExportaExcel.importaExcel("C:\\Desenvolvimento\\MDriver.xlsx");
		repository.saveAll(eventos);
	}

}
