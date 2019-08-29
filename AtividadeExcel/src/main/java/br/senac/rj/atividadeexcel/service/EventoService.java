package br.senac.rj.atividadeexcel.service;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import br.senac.rj.atividadeexcel.domain.Evento;
import br.senac.rj.atividadeexcel.repository.EventoRepository;

@Service
public class EventoService {

	@Autowired
	EventoRepository repository;
	
	public List<Evento> findAll() {
		return repository.findAll();
	}
}
