package br.senac.rj.atividadeexcel.repository;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import br.senac.rj.atividadeexcel.domain.Evento;

@Repository
public interface EventoRepository extends JpaRepository<Evento, Integer> {

}
