package br.senac.rj.atividadeexcel.util;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import br.senac.rj.atividadeexcel.domain.Evento;

public class ImportaExportaExcel {

	public List<Evento> importaExcel(String filePath) {
		List<Evento> eventos = new ArrayList<>();
		try {
			FileInputStream file = new FileInputStream(new File(filePath));
			//Pasta com as tabelas do excel.
			Workbook workbook = new XSSFWorkbook(file);
			//Abre a primeira tabela da pasta do excel.
			Sheet sheet = workbook.getSheetAt(0);
			//Criando um objeto evento para cada linha e adicionando na lista...
			SimpleDateFormat sdf = new SimpleDateFormat("dd/MM");
			DataFormatter formatter = new DataFormatter();
			sheet.forEach(linha->{
				if(linha.getRowNum()>0) {
					Evento evento = new Evento();
					evento.setId(null);
					evento.setAlvo(formatter.formatCellValue(linha.getCell(0)));
					try {
						evento.setData(sdf.parse(linha.getCell(1).getRichStringCellValue().getString()));
					} catch (ParseException e) {
						e.printStackTrace();
					}
					evento.setHora(formatter.formatCellValue(linha.getCell(2)));
					evento.setCodigo(formatter.formatCellValue(linha.getCell(3)));
					evento.setDescricao(formatter.formatCellValue(linha.getCell(4)));
					evento.setEvento(formatter.formatCellValue(linha.getCell(5)));
					evento.setLocal(formatter.formatCellValue(linha.getCell(6)));
					evento.setReferencia(formatter.formatCellValue(linha.getCell(7)));
					evento.setUsuario(formatter.formatCellValue(linha.getCell(8)));
					eventos.add(evento);
				}
			});
			workbook.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {			
			e.printStackTrace();
		}
		return eventos;	
	}
	
	public ByteArrayInputStream exportaExcel(List<Evento> eventos) {
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("Eventos SAP");
		
		//Criando a linha de cabecalho
		Row header = sheet.createRow(0);
		
		//Definindo o estilo do cabecalho
		CellStyle headerStyle = workbook.createCellStyle();
		headerStyle.setFillForegroundColor(IndexedColors.GREEN.getIndex());
		headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		//Definindo o nome das colunas...
		Cell headerCell = header.createCell(0);
		headerCell.setCellValue("Alvo");
		headerCell.setCellStyle(headerStyle);
		
		headerCell = header.createCell(1);
		headerCell.setCellValue("Data");
		headerCell.setCellStyle(headerStyle);
		
		headerCell = header.createCell(2);
		headerCell.setCellValue("Hora");
		headerCell.setCellStyle(headerStyle);
		
		headerCell = header.createCell(3);
		headerCell.setCellValue("Código");
		headerCell.setCellStyle(headerStyle);
		
		headerCell = header.createCell(4);
		headerCell.setCellValue("Descrição");
		headerCell.setCellStyle(headerStyle);
		
		headerCell = header.createCell(5);
		headerCell.setCellValue("Evento");
		headerCell.setCellStyle(headerStyle);
		
		headerCell = header.createCell(6);
		headerCell.setCellValue("Local");
		headerCell.setCellStyle(headerStyle);
		
		headerCell = header.createCell(7);
		headerCell.setCellValue("Referência");
		headerCell.setCellStyle(headerStyle);
		
		headerCell = header.createCell(8);
		headerCell.setCellValue("Usuário");
		headerCell.setCellStyle(headerStyle);
		
		//Definindo o estilo das linhas impares...
		CellStyle estiloLinhaImpar = workbook.createCellStyle();
		estiloLinhaImpar.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		estiloLinhaImpar.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		Row linha;
		Cell cell;
		for(int i=0;i<eventos.size();i++) {
			linha = sheet.createRow(i+1);
			
			cell = linha.createCell(0);
			cell.setCellValue(eventos.get(i).getAlvo());
			
			cell = linha.createCell(1);
			cell.setCellValue(eventos.get(i).getData());
			
			cell = linha.createCell(2);
			cell.setCellValue(eventos.get(i).getHora());
			
			cell = linha.createCell(3);
			cell.setCellValue(eventos.get(i).getCodigo());
			
			cell = linha.createCell(4);
			cell.setCellValue(eventos.get(i).getDescricao());
			
			cell = linha.createCell(5);
			cell.setCellValue(eventos.get(i).getEvento());
			
			cell = linha.createCell(6);
			cell.setCellValue(eventos.get(i).getLocal());
			
			cell = linha.createCell(7);
			cell.setCellValue(eventos.get(i).getReferencia());
			
			cell = linha.createCell(8);
			cell.setCellValue(eventos.get(i).getUsuario());
			
			//Em caso de linha impar, pinta a linha de cinza claro...
			if(i%2!=0) {
				for(int j=0; j<9;j++) {
					linha.getCell(j).setCellStyle(estiloLinhaImpar);
				}
			}
		}
		
		ByteArrayOutputStream out = new ByteArrayOutputStream();
		try {
			workbook.write(out);
			workbook.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		
		return new ByteArrayInputStream(out.toByteArray());
	}
}
