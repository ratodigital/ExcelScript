package com.ratodigital.excelscript.util;

import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Collections;
import java.util.List;
import java.util.TimeZone;

import jxl.CellType;
import jxl.DateCell;
import jxl.NumberCell;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.NumberFormats;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import com.google.common.primitives.Ints;

public class Excel {
	private WritableWorkbook out;
	private WritableSheet sheet;

	private String saida;

	public Excel abrirParaEscrita(String entrada, String saida) {
		System.out.println("\nProcessando planilha " + entrada + "...");
		Workbook in;
		try {
			in = Workbook.getWorkbook(new File(entrada));
			out = Workbook.createWorkbook(new File(saida), in);
			sheet = out.getSheet(0);
			this.saida = saida;
			return this;
		} catch (BiffException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			System.out.println("ERRO: Arquivo de entrada " + entrada
					+ " não encontrado!");
		}
		return null;
	}

	public Excel excluirUltimaLinha() {
		System.out.println("Removendo última linha: " + getTotalLinhas());
		sheet.removeRow(getTotalLinhas());

		return this;
	}

	public Excel inserirColuna(int coluna, String titulo) {
		System.out.println("Inserindo coluna " + coluna + ": " + titulo);
		sheet.insertColumn(coluna);
		alterarCelula(coluna, 0, titulo);
		return this;
	}

	public Excel removerTodasColunasExceto(int[] colunasAManter) {
		List<Integer> cols = Ints.asList(colunasAManter);
		Collections.reverse(cols);
		System.out.println("Removendo todas as colunas, exceto " + cols);
		for (int col = sheet.getColumns() - 1; col >= 0; col--) {
			if (!cols.contains(new Integer(col))) {
				sheet.removeColumn(col);
			}
		}
		return this;
	}

	public Excel removerColuna(int colunaARemover) {
		return removerColunas(new int[] { colunaARemover });
	}

	public Excel removerColunas(int[] colunasARemover) {
		List<Integer> cols = Ints.asList(colunasARemover);
		Collections.reverse(cols);
		System.out.println("Removendo colunas " + cols);
		for (int i = 0; i < cols.size(); i++) {
			sheet.removeColumn(cols.get(i));
		}
		return this;
	}

	public double obtemValorCelula(int coluna, int linha) {
		return sheet.getCell(coluna, linha).getType() == CellType.NUMBER ? ((NumberCell) sheet
				.getCell(coluna, linha)).getValue() : 0.0;
	}

	public String obtemTextoCelula(int coluna, int linha) {
		return sheet.getCell(coluna, linha).getContents();
	}

	public String obtemDataCelula(int coluna, int linha) {
		TimeZone gmtZone = TimeZone.getTimeZone("GMT");
		SimpleDateFormat format = new SimpleDateFormat("dd/MM/yyyy");
		format.setTimeZone(gmtZone);
	
		DateCell dateCell = (DateCell) sheet.getCell(coluna, linha);
		String dateString = format.format(dateCell.getDate());
		return dateString;
	}
	
	public Excel salvarEFechar() {
		System.out
				.println("Arquivo de saída " + saida + " gerado com sucesso!");
		try {
			out.write();
			out.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return this;
	}

	public int getTotalLinhas() {
		return sheet.getRows() - 1;
	}

	public Excel alterarCelula(int coluna, int linha, String valor) {
		try {
			sheet.addCell(new Label(coluna, linha, valor));
		} catch (RowsExceededException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return this;
	}

	public Excel alterarCelula(int coluna, int linha, double valor) {
		return alterarCelula(coluna, linha, valor, true);
	}

	public Excel alterarCelula(int coluna, int linha, double valor,
			boolean decimals) {
		try {
			if (decimals) {
				sheet.addCell(new Number(coluna, linha, valor,
						new WritableCellFormat(NumberFormats.FORMAT3)));
			} else {
				sheet.addCell(new Number(coluna, linha, valor));
			}
		} catch (RowsExceededException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return this;
	}

	public void juntasArquivosEmUnicaPlanilha(String diretorio,
			String arquivoSaida) {
		System.out.println("\nJuntando planilhas de " + diretorio + " em "
				+ arquivoSaida);
		try {
			String fileName = arquivoSaida;
			WritableWorkbook workbookDestino = Workbook.createWorkbook(new File(
					fileName));

			File[] arquivos = Arquivo.listaArquivos(diretorio, ".xls");

			Workbook workbookOrigem;
			WritableWorkbook tempOrigem = null;
			File tempFile = new File(arquivoSaida + "_temp");
			
			int s = 0;
			for (File file : arquivos) {
				if (file.isFile()) {
					System.out.println(file.getAbsolutePath());
					try {
						workbookOrigem = Workbook.getWorkbook(file);
						tempOrigem = Workbook.createWorkbook(tempFile, workbookOrigem);
						WritableSheet sheetOrigem = tempOrigem.getSheet(0);
						WritableSheet sheetDestino = workbookDestino.createSheet(workbookOrigem.getSheet(0).getName(), s++);
						
						for (int row = 0; row < sheetOrigem.getRows(); row++) {
							//System.out.println("     linha" + i);
							for (int col = 0; col < sheetOrigem.getColumns(); col++) {
								WritableCell cellOrigem = sheetOrigem.getWritableCell(col, row);
								WritableCell cellDestino = cellOrigem.copyTo(col, row);
								sheetDestino.addCell(cellDestino); 	
							}
						}
					} catch (Exception e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			}

			tempOrigem.close();
			tempFile.delete();
			workbookDestino.write();
			workbookDestino.close();
		} catch (WriteException e) {
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
}
