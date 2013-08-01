import com.ratodigital.excelscript.util.Arquivo
import com.ratodigital.excelscript.util.Excel

Planilhas bsoft = new Planilhas(DIR_ENTRADA, DIR_SAIDA, EXT, SAIDA)

bsoft.cpagar("cpagar")
bsoft.creceber("creceber")
bsoft.manifesto("manifesto")
bsoft.ctrc("ctrc")
bsoft.manifestoCtrc("manifestoctrc")
bsoft.juntaTudo("zoho")
		
	
class Planilhas {
	String DIR_ENTRADA
	String DIR_SAIDA
	String EXT
	String SAIDA
	
	Planilhas(dirEntrada, dirSaida, ext, saida) {
		DIR_ENTRADA = dirEntrada
		DIR_SAIDA = dirSaida
		EXT = ext
		SAIDA = saida
	}
	
	/*
	 * cpagar.xls - ficam iguais as colunas Documento, Parcela, Emissão,
	 * Vencimento, Pagamento, Valor Pago, Centro de Custo - entre as colunas
	 * Vencimento e Pagamento, cria-se a coluna Valor Total =
	 * Valor+Multa+Juros-Desconto. Esta coluna deve ter o valor copiado e colado
	 * "especial" como "valores". Assim, as colunas que dao origem ao valor
	 * podem ser excluidas - exclui a ultima linha que tem os totais
	 */
	public void cpagar(String file) {
		String entrada = DIR_ENTRADA + file + EXT
		String saida = DIR_SAIDA + file + SAIDA
		
		Excel e = new Excel()
		
		e.abrirParaEscrita(entrada, saida)
				.excluirUltimaLinha()
				.removerTodasColunasExceto(
						[0, 1, 3, 4, 5, 6, 7, 8, 9, 10, 11, 18, 19] as int[])
				.inserirColuna(6, "Valor Total")
		
		println "Adicionando valores na coluna Valor Total"
		1.upto(e.getTotalLinhas()) { row ->
			double valor = e.obtemValorCelula(6, row)
			double multa = e.obtemValorCelula(7, row)
			double juros = e.obtemValorCelula(8, row)
			double desconto = e.obtemValorCelula(9, row)
			double valorTotal = valor + multa + juros - desconto
			e.alterarCelula(6, row, valorTotal)
		}
		
		e.removerColunas([7, 8, 9, 10] as int[]).salvarEFechar()
		//e.salvarEFechar()
	}
	
	/*
	 * creceber.xls - fica igual o conteudo da coluna Documento com o novo nome
	 * "Nº do Documento a Receber" - ficam iguais as colunas Documento, Parcela,
	 * Nome do Cliente, Emissão, Vencimento, Pagamento, Valor Pago - entre as
	 * colunas Vencimento e Pagamento, cria-se a coluna Valor Total =
	 * Valor+Multa+Juros-Desconto. Esta coluna deve ter o valor copiado e colado
	 * "especial" como "valores". Assim, as colunas que dao origem ao valor
	 * podem ser excluidas - exclui a ultima linha que tem os totais
	 */
	public void creceber(String file) {
		String entrada = DIR_ENTRADA + file + EXT
		String saida = DIR_SAIDA + file + SAIDA

		Excel e = new Excel()

		e.abrirParaEscrita(entrada, saida)
				.excluirUltimaLinha()
				.alterarCelula(0, 0, "Nº do Documento a Receber")
				.removerTodasColunasExceto([0, 1, 3, 4, 5, 6, 7, 8, 9, 10, 11, 18, 19] as int[])
				.inserirColuna(10, "Valor Total")

		println "Adicionando valores na coluna Valor Total"
		1.upto(e.getTotalLinhas()) { row ->
			double valor = e.obtemValorCelula(6, row)
			double multa = e.obtemValorCelula(7, row)
			double juros = e.obtemValorCelula(8, row)
			double desconto = e.obtemValorCelula(9, row)
			double valorTotal = valor + multa + juros - desconto
			e.alterarCelula(10, row, valorTotal)
		}

		e.removerColunas([6, 7, 8, 9] as int[])
			.salvarEFechar()
	}

	/*
	 * manifesto.xls - ficam iguais as colunas Nº Manifesto, Data, Soma de Peso,
	 * Soma Valor Total, Frete Motorista
	 * 
	 * ctrc.xls - ficam iguais as colunas Conhecimento, Data, Remetente, Soma
	 * dos Pesos, Total, Nº do Documento a Receber
	 * 
	 * manifesto-ctrc.xls - ficam iguais as colunas Nº Manifesto, Destino - a
	 * coluna CTRC precisa ser explodida em ctrc1, ctrc2, ctrc3, ctrc4, ctrc5,
	 * ctrc6, ctrc7, ctrc8 - cria-se a coluna Frete Motorista=Sub
	 * Total-Desconto-Outras Despesas. Esta coluna deve ter o valor copiado e
	 * colado "especial" como "valores". Assim, as colunas que dao origem ao
	 * valor podem ser excluidas - exclui a ultima linha que tem os totais
	 */
	public void manifesto(String file) {
		String entrada = DIR_ENTRADA + file + EXT
		String saida = DIR_SAIDA + file + SAIDA

		Excel e = new Excel()

		e.abrirParaEscrita(entrada, saida)
				.removerTodasColunasExceto([0, 1, 7, 9, 11] as int[])
				.salvarEFechar()
	}

	/*
	 * ctrc.xls - ficam iguais as colunas Conhecimento, Data, Remetente, Soma
	 * dos Pesos, Total, Nº do Documento a Receber
	 */
	public void ctrc(String file) {
		String entrada = DIR_ENTRADA + file + EXT
		String saida = DIR_SAIDA + file + SAIDA

		Excel e = new Excel()

		e.abrirParaEscrita(entrada, saida)
				.removerTodasColunasExceto([2, 4, 7, 13, 16, 20, 22] as int[])
				.salvarEFechar()
	}

	/*
	 * manifesto-ctrc.xls - ficam iguais as colunas Nº Manifesto, Destino - a
	 * coluna CTRC precisa ser explodida em ctrc1, ctrc2, ctrc3, ctrc4, ctrc5,
	 * ctrc6, ctrc7, ctrc8 - cria-se a coluna Frete Motorista=Sub
	 * Total-Desconto-Outras Despesas. Esta coluna deve ter o valor copiado e
	 * colado "especial" como "valores". Assim, as colunas que dao origem ao
	 * valor podem ser excluidas - exclui a ultima linha que tem os totais
	 */
	public void manifestoCtrc(String file) {
		String entrada = DIR_ENTRADA + file + EXT
		String saida = DIR_SAIDA + file + SAIDA

		Excel e = new Excel()

		e.abrirParaEscrita(entrada, saida)
			.excluirUltimaLinha()
			.removerTodasColunasExceto([1, 4, 12, 27, 30, 37] as int[])

		println "Insere colunas CTRC 1 a 8"
		(1..8).each { i ->
			e.inserirColuna(5 + i, "ctrc" + i)
		}

		e.inserirColuna(14, "Frete Motorista")

		println "Adicionando valores na coluna Frete Motorista"
		1.upto(e.getTotalLinhas()) { row ->
			double subtotal = e.obtemValorCelula(3, row)
			double desconto = e.obtemValorCelula(4, row)
			double outrasDespesas = e.obtemValorCelula(5, row)
			double freteMotorista = subtotal - desconto - outrasDespesas
			e.alterarCelula(14, row, freteMotorista)
		}

		// Remove as colunas usadas para calcular Frete Motorista
		e.removerColunas([3, 4, 5] as int[])

		// Quebra CTRC em várias colunas
		println "Quebrando coluna CTRC"
		1.upto(e.getTotalLinhas()) { row ->	
			String ctrcs = e.obtemTextoCelula(2, row)
			String[] ctrc = ctrcs.split(",")
			(0..<ctrc.length).each { c ->
				e.alterarCelula(3 + c, row, ctrc[c] as int, false)
			}
		}

		// Remove coluna CTRC e salva
		e.removerColuna(2)
			.salvarEFechar()
	}
	
	public void juntaTudo(String saida) {
		Excel e = new Excel()
		Arquivo.criaDiretorio(DIR_SAIDA + "/zoho/")		
		e.juntasArquivosEmUnicaPlanilha(DIR_SAIDA, DIR_SAIDA + "zoho/" + saida + EXT)
		println "TUDO OK!!!"
	}

}