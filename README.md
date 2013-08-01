ExcelScript
===========

Este é um projeto Java/Eclipse comum (sem Maven) que serve para leitura e manipulação de planilhas Excel. 

Nos bastidores, o projeto usa a simples e direta [Java Excel API](http://jexcelapi.sourceforge.net).

A melhor forma de entender seu funcionamento é executando o [target/excelscript.sh](https://github.com/ratodigital/ExcelScript/blob/master/target/excelscript.sh) (ou .bat, se estiver no windows). 
São realizadas várias transformações nas planilhas localizadas na pasta */sample*

Todas os comandos devem ser especificados no arquivo [script.groovy](https://github.com/ratodigital/ExcelScript/blob/master/sample/normando/script.groovy), usando-se a sintaxe da linguagem Groovy. 
Usando um editor de texto comum você poderá automatizar várias operações de leitura e escrita em planilhas Excel.

Abaixo um trecho, para ilustrar o potencial da linguagem. 

```java
    import com.ratodigital.excelscript.util.Arquivo
    import com.ratodigital.excelscript.util.Excel

    Planilhas planilhas = new Planilhas(DIR_ENTRADA, DIR_SAIDA, EXT, SAIDA)

    planilhas.cpagar("cpagar")
  	
 
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
		    1.upto(*e.getTotalLinhas()*) { row ->
			    double valor = e.obtemValorCelula(6, row)
			    double multa = e.obtemValorCelula(7, row)
			    double juros = e.obtemValorCelula(8, row)
			    double desconto = e.obtemValorCelula(9, row)
			    double valorTotal = valor + multa + juros - desconto
		      e.alterarCelula(6, row, valorTotal)
    		}
		
		    e.removerColunas([7, 8, 9, 10] as int[]).salvarEFechar()
	    }
    }
```

Eu pensei em criar uma DSL mais fácil, mas cadê tempo? Quem tiver a fim de colaborar, fica à vontade!
