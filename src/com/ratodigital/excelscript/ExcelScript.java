package com.ratodigital.excelscript;

import groovy.lang.Binding;
import groovy.util.GroovyScriptEngine;
import groovy.util.ResourceException;
import groovy.util.ScriptException;

import java.io.File;
import java.io.IOException;

import com.ratodigital.excelscript.util.Arquivo;


public class ExcelScript {

	private static String DIR_ENTRADA = System.getProperty("user.dir");

	private static String DIR_SAIDA = DIR_ENTRADA;

	private final static String EXT = ".xls";

	private final static String SAIDA = "_saida" + EXT;

	public static void main(String[] args) {
		if (args.length != 3) {
			System.out.println("USO: java -jar bittle.jar <SCRIPT.groovy> <DIR_ENTRADA> <DIR_SAIDA>" + args.length);
			System.out.println("     SCRIPT      - caminho e nome do arquivo do script, com extensão .groovy");
			System.out.println("     DIR_ENTRADA - caminho das planilhas de origem.");
			System.out.println("     DIR_SAIDA   - caminho das planilhas de saida.");
			System.exit(0);
		}

		String script = args[0];
		DIR_ENTRADA = args[1];
		DIR_SAIDA = args[2];

		Arquivo.criaDiretorio(DIR_SAIDA);

		System.out.println("Executando script   : " + script);
		System.out.println("Diretorio de entrada: " + DIR_ENTRADA);
		System.out.println("Diretorio de saida  : " + DIR_SAIDA);

		run(script);
	}

	public static void run(String scriptFileName) {
		File script = new File(scriptFileName);
		Binding binding = new Binding();
		binding.setProperty("DIR_ENTRADA", DIR_ENTRADA);
		binding.setProperty("DIR_SAIDA", DIR_SAIDA);
		binding.setProperty("EXT", EXT);
		binding.setProperty("SAIDA", SAIDA);
		GroovyScriptEngine gse;
		try {
			gse = new GroovyScriptEngine(script.getPath());
			gse.run(script.getName(), binding);
		} catch (ResourceException ex) {
			ex.printStackTrace();
		} catch (ScriptException ex) {
			ex.printStackTrace();
		} catch (IOException ex) {
			ex.printStackTrace();
		}

	}
}
