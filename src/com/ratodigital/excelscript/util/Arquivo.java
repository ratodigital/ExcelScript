package com.ratodigital.excelscript.util;

import java.io.File;
import java.io.FilenameFilter;

public class Arquivo {
	public static boolean diretorioExiste(String diretorio) {
		return new File(diretorio).exists();
	}

	public static void criaDiretorio(String diretorio) {
		File dir = new File(diretorio);

		if (!dir.exists()) {
			System.out.println("Criando pasta: " + diretorio);
			dir.mkdir();
		}
	}

	public static File[] listaArquivos(String diretorio, final String filtro) {
		FilenameFilter filter = new FilenameFilter() {
		    public boolean accept(File dir, String name) {
		        return name.endsWith(filtro);
		    }
		};

		File folder = new File(diretorio);
		
		return(folder.listFiles(filter));
	}
}
