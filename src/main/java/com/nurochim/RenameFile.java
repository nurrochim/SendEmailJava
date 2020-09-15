package com.nurochim;

import java.io.File;

public class RenameFile {
	public String RenameFile(String oldFilePath, String newFilePath) {
		try {
			String dir = System.getProperty("user.dir");
			System.out.println(oldFilePath+" to "+ newFilePath);
			
			File oldfile =new File(oldFilePath);
	        File newfile =new File(newFilePath);

	        if(oldfile.renameTo(newfile)){
	            System.out.println("File renamed");
	            return "Succes rename";
	        }else{
	            System.out.println("Sorry! the file can't be renamed");
	            return "Failed rename";
	        }
	        
		} catch (Exception e) {
			e.printStackTrace();
			return "Failed rename file "+oldFilePath;
		}
	}
}
