package json2xlsm.cli;

import json2xlsm.lib.json2xlsm;
import org.apache.commons.cli.*;

import java.io.IOException;
import java.security.GeneralSecurityException;

/**
 * Command Line main class to call json2xlsm object.
 * @author Pep Marxuach, jmarxuach
 * @version 1.0.0.0
 */
public class Main {
    
	/**
	 * This method is called when the jar is executed. It shows the help if parameters weren't given correctly.
	 * It creates a json2xlsm Object and calls the ExecuteExport(); function on it.
	 * @author Pep Marxuach, jmarxuach
	 */
	public static void main(String[] args){

       String errorMiss = "Command line example :"
			+ "\n\tjava -jar json2xlsm.jar <strFileJSON> <strMacroExcelFileIn> <strMacroExcelFileOut>";
		
	  if (args.length>0) {
		
			try {
				json2xlsm Exp; 
				Exp = new json2xlsm();
			    Exp.ExecuteExport(args[0],args[1],args[2]);
			    System.out.println("\n\t" + errorMiss);
			    System.exit(0);
				
			} catch (Exception e2) {
				e2.printStackTrace();
				System.out.println("\n\t" + errorMiss);
				System.exit(-1);					
			}			
			
		} else {
			System.out.println(errorMiss);
			System.exit(-1);
		}		
    }
}
