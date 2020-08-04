package json2xlsm.lib;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Iterator;
import java.util.Map;
import java.util.List;
import java.util.Set;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import com.google.gson.*;

/**
 * Adds json data into xlsm first sheet macro excel file
 * 
 * @author Pep Marxuach, jmarxuach
 * @version 1.0.0.1
 */
public class json2xlsm {

	/**
	 * Constructor : Nothing to do.
	 * 
	 * @author Pep Marxuach, jmarxuach
	 */
	public json2xlsm() throws Exception {

	}

	/**
	 * Parse JSON and adds json data into xlsm first sheet macro excel file.
	 * 
	 * @param strFileJSON
	 * @param strExcelFileIn
	 * @param strExcelFileOut
	 * @throws InvalidFormatException
	 * @throws Exception
	 * @author Pep Marxuach, jmarxuach
	 */
	public void ExecuteExport(String strFileJSON, String strExcelFileIn, String strExcelFileOut) throws Exception {

		DateFormat dateFormat = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
		java.util.Date dateStart;
		java.util.Date dateEnd;
		DecimalFormat df1 = new DecimalFormat("###0.00");
		try {
			dateStart = new java.util.Date();

			String extension = "";

			int i = strFileJSON.lastIndexOf('.');
			if (i > 0) {
				extension = strFileJSON.substring(i + 1).toLowerCase();
			}

			if (extension.equals("json"))
				this.json2excel(strFileJSON, strExcelFileIn, strExcelFileOut);

			dateEnd = new java.util.Date();

		} catch (Exception e) {
			throw e;
		}

	}

	/**
	 * Checks if a string is numeric representation
	 * 
	 * @returns True is string is numeric, or False otherwise
	 */
	public boolean isNumber(String num) {
		try {
			Integer.parseInt(num);
			return true;
		} catch (NumberFormatException nfe) {
			try {
				num.replace(",", ".");
				Double.parseDouble(num);
				return true;
			} catch (NumberFormatException nfe2) {
				return false;
			}
		}
	}

	/**
	 * 
	 * @param strFileJSON
	 * @param strExcelFileIn
	 * @param strExcelFileOut
	 * @throws InvalidFormatException
	 * @throws IOException, InvalidFormatException
	 * @author Pep Marxuach, jmarxuach
	 */
	private void json2excel(String strFileJSON, String strExcelFileIn, String strExcelFileOut)
			throws InvalidFormatException, IOException {
		InputStream inp = new FileInputStream(strExcelFileIn);

		Workbook wb = WorkbookFactory.create(inp);
		Sheet sheet = wb.getSheetAt(0);
		Row row;
		Cell cell;
		int linea = 0;
		int columna;

		
		Map[] records = this.readGSONFile(strFileJSON);
		String k;
		String FieldValue;
		
		if (records != null) {	
		   	for (int i = 0; i < records.length; i++) {
		   		
				row = sheet.getRow(linea);
				if (row == null)
					row = sheet.createRow(linea);
				
		   		Iterator<Map.Entry> itrKeys = records[i].entrySet().iterator();
		   		if (i==0) {
		   			columna = 0;
		   			while (itrKeys.hasNext()) {
						Map.Entry pair = itrKeys.next();
						k = pair.getKey().toString();
						cell = row.getCell(columna);
						if (cell == null)
							cell = row.createCell(columna);
			        	cell.setCellValue(k);
			        	columna++;
						
			   	 	}
		   			linea++;
		   			row = sheet.getRow(linea);
					if (row == null)
						row = sheet.createRow(linea);
		   			
		   		} 
		   		
		   		Iterator<Map.Entry> itr1 = records[i].entrySet().iterator();
		   		columna = 0;		
		   		while (itr1.hasNext()) {
					Map.Entry pair = itr1.next();
					if (pair.getValue()==null)
		                	FieldValue = "";
		                else FieldValue = pair.getValue().toString();
					 
					  cell = row.getCell(columna);
					if (cell == null)
						cell = row.createCell(columna);

					if (this.isNumber(FieldValue)) {
						// cell.setCellType(Cell.CELL_TYPE_NUMERIC);
						cell.setCellValue(Double.parseDouble(FieldValue));
					} else
						cell.setCellValue(FieldValue);
					columna++;
					
		   	 	}
		   		
		   	 	linea++;
	
		   	}
		}

		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream(strExcelFileOut);
		wb.write(fileOut);
		fileOut.close();

	}

	/**
	 *
	 * @author Pep Marxuach, jmarxuach
	 */
	private Map[] readGSONFile(String jsonFilename) {
		
       try {
       	
    	   Gson gson = new Gson();
    	   Map[] map = gson.fromJson(new FileReader(jsonFilename), Map[].class);
    	   
    	   return map;    	   	
       	   
       } catch (IOException e) {
           e.printStackTrace();
       } 
       
       
       return null;
		
	}

}
