import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;


public class ExcelToJson {

	/**
	 * @param args
	 * @throws IOException 
	 * @throws InvalidFormatException 
	 * @throws JSONException 
	 */
	public static void main(String[] args) {
		//	String file = "/Users/miniuno/Desktop/cameras.xls";
		if(args.length < 2){
			System.out.println("**Error: Invalid usage**");
			System.out.println("\texample usage");
			System.out.println("\t\tjava -jar ExcelToJson.jar Cameras.xlsx cameras.json");
			System.out.println("\t\tjava -jar ExcelToJson.jar Cameras.xls cameras.json");
			return;
		}
		String file = args[0]; 

		// TODO Auto-generated method stub
		FileInputStream inp;
		Workbook workbook;
		try {
			System.out.println("Loading the Excel File...");
			inp = new FileInputStream( file  );
			workbook = WorkbookFactory.create( inp );
			
			// Get the first Sheet.
			Sheet sheet = workbook.getSheetAt(0);
			System.out.println("File Load successfull");
			System.out.println("Processing...");
			// Start constructing JSON.
			JSONArray json = new JSONArray();
			boolean isFirstRow = true;
			ArrayList<String> rowName = new ArrayList<String>();
			for ( Iterator<Row> rowsIT = sheet.rowIterator(); rowsIT.hasNext(); )
			{
				Row row = rowsIT.next();
				JSONObject jRow = new JSONObject();
				// Iterate through the cells.

				if(isFirstRow){
					for ( Iterator<Cell> cellsIT = row.cellIterator(); cellsIT.hasNext(); )
					{
						Cell cell = cellsIT.next();
						if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC){
							rowName.add(String.valueOf(cell.getNumericCellValue()));
						}
						else{			            	
							rowName.add(cell.getStringCellValue());
						}

					}
					isFirstRow = false;
				}
				else{
					int cellCount = 0;
					for ( Iterator<Cell> cellsIT = row.cellIterator(); cellsIT.hasNext(); )
					{		        		
						Cell cell = cellsIT.next();
						if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC)
							jRow.put(rowName.get(cellCount++), String.valueOf(cell.getNumericCellValue() ));		            	
						else
							jRow.put(rowName.get(cellCount++), cell.getStringCellValue());     		        		
					}
					json.put(jRow);
				}
			}
			File outputFile = new File(args[1]);
			FileOutputStream outputFileStream = new FileOutputStream(outputFile);
			outputFileStream.write(json.toString().getBytes());
			System.out.println("Excel file successfully exported to Json");
		} catch (InvalidFormatException e) {
			// TODO Auto-generated catch block
			System.out.println("Invalid Format, Only Excel files are supported");
			e.printStackTrace();
		} catch (IOException e) {
			System.out.println("Check if the input file exists and the path is correct");
			e.printStackTrace();
		} catch (JSONException e) {
			// TODO Auto-generated catch block
			System.out.println("Unable to generate Json");
			e.printStackTrace();
		} 
	}
}
