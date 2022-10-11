package dom5utils;
import java.io.BufferedWriter;
import java.io.FileOutputStream;
import java.io.FileWriter;
/* This file is part of dom5utils.
*
* dom5utils is free software: you can redistribute it and/or modify
* it under the terms of the GNU General Public License as published by
* the Free Software Foundation, either version 3 of the License, or
* (at your option) any later version.
*
* dom5utils is distributed in the hope that it will be useful,
* but WITHOUT ANY WARRANTY; without even the implied warranty of
* MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
* GNU General Public License for more details.
*
* You should have received a copy of the GNU General Public License
* along with dom5utils.  If not, see <http://www.gnu.org/licenses/>.
*/
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class CSVWriter {

	enum Delimiter {
		COMMA, TAB,
	};
	
	//Spreadsheet type
	enum SSType {
		CSV, XLSX,
	}
	
	private static String getFilePathWithExtension(String basename, SSType ssType) {
		if (basename == null || basename.isEmpty()) {
			return null;
		}
		switch (ssType) {
		case CSV:
			return Paths.get(CSV_OUTPUT_DIR_NAME, basename + ".csv").toString();
		case XLSX:
		default:
			return Paths.get(basename + ".xlsx").toString();
		}
	}
	
	
	public static void createCSVOutputDirectory() throws IOException {
		Files.createDirectories(Paths.get(CSV_OUTPUT_DIR_NAME));
	}
	
	public static FileOutputStream getFOS (String basename, SSType ssType) throws IOException {
		if (basename != null && !basename.isEmpty()) {
			return new FileOutputStream(CSVWriter.getFilePathWithExtension(basename, ssType));
		}
		return null;
	}
	
	public static BufferedWriter getBFW (String basename, SSType ssType) throws IOException {
		if (basename != null && !basename.isEmpty()) {
			return new BufferedWriter(new FileWriter(CSVWriter.getFilePathWithExtension(basename, ssType)));
		}
		return null;
	}
	
	public static void writeSimpleCSV(XSSFSheet sheet, BufferedWriter writer, Delimiter delim) throws IOException {
		final char delimChar = getDelimeterChar(delim);
		
		if (sheet != null)
		{
			int columnsInWidestRow = 0;
			
			//first, we need to get the number of columns in the row that has the most defined columns
			for (Row curRow : sheet) {
				columnsInWidestRow = Math.max(columnsInWidestRow, curRow.getLastCellNum());
			}
			
			DataFormatter df = new DataFormatter(true);
			for (Row curRow : sheet) {
				if (curRow != null) {
					boolean firstColumn = true;
					for (short cellNum = 0; cellNum < columnsInWidestRow; ++cellNum) {
						Cell curCell = curRow.getCell(cellNum); //it's safe to call this even for empty cells... we're just passed null back
						if (firstColumn) {
							//for the first column, we don't need to prepend a delimiter
							firstColumn = false;
						}
						else {
							//for all non-first columns, we need to add a delimiter
							writer.write(delimChar);
						}
						
						if (curCell != null) {
							writer.write(df.formatCellValue(curCell));
						}
					}
				}
				///The dom5inspector is expecting unix newlines for the CSV files
				writer.write('\n');
			}
		}
	}
	
	private static char getDelimeterChar(Delimiter delim) {
		switch (delim) {
		case COMMA:
			return ',';
		case TAB:
		default:
			return '\t';
			
		}
	}
	
	private static final String CSV_OUTPUT_DIR_NAME = "csv_output";
	
	
}
