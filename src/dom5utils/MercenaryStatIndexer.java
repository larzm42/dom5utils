package dom5utils;
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
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.Reader;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import dom5utils.CSVWriter.Delimiter;
import dom5utils.CSVWriter.SSType;

public class MercenaryStatIndexer extends AbstractStatIndexer {
	public static String[] mercenary_columns = {"id", "name", "bossname", "com", "unit", "nrunits", "level", "minmen", "minpay", "xp", "randequip", "recrate", "item1", "item2", "eramask", "end"};
	
	public static void main(String[] args) {
		run();
	}
	
	public static void run() {
		FileInputStream stream = null;
        List<Mercenary> mercenaryList = new ArrayList<Mercenary>();

		try {
	        long startIndex = Starts.MERCENARY;
	        int ch;
			stream = new FileInputStream(EXE_NAME);			
			stream.skip(startIndex);
			
			// name
			InputStreamReader isr = new InputStreamReader(stream, "ISO-8859-1");
	        Reader in = new BufferedReader(isr);
	        int rowNumber = 1;
			while ((ch = in.read()) > -1) {
				StringBuffer name = new StringBuffer();
				while (ch != 0) {
					name.append((char)ch);
					ch = in.read();
				}
				if (name.length() == 0) {
					continue;
				}
				if (name.toString().equals("end")) {
					break;
				}
				in.close();
				
				Mercenary merc = new Mercenary();
				merc.parameters = new HashMap<String, Object>();
				merc.parameters.put("id", rowNumber);
				merc.parameters.put("name", name.toString());
				merc.parameters.put("bossname", getString(startIndex + 36l));
				merc.parameters.put("com", getBytes2(startIndex + 74));
				merc.parameters.put("unit", getBytes2(startIndex + 78));
				merc.parameters.put("nrunits", getBytes2(startIndex + 82));
				merc.parameters.put("minmen", getBytes2(startIndex + 86));
				merc.parameters.put("minpay", getBytes2(startIndex + 90));
				merc.parameters.put("xp", getBytes2(startIndex + 94));
				merc.parameters.put("randequip", getBytes2(startIndex + 158));
				merc.parameters.put("recrate", getBytes2(startIndex + 306));
				merc.parameters.put("item1", getString(startIndex + 162l));
				merc.parameters.put("item2", getString(startIndex + 198l));
				merc.parameters.put("eramask", getBytes1(startIndex-2));
				merc.parameters.put("level", getBytes1(startIndex-1));
				
				mercenaryList.add(merc);
				
				stream = new FileInputStream(EXE_NAME);		
				startIndex = startIndex + 312l;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);

				rowNumber++;
			}
			in.close();
			stream.close();
			
			//make sure there's a place to put csv files
			CSVWriter.createCSVOutputDirectory();
			
			XSSFWorkbook wb = new XSSFWorkbook();
			FileOutputStream fos = CSVWriter.getFOS("Mercenary", SSType.XLSX);
			BufferedWriter   csv = CSVWriter.getBFW("Mercenary", SSType.CSV);
			XSSFSheet sheet = wb.createSheet();
			
			int rowNum = 0;
			for (Mercenary merc : mercenaryList) {
				// Mercenary
				if (rowNum == 0) {
					XSSFRow row = sheet.createRow(rowNum);
					for (int i = 0; i < mercenary_columns.length; i++) {
						row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(mercenary_columns[i]);
					}
					rowNum++;
				}
				XSSFRow row = sheet.createRow(rowNum);
				for (int i = 0; i < mercenary_columns.length; i++) {
					Object object = merc.parameters.get(mercenary_columns[i]);
					if (object != null) {
						row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(object.toString());
					}
				}
				rowNum++;
			}
			wb.write(fos);
			fos.close();
			CSVWriter.writeSimpleCSV(sheet, csv, Delimiter.TAB);
			csv.close();
			wb.close();

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (stream != null) {
				try {
					stream.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}
	
	private static class Mercenary {
		Map<String, Object> parameters;
	}

}
