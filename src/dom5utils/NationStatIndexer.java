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
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.Reader;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class NationStatIndexer {
	
	enum unitType {
		PRETENDER(2, "pretender_types_by_nation.xlsx"),
		UNPRETENDER(1, "unpretender_types_by_nation.xlsx"),
		TROOP(0, "fort_troop_types_by_nation.xlsx"),
		UNKNOWN(-1, ""),
		LEADER(-2, "fort_leader_types_by_nation.xlsx"),
		NONFORT_TROOP(-3, "nonfort_troop_types_by_nation.xlsx"),
		NONFORT_LEADER(-4, "nonfort_leader_types_by_nation.xlsx"),
		COAST_TROOP(-5, "coast_troop_types_by_nation.xlsx"),
		COAST_LEADER(-6, "coast_leader_types_by_nation.xlsx");
		
		private int id;
		private String filename;
		
		unitType(int i, String filename) {
			id = i;
			this.filename = filename;
		}
		public static unitType fromValue(int id) {
	        for (unitType aip: values()) {
	            if (aip.getId() == id) {
	                return aip;
	            }
	        }
	        return null;
	    }
		public int getId() { return id;}
		public String getFilename() { return filename;}
	}
	
	private static XSSFWorkbook readFile(String filename) throws IOException {
		return new XSSFWorkbook(new FileInputStream(filename));
	}
	
	private static int doit3(long skip) throws IOException {
		FileInputStream stream = new FileInputStream("Dominions5.exe");			
		int value = 0;
		byte[] c = new byte[4];
		stream.skip(skip);
		while ((stream.read(c, 0, 4)) != -1) {
			String high1 = String.format("%02X", c[3]);
			String low1 = String.format("%02X", c[2]);
			String high = String.format("%02X", c[1]);
			String low = String.format("%02X", c[0]);
			
			value = new BigInteger(high1 + low1 + high + low, 16).intValue();

			break;
		}
		stream.close();
		return value;
	}
	
	public static void main(String[] args) {
		run();
	}
	
	public static void run() {
		FileInputStream stream = null;
		try {
	        long startIndex = Starts.NATION;
	        int ch;

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(startIndex);
			
			XSSFWorkbook wb = NationStatIndexer.readFile("BaseN.xlsx");
			
			FileOutputStream fos = new FileOutputStream("NewBaseN.xlsx");
			XSSFSheet sheet = wb.getSheetAt(0);

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

				stream = new FileInputStream("Dominions5.exe");		
				startIndex = startIndex + 1752l;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);

				//System.out.println(name);
				
				XSSFRow row = sheet.getRow(rowNumber);
				XSSFCell cell1 = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell1.setCellValue(rowNumber-1);
				rowNumber++;
				XSSFCell cell = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(name.toString());
			}
			in.close();
			stream.close();

			// epithet
	        startIndex = Starts.NATION;
			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(startIndex + 36l);
			isr = new InputStreamReader(stream, "ISO-8859-1");
	        in = new BufferedReader(isr);
	        rowNumber = 1;
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

				stream = new FileInputStream("Dominions5.exe");		
				startIndex = startIndex + 1752l;
				stream.skip(startIndex + 36l);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);

				//System.out.println(name);
				
				XSSFRow row = sheet.getRow(rowNumber);
				XSSFCell cell = row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(name.toString());
				rowNumber++;
			}
			in.close();
			stream.close();
			
			// abbreviation
	        startIndex = Starts.NATION;
			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(startIndex + 72l);
			isr = new InputStreamReader(stream, "ISO-8859-1");
	        in = new BufferedReader(isr);
	        rowNumber = 1;
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

				stream = new FileInputStream("Dominions5.exe");		
				startIndex = startIndex + 1752l;
				stream.skip(startIndex + 72l);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);

				//System.out.println(name);
				
				XSSFRow row = sheet.getRow(rowNumber);
				XSSFCell cell = row.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(name.toString());
				rowNumber++;
			}
			in.close();
			stream.close();

			// file_name_base
	        startIndex = Starts.NATION;
			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(startIndex + 77l);
			isr = new InputStreamReader(stream, "ISO-8859-1");
	        in = new BufferedReader(isr);
	        rowNumber = 1;
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

				stream = new FileInputStream("Dominions5.exe");		
				startIndex = startIndex + 1752l;
				stream.skip(startIndex + 77l);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);

				//System.out.println(name);
				
				XSSFRow row = sheet.getRow(rowNumber);
				XSSFCell cell = row.getCell(4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(name.toString());
				rowNumber++;
			}
			in.close();
			stream.close();
			
			wb.write(fos);
			fos.close();

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
				
		// attributes_by_nation
		try {
	        long startIndex = Starts.NATION;
	        int ch;

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(startIndex);
			
			XSSFWorkbook wb = NationStatIndexer.readFile("attributes_by_nation.xlsx");
			
			FileOutputStream fos = new FileOutputStream("Newattributes_by_nation.xlsx");
			XSSFSheet sheet = wb.getSheetAt(0);

			// name
			InputStreamReader isr = new InputStreamReader(stream, "ISO-8859-1");
	        Reader in = new BufferedReader(isr);
	        int rowNumber = 1;
	        List<Attributes> attributes = new ArrayList<Attributes>();
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
				
				long newIndex = startIndex+172l;
				
				int attrib = doit3(newIndex);
				long valueIndex = newIndex + 388l;
				long value = doit3(valueIndex);
				while (attrib != 0) {
					attributes.add(new Attributes(rowNumber-1, attrib, value));
					newIndex+=4;
					valueIndex+=8;
					attrib = doit3(newIndex);
					value = doit3(valueIndex);
				}
				rowNumber++;
				//System.out.println(name);

				stream = new FileInputStream("Dominions5.exe");		
				startIndex = startIndex + 1752l;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);
				
			}
			
			int rowNum = 1;
			for (Attributes attribute : attributes) {
				XSSFRow row = sheet.getRow(rowNum);
				XSSFCell cell1 = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell1.setCellValue(attribute.nation_number);
				XSSFCell cell2 = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell2.setCellValue(attribute.attribute);
				cell2 = row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell2.setCellValue(attribute.raw_value);
				rowNum++;
			}

			in.close();
			stream.close();

			wb.write(fos);
			fos.close();

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
		
		// *_types_by_nation
		try {
	        long startIndex = Starts.NATION;
	        int ch;

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(startIndex);
			
			// name
			InputStreamReader isr = new InputStreamReader(stream, "ISO-8859-1");
	        Reader in = new BufferedReader(isr);
	        int rowNumber = 1;
	        
	        Map<unitType, List<Troops>> unitMap = new HashMap<unitType, List<Troops>>();
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
				
				long newIndex = startIndex+1328l;
				
				unitType type = unitType.TROOP;
				if (unitMap.get(type) == null) {
					unitMap.put(type, new ArrayList<Troops>());
				}
				
				int attrib = doit3(newIndex);
				while (attrib != 0) {
					if (attrib < 0) {
						if (attrib != -1) {
							type = unitType.fromValue(attrib);
							if (unitMap.get(type) == null) {
								unitMap.put(type, new ArrayList<Troops>());
							}
						}
					} else {
						unitMap.get(type).add(new Troops(rowNumber-1, attrib));
					}
					newIndex+=4;
					attrib = doit3(newIndex);
				}
				rowNumber++;
				//System.out.println(name);

				stream = new FileInputStream("Dominions5.exe");		
				startIndex = startIndex + 1752l;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);
				
			}
			
			for (Map.Entry<unitType, List<Troops>> entry : unitMap.entrySet()) {
				if (entry.getKey() == unitType.UNKNOWN) { continue; }
				XSSFWorkbook wb = NationStatIndexer.readFile(entry.getKey().getFilename());
				FileOutputStream fos = new FileOutputStream("New" + entry.getKey().getFilename());
				XSSFSheet sheet = wb.getSheetAt(0);
				int rowNum = 1;
				for (Troops troop : entry.getValue()) {
					XSSFRow row = sheet.getRow(rowNum);
					XSSFCell cell1 = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					cell1.setCellValue(troop.monster_number);
					XSSFCell cell2 = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					cell2.setCellValue(troop.nation_number);
					rowNum++;
				}
				wb.write(fos);
				fos.close();
			}

			in.close();
			stream.close();

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

		// un/pretender_types_by_nation
		try {
	        long startIndex = Starts.NATION;
	        int ch;

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(startIndex);
			
			// name
			InputStreamReader isr = new InputStreamReader(stream, "ISO-8859-1");
	        Reader in = new BufferedReader(isr);
	        int rowNumber = 1;
	        
	        Map<unitType, List<Troops>> unitMap = new HashMap<unitType, List<Troops>>();
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
				
				long newIndex = startIndex+1496l;
				
				int attrib = doit3(newIndex);
				while (attrib != 0) {
					if (attrib < 0) {
						if (unitMap.get(unitType.UNPRETENDER) == null) {
							unitMap.put(unitType.UNPRETENDER, new ArrayList<Troops>());
						}
						unitMap.get(unitType.UNPRETENDER).add(new Troops(rowNumber-1, Math.abs(attrib)));
					} else {
						if (unitMap.get(unitType.PRETENDER) == null) {
							unitMap.put(unitType.PRETENDER, new ArrayList<Troops>());
						}
						unitMap.get(unitType.PRETENDER).add(new Troops(rowNumber-1, attrib));
					}
					newIndex+=4;
					attrib = doit3(newIndex);
				}
				rowNumber++;
				//System.out.println(name);

				stream = new FileInputStream("Dominions5.exe");		
				startIndex = startIndex + 1752l;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);
				
			}
			
			for (Map.Entry<unitType, List<Troops>> entry : unitMap.entrySet()) {
				if (entry.getKey() == unitType.UNKNOWN) { continue; }
				XSSFWorkbook wb = NationStatIndexer.readFile(entry.getKey().getFilename());
				FileOutputStream fos = new FileOutputStream("New" + entry.getKey().getFilename());
				XSSFSheet sheet = wb.getSheetAt(0);
				int rowNum = 1;
				for (Troops troop : entry.getValue()) {
					XSSFRow row = sheet.getRow(rowNum);
					XSSFCell cell1 = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					cell1.setCellValue(troop.monster_number);
					XSSFCell cell2 = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					cell2.setCellValue(troop.nation_number);
					rowNum++;
				}
				wb.write(fos);
				fos.close();
			}

			in.close();
			stream.close();

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
	
	private static class Attributes {
		int nation_number;
		int attribute;
		long raw_value;

		public Attributes(int nation_number, int attribute_number, long raw_value) {
			super();
			this.nation_number = nation_number;
			this.attribute = attribute_number;
			this.raw_value = raw_value;
		}
		
	}

	private static class Troops {
		int nation_number;
		int monster_number;

		public Troops(int nation_number, int monster_number) {
			super();
			this.nation_number = nation_number;
			this.monster_number = monster_number;
		}
		
	}
}
