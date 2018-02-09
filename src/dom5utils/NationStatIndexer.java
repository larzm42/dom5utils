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
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class NationStatIndexer extends AbstractStatIndexer {
	public static String[] nations_columns = {"id", "name", "epithet", "abbreviation", "file_name_base", "era", "end"};
	public static String[] attributes_by_nation_columns = {"nation_number", "attribute", "raw_value", "end"};
	public static String[] troops_by_nation_columns = {"monster_number", "nation_number", "end"};

	enum unitType {
		PRETENDER(2, "pretender_types_by_nation"),
		UNPRETENDER(1, "unpretender_types_by_nation"),
		TROOP(0, "fort_troop_types_by_nation"),
		UNKNOWN(-1, ""),
		LEADER(-2, "fort_leader_types_by_nation"),
		NONFORT_TROOP(-3, "nonfort_troop_types_by_nation"),
		NONFORT_LEADER(-4, "nonfort_leader_types_by_nation"),
		COAST_TROOP(-5, "coast_troop_types_by_nation"),
		COAST_LEADER(-6, "coast_leader_types_by_nation");
		
		private int id;
		private String filename;
		private int rowNum = 0;
		XSSFWorkbook wb;
		FileOutputStream fos;
		XSSFSheet sheet;
		
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
		public int getRowNum() { return rowNum; }
		public void incrementRowNum() { rowNum++; }
	}
	
	public static void main(String[] args) {
		run();
	}
	
	public static void run() {
		FileInputStream stream = null;
        List<Nation> nationList = new ArrayList<Nation>();

        try {
	        long startIndex = Starts.NATION;
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

				Nation nation = new Nation();
				nation.parameters = new HashMap<String, Object>();
				nation.parameters.put("id", rowNumber-1);
				nation.parameters.put("name", name.toString());
				nation.parameters.put("epithet", getString(startIndex + 36l));
				nation.parameters.put("abbreviation", getString(startIndex + 72l));
				String fileNameBase = getString(startIndex + 77l);
				nation.parameters.put("file_name_base", fileNameBase);
				nation.parameters.put("era", getBytes2(startIndex + 168l));
				
		        List<Attribute> attributes = new ArrayList<Attribute>();
				long newIndex = startIndex+172l;
				
				int attrib = getBytes4(newIndex);
				long valueIndex = newIndex + 388l;
				long value = getBytes4(valueIndex);
				while (attrib != 0) {
					attributes.add(new Attribute(rowNumber-1, attrib, value));
					newIndex+=4;
					valueIndex+=8;
					attrib = getBytes4(newIndex);
					value = getBytes4(valueIndex);
				}
				nation.attributes = attributes;
				
		        Map<unitType, List<Troops>> unitMap = new HashMap<unitType, List<Troops>>();
				newIndex = startIndex+1328l;
				unitType type = unitType.TROOP;
				if (unitMap.get(type) == null) {
					unitMap.put(type, new ArrayList<Troops>());
				}
				
				attrib = getBytes4(newIndex);
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
					attrib = getBytes4(newIndex);
				}
				
				newIndex = startIndex+1496l;
				attrib = getBytes4(newIndex);
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
					attrib = getBytes4(newIndex);
				}

				nation.unitMap = unitMap;
				
				nationList.add(nation);

				stream = new FileInputStream(EXE_NAME);		
				startIndex = startIndex + Starts.NATION_SIZE;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);
		        
				rowNumber++;
			}
			in.close();
			stream.close();
			
			XSSFWorkbook wb = new XSSFWorkbook();
			FileOutputStream fos = new FileOutputStream("nations.xlsx");
			XSSFSheet sheet = wb.createSheet();
			XSSFWorkbook wb2 = new XSSFWorkbook();
			FileOutputStream fos2 = new FileOutputStream("attributes_by_nation.xlsx");
			XSSFSheet sheet2 = wb2.createSheet();
			
			int rowNum = 0;
			int attributesNum = 0;
			for (Nation nation : nationList) {
				// nations
				if (rowNum == 0) {
					XSSFRow row = sheet.createRow(rowNum);
					for (int i = 0; i < nations_columns.length; i++) {
						row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(nations_columns[i]);
					}
					rowNum++;
				}
				XSSFRow row = sheet.createRow(rowNum);
				for (int i = 0; i < nations_columns.length; i++) {
					Object object = nation.parameters.get(nations_columns[i]);
					if (object != null) {
						row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(object.toString());
					}
				}
				
				// attributes_by_nation
				for (Attribute attribute : nation.attributes) {
					if (attributesNum == 0) {
						row = sheet2.createRow(attributesNum);
						for (int i = 0; i < attributes_by_nation_columns.length; i++) {
							row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(attributes_by_nation_columns[i]);
						}
						attributesNum++;
					}
					row = sheet2.createRow(attributesNum);
					row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(attribute.object_number);
					row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(attribute.attribute);
					row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(attribute.raw_value);
					attributesNum++;
				}
				
				for (Map.Entry<unitType, List<Troops>> entry : nation.unitMap.entrySet()) {
					if (entry.getKey() == unitType.UNKNOWN) { continue; }
					if (entry.getKey().getRowNum() == 0) {
						entry.getKey().wb = new XSSFWorkbook();
						entry.getKey().fos = new FileOutputStream(entry.getKey().getFilename() + ".xlsx");
						entry.getKey().sheet = entry.getKey().wb.createSheet();
						row = entry.getKey().sheet.createRow(0);
						for (int i = 0; i < troops_by_nation_columns.length; i++) {
							row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(troops_by_nation_columns[i]);
						}
						entry.getKey().incrementRowNum();
					}
					for (Troops troop : entry.getValue()) {
						row = entry.getKey().sheet.createRow(entry.getKey().getRowNum());
						row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(troop.monster_number);
						row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(troop.nation_number);
						entry.getKey().incrementRowNum();
					}
				}
				
				rowNum++;
			}
			wb.write(fos);
			fos.close();
			wb.close();

			wb2.write(fos2);
			fos2.close();
			wb2.close();
			
			for (Nation nation : nationList) {
				for (Map.Entry<unitType, List<Troops>> entry : nation.unitMap.entrySet()) {
					if (entry.getKey().rowNum != -1) {
						entry.getKey().wb.write(entry.getKey().fos);
						entry.getKey().fos.close();
						entry.getKey().wb.close();
						entry.getKey().rowNum = -1;
					}
				}
			}

			
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
	
	private static class Troops {
		int nation_number;
		int monster_number;

		public Troops(int nation_number, int monster_number) {
			super();
			this.nation_number = nation_number;
			this.monster_number = monster_number;
		}
		
	}
	
	private static class Nation {
		Map<String, Object> parameters;
		List<Attribute> attributes;
		Map<unitType, List<Troops>> unitMap;
	}

}
