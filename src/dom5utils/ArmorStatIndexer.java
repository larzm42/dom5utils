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
import java.io.FileWriter;
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

public class ArmorStatIndexer extends AbstractStatIndexer {
	public static String[] armors_columns = {"id", "name", "type", "def", "enc", "rcost", "end"};
	public static String[] attributes_by_armor_columns = {"armor_number", "attribute", "raw_value", "end"};
	public static String[] protections_by_armor_columns = {"zone_number", "protection", "armor_number", "end"};

	public static void main(String[] args) {
		run();
	}
	
	public static void run() {
		FileInputStream stream = null;
        List<Armor> armorList = new ArrayList<Armor>();

		try {
			BufferedWriter writer = new BufferedWriter(new FileWriter("armors.txt"));
	        long startIndex = Starts.ARMOR;
	        int ch;
			stream = new FileInputStream(EXE_NAME);			
			stream.skip(startIndex);
			
			// name
			InputStreamReader isr = new InputStreamReader(stream, "ISO-8859-1");
	        Reader in = new BufferedReader(isr);
	        int armorNumber = 1;
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

				Armor armor = new Armor();
				armor.parameters = new HashMap<String, Object>();
				armor.parameters.put("id", armorNumber);
				armor.parameters.put("name", name.toString());
				armor.parameters.put("def", getBytes2(startIndex + 62));
				armor.parameters.put("enc", getBytes2(startIndex + 64));
				armor.parameters.put("type", getBytes2(startIndex + 66));
				armor.parameters.put("rcost", getBytes2(startIndex + 68));
				
		        List<Protection> protections = new ArrayList<Protection>();
				long newIndex = startIndex+36;
				int zone = getBytes2(newIndex);
				while (zone != 0) {
					newIndex+=2;
					int prot = getBytes2(newIndex);
					protections.add(new Protection(zone, prot, armorNumber));
					newIndex+=2;					
					zone = getBytes2(newIndex);
				}
				armor.protections = protections;
				
		        List<Attribute> attributes = new ArrayList<Attribute>();
				newIndex = startIndex+72;
				int attrib = getBytes4(newIndex);
				long valueIndex = newIndex + 16l;
				long value = getBytes4(valueIndex);
				while (attrib != 0) {
					attributes.add(new Attribute(armorNumber, attrib, value));
					newIndex+=4;
					valueIndex+=4;
					attrib = getBytes4(newIndex);
					value = getBytes4(valueIndex);
				}
				armor.attributes = attributes;
				
				armorList.add(armor);

				stream = new FileInputStream(EXE_NAME);		
				startIndex = startIndex + Starts.ARMOR_SIZE;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);

		        armorNumber++;
			}
			in.close();
			stream.close();
			
			XSSFWorkbook wb = new XSSFWorkbook();
			FileOutputStream fos = new FileOutputStream("armors.xlsx");
			XSSFSheet sheet = wb.createSheet();
			XSSFWorkbook wb2 = new XSSFWorkbook();
			FileOutputStream fos2 = new FileOutputStream("protections_by_armor.xlsx");
			XSSFSheet sheet2 = wb2.createSheet();
			XSSFWorkbook wb3 = new XSSFWorkbook();
			FileOutputStream fos3 = new FileOutputStream("attributes_by_armor.xlsx");
			XSSFSheet sheet3 = wb3.createSheet();

			int rowNum = 0;
			int protectionsNum = 0;
			int attributesNum = 0;
			for (Armor armor : armorList) {
				// BaseA
				if (rowNum == 0) {
					XSSFRow row = sheet.createRow(rowNum);
					for (int i = 0; i < armors_columns.length; i++) {
						row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(armors_columns[i]);
					}
					rowNum++;
				}
				XSSFRow row = sheet.createRow(rowNum);
				for (int i = 0; i < armors_columns.length; i++) {
					Object object = armor.parameters.get(armors_columns[i]);
					if (object != null) {
						row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(object.toString());
					}
				}
				
				
				// protections_by_armor
				for (Protection prot : armor.protections) {
					if (protectionsNum == 0) {
						row = sheet2.createRow(protectionsNum);
						for (int i = 0; i < protections_by_armor_columns.length; i++) {
							row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(protections_by_armor_columns[i]);
						}
						protectionsNum++;
					}
					row = sheet2.createRow(protectionsNum);
					row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(prot.zone_number);
					row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(prot.protection);
					row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(prot.armor_number);
					protectionsNum++;
				}
				
				// attributes_by_armor
				for (Attribute attribute : armor.attributes) {
					if (attributesNum == 0) {
						row = sheet3.createRow(attributesNum);
						for (int i = 0; i < attributes_by_armor_columns.length; i++) {
							row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(attributes_by_armor_columns[i]);
						}
						attributesNum++;
					}
					row = sheet3.createRow(attributesNum);
					row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(attribute.object_number);
					row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(attribute.attribute);
					row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(attribute.raw_value);
					attributesNum++;
				}
				dumpTextFile(armor, writer);
				rowNum++;
			}
			
			wb.write(fos);
			fos.close();
			wb.close();

			wb2.write(fos2);
			fos2.close();
			wb2.close();

			wb3.write(fos3);
			fos3.close();
			wb3.close();

			writer.close();

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
	
	private static void dumpTextFile(Armor armor, BufferedWriter writer) throws IOException {
		Object name = armor.parameters.get("name");
		Object id = armor.parameters.get("id");
		writer.write(name.toString() + "(" + id + ")");
		writer.newLine();
		for (Map.Entry<String, Object> entry : armor.parameters.entrySet()) {
			if (!entry.getKey().equals("name") && !entry.getKey().equals("id")) {
				writer.write("\t" + entry.getKey() + ": " + entry.getValue());
				writer.newLine();
			}
		}
		printAttributes(armor.attributes, writer);
		for (Protection prot : armor.protections) {
			writer.write("\tZone " + prot.zone_number + ": " + prot.protection);
			writer.newLine();
		}
		writer.newLine();
	}
	
	private static class Protection {
		int zone_number;
		int protection;
		int armor_number;
		public Protection(int zone_number, int protection, int armor_number) {
			super();
			this.zone_number = zone_number;
			this.protection = protection;
			this.armor_number = armor_number;
		}
	}
	
	private static class Armor {
		Map<String, Object> parameters;
		List<Protection> protections;
		List<Attribute> attributes;
	}
}
