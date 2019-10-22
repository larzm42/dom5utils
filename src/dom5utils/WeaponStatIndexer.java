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

public class WeaponStatIndexer extends AbstractStatIndexer {
	public static String[] weapons_columns = {"id", "name", "effect_record_id", "att", "def", "len", "nratt", "ammo", "secondaryeffect", "secondaryeffectalways", "rcost", "weapon", "end"};																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																						
	public static String[] effects_weapons_columns = {"record_id", "effect_number", "duration", "ritual", "object_type", "raw_argument", "modifiers_mask", "range_base", "range_per_level", "range_strength_divisor", "area_base", "area_per_level", "area_battlefield_pct", "end"};
	public static String[] attributes_by_weapons_columns = {"weapon_number", "attribute", "raw_value", "end"};

	public static void main(String[] args) {
		run();
	}
	
	public static void run() {
		FileInputStream stream = null;
        List<Weapon> weaponList = new ArrayList<Weapon>();

        try {
			BufferedWriter writer = new BufferedWriter(new FileWriter("weapons.txt"));
	        long startIndex = Starts.WEAPON;
	        int ch;
			stream = new FileInputStream(EXE_NAME);			
			stream.skip(startIndex);
			
			// name
			InputStreamReader isr = new InputStreamReader(stream, "ISO-8859-1");
	        Reader in = new BufferedReader(isr);
	        int weaponNumber = 1;
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
				
				Weapon weapon = new Weapon();
				weapon.parameters = new HashMap<String, Object>();
				weapon.parameters.put("id", weaponNumber);
				weapon.parameters.put("weapon", weaponNumber);
				weapon.parameters.put("name", name.toString());
				weapon.parameters.put("effect_record_id", weaponNumber);
				weapon.parameters.put("att", getBytes2(startIndex + 48));
				weapon.parameters.put("def", getBytes2(startIndex + 50));
				weapon.parameters.put("len", getBytes2(startIndex + 54));
				weapon.parameters.put("nratt", getBytes2(startIndex + 58));
				weapon.parameters.put("ammo", getBytes2(startIndex + 60));
				int bytes2 = getBytes2(startIndex + 72);
				weapon.parameters.put("secondaryeffect", bytes2 > 0 ? bytes2 : 0);
				weapon.parameters.put("secondaryeffectalways", bytes2 < 0 ? Math.abs(bytes2) : 0);
				weapon.parameters.put("rcost", getBytes2(startIndex + 86));
				
		        List<Attribute> attributes = new ArrayList<Attribute>();
				long newIndex = startIndex+88;
				int attrib = getBytes4(newIndex);
				long valueIndex = newIndex + 12l;
				long value = getBytes4(valueIndex);
				while (attrib != 0) {
					attributes.add(new Attribute(weaponNumber, attrib, value));
					newIndex+=4;
					valueIndex+=4;
					attrib = getBytes4(newIndex);
					value = getBytes4(valueIndex);
				}
				weapon.attributes = attributes;

				newIndex = startIndex+52;
				int effect_number = getBytes2(newIndex);
				if (effect_number != 0) {
					Effect effect = new Effect();
					effect.effect_number = effect_number;
					effect.record_number = weaponNumber;
					effect.object_type = "Weapon";
					effect.ritual = 0;
					effect.modifiers_mask = getBytes6(startIndex+64);
					effect.raw_argument = getBytes2(startIndex+40);
					effect.range_base = getBytes2(startIndex+56);
					effect.area_base = getBytes2(startIndex+82);
					weapon.effect = effect;
				}
				weaponList.add(weapon);

				stream = new FileInputStream(EXE_NAME);		
				startIndex = startIndex + Starts.WEAPON_SIZE;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);

		        weaponNumber++;
			}
			in.close();
			stream.close();
			
			XSSFWorkbook wb = new XSSFWorkbook();
			FileOutputStream fos = new FileOutputStream("weapons.xlsx");
			XSSFSheet sheet = wb.createSheet();
			XSSFWorkbook wb2 = new XSSFWorkbook();
			FileOutputStream fos2 = new FileOutputStream("effects_weapons.xlsx");
			XSSFSheet sheet2 = wb2.createSheet();
			XSSFWorkbook wb3 = new XSSFWorkbook();
			FileOutputStream fos3 = new FileOutputStream("attributes_by_weapon.xlsx");
			XSSFSheet sheet3 = wb3.createSheet();

			int rowNum = 0;
			int effectsNum = 0;
			int attributesNum = 0;
			for (Weapon weapon : weaponList) {
				// BaseW
				if (rowNum == 0) {
					XSSFRow row = sheet.createRow(rowNum);
					for (int i = 0; i < weapons_columns.length; i++) {
						row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(weapons_columns[i]);
					}
					rowNum++;
				}
				XSSFRow row = sheet.createRow(rowNum);
				for (int i = 0; i < weapons_columns.length; i++) {
					Object object = weapon.parameters.get(weapons_columns[i]);
					if (object != null) {
						row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(object.toString());
					}
				}
				
				// effects_weapons
				if (weapon.effect != null) {
					if (effectsNum == 0) {
						row = sheet2.createRow(effectsNum);
						for (int i = 0; i < effects_weapons_columns.length; i++) {
							row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(effects_weapons_columns[i]);
						}
						effectsNum++;
					}
					row = sheet2.createRow(effectsNum);
					row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(weapon.effect.record_number);
					row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(weapon.effect.effect_number);
					row.getCell(4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(weapon.effect.object_type);
					row.getCell(5, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(weapon.effect.raw_argument);
					row.getCell(6, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(Long.toString(weapon.effect.modifiers_mask));
					if (weapon.effect.range_base >= 0) {
						row.getCell(7, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(weapon.effect.range_base);
					} else {
						row.getCell(9, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(Math.abs(weapon.effect.range_base));
					}
					row.getCell(10, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(Math.abs(weapon.effect.area_base));
					effectsNum++;
				}
				
				// attributes_by_weapon
				for (Attribute attribute : weapon.attributes) {
					if (attributesNum == 0) {
						row = sheet3.createRow(attributesNum);
						for (int i = 0; i < attributes_by_weapons_columns.length; i++) {
							row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(attributes_by_weapons_columns[i]);
						}
						attributesNum++;
					}
					row = sheet3.createRow(attributesNum);
					row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(attribute.object_number);
					row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(attribute.attribute);
					row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(attribute.raw_value);
					attributesNum++;
				}
				dumpTextFile(weapon, writer);
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
	
	private static void dumpTextFile(Weapon weapon, BufferedWriter writer) throws IOException {
		Object name = weapon.parameters.get("name");
		Object id = weapon.parameters.get("id");
		writer.write(name.toString() + "(" + id + ")");
		writer.newLine();
		for (Map.Entry<String, Object> entry : weapon.parameters.entrySet()) {
			if (!entry.getKey().equals("name") && !entry.getKey().equals("id")) {
				writer.write("\t" + entry.getKey() + ": " + entry.getValue());
				writer.newLine();
			}
		}
		printAttributes(weapon.attributes, writer);
		printEffect(weapon.effect, writer);
		writer.newLine();
	}
		
	private static class Weapon {
		Map<String, Object> parameters;
		Effect effect;
		List<Attribute> attributes;
	}
	
}
