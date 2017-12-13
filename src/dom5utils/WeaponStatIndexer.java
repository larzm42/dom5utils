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
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WeaponStatIndexer extends AbstractStatIndexer {
	
	private static void putBytes2(XSSFSheet sheet, int skip, int column) throws IOException {
		putBytes2(sheet, skip, column, Starts.WEAPON, Starts.WEAPON_SIZE, Starts.WEAPON_COUNT);
	}
	
	private static void putBytes2(XSSFSheet sheet, int skip, int column, Callback callback) throws IOException {
		putBytes2(sheet, skip, column, Starts.WEAPON, Starts.WEAPON_SIZE, Starts.WEAPON_COUNT, callback);
	}
	
	public static void main(String[] args) {
		run();
	}
	
	public static void run() {
		FileInputStream stream = null;
		try {
	        long startIndex = Starts.WEAPON;
	        int ch;

			stream = new FileInputStream(EXE_NAME);			
			stream.skip(startIndex);
			
			XSSFWorkbook wb = WeaponStatIndexer.readFile("BaseW_Template.xlsx");
			
			FileOutputStream fos = new FileOutputStream("BaseW.xlsx");
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

				stream = new FileInputStream(EXE_NAME);		
				startIndex = startIndex + Starts.WEAPON_SIZE;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);

				XSSFRow row = sheet.createRow(rowNumber);
				XSSFCell cell1 = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell1.setCellValue(rowNumber);
				XSSFCell cell = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(name.toString());
				XSSFCell cell2 = row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell2.setCellValue(rowNumber);
				rowNumber++;
			}
			in.close();
			stream.close();

			// att
			putBytes2(sheet, 48, 3);

			// def
			putBytes2(sheet, 50, 4);

			// len		
			putBytes2(sheet, 56, 5);

			// nratt
			putBytes2(sheet, 58, 6);

			// ammo
			putBytes2(sheet, 60, 7);
			
			// secondaryeffect
			putBytes2(sheet, 72, 8, new CallbackAdapter() {
				@Override
				public String found(String value) {
					if (Integer.parseInt(value) > 0) {
						return value;
					}
					return "0";
				}
				
			});
			
			// secondaryeffectalways
			putBytes2(sheet, 72, 9, new CallbackAdapter() {
				@Override
				public String found(String value) {
					if (Integer.parseInt(value) < 0) {
						return Integer.toString(Math.abs(Integer.valueOf(value)));
					}
					return "0";
				}
				
			});
			
			// rcost
			putBytes2(sheet, 86, 10);
			
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
		
		// attributes_by_weapon
		try {
	        long startIndex = Starts.WEAPON;
	        int ch;

			stream = new FileInputStream(EXE_NAME);			
			stream.skip(startIndex);
			
			XSSFWorkbook wb = WeaponStatIndexer.readFile("attributes_by_weapon_Template.xlsx");
			
			FileOutputStream fos = new FileOutputStream("attributes_by_weapon.xlsx");
			XSSFSheet sheet = wb.getSheetAt(0);

			// name
			InputStreamReader isr = new InputStreamReader(stream, "ISO-8859-1");
	        Reader in = new BufferedReader(isr);
	        int weaponNumber = 1;
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
				
				long newIndex = startIndex+88;
				
				int attrib = getBytes4(newIndex);
				while (attrib != 0) {
					attributes.add(new Attributes(attrib, weaponNumber));
					newIndex+=4;					
					attrib = getBytes4(newIndex);
				}
				weaponNumber++;

				stream = new FileInputStream(EXE_NAME);		
				startIndex = startIndex + Starts.WEAPON_SIZE;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);
				
			}
			
			int rowNum = 1;
			for (Attributes attribute : attributes) {
				XSSFRow row = sheet.createRow(rowNum);
				XSSFCell cell1 = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell1.setCellValue(attribute.weapon_number);
				XSSFCell cell2 = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell2.setCellValue(attribute.attribute);
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

		// effects_weapons
		try {
	        long startIndex = Starts.WEAPON;
	        int ch;

			stream = new FileInputStream(EXE_NAME);			
			stream.skip(startIndex);
			
			XSSFWorkbook wb = WeaponStatIndexer.readFile("effects_weapons_Template.xlsx");
			
			FileOutputStream fos = new FileOutputStream("effects_weapons.xlsx");
			XSSFSheet sheet = wb.getSheetAt(0);

			// name
			InputStreamReader isr = new InputStreamReader(stream, "ISO-8859-1");
	        Reader in = new BufferedReader(isr);
	        int weaponNumber = 1;
	        List<Effect> effects = new ArrayList<Effect>();
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
				
				long newIndex = startIndex+52;
				
				short effect_number = getBytes2(newIndex);
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
					
					effects.add(effect);

				}
				weaponNumber++;

				stream = new FileInputStream(EXE_NAME);		
				startIndex = startIndex + Starts.WEAPON_SIZE;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);
				
			}
			
			int rowNum = 1;
			for (Effect eff : effects) {
				XSSFRow row = sheet.createRow(rowNum);
				XSSFCell cell1 = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell1.setCellValue(eff.record_number);
				XSSFCell cell2 = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell2.setCellValue(eff.effect_number);
				XSSFCell cell3 = row.getCell(4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell3.setCellValue(eff.object_type);
				XSSFCell cell4 = row.getCell(5, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell4.setCellValue(eff.raw_argument);
				XSSFCell cell5 = row.getCell(6, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell5.setCellValue(Long.toString(eff.modifiers_mask));
				if (eff.range_base >= 0) {
					cell5 = row.getCell(7, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					cell5.setCellValue(eff.range_base);
				} else {
					cell5 = row.getCell(9, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					cell5.setCellValue(Math.abs(eff.range_base));
				}
				cell5 = row.getCell(10, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell5.setCellValue(Math.abs(eff.area_base));
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
		
	}	
	
	private static class Effect {
		int record_number;
		int effect_number;
		int duration;
		int ritual;
		String object_type;
		int raw_argument;
		long modifiers_mask;
		int range_base;
		int range_per_level;
		int range_strength_divisor;
		int area_base;
		int area_per_level;
		int area_battlefield_pct;
		int sound_number;
		int flight_sprite_number;
		int flight_sprite_length;
		int explosion_sprite_number;
		int explosion_sprite_length;

	}
	
	private static class Attributes {
		int attribute;
		int weapon_number;
		public Attributes(int attribute, int weapon_number) {
			super();
			this.attribute = attribute;
			this.weapon_number = weapon_number;
		}
	}
	
}
