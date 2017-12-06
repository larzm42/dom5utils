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
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WeaponStatIndexer {
	
	private static XSSFWorkbook readFile(String filename) throws IOException {
		return new XSSFWorkbook(new FileInputStream(filename));
	}
	
	private static void doit(XSSFSheet sheet, int skip, int column) throws IOException {
		doit(sheet, skip, column, null);
	}
	
	private static void doit(XSSFSheet sheet, int skip, int column, Callback callback) throws IOException {
		FileInputStream stream = new FileInputStream("Dominions5.exe");			
		stream.skip(Starts.WEAPON);
		int rowNumber = 1;
		int i = 0;
		byte[] c = new byte[2];
		stream.skip(skip);
		while ((stream.read(c, 0, 2)) != -1) {
			XSSFRow row = sheet.getRow(rowNumber);
			rowNumber++;
			XSSFCell cell = row.getCell(column, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
			String high = String.format("%02X", c[1]);
			String low = String.format("%02X", c[0]);
			short value = Integer.decode("0X" + high + low).shortValue();

			if (callback != null) {
				cell.setCellValue(callback.found(Short.toString(value)));
			} else {
				cell.setCellValue(value);
			}

			stream.skip(110l);
			i++;
			if (i >= Starts.WEAPON_COUNT) {
				break;
			}
		}
		stream.close();
	}
	
	private static short doit2(long skip) throws IOException {
		FileInputStream stream = new FileInputStream("Dominions5.exe");			
		short value = 0;
		byte[] c = new byte[2];
		stream.skip(skip);
		while ((stream.read(c, 0, 2)) != -1) {
			String high = String.format("%02X", c[1]);
			String low = String.format("%02X", c[0]);
			 value = Integer.decode("0X" + high + low).shortValue();
			break;
		}
		stream.close();
		return value;
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

	private static long doit4(long skip) throws IOException {
		FileInputStream stream = new FileInputStream("Dominions5.exe");			
		long value = 0;
		byte[] c = new byte[6];
		stream.skip(skip);
		while ((stream.read(c, 0, 6)) != -1) {
			String high2 = String.format("%02X", c[5]);
			String low2 = String.format("%02X", c[4]);
			String high1 = String.format("%02X", c[3]);
			String low1 = String.format("%02X", c[2]);
			String high = String.format("%02X", c[1]);
			String low = String.format("%02X", c[0]);
			
			value = new BigInteger(high2 + low2 + high1 + low1 + high + low, 16).longValue();

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
	        long startIndex = Starts.WEAPON;
	        int ch;

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(startIndex);
			
			XSSFWorkbook wb = WeaponStatIndexer.readFile("BaseW.xlsx");
			
			FileOutputStream fos = new FileOutputStream("NewBaseW.xlsx");
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
				startIndex = startIndex + 112l;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);

				//System.out.println(name);
				
				XSSFRow row = sheet.getRow(rowNumber);
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
			doit(sheet, 48, 3);

			// def
			doit(sheet, 50, 4);

			// len		
			doit(sheet, 56, 5);

			// nratt
			doit(sheet, 58, 6);

			// ammo
			doit(sheet, 60, 7);
			
			// secondaryeffect
			doit(sheet, 72, 8, new CallbackAdapter() {
				@Override
				public String found(String value) {
					if (Integer.parseInt(value) > 0) {
						return value;
					}
					return "0";
				}
				
			});
			
			// secondaryeffectalways
			doit(sheet, 72, 9, new CallbackAdapter() {
				@Override
				public String found(String value) {
					if (Integer.parseInt(value) < 0) {
						return Integer.toString(Math.abs(Integer.valueOf(value)));
					}
					return "0";
				}
				
			});
			
			// rcost
			doit(sheet, 86, 10);
			
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

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(startIndex);
			
			XSSFWorkbook wb = WeaponStatIndexer.readFile("attributes_by_weapon.xlsx");
			
			FileOutputStream fos = new FileOutputStream("Newattributes_by_weapon.xlsx");
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
				
				int attrib = doit3(newIndex);
				while (attrib != 0) {
					attributes.add(new Attributes(attrib, weaponNumber));
					newIndex+=4;					
					attrib = doit3(newIndex);
				}
				weaponNumber++;
				//System.out.println(name);

				stream = new FileInputStream("Dominions5.exe");		
				startIndex = startIndex + 112l;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);
				
			}
			
			int rowNum = 1;
			for (Attributes attribute : attributes) {
				XSSFRow row = sheet.getRow(rowNum);
				XSSFCell cell1 = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell1.setCellValue(attribute.armor_number);
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

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(startIndex);
			
			XSSFWorkbook wb = WeaponStatIndexer.readFile("effects_weapons.xlsx");
			
			FileOutputStream fos = new FileOutputStream("Neweffects_weapons.xlsx");
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
				
				short effect_number = doit2(newIndex);
				if (effect_number != 0) {
					Effect effect = new Effect();
					effect.effect_number = effect_number;
					effect.record_number = weaponNumber;
					effect.object_type = "Weapon";
					effect.ritual = 0;
					effect.modifiers_mask = doit4(startIndex+64);
					effect.raw_argument = doit2(startIndex+40);
					effect.range_base = doit2(startIndex+56);
					effect.area_base = doit2(startIndex+82);
					
					effects.add(effect);

				}
				weaponNumber++;
				System.out.println(name);

				stream = new FileInputStream("Dominions5.exe");		
				startIndex = startIndex + 112l;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);
				
			}
			
			int rowNum = 1;
			for (Effect eff : effects) {
				XSSFRow row = sheet.getRow(rowNum);
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
		int armor_number;
		public Attributes(int attribute, int armor_number) {
			super();
			this.attribute = attribute;
			this.armor_number = armor_number;
		}
		
	}
}
