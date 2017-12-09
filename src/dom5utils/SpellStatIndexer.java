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


public class SpellStatIndexer {
	
	
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
	
	private static short doit1(long skip) throws IOException {
		FileInputStream stream = new FileInputStream("Dominions5.exe");			
		short value = 0;
		byte[] c = new byte[1];
		stream.skip(skip);
		while ((stream.read(c, 0, 1)) != -1) {
			//String high = String.format("%02X", c[1]);
			String low = String.format("%02X", c[0]);
			value = new BigInteger(low, 16).shortValue();
			if (value > 100) {
				value = new BigInteger("FF" + low, 16).shortValue();
			}
			break;
		}
		stream.close();
		return value;
	}
	
	private static int doit2(long skip) throws IOException {
		FileInputStream stream = new FileInputStream("Dominions5.exe");			
		int value = 0;
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
	        long startIndex = Starts.SPELL;
	        int ch;

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(startIndex);
			
			XSSFWorkbook wb = SpellStatIndexer.readFile("BaseS.xlsx");
			
			FileOutputStream fos = new FileOutputStream("NewBaseS.xlsx");
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

				XSSFRow row = sheet.getRow(rowNumber);
				XSSFCell cell1 = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell1.setCellValue(rowNumber-1);
				XSSFCell cell = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(name.toString());

				cell = row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(doit1(startIndex + 36l));
				cell = row.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(doit1(startIndex + 37l));
				cell = row.getCell(4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(doit1(startIndex + 38l));
				cell = row.getCell(5, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(doit1(startIndex + 40l));
				cell = row.getCell(6, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(doit1(startIndex + 39l));
				cell = row.getCell(7, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(doit1(startIndex + 41l));
				cell = row.getCell(8, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(rowNumber-1);

				cell = row.getCell(9, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(doit2(startIndex + 64l));
				cell = row.getCell(10, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(doit1(startIndex + 50l));
				cell = row.getCell(11, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				int fat = doit2(startIndex + 42l);
				cell.setCellValue(fat%100);
				cell = row.getCell(12, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(fat/100);
				cell = row.getCell(13, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(doit1(startIndex + 88l));

				rowNumber++;

				stream = new FileInputStream("Dominions5.exe");		
				startIndex = startIndex + 216l;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);

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
				
		// effects_spells
		try {
	        long startIndex = Starts.SPELL;
	        int ch;

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(startIndex);
			
			XSSFWorkbook wb = SpellStatIndexer.readFile("effects_spells.xlsx");
			
			FileOutputStream fos = new FileOutputStream("Neweffects_spells.xlsx");
			XSSFSheet sheet = wb.getSheetAt(0);

			// name
			InputStreamReader isr = new InputStreamReader(stream, "ISO-8859-1");
	        Reader in = new BufferedReader(isr);
	        int spellNumber = 0;
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
				
				
				Effect effect = new Effect();
				int effect_number = doit2(startIndex+46);
				if (effect_number > 10000) {
					effect.ritual = 1;
					//effect_number = effect_number % 1000;
				} else if (effect_number > 1000) {
					effect.duration = effect_number / 1000;
				}
				effect_number = effect_number % 1000;
				effect.effect_number = effect_number;
				effect.record_number = spellNumber;
				effect.object_type = "Spell";
				effect.modifiers_mask = doit4(startIndex+80);
				effect.raw_argument = doit2(startIndex+56);
				int raw_range = doit2(startIndex+48);
				if (raw_range > 0) {
					effect.range_base = raw_range % 1000;
					effect.range_per_level = raw_range / 1000;
				} else {
					effect.range_strength_divisor = -raw_range;
				}
				int raw_area = doit2(startIndex+44);
				switch (raw_area) {
				case 666:
					effect.area_battlefield_pct = 100;
					break;
				case 663:
					effect.area_battlefield_pct = 50;
					break;
				case 665:
					effect.area_battlefield_pct = 25;
					break;
				case 664:
					effect.area_battlefield_pct = 10;
					break;
				case 662:
					effect.area_battlefield_pct = 5;
					break;
				default:
					effect.area_base = raw_area % 1000;
					effect.area_per_level = raw_area / 1000;
					break;
				}

				effects.add(effect);

				spellNumber++;
//				System.out.println(name);

				stream = new FileInputStream("Dominions5.exe");		
				startIndex = startIndex + 216l;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);
				
			}
			
			int rowNum = 1;
			for (Effect eff : effects) {
				XSSFRow row = sheet.getRow(rowNum);
				XSSFCell cell = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(eff.record_number);
				cell = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(eff.effect_number);
				cell = row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(eff.duration != 0 ? eff.duration+"" : "");
				cell = row.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(eff.ritual);
				cell = row.getCell(4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(eff.object_type);
				cell = row.getCell(5, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(eff.raw_argument);
				cell = row.getCell(6, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(Long.toString(eff.modifiers_mask));
				cell = row.getCell(7, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(eff.range_base);
				cell = row.getCell(8, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(eff.range_per_level);
				cell = row.getCell(9, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(eff.range_strength_divisor);
				cell = row.getCell(10, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(eff.area_base);
				cell = row.getCell(11, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(eff.area_per_level);
				cell = row.getCell(12, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(eff.area_battlefield_pct != 0 ? eff.area_battlefield_pct+"" : "");
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

		// attributes_by_nation
		try {
	        long startIndex = Starts.SPELL;
	        int ch;

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(startIndex);
			
			XSSFWorkbook wb = SpellStatIndexer.readFile("attributes_by_spell.xlsx");
			
			FileOutputStream fos = new FileOutputStream("Newattributes_by_spell.xlsx");
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
				
				long newIndex = startIndex+92l;
				
				int attrib = doit3(newIndex);
				long valueIndex = newIndex + 60l;
				long value = doit3(valueIndex);
				while (attrib != 0) {
					attributes.add(new Attributes(rowNumber-1, attrib, value));
					newIndex+=4;
					valueIndex+=4;
					attrib = doit3(newIndex);
					value = doit3(valueIndex);
				}
				rowNumber++;
				//System.out.println(name);

				stream = new FileInputStream("Dominions5.exe");		
				startIndex = startIndex + 216l;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);
				
			}
			
			int rowNum = 1;
			for (Attributes attribute : attributes) {
				XSSFRow row = sheet.getRow(rowNum);
				XSSFCell cell1 = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell1.setCellValue(attribute.spell_number);
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
		/*try {
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
				XSSFWorkbook wb = SpellStatIndexer.readFile(entry.getKey().getFilename());
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
				XSSFWorkbook wb = SpellStatIndexer.readFile(entry.getKey().getFilename());
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
		}*/
	}	
	
	private static class Attributes {
		int spell_number;
		int attribute;
		long raw_value;

		public Attributes(int spell_number, int attribute_number, long raw_value) {
			super();
			this.spell_number = spell_number;
			this.attribute = attribute_number;
			this.raw_value = raw_value;
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
}
