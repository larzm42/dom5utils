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
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import dom5utils.CSVWriter.Delimiter;
import dom5utils.CSVWriter.SSType;

public class SpellStatIndexer extends AbstractStatIndexer {
	public static String[] spells_columns = {"id","name", "school", "researchlevel", "path1", "pathlevel1", "path2", "pathlevel2", "effect_record_id", "effects_count", "precision", "fatiguecost", "gemcost", "next_spell", "damage", "end"};																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																						
	public static String[] effects_spells_columns = {"record_id", "effect_number", "duration", "ritual", "object_type", "raw_argument", "modifiers_mask", "range_base", "range_per_level", "range_strength_divisor", "area_base", "area_per_level", "area_battlefield_pct", "end"};
	public static String[] attributes_by_spell_columns = {"spell_number", "attribute", "raw_value", "end"};

	public static void main(String[] args) {
		run();
	}
	
	public static void run() {
		FileInputStream stream = null;
        List<Spell> spellList = new ArrayList<Spell>();

		try {
			BufferedWriter writer = new BufferedWriter(new FileWriter("spells.txt"));
	        long startIndex = Starts.SPELL;
	        int ch;
			stream = new FileInputStream(EXE_NAME);			
			stream.skip(startIndex);
			
			// name
			InputStreamReader isr = new InputStreamReader(stream, "ISO-8859-1");
	        Reader in = new BufferedReader(isr);
	        int spellNumber = 0;
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
				
				Spell spell = new Spell();
				spell.parameters = new TreeMap<String, Object>();
				spell.parameters.put("id", spellNumber);
				spell.parameters.put("effect_record_id", spellNumber);
				spell.parameters.put("name", name.toString());
				spell.parameters.put("school", getBytes1(startIndex + 36l));
				spell.parameters.put("researchlevel", getBytes1(startIndex + 37l));
				spell.parameters.put("path1", getBytes1(startIndex + 38l));
				spell.parameters.put("path2", getBytes1(startIndex + 39l));
				spell.parameters.put("pathlevel1", getBytes1(startIndex + 40l));
				spell.parameters.put("pathlevel2", getBytes1(startIndex + 41l));
				int fat = getBytes2(startIndex + 42l);
				spell.parameters.put("fatiguecost", fat%100);
				spell.parameters.put("gemcost", fat/100);
				
				Effect effect = new Effect();
				int raw_area = getBytes2(startIndex+44);
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
				
				int effect_number = getBytes2(startIndex+46);
				if (effect_number > 10000) {
					effect.ritual = 1;
				} else if (effect_number > 1000) {
					effect.duration = effect_number / 1000;
				}
				effect_number = effect_number % 1000;
				effect.effect_number = effect_number;

				int raw_range = getBytes2(startIndex+48);
				if (raw_range > 0) {
					effect.range_base = raw_range % 1000;
					effect.range_per_level = raw_range / 1000;
				} else {
					effect.range_strength_divisor = -raw_range;
				}

				spell.parameters.put("precision", getBytes1(startIndex + 50l));
				spell.parameters.put("damage", getBytes8(startIndex + 56l));

//				spell.parameters.put("Unknown51", getBytes1(startIndex + 51l));
//				spell.parameters.put("Unknown52", getBytes2(startIndex + 52l));
//				spell.parameters.put("Unknown54", getBytes2(startIndex + 54l));
				
				effect.raw_argument = getBytes8(startIndex+56);

//				spell.parameters.put("Unknown66", getBytes2(startIndex + 66l));
//				spell.parameters.put("Unknown68", getBytes2(startIndex + 68l));
//				spell.parameters.put("Unknown70", getBytes2(startIndex + 70l));
//				spell.parameters.put("Unknown72", getBytes2(startIndex + 72l));
//				spell.parameters.put("Unknown74", getBytes2(startIndex + 74l));
//				spell.parameters.put("Unknown76", getBytes2(startIndex + 76l));
//				spell.parameters.put("Unknown78", getBytes2(startIndex + 78l));
//				spell.parameters.put("Unknown90", getBytes2(startIndex + 90l));
//
//				spell.parameters.put("Unknown66-1", getBytes1(startIndex + 66l));
//				spell.parameters.put("Unknown67-1", getBytes1(startIndex + 67l));
//				spell.parameters.put("Unknown68-1", getBytes1(startIndex + 68l));
//				spell.parameters.put("Unknown69-1", getBytes1(startIndex + 69l));
//				spell.parameters.put("Unknown70-1", getBytes1(startIndex + 70l));
//				spell.parameters.put("Unknown71-1", getBytes1(startIndex + 71l));
//				spell.parameters.put("Unknown72-1", getBytes1(startIndex + 72l));
//
				spell.parameters.put("effects_count", getBytes2(startIndex + 64l));
				
				effect.modifiers_mask = getBytes8(startIndex+80);
				spell.parameters.put("next_spell", getBytes2(startIndex + 88l));

				effect.record_number = spellNumber;
				effect.object_type = "Spell";

				spell.effect = effect;
				
		        List<Attribute> attributes = new ArrayList<Attribute>();
				long newIndex = startIndex+92l;
				int attrib = getBytes4(newIndex);
				long valueIndex = newIndex + 60l;
				long value = getBytes4(valueIndex);
				while (attrib != 0) {
					attributes.add(new Attribute(spellNumber, attrib, value));
					newIndex+=4;
					valueIndex+=4;
					attrib = getBytes4(newIndex);
					value = getBytes4(valueIndex);
				}
				spell.attributes = attributes;

				spellList.add(spell);

				stream = new FileInputStream(EXE_NAME);		
				startIndex = startIndex + Starts.SPELL_SIZE;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);

				spellNumber++;
			}
			in.close();
			stream.close();
			
			//make sure there's a place to put csv files
			CSVWriter.createCSVOutputDirectory();
			
			XSSFWorkbook wb = new XSSFWorkbook();
			FileOutputStream fos = CSVWriter.getFOS("spells", SSType.XLSX);
			BufferedWriter   csv = CSVWriter.getBFW("spells", SSType.CSV);
			XSSFSheet sheet = wb.createSheet();
			
			XSSFWorkbook wb2 = new XSSFWorkbook();
			FileOutputStream fos2 = CSVWriter.getFOS("effects_spells", SSType.XLSX);
			BufferedWriter   csv2 = CSVWriter.getBFW("effects_spells", SSType.CSV);
			XSSFSheet sheet2 = wb2.createSheet();
			
			XSSFWorkbook wb3 = new XSSFWorkbook();
			FileOutputStream fos3 = CSVWriter.getFOS("attributes_by_spell", SSType.XLSX);
			BufferedWriter   csv3 = CSVWriter.getBFW("attributes_by_spell", SSType.CSV);
			XSSFSheet sheet3 = wb3.createSheet();

			int rowNum = 0;
			int effectsNum = 0;
			int attributesNum = 0;
			for (Spell spell : spellList) {
				// spells
				if (rowNum == 0) {
					XSSFRow row = sheet.createRow(rowNum);
					for (int i = 0; i < spells_columns.length; i++) {
						row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(spells_columns[i]);
					}
					rowNum++;
				}
				XSSFRow row = sheet.createRow(rowNum);
				for (int i = 0; i < spells_columns.length; i++) {
					Object object = spell.parameters.get(spells_columns[i]);
					if (object != null) {
						row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(object.toString());
					}
				}
				
				// effects_spells
				if (spell.effect != null) {
					if (effectsNum == 0) {
						row = sheet2.createRow(effectsNum);
						for (int i = 0; i < effects_spells_columns.length; i++) {
							row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(effects_spells_columns[i]);
						}
						effectsNum++;
					}
					
					row = sheet2.createRow(effectsNum);
					row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(spell.effect.record_number);
					row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(spell.effect.effect_number);
					row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(spell.effect.duration != 0 ? spell.effect.duration+"" : "");
					row.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(spell.effect.ritual);
					row.getCell(4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(spell.effect.object_type);
					row.getCell(5, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(Long.toString(spell.effect.raw_argument));
					row.getCell(6, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(Long.toString(spell.effect.modifiers_mask));
					row.getCell(7, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(spell.effect.range_base);
					row.getCell(8, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(spell.effect.range_per_level);
					row.getCell(9, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(spell.effect.range_strength_divisor);
					row.getCell(10, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(spell.effect.area_base);
					row.getCell(11, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(spell.effect.area_per_level);
					row.getCell(12, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(spell.effect.area_battlefield_pct != 0 ? spell.effect.area_battlefield_pct+"" : "");
					effectsNum++;
				}
				
				
				// attributes_by_spell
				for (Attribute attribute : spell.attributes) {
					if (attributesNum == 0) {
						row = sheet3.createRow(attributesNum);
						for (int i = 0; i < attributes_by_spell_columns.length; i++) {
							row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(attributes_by_spell_columns[i]);
						}
						attributesNum++;
					}
					row = sheet3.createRow(attributesNum);
					row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(attribute.object_number);
					row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(attribute.attribute);
					row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(attribute.raw_value);
					attributesNum++;
				}
				
				dumpTextFile(spell, writer);

				rowNum++;
			}
			
			writer.close();

			wb.write(fos);
			fos.close();
			CSVWriter.writeSimpleCSV(sheet, csv, Delimiter.TAB);
			csv.close();
			wb.close();

			wb2.write(fos2);
			fos2.close();
			CSVWriter.writeSimpleCSV(sheet2, csv2, Delimiter.TAB);
			csv2.close();
			wb2.close();

			wb3.write(fos3);
			fos3.close();
			CSVWriter.writeSimpleCSV(sheet3, csv3, Delimiter.TAB);
			csv3.close();
			wb3.close();

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
	
	private static void dumpTextFile(Spell spell, BufferedWriter writer) throws IOException {
		Object name = spell.parameters.get("name");
		Object id = spell.parameters.get("id");
		writer.write(name.toString() + "(" + id + ")");
		writer.newLine();
		for (Map.Entry<String, Object> entry : spell.parameters.entrySet()) {
			if (!entry.getKey().equals("name") && !entry.getKey().equals("id")) {
				writer.write("\t" + entry.getKey() + ": " + entry.getValue());
				writer.newLine();
			}
		}
		printAttributes(spell.attributes, writer);
		printEffect(spell.effect, writer);
		writer.newLine();
	}
	
	private static class Spell {
		Map<String, Object> parameters;
		Effect effect;
		List<Attribute> attributes;
	}

}
