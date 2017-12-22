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
import java.util.Set;
import java.util.TreeMap;
import java.util.TreeSet;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ItemStatIndexer extends AbstractStatIndexer {
	public static String[] items_columns = {"id", "name", "type", "constlevel", "mainpath", "mainlevel", "secondarypath", "secondarylevel", "weapon", "armor", 
			"shockres", "fireres", "coldres", "poisonres", "darkvision", "airshield", "limitedregeneration", "woundfend", "fixforge", "autodishealer", "healer", 
			"alch", "insa", "hp", "protf", "protb", "mr", "morale", "str", "att", "def", "prec", "enc", "mapspeed", "ap", "reinvigoration", "pen", "stealth", 
			"stealthb", "forest", "mount", "waste", "swamp", "fly", "float", "sailingshipsize", "sailingmaxunitsize", "waterbreathing", "giftofwater", "airbr", 
			"flytr", "quick", "eth", "trample", "bless", "luck", "fluck", "curse", "disease", "cursed", "taint", "ldr-n", "ldr-u", "ldr-m", "inspirational", 
			"taskmaster", "fear", "awe", "animalawe", "exp", "chill", "heat", "gold", "F", "A", "W", "E", "S", "D", "N", "B", "H", "firerange", "airrange", 
			"waterrange", "earthrange", "astralrange", "deathrange", "naturerange", "bloodrange", "tmpfiregems", "tmpairgems", "tmpwatergems", "tmpearthgems", 
			"tmpastralgems", "tmpdeathgems", "tmpnaturegems", "tmpbloodslaves", "gf", "ga", "gw", "ge", "gs", "gd", "gn", "gb", "berserk", "bers", "fireshield", 
			"banefireshield", "iceprot", "invul", "bloodvengeance", "pillagebonus", "patrolbonus", "castledef", "siegebonus", "supplybonus", "researchbonus", 
			"heretic", "douse", "void", "diseasecloud", "poisoncloud", "reaper", "crossbreeder", "ivylord", "spelleffect", "startbattlespell", "autocombatspell", 
			"itemspell", "ritual", "sumrit", "#sumrit", "sumauto", "#sumauto", "sumbat", "#sumbat", "affliction", "restrictions", "special", "regeneration", 
			"restricted1", "restricted2", "restricted3", "restricted4", "restricted5", "restricted6", "aging", "corpselord", "lictorlord", "bloodsac", 
			"mastersmith", "eyeloss", "armysize", "defender", "cannotwear", "reanimH", "reanimD", "dragonmastery", "patience", "retinue", "end"};																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																				

	private static String[][] KNOWN_ITEM_ATTRS = {
			{"C600", "fireres"},
			{"C900", "coldres"},
			{"C800", "poisonres"},
			{"C700", "shockres"},
			{"9D00", "ldr-n"},
			{"9700", "str"},
			{"C501", "fixforge"},
			{"9E00", "ldr-m"},
			{"9F00", "ldr-u"},
			{"7001", "inspirational"},
			{"3401", "morale"},
			{"A100", "pen"},
			{"8300", "pillagebonus"},
			{"B700", "fear"},
			{"A000", "mr"},
			{"0601", "taint"},
			{"7500", "reinvigoration"},
			{"6900", "awe"},
			{"8500", "autocombatspell"},
			{"0A00", "F"},
			{"0B00", "A"},
			{"0C00", "W"},
			{"0D00", "E"},
			{"0E00", "S"},
			{"0F00", "D"},
			{"1000", "N"},
			{"1100", "B"},
			{"1200", "H"},
			{"1400", "elemental"},
			{"1500", "sorcery"},
			{"1600", "all"},
			{"1700", "elemental range"},
			{"1800", "sorcery range"},
			{"1900", "all range"},
			{"2800", "firerange"},
			{"2900", "airrange"},
			{"2A00", "waterrange"},
			{"2B00", "earthrange"},
			{"2C00", "astralrange"},
			{"2D00", "deathrange"},
			{"2E00", "naturerange"},
			{"2F00", "bloodrange"},
			{"1901", "darkvision"},
			{"CE01", "limitedregeneration"},
			{"BD00", "regeneration"},
			{"6E00", "waterbreathing"},
			{"8601", "stealthb"},
			{"6C00", "stealth"},
			{"9600", "att"},
			{"7901", "def"},
			{"9601", "woundfend"},
			{"1601", "restricted#"},
			{"BE00", "berserk"},
			{"2F01", "aging"},
			{"6500", "ivylord"},
			{"A501", "mount"},
			{"A601", "forest"},
			{"A701", "waste"},
			{"A801", "swamp"},
			{"7900", "researchbonus"},
			{"6F00", "giftofwater"},
			{"9A00", "corpselord"},
			{"6600", "lictorlord"},
			{"8B00", "sumauto"},
			{"D800", "bloodsac"},
			{"6B01", "mastersmith"},
			{"8400", "alch"},
			{"7E00", "eyeloss"},
			{"A301", "armysize"},
			{"8900", "defender"},
			{"7700", "forbidden light"},
			{"C701", "cannotwear"},
			{"C801", "cannotwear"},
			{"7000", "sailingshipsize"},
			{"9A01", "sailingmaxunitsize"},
			{"7100", "flytr"},
			{"7F01", "protf"},
			{"B800", "heretic"},
			{"6301", "autodishealer"},
			{"AA00", "patrolbonus"},
			{"B500", "prec"},
			{"5801", "tmpfiregems"},
			{"5901", "tmpairgems"},
			{"5A01", "tmpwatergems"},
			{"5B01", "tmpearthgems"},
			{"5C01", "tmpastralgems"},
			{"5D01", "tmpdeathgems"},
			{"5E01", "tmpnaturegems"},
			{"5F01", "tmpbloodgems"},
			{"6201", "healer"},
			{"7A00", "supplybonus"},
			{"8100", "airbr"},
			{"A901", "mapspeed"},
			{"1E00", "gf"},
			{"1F00", "ga"},
			{"2000", "gw"},
			{"2100", "ge"},
			{"2200", "gs"},
			{"2300", "gd"},
			{"2400", "gn"},
			{"2500", "gb"},
			{"6C01", "reanimH"},
			{"6D01", "reanimD"},
			{"0E01", "dragonmastery"},
			{"AF01", "crossbreeder"},
			{"CD01", "patience"},
			{"B401", "retinue"}
	};

	
	private static String values[][] = {{"bless", "luck", "", "airshield", "barkskin", "", "", "", "bers", "", "", "", "", "", "", "" },
										{"stoneskin", "fly", "quick", "", "", "", "", "", "", "", "", "eth", "ironskin", "", "", "" },
										{"", "", "", "float", "", "", "", "", "", "", "", "", "", "", "", "" },
										{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" },
										{"", "", "", "", "", "", "trample", "", "", "", "", "", "", "", "", "fireshield?" },
										{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" },
										{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" },
										{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" },
										{"disease", "curse", "", "", "", "", "", "", "", "", "", "", "", "", "", "cursed" },
										{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" },
										{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" },
										{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" }};

	public static void main(String[] args) {
		run();
	}
	
	public static void run() {
		FileInputStream stream = null;
        List<Item> itemList = new ArrayList<Item>();

		try {
			BufferedWriter writer = new BufferedWriter(new FileWriter("items.txt"));
			BufferedWriter writerUnknown = new BufferedWriter(new FileWriter("itemsUnknown.txt"));
	        long startIndex = Starts.ITEM;
	        int ch;
			stream = new FileInputStream(EXE_NAME);			
			stream.skip(startIndex);
			Set<String> unknownAttributes = new TreeSet<String>();
			
			// Name
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
				
				Item item = new Item();
				item.parameters = new TreeMap<String, Object>();
				item.parameters.put("id", rowNumber);
				item.parameters.put("name", name.toString());
				short constlevel = getBytes1(startIndex + 36);
				item.parameters.put("constlevel", constlevel == -1 ? "" : constlevel*2);
				short mainpath = getBytes1(startIndex + 37);
				String[] paths = {"F", "A", "W", "E", "S", "D", "N", "B"};
				item.parameters.put("mainpath", mainpath == -1 ? "" : paths[mainpath]);
				item.parameters.put("mainlevel", getBytes1(startIndex + 39));
				short secondarypath = getBytes1(startIndex + 38);
				item.parameters.put("secondarypath", secondarypath == -1 ? "" : paths[secondarypath]);
				short secondarylevel = getBytes1(startIndex + 40);
				item.parameters.put("secondarylevel", secondarylevel == 0 ? "" : secondarylevel);
				short type = getBytes1(startIndex + 41);
				String[] types = {"1-h wpn", "2-h wpn", "missile", "shield", "armor", "helm", "boots", "misc"};
				item.parameters.put("type", type < 1 ? "" : types[type-1]);
				int weapon = getBytes2(startIndex + 44);
				item.parameters.put("weapon", weapon == 0 ? "" : weapon);
				int armor = getBytes2(startIndex + 46);
				item.parameters.put("armor", armor == 0 ? "" : armor);
				item.parameters.put("itemspell", getString(startIndex + 48));
				
				List<AttributeValue> attributes = getAttributes(startIndex + Starts.ITEM_ATTRIBUTE_OFFSET, Starts.ITEM_ATTRIBUTE_GAP);
				
				for (AttributeValue attr : attributes) {
					for (int x = 0; x < KNOWN_ITEM_ATTRS.length; x++) {
						if (KNOWN_ITEM_ATTRS[x][0].equals(attr.attribute)) {
							if (KNOWN_ITEM_ATTRS[x][1].endsWith("#")) {
								int i = 1;
								for (String value : attr.values) {
									item.parameters.put(KNOWN_ITEM_ATTRS[x][1].replace("#", i+""), Integer.parseInt(value)-100);
									i++;
								}
							} else {
								switch (KNOWN_ITEM_ATTRS[x][1]) {
								case ("elemental"):
									item.parameters.put("F", attr.values.get(0));
									item.parameters.put("A", attr.values.get(0));
									item.parameters.put("W", attr.values.get(0));
									item.parameters.put("E", attr.values.get(0));
									break;
								case ("sorcery"):
									item.parameters.put("S", attr.values.get(0));
									item.parameters.put("D", attr.values.get(0));
									item.parameters.put("N", attr.values.get(0));
									item.parameters.put("B", attr.values.get(0));
									break;
								case ("all"):
									item.parameters.put("F", attr.values.get(0));
									item.parameters.put("A", attr.values.get(0));
									item.parameters.put("W", attr.values.get(0));
									item.parameters.put("E", attr.values.get(0));
									item.parameters.put("S", attr.values.get(0));
									item.parameters.put("D", attr.values.get(0));
									item.parameters.put("N", attr.values.get(0));
									item.parameters.put("B", attr.values.get(0));
									break;
								case ("elemental range"):
									item.parameters.put("firerange", attr.values.get(0));
									item.parameters.put("airrange", attr.values.get(0));
									item.parameters.put("waterrange", attr.values.get(0));
									item.parameters.put("earthrange", attr.values.get(0));
									break;
								case ("sorcery range"):
									item.parameters.put("astralrange", attr.values.get(0));
									item.parameters.put("deathrange", attr.values.get(0));
									item.parameters.put("naturerange", attr.values.get(0));
									item.parameters.put("bloodrange", attr.values.get(0));
									break;
								case ("all range"):
									item.parameters.put("firerange", attr.values.get(0));
									item.parameters.put("airrange", attr.values.get(0));
									item.parameters.put("waterrange", attr.values.get(0));
									item.parameters.put("earthrange", attr.values.get(0));
									item.parameters.put("astralrange", attr.values.get(0));
									item.parameters.put("deathrange", attr.values.get(0));
									item.parameters.put("naturerange", attr.values.get(0));
									item.parameters.put("bloodrange", attr.values.get(0));
									break;
								case ("forbidden light"): //HACK
									item.parameters.put("F", 2);
									item.parameters.put("S", 2);
									break;
								default:
									item.parameters.put(KNOWN_ITEM_ATTRS[x][1], attr.values.get(0));
								}
							}
						} else {
 							item.parameters.put("\tUnknown Attribute<" + attr.attribute + ">", attr.values.get(0));	
 							unknownAttributes.add(attr.attribute);
						}
					}
				}
				
				if (attributes.contains(new AttributeValue("8500"))) {
					item.parameters.put("startbattlespell", "");
					item.parameters.put("autocombatspell", getString(startIndex + 84));
				} else {
					item.parameters.put("startbattlespell", getString(startIndex + 84));
					item.parameters.put("autocombatspell", "");
				}
				
				List<String> largeBitmap = largeBitmap(startIndex + 208, values);
				for (String bit : largeBitmap) {
					if (bit.equals("airshield")) {
						item.parameters.put(bit, 80);
					} else {
						item.parameters.put(bit, 1);
					}
				}
				
				itemList.add(item);

				stream = new FileInputStream(EXE_NAME);		
				startIndex = startIndex + 232l;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);

		        rowNumber++;
			}
			in.close();
			stream.close();

			XSSFWorkbook wb = new XSSFWorkbook();
			FileOutputStream fos = new FileOutputStream("BaseI.xlsx");
			XSSFSheet sheet = wb.createSheet();

			int rowNum = 0;
			for (Item item : itemList) {
				if (rowNum == 0) {
					XSSFRow row = sheet.createRow(rowNum);
					for (int i = 0; i < items_columns.length; i++) {
						row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(items_columns[i]);
					}
					rowNum++;
				}
				XSSFRow row = sheet.createRow(rowNum);
				for (int i = 0; i < items_columns.length; i++) {
					Object object = item.parameters.get(items_columns[i]);
					if (object != null) {
						row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(object.toString());
					}
				}
				dumpTextFile(item, writer);
				rowNum++;
			}

			dumpUnknown(unknownAttributes, writerUnknown);
			
			writerUnknown.close();
			writer.close();
			
			wb.write(fos);
			fos.close();
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
	
	private static void dumpTextFile(Item item, BufferedWriter writer) throws IOException {
		Object name = item.parameters.get("name");
		Object id = item.parameters.get("id");
		writer.write(name.toString() + "(" + id + ")");
		writer.newLine();
		for (Map.Entry<String, Object> entry : item.parameters.entrySet()) {
			if (!entry.getKey().equals("name") && !entry.getKey().equals("id") && entry.getValue() != null && !entry.getValue().equals("")) {
				writer.write("\t" + entry.getKey() + ": " + entry.getValue());
				writer.newLine();
			}
		}
		writer.newLine();
	}

	private static class Item {
		Map<String, Object> parameters;
	}

}
