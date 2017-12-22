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
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SiteStatIndexer extends AbstractStatIndexer {
	public static String[] site_columns = {"id", "name", "rarity", "loc", "level", "path", "F", "A", "W", "E", "S", "D", "N", "B", "gold", "res", 
			"sup", "unr", "exp", "lab", "fort", "scale1", "scale2", "domspread", "turmoil", "sloth", "cold", "death", "misfortune", "drain", 
			"fireres", "coldres", "shockres", "poisonres", "str", "prec", "mor", "undying", "att", "darkvision", "aawe", "rit", "ritrng", 
			"hmon1", "hmon2", "hmon3", "hmon4", "hmon5", "voidgate", "sum1", "n_sum1", "sum2", "n_sum2", "sum3", "n_sum3", "conj", "alter", "evo", 
			"const", "ench", "thau", "blood", "heal", "disease", "curse", "horror", "holyfire", "holypow", "scry", "adventure", "other", "sum4", "n_sum4", 
			"hcom1", "hcom2", "hcom3", "hcom4", "hcom5", "mon1", "mon2", "mon3", "mon4", "mon5", "com1", "com2", "com3", "com4", "com5", "reveal", 
			"provdef1", "provdef2", "def", "F2", "A2", "W2", "E2", "S2", "D2", "N2", "B2", "awe", "reinvigoration", "airshield", "provdefcom", 
			"domconflict", "sprite", "end"};																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																																									

	private static String[][] KNOWN_SITE_ATTRS = {
			{"0100", "F"},
			{"0200", "A"},
			{"0300", "W"},
			{"0400", "E"},
			{"0500", "S"},
			{"0600", "D"},
			{"0700", "N"},
			{"0800", "B"},
			{"0D00", "gold"},
			{"0E00", "res"},
			{"1400", "sup"},
			{"1300", "unr"},
			{"1600", "exp"},
			{"0F00", "lab"},
			{"1100", "fort"},
			{"1E00", "hcom#"},
			{"1D00", "hmon#"},
			{"0C00", "com#"},
			{"0B00", "mon#"},
			{"3C00", "conj"},
			{"3D00", "alter"},
			{"3E00", "evo"},
			{"3F00", "const"},
			{"4000", "ench"},
			{"4100", "thau"},
			{"4200", "blood"},
			{"4600", "heal"},
			{"1500", "disease"},
			{"4700", "curse"},
			{"1800", "horror"},
			{"4400", "holyfire"},
			{"4300", "holypow"},
			{"4800", "scry"},
			{"C000", "adventure"},
			{"3900", "voidgate"},
			{"1200", "summoning"},
			{"1501", "domspread"},
			{"1901", "turmoil"},
			{"1A01", "sloth"},
			{"1B01", "cold"},
			{"1C01", "death"},
			{"1D01", "misfortune"},
			{"1E01", "drain"},
			{"FB01", "fireres"},
			{"FC01", "coldres"},
			{"FA01", "str"},
			{"0402", "prec"},
			{"F401", "mor"},
			{"FD01", "shockres"},
			{"F801", "undying"},
			{"F501", "att"},
			{"FE01", "poisonres"},
			{"0302", "darkvision"},
			{"0102", "aawe"},
			{"1401", "throne?"},
			{"0A01", "fortparts"},
			{"0601", "reveal"},
			{"E000", "provdef#"},
			{"F601", "def"},
			{"0202", "awe"},
			{"FF01", "reinvigoration"},
			{"0002", "airshield"},
			{"4A00", "provdefcom"},
	};

	public static void main(String[] args) {
		run();
	}
	
	public static void run() {
		FileInputStream stream = null;
        List<Site> siteList = new ArrayList<Site>();

		try {
			BufferedWriter writer = new BufferedWriter(new FileWriter("items.txt"));
			BufferedWriter writerUnknown = new BufferedWriter(new FileWriter("itemsUnknown.txt"));
	        long startIndex = Starts.SITE;
	        int ch;
			stream = new FileInputStream(EXE_NAME);			
			stream.skip(Starts.SITE);
			Set<String> unknown = new HashSet<String>();

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

				Site site = new Site();
				site.parameters = new TreeMap<String, Object>();
				site.parameters.put("id", rowNumber);
				site.parameters.put("name", name.toString());
				short rarity = getBytes1(startIndex + 42);
				site.parameters.put("rarity", rarity == -1 ? "0" : rarity);
				site.parameters.put("loc", getBytes4(startIndex + 208));
				site.parameters.put("level", getBytes2(startIndex + 40));
				String[] paths = {"Fire", "Air", "Water", "Earth", "Astral", "Death", "Nature", "Blood", "Holy"};
				int[] spriteOffset = {1, 9, 18, 26, 35, 42, 50, 59, 68};
				short path = getBytes1(startIndex + 38);
				short sprite = getBytes1(startIndex + 36);
				site.parameters.put("path", path == -1 ? "" : paths[path]);
				site.parameters.put("sprite", path == -1 ? "" : spriteOffset[path] + sprite);
				
				List<AttributeValue> attributes = getAttributes(startIndex + Starts.SITE_ATTRIBUTE_OFFSET, Starts.SITE_ATTRIBUTE_GAP, 8);
				for (AttributeValue attr : attributes) {
					for (int x = 0; x < KNOWN_SITE_ATTRS.length; x++) {
						if (KNOWN_SITE_ATTRS[x][0].equals(attr.attribute)) {
							if (KNOWN_SITE_ATTRS[x][1].endsWith("#")) {
								int i = 1;
								for (String value : attr.values) {
									site.parameters.put(KNOWN_SITE_ATTRS[x][1].replace("#", i+""), Integer.parseInt(value));
									i++;
								}
							} else {
								switch (KNOWN_SITE_ATTRS[x][1]) {
								case ("conj"):
								case ("alter"):
								case ("evo"):
								case ("const"):
								case ("ench"):
								case ("thau"):
								case ("blood"):
								case ("heal"):
								case ("disease"):
								case ("curse"):
								case ("horror"):
								case ("holyfire"):
								case ("holypow"):
								case ("voidgate"):
									site.parameters.put(KNOWN_SITE_ATTRS[x][1], attr.values.get(0)+"%");
									break;
								case ("lab"):
									site.parameters.put(KNOWN_SITE_ATTRS[x][1], "lab");
									break;
								case ("unr"):
									site.parameters.put(KNOWN_SITE_ATTRS[x][1], -Integer.parseInt(attr.values.get(0)));
									break;
								default:
									site.parameters.put(KNOWN_SITE_ATTRS[x][1], attr.values.get(0));
								}

							}
						} else {
							site.parameters.put("\tUnknown Attribute<" + attr.attribute + ">", attr.values.get(0));							
 							unknown.add(attr.attribute);
						}
					}
				}
				
				// scales
				String[] scales = {"Turmoil", "Sloth", "Cold", "Death", "Misfortune", "Drain"};
				String[] opposite = {"Order", "Productivity", "Heat", "Growth", "Luck", "Magic"};
				String scalesValue[] = {"", ""};
				int index = 0;
				if (attributes.contains(new AttributeValue("1F00"))) {
					AttributeValue attributeValue = attributes.get(attributes.indexOf(new AttributeValue("1F00")));
					for (String vals : attributeValue.values) {
						scalesValue[index++] = opposite[Integer.parseInt(vals)];
					}
				}
				if (attributes.contains(new AttributeValue("2000"))) {
					AttributeValue attributeValue = attributes.get(attributes.indexOf(new AttributeValue("2000")));
					for (String vals : attributeValue.values) {
						scalesValue[index++] = scales[Integer.parseInt(vals)];
					}
				}
				site.parameters.put("scale1", scalesValue[0]);
				site.parameters.put("scale2", scalesValue[1]);
				
				// rit/ritrng
				boolean[] boolPaths = {false, false, false, false, false, false, false, false};
				String[] boolPathsStr = {"F", "A", "W", "E", "S", "D", "N", "B"};
				String value = "";
				if (attributes.contains(new AttributeValue("FA00"))) {
					AttributeValue attributeValue = attributes.get(attributes.indexOf(new AttributeValue("FA00")));
					value = attributeValue.values.get(0);
					boolPaths[0] = true;
				}
				if (attributes.contains(new AttributeValue("FB00"))) {
					AttributeValue attributeValue = attributes.get(attributes.indexOf(new AttributeValue("FB00")));
					value = attributeValue.values.get(0);
					boolPaths[1] = true;
				}
				if (attributes.contains(new AttributeValue("FC00"))) {
					AttributeValue attributeValue = attributes.get(attributes.indexOf(new AttributeValue("FC00")));
					value = attributeValue.values.get(0);
					boolPaths[2] = true;
				}
				if (attributes.contains(new AttributeValue("FD00"))) {
					AttributeValue attributeValue = attributes.get(attributes.indexOf(new AttributeValue("FD00")));
					value = attributeValue.values.get(0);
					boolPaths[3] = true;
				}
				if (attributes.contains(new AttributeValue("FE00"))) {
					AttributeValue attributeValue = attributes.get(attributes.indexOf(new AttributeValue("FE00")));
					value = attributeValue.values.get(0);
					boolPaths[4] = true;
				}
				if (attributes.contains(new AttributeValue("FF00"))) {
					AttributeValue attributeValue = attributes.get(attributes.indexOf(new AttributeValue("FF00")));
					value = attributeValue.values.get(0);
					boolPaths[5] = true;
				}
				if (attributes.contains(new AttributeValue("0001"))) {
					AttributeValue attributeValue = attributes.get(attributes.indexOf(new AttributeValue("0001")));
					value = attributeValue.values.get(0);
					boolPaths[6] = true;
				}
				if (attributes.contains(new AttributeValue("0101"))) {
					AttributeValue attributeValue = attributes.get(attributes.indexOf(new AttributeValue("0101")));
					value = attributeValue.values.get(0);
					boolPaths[7] = true;
				}
				if (attributes.contains(new AttributeValue("0401"))) {
					AttributeValue attributeValue = attributes.get(attributes.indexOf(new AttributeValue("0401")));
					value = attributeValue.values.get(0);
					boolPaths = new boolean[]{true, true, true, true, true, true, true, true};
				}
				StringBuffer rit = new StringBuffer();
				for (int x = 0; x < boolPaths.length; x++) {
					if (boolPaths[x]) {
						rit.append(boolPathsStr[x]);
					}
				}
				site.parameters.put("rit", rit.toString());
				site.parameters.put("ritrng", value);
				
				// summoning
				String sum1 = null;
				int sum1count = 0;
				String sum2 = null;
				int sum2count = 0;
				String sum3 = null;
				int sum3count = 0;
				String sum4 = null;
				int sum4count = 0;
				if (attributes.contains(new AttributeValue("1200"))) {
					AttributeValue attributeValue = attributes.get(attributes.indexOf(new AttributeValue("1200")));
					for (String val : attributeValue.values) {
						if (sum1 == null || sum1.equals(val)) {
							sum1 = val;
							sum1count++;
						} else if (sum2 == null || sum2.equals(val)) {
							sum2 = val;
							sum2count++;
						} else if (sum3 == null || sum3.equals(val)) {
							sum3 = val;
							sum3count++;
						} else if (sum4 == null || sum4.equals(val)) {
							sum4 = val;
							sum4count++;
						}
					}
				}
				if (sum1 != null) {
					site.parameters.put("sum1", sum1);
					site.parameters.put("n_sum1", sum1count);
				}
				if (sum2 != null) {
					site.parameters.put("sum2", sum2);
					site.parameters.put("n_sum2", sum2count);
				}
				if (sum3 != null) {
					site.parameters.put("sum3", sum3);
					site.parameters.put("n_sum3", sum3count);
				}
				if (sum4 != null) {
					site.parameters.put("sum4", sum4);
					site.parameters.put("n_sum4", sum4count);
				}

				siteList.add(site);

				stream = new FileInputStream(EXE_NAME);		
				startIndex = startIndex + Starts.SITE_SIZE;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);

				rowNumber++;
			}
			in.close();
			stream.close();

			XSSFWorkbook wb = new XSSFWorkbook();
			FileOutputStream fos = new FileOutputStream("MagicSites.xlsx");
			XSSFSheet sheet = wb.createSheet();

			int rowNum = 0;
			for (Site site : siteList) {
				if (rowNum == 0) {
					XSSFRow row = sheet.createRow(rowNum);
					for (int i = 0; i < site_columns.length; i++) {
						row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(site_columns[i]);
					}
					rowNum++;
				}
				XSSFRow row = sheet.createRow(rowNum);
				for (int i = 0; i < site_columns.length; i++) {
					Object object = site.parameters.get(site_columns[i]);
					if (object != null) {
						row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(object.toString());
					}
				}
				
				dumpTextFile(site, writer);
				rowNum++;
			}
			
			dumpUnknown(unknown, writerUnknown);

			writer.close();
			writerUnknown.close();
			
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
	
	private static void dumpTextFile(Site site, BufferedWriter writer) throws IOException {
		Object name = site.parameters.get("name");
		Object id = site.parameters.get("id");
		writer.write(name.toString() + "(" + id + ")");
		writer.newLine();
		for (Map.Entry<String, Object> entry : site.parameters.entrySet()) {
			if (!entry.getKey().equals("name") && !entry.getKey().equals("id") && entry.getValue() != null && !entry.getValue().equals("")) {
				writer.write("\t" + entry.getKey() + ": " + entry.getValue());
				writer.newLine();
			}
		}
		writer.newLine();
	}

	private static class Site {
		Map<String, Object> parameters;
	}

}
