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
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MonsterStatIndexer {
	private static Map<String, boolean[]> largeBitmapCache = new HashMap<String, boolean[]>();

	static class CacheKey {
		String key;
		int id;
		public CacheKey(String key, int id) {
			super();
			this.key = key;
			this.id = id;
		}
		
	}
	private static Map<CacheKey, String> attrCache = new HashMap<CacheKey, String>();
	
	private static Map<Integer, String> columnsUsed = new HashMap<Integer, String>();

	private static XSSFWorkbook readFile(String filename) throws IOException {
		return new XSSFWorkbook(new FileInputStream(filename));
	}
	
	private static int MASK[] = {0x0001, 0x0002, 0x0004, 0x0008, 0x0010, 0x0020, 0x0040, 0x0080,
			0x0100, 0x0200, 0x0400, 0x0800, 0x1000, 0x2000, 0x4000, 0x8000};
	
	private static String value1[] = {"heal", "mounted", "animal", "amphibian", "wastesurvival", "undead", "coldres15", "heat", "neednoteat", "fireres15", "poisonres15", "aquatic", "flying", "trample", "immobile", "immortal" };
	private static String value2[] = {"cold", "forestsurvival", "shockres15", "swampsurvival", "demon", "sacred", "mountainsurvival", "illusion", "noheal", "ethereal", "pooramphibian", "stealthy40", "misc2", "coldblood", "inanimate", "female" };
	private static String value3[] = {"bluntres", "slashres", "pierceres", "slow_to_recruit", "float", "", "teleport", "", "", "", "", "", "", "", "", "" };
	private static String value4[] = {"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
	private static String value5[] = {"magicbeing", "", "", "poormagicleader", "okmagicleader", "goodmagicleader", "expertmagicleader", "superiormagicleader", "poorundeadleader", "okundeadleader", "goodundeadleader", "expertundeadleader", "superiorundeadleader", "", "", "" };
	private static String value6[] = {"", "", "", "", "", "", "", "", "noleader", "poorleader", "goodleader", "expertleader", "superiorleader", "", "", "" };
	private static String value7[] = {"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
	private static String value8[] = {"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
	
	
	private static String DOTHESE[][] = {
		{"heal", "108"}, 
		{"mounted", "33"},
		{"animal", "81"},
		{"amphibian", "90"},
		{"wastesurvival", "86"},
		{"undead", "77"},
		{"heat", "141"},
		{"neednoteat", "112"},
		{"aquatic", "89"},
		{"flying", "93"},
		{"trample", "159"},
		{"immortal", "109"},
		{"cold", "140"},
		{"forestsurvival", "84"},
		{"swampsurvival", "87"},
		{"demon", "78"},
		{"sacred", "73"},
		{"mountainsurvival", "85"},
		{"illusion", "101"},
		{"noheal", "111"},
		{"ethereal", "156"},
		{"pooramphibian", "91"},
		{"coldblood", "82"},
		{"inanimate", "76"},
		{"female", "83"},
		{"bluntres", "126"},
		{"slashres", "125"},
		{"pierceres", "127"},
		{"slow_to_recruit", "13"},
		{"float", "92"},
		{"teleport", "95"},
		{"magicbeing", "79"},
		{"immobile", "96"},
		{"stealthy40", "100"},
		{"coldres15", "130"},
		{"fireres15", "129"},
		{"poisonres15", "131"},
		{"shockres15", "128"},
		{"noleader", "35"},
		{"poorleader", "35"}, 
		{"goodleader", "35"}, 
		{"expertleader", "35"}, 
		{"superiorleader", "35"},
		{"poorundeadleader", "36"},
		{"okundeadleader", "36"}, 
		{"goodundeadleader", "36"}, 
		{"expertundeadleader", "36"}, 
		{"superiorundeadleader", "36"},
		{"poormagicleader", "37"},
		{"okmagicleader", "37"}, 
		{"goodmagicleader", "37"}, 
		{"expertmagicleader", "37"}, 
		{"superiormagicleader", "37"},
		{"misc2", "40"},
		};
	
	private static String SkipColumns[] = {
		"id",
		"name",
		"armor4",
		"gcom",
		"gmon",
		"rcost",
		"forgebonus",
		"gem",
		"mind",
		"F",
		"A",
		"W",
		"E",
		"S",
		"D",
		"N",
		"B",
		"H",
		"link1",
		"nbr1",
		"rand1",
		"mask1",
		"link2",
		"nbr2",
		"rand2",
		"mask2",
		"link3",
		"nbr3",
		"rand3",
		"mask3",
		"link4",
		"nbr4",
		"rand4",
		"mask4",
		"type",
		"typeclass",
		"from",
		"unique",
		"fixedname",
		"special",
		"realm3",
		"realm2",
		"realm1",
		"baseleadership",
		"pathboost",
		"test"
	};
	
	private static class Magic {
		public int F;
		public int A;
		public int W;
		public int E;
		public int S;
		public int D;
		public int N;
		public int B;
		public int H;
		public List<RandomMagic> rand;
	}
	
	private static class RandomMagic {
		public int rand;
		public int nbr;
		public int link;
		public int mask;
	}
	
	private static Map<Integer, Magic> monsterMagic = new HashMap<Integer, Magic>();
	
	private static String magicStrip(int mag) {
		return mag > 0 ? Integer.toString(mag) : "";
	}

	private static void doit1(int skip, int column, XSSFSheet sheet) throws IOException {
		doit1(skip, column, sheet, null);
	}
	
	private static void doit1(int skip, int column, XSSFSheet sheet, Callback callback) throws IOException {
        columnsUsed.remove(column);
        
		FileInputStream stream = new FileInputStream("Dominions5.exe");			
		stream.skip(Starts.MONSTER);
		int rowNumber = 1;
		int i = 0;
		byte[] c = new byte[2];
		stream.skip(skip);
		while ((stream.read(c, 0, 2)) != -1) {
			String high = String.format("%02X", c[1]);
			String low = String.format("%02X", c[0]);
			XSSFRow row = sheet.getRow(rowNumber);
			rowNumber++;
			XSSFCell cell = row.getCell(column, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
			int value = Integer.decode("0X" + high + low);
			if (callback != null) {
				cell.setCellValue(callback.found(Integer.toString(value)));
			} else {
				cell.setCellValue(Integer.decode("0X" + high + low));
			}
			stream.skip(262l);
			i++;
			if (i >= Starts.MONSTER_COUNT) {
				break;
			}
		}
		stream.close();
	}
	
	private static boolean[] largeBitmap(String fieldName) throws IOException {
		boolean[] boolArray = largeBitmapCache.get(fieldName);
		if (boolArray != null) {
			return boolArray;
		}
		FileInputStream stream = new FileInputStream("Dominions5.exe");			
		stream.skip(Starts.MONSTER);
		int i = 0;
		byte[] c = new byte[16];
		boolArray = new boolean[Starts.MONSTER_COUNT];
		stream.skip(248);
		while ((stream.read(c, 0, 16)) != -1) {
			boolean found = false;
			String high = String.format("%02X", c[1]);
			String low = String.format("%02X", c[0]);
			int val = Integer.decode("0X" + high + low);
			if (val > 0) {
				for (int j=0; j < 16; j++) {
					if ((val & MASK[j]) != 0) {
						if (value1[j].equals(fieldName)) {
							found = true;
						}
					}
				}
			}
			high = String.format("%02X", c[3]);
			low = String.format("%02X", c[2]);
			val = Integer.decode("0X" + high + low);
			if (val > 0) {
				for (int j=0; j < 16; j++) {
					if ((val & MASK[j]) != 0) {
						if (value2[j].equals(fieldName)) {
							found = true;
						}
					}
				}
			}
			high = String.format("%02X", c[5]);
			low = String.format("%02X", c[4]);
			val = Integer.decode("0X" + high + low);
			if (val > 0) {
				for (int j=0; j < 16; j++) {
					if ((val & MASK[j]) != 0) {
						if (value3[j].equals(fieldName)) {
							found = true;
						}
					}
				}
			}
			high = String.format("%02X", c[7]);
			low = String.format("%02X", c[6]);
			val = Integer.decode("0X" + high + low);
			if (val > 0) {
				for (int j=0; j < 16; j++) {
					if ((val & MASK[j]) != 0) {
						if (value4[j].equals(fieldName)) {
							found = true;
						}
					}
				}
			}
			high = String.format("%02X", c[9]);
			low = String.format("%02X", c[8]);
			val = Integer.decode("0X" + high + low);
			if (val > 0) {
				for (int j=0; j < 16; j++) {
					if ((val & MASK[j]) != 0) {
						if (value5[j].equals(fieldName)) {
							found = true;
						}
					}
				}
			}
			high = String.format("%02X", c[11]);
			low = String.format("%02X", c[10]);
			val = Integer.decode("0X" + high + low);
			if (val > 0) {
				for (int j=0; j < 16; j++) {
					if ((val & MASK[j]) != 0) {
						if (value6[j].equals(fieldName)) {
							found = true;
						}
					}
				}
			}
			high = String.format("%02X", c[13]);
			low = String.format("%02X", c[12]);
			val = Integer.decode("0X" + high + low);
			if (val > 0) {
				for (int j=0; j < 16; j++) {
					if ((val & MASK[j]) != 0) {
						if (value7[j].equals(fieldName)) {
							found = true;
						}
					}
				}
			}
			high = String.format("%02X", c[15]);
			low = String.format("%02X", c[14]);
			val = Integer.decode("0X" + high + low);
			if (val > 0) {
				for (int j=0; j < 16; j++) {
					if ((val & MASK[j]) != 0) {
						if (value8[j].equals(fieldName)) {
							found = true;
						}
					}
				}
			}
			boolArray[i] = found;
			stream.skip(248l);
			i++;
			if (i >= Starts.MONSTER_COUNT) {
				break;
			}
		}
		stream.close();
		largeBitmapCache.put(fieldName, boolArray);
		return boolArray;
	}
	
	private static String getAttr(String key, int id) throws IOException {
		String attributeString = attrCache.get(new CacheKey(key, id));
		if (attributeString != null) {
			return attributeString;
		}

		FileInputStream stream = new FileInputStream("Dominions5.exe");			
		stream.skip(Starts.MONSTER);
		int k = 0;
		int pos = -1;
		long numFound = 0;
		byte[] c = new byte[2];
		stream.skip(64);
		stream.skip(264*(id-1));
		while ((stream.read(c, 0, 2)) != -1) {
			String high = String.format("%02X", c[1]);
			String low = String.format("%02X", c[0]);
			int weapon = Integer.decode("0X" + high + low);
			if (weapon == 0) {
				boolean found = false;
				int value = 0;
				stream.skip(46l - numFound*2l);
				// Values
				for (int x = 0; x < numFound; x++) {
					byte[] d = new byte[4];
					stream.read(d, 0, 4);
					String high1 = String.format("%02X", d[3]);
					String low1 = String.format("%02X", d[2]);
					high = String.format("%02X", d[1]);
					low = String.format("%02X", d[0]);
					//System.out.print(low + high + " ");
					if (x == pos) {
						value = new BigInteger(high1 + low1 + high + low, 16).intValue();
						//System.out.print(fire);
						found = true;
					}
					//stream.skip(2);
				}
				
				stream.close();
				if (found) {
					attrCache.put(new CacheKey(key, id), Integer.toString(value));
					return Integer.toString(value);
				} else {
					attrCache.put(new CacheKey(key, id), "");
					return "";
				}
			} else {
				//System.out.print(low + high + " ");
				if ((low + high).equals(key)) {
					pos = k;
				}
				k++;
				numFound++;
			}				
		}
		stream.close();
		return null;
	}
	
	private static void doit2(XSSFSheet sheet, String attr, int column) throws IOException {
		doit2(sheet, attr, column, null);
	}
	
	private static void doit2(XSSFSheet sheet, String attr, int column, Callback callback) throws IOException {
		doit2(sheet, attr, column, callback, false);
	}
	private static void doit2(XSSFSheet sheet, String attr, int column, Callback callback, boolean append) throws IOException {
        columnsUsed.remove(column);

        FileInputStream stream = new FileInputStream("Dominions5.exe");			
		stream.skip(Starts.MONSTER);
		int rowNumber = 1;
		int i = 0;
		int k = 0;
		int pos = -1;
		long numFound = 0;
		byte[] c = new byte[2];
		stream.skip(64);
		while ((stream.read(c, 0, 2)) != -1) {
			String high = String.format("%02X", c[1]);
			String low = String.format("%02X", c[0]);
			int weapon = Integer.decode("0X" + high + low);
			if (weapon == 0) {
				boolean found = false;
				int value = 0;
				stream.skip(46l - numFound*2l);
				// Values
				for (int x = 0; x < numFound; x++) {
					byte[] d = new byte[4];
					stream.read(d, 0, 4);
					String high1 = String.format("%02X", d[3]);
					String low1 = String.format("%02X", d[2]);
					high = String.format("%02X", d[1]);
					low = String.format("%02X", d[0]);
					//System.out.print(low + high + " ");
					if (x == pos) {
						value = new BigInteger(high1 + low1 + high + low, 16).intValue();
						//System.out.print(fire);
						found = true;
					}
					//stream.skip(2);
				}
				
				//System.out.println("");
				XSSFRow row = sheet.getRow(rowNumber);
				rowNumber++;
				XSSFCell cell = row.getCell(column, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				if (found) {
					if (callback == null) {
						if (append) {
							String origVal = cell.getStringCellValue();
							cell.setCellValue(origVal + value);
						} else {
							cell.setCellValue(value);
						}
					} else {
						if (append) {
							String origVal = cell.getStringCellValue();
							if (callback.found(Integer.toString(value)) != null) {
								cell.setCellValue(origVal + callback.found(Integer.toString(value)));
							}
						} else {
							if (callback.found(Integer.toString(value)) != null) {
								cell.setCellValue(callback.found(Integer.toString(value)));
							}
						}
					}
				} else {
					if (callback == null) {
						cell.setCellValue("");
					} else {
						if (callback.notFound() != null) {
							cell.setCellValue(callback.notFound());
						}
					}
				}
				stream.skip(262l - 46l - numFound*4l);
				numFound = 0;
				pos = -1;
				k = 0;
				i++;
			} else {
				//System.out.print(low + high + " ");
				if (attr.indexOf(low + high) != -1) {
					if (pos == -1) {
						pos = k;
					}
				}
				k++;
				numFound++;
			}				
			if (i >= Starts.MONSTER_COUNT) {
				break;
			}
		}
		stream.close();
	}
	
	public static void main(String[] args) {
		run();
	}
	
	public static void run() {
		FileInputStream stream = null;
		try {
	        long startIndex = Starts.MONSTER;
	        int ch;

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(startIndex);
			
			XSSFWorkbook wb = MonsterStatIndexer.readFile("BaseU.xlsx");
			FileOutputStream fos = new FileOutputStream("NewBaseU.xlsx");
			XSSFSheet sheet = wb.getSheetAt(0);

			XSSFRow titleRow = sheet.getRow(0);
			int cellNum = 0;
			XSSFCell titleCell = titleRow.getCell(cellNum);
			Set<String> skip = new HashSet<String>(Arrays.asList(SkipColumns));
			while (titleCell != null) {
				String stringCellValue = titleCell.getStringCellValue();
				if (!skip.contains(stringCellValue)) {
					columnsUsed.put(cellNum, stringCellValue);
				}
				cellNum++;
				titleCell = titleRow.getCell(cellNum);
			}
			
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

				stream = new FileInputStream("Dominions5.exe");		
				startIndex = startIndex + 264l;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);

				//System.out.println(name);

				XSSFRow row = sheet.getRow(rowNumber);
				XSSFCell cell1 = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell1.setCellValue(rowNumber);
				rowNumber++;
				XSSFCell cell = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(name.toString());
			}
			in.close();
			stream.close();

	        // AP
			doit1(40, 31, sheet);

			// MM
			doit1(42, 30, sheet);

			// Size
			doit1(44, 19, sheet);
			
			// ressize
			doit1(44, 20, sheet);
			doit2(sheet, "0901", 20, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			
			// HP
			doit1(46, 21, sheet);

			// Prot
			doit1(48, 22, sheet);

			// STR
			doit1(50, 25, sheet);

			// ENC
			doit1(52, 29, sheet);

			// Prec
			doit1(54, 28, sheet);

			// ATT
			doit1(56, 26, sheet);

			// Def
			doit1(58, 27, sheet);

			// MR
			doit1(60, 23, sheet);

			// Mor
			doit1(62, 24, sheet);

			// wpn1
			doit1(208, 2, sheet, new CallbackAdapter(){
				@Override
				public String found(String value) {
					if (Integer.parseInt(value) == 0) {
						return "";
					}
					return value;
				}
			});

			// wpn2
			doit1(210, 3, sheet, new CallbackAdapter(){
				@Override
				public String found(String value) {
					if (Integer.parseInt(value) == 0) {
						return "";
					}
					return value;
				}
			});

			// wpn3
			doit1(212, 4, sheet, new CallbackAdapter(){
				@Override
				public String found(String value) {
					if (Integer.parseInt(value) == 0) {
						return "";
					}
					return value;
				}
			});

			// wpn4
			doit1(214, 5, sheet, new CallbackAdapter(){
				@Override
				public String found(String value) {
					if (Integer.parseInt(value) == 0) {
						return "";
					}
					return value;
				}
			});

			// wpn5
			doit1(216, 6, sheet, new CallbackAdapter(){
				@Override
				public String found(String value) {
					if (Integer.parseInt(value) == 0) {
						return "";
					}
					return value;
				}
			});

			// wpn6
			doit1(218, 7, sheet, new CallbackAdapter(){
				@Override
				public String found(String value) {
					if (Integer.parseInt(value) == 0) {
						return "";
					}
					return value;
				}
			});

			// wpn7
			doit1(220, 8, sheet, new CallbackAdapter(){
				@Override
				public String found(String value) {
					if (Integer.parseInt(value) == 0) {
						return "";
					}
					return value;
				}
			});

			// armor1
			doit1(228, 9, sheet, new CallbackAdapter(){
				@Override
				public String found(String value) {
					if (Integer.parseInt(value) == 0) {
						return "";
					}
					return value;
				}
			});

			// armor2
			doit1(230, 10, sheet, new CallbackAdapter(){
				@Override
				public String found(String value) {
					if (Integer.parseInt(value) == 0) {
						return "";
					}
					return value;
				}
			});

			// armor3
			doit1(232, 11, sheet, new CallbackAdapter(){
				@Override
				public String found(String value) {
					if (Integer.parseInt(value) == 0) {
						return "";
					}
					return value;
				}
			});

			// basecost
			doit1(234, 15, sheet);

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.MONSTER);
			rowNumber = 1;
			// res
			int i = 0;
			byte[] c = new byte[2];
			stream.skip(236);
			while ((stream.read(c, 0, 2)) != -1) {
				String high = String.format("%02X", c[1]);
				String low = String.format("%02X", c[0]);
				//System.out.println(Integer.decode("0X" + high + low));
				XSSFRow row = sheet.getRow(rowNumber);
				rowNumber++;
				XSSFCell cell = row.getCell(18, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				int tmp = new BigInteger(high + low, 16).intValue();
				if (tmp < 1000) {
					cell.setCellValue(Integer.decode("0X" + high + low));
				} else {
					cell.setCellValue(new BigInteger("FFFF" + high + low, 16).intValue());
				}
				stream.skip(262l);
				i++;
				if (i >= Starts.MONSTER_COUNT) {
					break;
				}
			}
			stream.close();

			// additional leadership
			doit2(sheet, "9D00", 35, new Callback() {
				
				@Override
				public String notFound() {
					return "40";
				}
				
				@Override
				public String found(String value) {
					return Integer.toString(Integer.parseInt(value)+40);
				}
			});
			
			// itemslots defaults
			doit2(sheet, "B600", 40, new CallbackAdapter(){
				// hand
				@Override
				public String notFound() {
					return "2";
				}
			});
			doit2(sheet, "B600", 41, new CallbackAdapter(){
				// head
				@Override
				public String notFound() {
					return "1";
				}
			});
			doit2(sheet, "B600", 42, new CallbackAdapter(){
				// body
				@Override
				public String notFound() {
					return "1";
				}
			});
			doit2(sheet, "B600", 43, new CallbackAdapter(){
				// foot
				@Override
				public String notFound() {
					return "1";
				}
			});
			doit2(sheet, "B600", 44, new CallbackAdapter(){
				// misc
				@Override
				public String notFound() {
					return "2";
				}
			});
			
			// Large bitmap
			for (String[] pair : DOTHESE) {
				columnsUsed.remove(Integer.parseInt(pair[1]));
				rowNumber = 1;
				boolean[] boolArray = largeBitmap(pair[0]);
				for (boolean found : boolArray) {
					XSSFRow row = sheet.getRow(rowNumber);
					rowNumber++;
					XSSFCell cell = row.getCell(Integer.parseInt(pair[1]), Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					if (found) {
						if (pair[0].equals("slow_to_recruit")) {
							cell.setCellValue(2);
						} else if (pair[0].equals("heat")) {
							cell.setCellValue(3);
						} else if (pair[0].equals("cold")) {
							cell.setCellValue(3);
						} else if (pair[0].equals("stealthy40")) {
							String additional = getAttr("6C00", rowNumber-1);
							boolean glamour = false;
							if (largeBitmap("illusion")[rowNumber-2]) {
								glamour = true;
							}
							cell.setCellValue(40 + Integer.parseInt(additional.equals("")?"0":additional) + (glamour?25:0));
						} else if (pair[0].equals("coldres15")) {
							String additional = getAttr("C900", rowNumber-1);
							boolean cold = false;
							if (largeBitmap("cold")[rowNumber-2] || getAttr("DC00", rowNumber-1).length() > 0) {
								cold = true;
							}
							cell.setCellValue(15 + Integer.parseInt(additional.equals("")?"0":additional) + (cold?10:0));
						} else if (pair[0].equals("fireres15")) {
							String additional = getAttr("C600", rowNumber-1);
							boolean heat = false;
							if (largeBitmap("heat")[rowNumber-2] || getAttr("3C01", rowNumber-1).length() > 0) {
								heat = true;
							}
							cell.setCellValue(15 + Integer.parseInt(additional.equals("")?"0":additional) + (heat?10:0));
						} else if (pair[0].equals("poisonres15")) {
							String additional = getAttr("C800", rowNumber-1);
							boolean poisoncloud = false;
							if (largeBitmap("undead")[rowNumber-2] || largeBitmap("inanimate")[rowNumber-2] || getAttr("6A00", rowNumber-1).length() > 0) {
								poisoncloud = true;
							}
							cell.setCellValue(15 + Integer.parseInt(additional.equals("")?"0":additional) + (poisoncloud?10:0));
						} else if (pair[0].equals("shockres15")) {
							String additional = getAttr("C700", rowNumber-1);
							cell.setCellValue(15 + Integer.parseInt(additional.equals("")?"0":additional));
						} else if (pair[0].equals("noleader")) {
							String additional = getAttr("9D00", rowNumber-1);
							XSSFCell baseLeaderCell = row.getCell(248, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
							baseLeaderCell.setCellValue("0");
							if (!"".equals(additional)) {
								cell.setCellValue(additional);
							} else {
								cell.setCellValue("0");
							}
						} else if (pair[0].equals("poorleader")) {
							String additional = getAttr("9D00", rowNumber-1);
							XSSFCell baseLeaderCell = row.getCell(248, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
							baseLeaderCell.setCellValue("10");
							if (!"".equals(additional)) {
								cell.setCellValue(Integer.toString(10+Integer.parseInt(additional)));
							} else {
								cell.setCellValue("10");
							}
						} else if (pair[0].equals("goodleader")) {
							String additional = getAttr("9D00", rowNumber-1);
							XSSFCell baseLeaderCell = row.getCell(248, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
							baseLeaderCell.setCellValue("80");
							if (!"".equals(additional)) {
								cell.setCellValue(Integer.toString(80+Integer.parseInt(additional)));
							} else {
								cell.setCellValue("80");
							}
						} else if (pair[0].equals("expertleader")) {
							String additional = getAttr("9D00", rowNumber-1);
							XSSFCell baseLeaderCell = row.getCell(248, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
							baseLeaderCell.setCellValue("120");
							if (!"".equals(additional)) {
								cell.setCellValue(Integer.toString(120+Integer.parseInt(additional)));
							} else {
								cell.setCellValue("120");
							}
						} else if (pair[0].equals("superiorleader")) {
							String additional = getAttr("9D00", rowNumber-1);
							XSSFCell baseLeaderCell = row.getCell(248, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
							baseLeaderCell.setCellValue("160");
							if (!"".equals(additional)) {
								cell.setCellValue(Integer.toString(160+Integer.parseInt(additional)));
							} else {
								cell.setCellValue("160");
							}
						} else if (pair[0].equals("poormagicleader")) {
							cell.setCellValue("10");
						} else if (pair[0].equals("okmagicleader")) {
							cell.setCellValue("40");
						} else if (pair[0].equals("goodmagicleader")) {
							cell.setCellValue("80");
						} else if (pair[0].equals("expertmagicleader")) {
							cell.setCellValue("120");
						} else if (pair[0].equals("superiormagicleader")) {
							cell.setCellValue("160");
						} else if (pair[0].equals("poorundeadleader")) {
							cell.setCellValue("10");
						} else if (pair[0].equals("okundeadleader")) {
							cell.setCellValue("40");
						} else if (pair[0].equals("goodundeadleader")) {
							cell.setCellValue("80");
						} else if (pair[0].equals("expertundeadleader")) {
							cell.setCellValue("120");
						} else if (pair[0].equals("superiorundeadleader")) {
							cell.setCellValue("160");
						} else if (pair[0].equals("misc2")) {
							XSSFCell handCell = row.getCell(40, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
							handCell.setCellValue(0);
							XSSFCell headCell = row.getCell(41, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
							headCell.setCellValue(0);
							XSSFCell bodyCell = row.getCell(42, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
							bodyCell.setCellValue(0);
							XSSFCell footCell = row.getCell(43, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
							footCell.setCellValue(0);
							XSSFCell miscCell = row.getCell(44, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
							miscCell.setCellValue(2);
						} else if (pair[0].equals("mounted")) {
							XSSFCell footCell = row.getCell(43, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
							footCell.setCellValue(0);
							cell.setCellValue(1);
						} else {
							cell.setCellValue(1);
						}
					} else {
						if (pair[0].equals("slow_to_recruit")) {
							cell.setCellValue(1);
						} else if (pair[0].equals("heat")) {
							cell.setCellValue(getAttr("3C01", rowNumber-1));
						} else if (pair[0].equals("cold")) {
							cell.setCellValue(getAttr("DC00", rowNumber-1));
						} else if (pair[0].equals("coldres15")) {
							boolean cold = false;
							if (largeBitmap("cold")[rowNumber-2] || getAttr("DC00", rowNumber-1).length() > 0) {
								cold = true;
							}
							String additional = getAttr("C900", rowNumber-1);
							int coldres = Integer.parseInt(additional.equals("")?"0":additional) + (cold?10:0);
							cell.setCellValue(coldres==0?"":Integer.toString(coldres));
						} else if (pair[0].equals("fireres15")) {
							boolean heat = false;
							if (largeBitmap("heat")[rowNumber-2] || getAttr("3C01", rowNumber-1).length() > 0) {
								heat = true;
							}
							String additional = getAttr("C600", rowNumber-1);
							int coldres = Integer.parseInt(additional.equals("")?"0":additional) + (heat?10:0);
							cell.setCellValue(coldres==0?"":Integer.toString(coldres));
						} else if (pair[0].equals("poisonres15")) {
							boolean poisoncloud = false;
							if (largeBitmap("undead")[rowNumber-2] || largeBitmap("inanimate")[rowNumber-2] || getAttr("6A00", rowNumber-1).length() > 0) {
								poisoncloud = true;
							}
							String additional = getAttr("C800", rowNumber-1);
							int poisonres = Integer.parseInt(additional.equals("")?"0":additional) + (poisoncloud?10:0);
							cell.setCellValue(poisonres==0?"":Integer.toString(poisonres));
						} else if (pair[0].equals("shockres15")) {
							String additional = getAttr("C700", rowNumber-1);
							int shockres = Integer.parseInt(additional.equals("")?"0":additional);
							cell.setCellValue(shockres==0?"":Integer.toString(shockres));
						} else if (pair[0].equals("stealthy40")) {
							String additional = getAttr("6C00", rowNumber-1);
							cell.setCellValue(additional == null || additional.equals("") ? "" : additional);
						} else if (pair[0].equals("immobile")
								|| pair[0].equals("teleport")
								|| pair[0].equals("float")
								|| pair[0].equals("bluntres")
								|| pair[0].equals("slashres")
								|| pair[0].equals("pierceres")
								) {
							cell.setCellValue("");
						} else if (pair[0].equals("noleader")
								|| pair[0].equals("poorleader")
								|| pair[0].equals("goodleader")
								|| pair[0].equals("expertleader")
								|| pair[0].equals("superiorleader")
								|| pair[0].equals("poormagicleader")
								|| pair[0].equals("okmagicleader")
								|| pair[0].equals("goodmagicleader")
								|| pair[0].equals("expertmagicleader")
								|| pair[0].equals("superiormagicleader")
								|| pair[0].equals("poorundeadleader")
								|| pair[0].equals("okundeadleader")
								|| pair[0].equals("goodundeadleader")
								|| pair[0].equals("expertundeadleader")
								|| pair[0].equals("superiorundeadleader")
								|| pair[0].equals("misc2")
								) {
						} else {
							cell.setCellValue("");
						}
					}
				}
			}
			
			/*stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.MONSTER);
			int i = 0;
			byte[] c = new byte[16];
			stream.skip(248);
			while ((stream.read(c, 0, 16)) != -1) {
				boolean found = false;
				System.out.print("(" + (i+1) + ") ");
				String high = String.format("%02X", c[1]);
				String low = String.format("%02X", c[0]);
				//System.out.print(high + low + " ");
				int val = Integer.decode("0X" + high + low);
				if (val > 0) {
					System.out.print("1:{");
					for (int j=0; j < 16; j++) {
						if ((val & MASK[j]) != 0) {
							System.out.print((found?",":"") + (value1[j].equals("")?("*****"+(j+1)+"*****"):value1[j]));
							found = true;
						}
					}
					System.out.print("}");
				}
				high = String.format("%02X", c[3]);
				low = String.format("%02X", c[2]);
				//System.out.print(high + low + " ");
				val = Integer.decode("0X" + high + low);
				if (val > 0) {
					System.out.print(" 2:{");
					found = false;
					for (int j=0; j < 16; j++) {
						if ((val & MASK[j]) != 0) {
							System.out.print((found?",":"") + (value2[j].equals("")?("*****"+(j+1)+"*****"):value2[j]));
							found = true;
						}
					}
					System.out.print("}");
				}
				high = String.format("%02X", c[5]);
				low = String.format("%02X", c[4]);
				//System.out.print(high + low + " ");
				val = Integer.decode("0X" + high + low);
				if (val > 0) {
					System.out.print(" 3:{");
					found = false;
					for (int j=0; j < 16; j++) {
						if ((val & MASK[j]) != 0) {
							System.out.print((found?",":"") + (value3[j].equals("")?("*****"+(j+1)+"*****"):value3[j]));
							found = true;
						}
					}
					System.out.print("}");
				}

				high = String.format("%02X", c[7]);
				low = String.format("%02X", c[6]);
				//System.out.print(high + low + " ");
				val = Integer.decode("0X" + high + low);
				if (val > 0) {
					System.out.print(" 4:{");
					found = false;
					for (int j=0; j < 16; j++) {
						if ((val & MASK[j]) != 0) {
							System.out.print((found?",":"") + (value4[j].equals("")?("*****"+(j+1)+"*****"):value4[j]));
							found = true;
						}
					}
					System.out.print("}");
				}

				high = String.format("%02X", c[9]);
				low = String.format("%02X", c[8]);
				//System.out.print(high + low + " ");
				val = Integer.decode("0X" + high + low);
				if (val > 0) {
					System.out.print(" 5:{");
					found = false;
					for (int j=0; j < 16; j++) {
						if ((val & MASK[j]) != 0) {
							System.out.print((found?",":"") + (value5[j].equals("")?("*****"+(j+1)+"*****"):value5[j]));
							found = true;
						}
					}
					System.out.print("}");
				}
				high = String.format("%02X", c[11]);
				low = String.format("%02X", c[10]);
				//System.out.print(high + low + " ");
				val = Integer.decode("0X" + high + low);
				if (val > 0) {
					System.out.print(" 6:{");
					found = false;
					for (int j=0; j < 16; j++) {
						if ((val & MASK[j]) != 0) {
							System.out.print((found?",":"") + (value6[j].equals("")?(j+1):value6[j]));
							found = true;
						}
					}
					System.out.print("}");
				}
				high = String.format("%02X", c[13]);
				low = String.format("%02X", c[12]);
				//System.out.print(high + low + " ");
				val = Integer.decode("0X" + high + low);
				if (val > 0) {
					System.out.print(" 7:{");
					found = false;
					for (int j=0; j < 16; j++) {
						if ((val & MASK[j]) != 0) {
							System.out.print((found?",":"") + (value7[j].equals("")?(j+1):value7[j]));
							found = true;
						}
					}
					System.out.print("}");
				}
				high = String.format("%02X", c[15]);
				low = String.format("%02X", c[14]);
				//System.out.print(high + low + " ");
				val = Integer.decode("0X" + high + low);
				if (val > 0) {
					System.out.print(" 8:{");
					found = false;
					for (int j=0; j < 16; j++) {
						if ((val & MASK[j]) != 0) {
							System.out.print((found?",":"") + (value8[j].equals("")?(j+1):value8[j]));
							found = true;
						}
					}
					System.out.print("}");
				}
				
				if (!found) {
					//System.out.println("");
				}
				System.out.println(" ");
				stream.skip(248l);
				i++;
				if (i >= Starts.MONSTER_COUNT) {
					break;
				}
			}
			stream.close();*/

			// realm
			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.MONSTER);
			i = 0;
			int k = 0;
			Set<Integer> posSet = new HashSet<Integer>();
			long numFound = 0;
			c = new byte[2];
			stream.skip(64);
			while ((stream.read(c, 0, 2)) != -1) {
				String high = String.format("%02X", c[1]);
				String low = String.format("%02X", c[0]);
				int weapon = Integer.decode("0X" + high + low);
				if (weapon == 0) {
					stream.skip(46l - numFound*2l);
					int numRealms = 0;
					// Values
					for (int x = 0; x < numFound; x++) {
						byte[] d = new byte[4];
						stream.read(d, 0, 4);
						String high1 = String.format("%02X", d[3]);
						String low1 = String.format("%02X", d[2]);
						high = String.format("%02X", d[1]);
						low = String.format("%02X", d[0]);
						//System.out.print(low + high + " ");
						if (posSet.contains(x)) {
							int fire = new BigInteger(high1 + low1 + high + low, 16).intValue();//Integer.decode("0X" + high + low);
							//System.out.print(i+1 + "\t" + fire);
							//System.out.println("");
							XSSFRow row = sheet.getRow(i+1);
							XSSFCell cell = row.getCell(233+numRealms, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);							
							cell.setCellValue(fire);
							numRealms++;
						}
						//stream.skip(2);
					}
					
//					System.out.println("");
					stream.skip(262l - 46l - numFound*4l);
					numFound = 0;
					posSet.clear();
					k = 0;
					i++;
				} else {
					//System.out.print(low + high + " ");
					if ((low + high).equals("5A02")) {
						posSet.add(k);
					}
					k++;
					numFound++;
				}				
				if (i >= Starts.MONSTER_COUNT) {
					break;
				}
			}
			stream.close();
			
			// patience
			doit2(sheet, "CD01", 104);

			// stormimmune
			doit2(sheet, "AF00", 94);
			
			// regeneration
			doit2(sheet, "BD00", 151);

			// secondshape
			doit2(sheet, "C200", 210);

			// firstshape
			doit2(sheet, "C300", 209);

			// shapechange
			doit2(sheet, "C100", 208);

			// secondtmpshape
			doit2(sheet, "C400", 211);
			
			// landshape
			doit2(sheet, "F500", 212);
			
			// watershape
			doit2(sheet, "F600", 213);
			
			// forestshape
			doit2(sheet, "4201", 214);
			
			// plainshape
			doit2(sheet, "4301", 215);
			
			// damagerev
			doit2(sheet, "CA00", 144, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return Integer.toString(Integer.parseInt(value)-1);
				}
			});

			// bloodvengeance
			doit2(sheet, "9E01", 231);
			
			// nobadevents
			doit2(sheet, "EB00", 178);
			
			// bringeroffortune
			doit2(sheet, "E400", 232);
			
			// darkvision
			doit2(sheet, "1901", 133);
			
			// fear
			doit2(sheet, "B700", 138);
			
			// voidsanity
			doit2(sheet, "1501", 132);
			
			// standard
			doit2(sheet, "6700", 117);
			
			// formationfighter
			doit2(sheet, "6E01", 115);
			
			// undisciplined
			doit2(sheet, "6F01", 114);
			
			// bodyguard
			doit2(sheet, "9801", 121);
			
			// summon
			doit2(sheet, "A400,A500,A600", 223);
			doit2(sheet, "A400", 224, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return "1";
				}
				@Override
				public String notFound() {
					return null;
				}
			});
			doit2(sheet, "A500", 224, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return "2";
				}
				@Override
				public String notFound() {
					return null;
				}
			});
			doit2(sheet, "A600", 224, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return "3";
				}
				@Override
				public String notFound() {
					return null;
				}
			});
			doit2(sheet, "A400,A500,A600", 224, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return null;
				}
			});

			// inspirational
			doit2(sheet, "7001", 118);
			
			// pillagebonus
			doit2(sheet, "8300", 187);
			
			// berserk
			doit2(sheet, "BE00", 139);
			
			// pathcost
			doit2(sheet, "F300", 45);
			
			// default startdom
			doit2(sheet, "F300", 46, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return "1";
				}
			});
			// startdom
			doit2(sheet, "F200", 46, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			
			// waterbreathing
			doit2(sheet, "6F00", 122);

			// batstartsum1
			doit2(sheet, "B401", 227);
			// batstartsum2
			doit2(sheet, "B501", 228);
			// batstartsum3
			doit2(sheet, "B601", 236);
			// batstartsum4
			doit2(sheet, "B701", 237);
			// batstartsum5
			doit2(sheet, "B801", 238);

			// batstartsum1d6
			doit2(sheet, "B901", 239);
			// batstartsum2d6
			doit2(sheet, "BA01", 240);
			// batstartsum3d6
			doit2(sheet, "BB01", 241);
			// batstartsum4d6
			doit2(sheet, "BC01", 242);
			// batstartsum5d6
			doit2(sheet, "BD01", 243);
			// batstartsum6d6
			doit2(sheet, "BE01", 244);

			// autosummon (#summon1-5)
			doit2(sheet, "6B00,8F00", 225);
			doit2(sheet, "6B00", 226, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return "?";
				}
				@Override
				public String notFound() {
					return null;
				}
			});
			doit2(sheet, "8F00", 226, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return "?";
				}
				@Override
				public String notFound() {
					return null;
				}
			});
			doit2(sheet, "6B00,8F00", 226, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return null;
				}
			});
			
			// domsummon
			doit2(sheet, "A101,DB00,F100", 229);
			doit2(sheet, "A101", 230, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return "?";
				}
				@Override
				public String notFound() {
					return null;
				}
			});
			doit2(sheet, "DB00", 230, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return "?";
				}
				@Override
				public String notFound() {
					return null;
				}
			});
			doit2(sheet, "F100", 230, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return "?";
				}
				@Override
				public String notFound() {
					return null;
				}
			});
			doit2(sheet, "A101,DB00,F100", 230, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return null;
				}
			});
			
			// turmoil summon
			doit2(sheet, "AD00", 245);
			
			// cold summon
			doit2(sheet, "9200", 246);
			
			// stormpower
			doit2(sheet, "AE00", 161);
			
			// firepower
			doit2(sheet, "B100", 162);

			// coldpower
			doit2(sheet, "B000", 163);

			// darkpower
			doit2(sheet, "2501", 164);
			
			// chaospower
			doit2(sheet, "A001", 165);
			
			// magicpower
			doit2(sheet, "4401", 166);
			
			// winterpower
			doit2(sheet, "EA00", 167);
			
			// springpower
			doit2(sheet, "E700", 168);
			
			// summerpower
			doit2(sheet, "E800", 169);
			
			// fallpower
			doit2(sheet, "E900", 170);
			
			// nametype
			doit2(sheet, "FB00", 219);
			
			// blind
			doit2(sheet, "AB00", 134);
			
			// eyes
			doit2(sheet, "B200", 279, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return Integer.toString((Integer.parseInt(value) + 2));
				}
			});
			
			// supplybonus
			doit2(sheet, "7A00", 192);
			
			// slave
			doit2(sheet, "7C01", 116);
			
			// awe
			doit2(sheet, "6900", 136);
			
			// siegebonus
			doit2(sheet, "7D00", 190);
			
			// researchbonus
			doit2(sheet, "7900", 195);
			
			// chaosrec
			doit2(sheet, "CA01", 186);
			
			// invulnerability
			doit2(sheet, "7E01", 124);
			
			// iceprot
			doit2(sheet, "BF00", 123);
			
			// reinvigoration
			doit2(sheet, "7500", 34);
			
			// ambidextrous
			doit2(sheet, "D900", 32);
			
			// spy
			doit2(sheet, "D600", 102);
			
			// scale walls
			doit2(sheet, "E301", 247);
			
			// dream seducer (succubus)
			doit2(sheet, "D200", 106);
			
			// seduction
			doit2(sheet, "2A01", 105);
			
			// assassin
			doit2(sheet, "D500", 103);
			
			// explode on death
			doit2(sheet, "2901", 249);
			
			// taskmaster
			doit2(sheet, "7B01", 119);
			
			// unique
			doit2(sheet, "1301", 216);
			
			// poisoncloud
			doit2(sheet, "6A00", 145);
			
			// startaff
			doit2(sheet, "B900", 250);
			
			// uwregen
			doit2(sheet, "DD00", 251);
			
			// patrolbonus
			doit2(sheet, "AA00", 188);
			
			// castledef
			doit2(sheet, "D700", 189);
			
			// sailingshipsize
			doit2(sheet, "7000", 98);
			
			// sailingmaxunitsize
			doit2(sheet, "9A01", 99);
			
			// incunrest
			doit2(sheet, "DF00", 193, new CallbackAdapter() {
				@Override
				public String found(String value) {
					double val = Double.parseDouble(value)/10d;
					if (val < 1 && val > -1) {
						return Double.toString(val);
					}
					return Integer.toString((int)val);
				}
			});
			
			// barbs
			doit2(sheet, "BC00", 153, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return "1";
				}
			});
			
			// inn
			doit2(sheet, "4E01", 47);
			
			// stonebeing
			doit2(sheet, "1801", 80);
			
			// shrinkhp
			doit2(sheet, "5001", 252);
			
			// growhp
			doit2(sheet, "4F01", 253);
			
			// transformation
			doit2(sheet, "FD01", 254);
			
			// heretic
			doit2(sheet, "B800", 206);
			
			// popkill
			doit2(sheet, "2001", 194, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return Integer.toString(Integer.parseInt(value)*10);
				}
			});
			
			// autohealer
			doit2(sheet, "6201", 175);
			
			// fireshield
			doit2(sheet, "A300", 142);
			
			// startingaff
			doit2(sheet, "E200", 255);
			
			// fixedresearch
			doit2(sheet, "F800", 256);
			
			// divineins
			doit2(sheet, "C201", 257);
			
			// halt
			doit2(sheet, "4701", 137);
			
			// crossbreeder
			doit2(sheet, "AF01", 201);

			// reclimit
			doit2(sheet, "7D01", 14);
			
			// fixforgebonus
			doit2(sheet, "C501", 172);

			// mastersmith
			doit2(sheet, "6B01", 173);

			// lamiabonus
			doit2(sheet, "A900", 258);

			// homesick
			doit2(sheet, "FD00", 113);

			// banefireshield
			doit2(sheet, "DE00", 143);

			// animalawe
			doit2(sheet, "A200", 135);

			// autodishealer
			doit2(sheet, "6301", 176);

			// shatteredsoul
			doit2(sheet, "4801", 181);

			// voidsum
			doit2(sheet, "CE00", 205);

			// makepearls
			doit2(sheet, "AE01", 202);

			// inspiringres
			doit2(sheet, "5501", 197);

			// drainimmune
			doit2(sheet, "1101", 196);
			
			// diseasecloud
			doit2(sheet, "AC00", 146);

			// inquisitor
			doit2(sheet, "D300", 74);

			// beastmaster
			doit2(sheet, "7101", 120);

			// douse
			doit2(sheet, "7400", 198);

			// preanimator
			doit2(sheet, "6C01", 259);

			// dreanimator
			doit2(sheet, "6D01", 260);

			// mummify
			doit2(sheet, "FF01", 261);
			
			// onebattlespell
			doit2(sheet, "D100", 262, new CallbackAdapter() {
				// Hack to get Pazuzu Natural Storm 
				@Override
				public String notFound() {
					return null;
				}
			});

			// fireattuned
			doit2(sheet, "F501", 263);									

			// airattuned
			doit2(sheet, "F601", 264);

			// waterattuned
			doit2(sheet, "F701", 265);

			// earthattuned
			doit2(sheet, "F801", 266);

			// astralattuned
			doit2(sheet, "F901", 267);

			// deathattuned
			doit2(sheet, "FA01", 268);

			// natureattuned
			doit2(sheet, "FB01", 269);

			// bloodattuned
			doit2(sheet, "FC01", 270);

			// magicboost F
			doit2(sheet, "0A00", 271);

			// magicboost A
			doit2(sheet, "0B00", 272);

			// magicboost W
			doit2(sheet, "0C00", 273);

			// magicboost E
			doit2(sheet, "0D00", 274);

			// magicboost S
			doit2(sheet, "0E00", 275);

			// magicboost D
			doit2(sheet, "0F00", 276);

			// magicboost N
			doit2(sheet, "1000", 277);

			// magicboost ALL
			doit2(sheet, "1600", 278);

			// heatrec
			doit2(sheet, "EA01", 280, new CallbackAdapter() {
				@Override
				public String found(String value) {
					if (value.equals("10")) {
						return "0";
					}
					return "1";
				}
			});

			// coldrec
			doit2(sheet, "EB01", 281, new CallbackAdapter() {
				@Override
				public String found(String value) {
					if (value.equals("10")) {
						return "0";
					}
					return "1";
				}
			});

			// spread chaos
			doit2(sheet, "D801", 282, new CallbackAdapter() {
				@Override
				public String found(String value) {
					if (value.equals("100")) {
						return "1";
					}
					return "";
				}
			});

			// spread death
			doit2(sheet, "D801", 283, new CallbackAdapter() {
				@Override
				public String found(String value) {
					if (value.equals("103")) {
						return "1";
					}
					return "";
				}
			});
			
			// spread order
			doit2(sheet, "D901", 288, new CallbackAdapter() {
				@Override
				public String found(String value) {
					if (value.equals("100")) {
						return "1";
					}
					return "";
				}
			});

			// spread growth
			doit2(sheet, "D901", 289, new CallbackAdapter() {
				@Override
				public String found(String value) {
					if (value.equals("103")) {
						return "1";
					}
					return "";
				}
			});
			
			// corpseeater
			doit2(sheet, "EC00", 284);

			// poisonskin
			doit2(sheet, "0602", 285);

			// bug
			doit2(sheet, "AB01", 286);

			// uwbug
			doit2(sheet, "AC01", 287);

			// spreaddom
			doit2(sheet, "9C00", 291);

			// reform
			doit2(sheet, "8A01", 292);

			// battlesum5
			doit2(sheet, "8E01", 293);
			
			// acidsplash
			doit2(sheet, "F101", 294);
			
			// drake
			doit2(sheet, "F201", 295);

			// prophetshape
			doit2(sheet, "0402", 296);

			// horror
			doit2(sheet, "8701", 297);

			// insane
			doit2(sheet, "3501", 180);
			
			// sacr
			doit2(sheet, "D800", 199);
			
			// enchrebate50
			doit2(sheet, "1002", 298);
			
			// leper
			doit2(sheet, "7C00", 182);
			
			// resources
			doit2(sheet, "9701", 174);

			// slimer
			doit2(sheet, "9C01", 147);

			// mindslime
			doit2(sheet, "9201", 148);

			// corrupt
			doit2(sheet, "E000", 107, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return "1";
				}
			});
			
			// petrify
			doit2(sheet, "6800", 154, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return "1";
				}
			});
			
			// eyeloss
			doit2(sheet, "F400", 155);
			
			// ethtrue
			doit2(sheet, "0401", 157, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return "1";
				}
			});			

			// gemprod fire
			doit2(sheet, "1E00", 185, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return value + "F";
				}
			});
			
			// gemprod air
			doit2(sheet, "1F00", 185, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return value + "A";
				}
				@Override
				public String notFound() {
					return null;
				}
			}, true);
			
			// gemprod water
			doit2(sheet, "2000", 185, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return value + "W";
				}
				@Override
				public String notFound() {
					return null;
				}
			}, true);
			
			// gemprod earth
			doit2(sheet, "2100", 185, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return value + "E";
				}
				@Override
				public String notFound() {
					return null;
				}
			}, true);
			
			// gemprod astral
			doit2(sheet, "2200", 185, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return value + "S";
				}
				@Override
				public String notFound() {
					return null;
				}
			}, true);
			
			// gemprod death
			doit2(sheet, "2300", 185, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return value + "D";
				}
				@Override
				public String notFound() {
					return null;
				}
			}, true);
			
			// gemprod nature
			doit2(sheet, "2400", 185, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return value + "N";
				}
				@Override
				public String notFound() {
					return null;
				}
			}, true);
			
			// gemprod blood
			doit2(sheet, "2500", 185, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return value + "B";
				}
				@Override
				public String notFound() {
					return null;
				}
			}, true);
			
			// itemslots
			doit2(sheet, "B600", 40, new CallbackAdapter(){
				// hand
				@Override
				public String found(String value) {
					int numHands = 0;
					int val = Integer.parseInt(value);
					if ((val & 0x0002) != 0) {
						numHands++;
					}
					if ((val & 0x0004) != 0) {
						numHands++;
					}
					if ((val & 0x0008) != 0) {
						numHands++;
					}
					if ((val & 0x0010) != 0) {
						numHands++;
					}
					return Integer.toString(numHands);
				}
				@Override
				public String notFound() {
					return null;
				}
			});
			doit2(sheet, "B600", 41, new CallbackAdapter(){
				// head
				@Override
				public String found(String value) {
					int numHeads = 0;
					int val = Integer.parseInt(value);
					if ((val & 0x0080) != 0) {
						numHeads++;
					}
					if ((val & 0x0100) != 0) {
						numHeads++;
					}
					if ((val & 0x0200) != 0) {
						numHeads++;
					}
					return Integer.toString(numHeads);
				}
				@Override
				public String notFound() {
					return null;
				}
			});
			doit2(sheet, "B600", 42, new CallbackAdapter(){
				// body
				@Override
				public String found(String value) {
					int numBody = 0;
					int val = Integer.parseInt(value);
					if ((val & 0x0400) != 0) {
						numBody++;
					}
					return Integer.toString(numBody);
				}
				@Override
				public String notFound() {
					return null;
				}
			});
			doit2(sheet, "B600", 43, new CallbackAdapter(){
				// foot
				@Override
				public String found(String value) {
					int numFoot = 0;
					int val = Integer.parseInt(value);
					if ((val & 0x0800) != 0) {
						numFoot++;
					}
					return Integer.toString(numFoot);
				}
				@Override
				public String notFound() {
					return null;
				}
			});
			doit2(sheet, "B600", 44, new CallbackAdapter(){
				// misc
				@Override
				public String found(String value) {
					int numMisc = 0;
					int val = Integer.parseInt(value);
					if ((val & 0x1000) != 0) {
						numMisc++;
					}
					if ((val & 0x2000) != 0) {
						numMisc++;
					}
					if ((val & 0x4000) != 0) {
						numMisc++;
					}
					if ((val & 0x8000) != 0) {
						numMisc++;
					}
					if ((val & 0x10000) != 0) {
						numMisc++;
					}
					return Integer.toString(numMisc);
				}
				@Override
				public String notFound() {
					return null;
				}
			});
			
			// startage
			doit2(sheet, "1D01", 38, new CallbackAdapter(){
				@Override
				public String found(String value) {
					int age = Integer.parseInt(value);
					if (age == -1) {
						age = 0;
					}
					return Integer.toString((int)(age+age*.1));
				}
				@Override
				public String notFound() {
					return null;
				}
			});
			
			// maxage
			doit2(sheet, "1C01", 39, new CallbackAdapter(){
				@Override
				public String notFound() {
					return null;
				}
			});
			
			// heroarrivallimit
			doit2(sheet, "1602", 299);
			
			// sailsize
			doit2(sheet, "1702", 300);
			
			// uwdamage
			doit2(sheet, "0201", 301);
			
			// landdamage
			doit2(sheet, "0E02", 302);
			
			// rpcost
			doit1(240, 303, sheet);

			// fixedname
			/*File heroesFile = new File("heroes.txt");
			Set<Integer> heroes = new HashSet<Integer>();
			File namesFile = new File("names.txt");
			List<String> names = new ArrayList<String>();
			FileReader herosFileReader = new FileReader(heroesFile);
			FileReader namesFileReader = new FileReader(namesFile);
			BufferedReader bufferedReader = new BufferedReader(herosFileReader);
			String line;
			while ((line = bufferedReader.readLine()) != null) {
				heroes.add(Integer.parseInt(line));
			}
			bufferedReader.close();
			bufferedReader = new BufferedReader(namesFileReader);
			while ((line = bufferedReader.readLine()) != null) {
				names.add(line);
			}
			bufferedReader.close();
			int nameIndex = 0;
			for (int row = 1; row <= Starts.MONSTER_COUNT; row++) {
				String unique = getAttr("1301", row);
				if (row == 621 || row == 980 ||row == 981||row==994||row==995||row==996||row==997||
					row==1484 || row==1485|| row==1486|| row==1487 || (row >= 2765 && row <=2781)) {
					unique = "0";
				}
				if (heroes.contains(row) || (unique != null && unique.equals("1"))) {
					System.out.println(names.get(nameIndex++));
				} else {
					System.out.println("");
				}
			}*/

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.MONSTER_MAGIC);
			
			// magic
			c = new byte[4];
			while ((stream.read(c, 0, 4)) != -1) {
				String high = String.format("%02X", c[1]);
				String low = String.format("%02X", c[0]);
				if ((high + low).equals("FFFF")) {
					break;
				}
				int id = Integer.decode("0X" + high + low);
				Magic monMagic = monsterMagic.get(id);
				if (monMagic == null) {
					monMagic = new Magic();
					monsterMagic.put(id, monMagic);
				}
				
				//System.out.print("id: " + Integer.decode("0X" + high + low));
				stream.read(c, 0, 4);
				high = String.format("%02X", c[1]);
				low = String.format("%02X", c[0]);
				int path = Integer.decode("0X" + high + low);
				//System.out.print(" path: " + Integer.decode("0X" + high + low));
				stream.read(c, 0, 4);
				high = String.format("%02X", c[1]);
				low = String.format("%02X", c[0]);
				int value = Integer.decode("0X" + high + low);
				//System.out.println(" value: " + Integer.decode("0X" + high + low));
				switch (path) {
				case 0:
					monMagic.F = value;
					break;
				case 1:
					monMagic.A = value;
					break;
				case 2:
					monMagic.W = value;
					break;
				case 3:
					monMagic.E = value;
					break;
				case 4:
					monMagic.S = value;
					break;
				case 5:
					monMagic.D = value;
					break;
				case 6:
					monMagic.N = value;
					break;
				case 7:
					monMagic.B = value;
					break;
				case 8:
					monMagic.H = value;
					break;
				default:
					RandomMagic monRandomMagic = null;
					List<RandomMagic> randomMagicList = monMagic.rand;
					if (randomMagicList == null) {
						randomMagicList = new ArrayList<RandomMagic>();
						monRandomMagic = new RandomMagic();
						if (path == 50) {
							monRandomMagic.mask = 32640;
							monRandomMagic.nbr = 1;
							monRandomMagic.link = value;
							monRandomMagic.rand = 100;
						} else if (path == 51) {
							monRandomMagic.mask = 1920;
							monRandomMagic.nbr = 1;
							monRandomMagic.link = value;
							monRandomMagic.rand = 100;
						} else {
							monRandomMagic.mask = path;
							monRandomMagic.nbr = 1;
							if (value > 100) {
								monRandomMagic.link = value / 100;
								monRandomMagic.rand = 100;
							} else {
								monRandomMagic.link = 1;
								monRandomMagic.rand = value;
							}
						}
						randomMagicList.add(monRandomMagic);
						monMagic.rand = randomMagicList;
					} else {
						boolean found = false;
						for (RandomMagic ranMagic : randomMagicList) {
							if (ranMagic.mask == path && ranMagic.rand == value) {
								ranMagic.nbr++;
								found = true;
							}
							if (ranMagic.mask == 32640 && path == 50 && ranMagic.link == value) {
								ranMagic.nbr++;
								found = true;
							}
							if (ranMagic.mask == 1920 && path == 51 && ranMagic.link == value) {
								ranMagic.nbr++;
								found = true;
							}
						}
						if (!found) {
							monRandomMagic = new RandomMagic();
							if (path == 50) {
								monRandomMagic.mask = 32640;
								monRandomMagic.nbr = 1;
								monRandomMagic.link = value;
								monRandomMagic.rand = 100;
							} else if (path == 51) {
								monRandomMagic.mask = 1920;
								monRandomMagic.nbr = 1;
								monRandomMagic.link = value;
								monRandomMagic.rand = 100;
							} else {
								monRandomMagic.mask = path;
								monRandomMagic.nbr = 1;
								if (value > 100) {
									monRandomMagic.link = value / 100;
									monRandomMagic.rand = 100;
								} else {
									monRandomMagic.link = 1;
									monRandomMagic.rand = value;
								}
							}
							randomMagicList.add(monRandomMagic);
						}
					}
				}
			}
			for (int j = 1; j < Starts.MONSTER_COUNT; j++) {
				Magic monMagic = monsterMagic.get(j);
				if (monMagic != null) {
					XSSFRow row = sheet.getRow(j);
					XSSFCell cell = row.getCell(48, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					cell.setCellValue(magicStrip(monMagic.F));
					cell = row.getCell(49, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					cell.setCellValue(magicStrip(monMagic.A));
					cell = row.getCell(50, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					cell.setCellValue(magicStrip(monMagic.W));
					cell = row.getCell(51, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					cell.setCellValue(magicStrip(monMagic.E));
					cell = row.getCell(52, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					cell.setCellValue(magicStrip(monMagic.S));
					cell = row.getCell(53, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					cell.setCellValue(magicStrip(monMagic.D));
					cell = row.getCell(54, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					cell.setCellValue(magicStrip(monMagic.N));
					cell = row.getCell(55, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					cell.setCellValue(magicStrip(monMagic.B));
					cell = row.getCell(56, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					cell.setCellValue(magicStrip(monMagic.H));
					//System.out.print(magout(monMagic.F) + magout(monMagic.A ) + magout(monMagic.W) + magout(monMagic.E) + magout(monMagic.S) + magout(monMagic.D) + magout(monMagic.N) + magout(monMagic.B) + magout(monMagic.H));
					
					if (monMagic.rand != null) {
						int count = 0;
						for (RandomMagic ranMag : monMagic.rand) {
							cell = row.getCell(57 + count*4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
							cell.setCellValue(magicStrip(ranMag.rand));
							cell = row.getCell(58 + count*4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
							cell.setCellValue(magicStrip(ranMag.nbr));
							cell = row.getCell(59 + count*4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
							cell.setCellValue(magicStrip(ranMag.link));
							cell = row.getCell(60 + count*4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
							cell.setCellValue(magicStrip(ranMag.mask));
							count++;
							//System.out.print(ranMag.rand + "\t" + ranMag.nbr + "\t" + ranMag.link + "\t" + ranMag.mask + "\t");
						}
					}
				}
				//System.out.println("");
			}
			stream.close();

			wb.write(fos);
			fos.close();
			
			for (String col :columnsUsed.values()) {
				System.out.println(col);
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

}
