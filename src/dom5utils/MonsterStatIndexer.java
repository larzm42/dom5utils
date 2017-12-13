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
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.Reader;
import java.math.BigInteger;
import java.util.ArrayList;
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

public class MonsterStatIndexer extends AbstractStatIndexer {
	private static String values[][] = {{"heal", "mounted", "animal", "amphibian", "wastesurvival", "undead", "coldres15", "heat", "neednoteat", "fireres15", "poisonres15", "aquatic", "flying", "trample", "immobile", "immortal" },
										{"cold", "forestsurvival", "shockres15", "swampsurvival", "demon", "sacred", "mountainsurvival", "illusion", "noheal", "ethereal", "pooramphibian", "stealthy40", "misc2", "coldblood", "inanimate", "female" },
										{"bluntres", "slashres", "pierceres", "slow_to_recruit", "float", "", "teleport", "", "", "", "", "", "", "", "", "" },
										{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" },
										{"magicbeing", "", "", "poormagicleader", "okmagicleader", "goodmagicleader", "expertmagicleader", "superiormagicleader", "poorundeadleader", "okundeadleader", "goodundeadleader", "expertundeadleader", "superiorundeadleader", "", "", "" },
										{"", "", "", "", "", "", "", "", "noleader", "poorleader", "goodleader", "expertleader", "superiorleader", "", "", "" },
										{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" },
										{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" }};
	
	
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

	private static void putBytes2(int skip, int column, XSSFSheet sheet) throws IOException {
		putBytes2(skip, column, sheet, null);
	}
	
	private static void putBytes2(int skip, int column, XSSFSheet sheet, Callback callback) throws IOException {
		putBytes2(sheet, skip, column, Starts.MONSTER, Starts.MONSTER_SIZE, Starts.MONSTER_COUNT, callback);
	}
	
	private static String getAttr(String key, int id) throws IOException {
		String attributeString = attrCache.get(new CacheKey(key, id));
		if (attributeString != null) {
			return attributeString;
		}

		FileInputStream stream = new FileInputStream(EXE_NAME);			
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
	
	private static void putAttribute(XSSFSheet sheet, String attr, int column) throws IOException {
		putAttribute(sheet, attr, column, null);
	}
	
	private static void putAttribute(XSSFSheet sheet, String attr, int column, Callback callback) throws IOException {
		putAttribute(sheet, attr, column, callback, false);
	}
	
	private static void putAttribute(XSSFSheet sheet, String attr, int column, Callback callback, boolean append) throws IOException {
		putAttribute(sheet, attr, 64, 46l, column, Starts.MONSTER, Starts.MONSTER_SIZE, Starts.MONSTER_COUNT, append, callback);
	}
	
	public static void main(String[] args) {
		run();
	}
	
	public static void run() {
		FileInputStream stream = null;
		try {
	        long startIndex = Starts.MONSTER;
	        int ch;

			stream = new FileInputStream(EXE_NAME);			
			stream.skip(startIndex);
			
			XSSFWorkbook wb = MonsterStatIndexer.readFile("BaseU_Template.xlsx");
			FileOutputStream fos = new FileOutputStream("BaseU.xlsx");
			XSSFSheet sheet = wb.getSheetAt(0);

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

				stream = new FileInputStream(EXE_NAME);		
				startIndex = startIndex + Starts.MONSTER_SIZE;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);

				XSSFRow row = sheet.createRow(rowNumber);
				XSSFCell cell1 = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell1.setCellValue(rowNumber);
				XSSFCell cell = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(name.toString());
				rowNumber++;
			}
			in.close();
			stream.close();

	        // AP
			putBytes2(40, 31, sheet);

			// MM
			putBytes2(42, 30, sheet);

			// Size
			putBytes2(44, 19, sheet);
			
			// ressize
			putBytes2(44, 20, sheet);
			putAttribute(sheet, "0901", 20, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			
			// HP
			putBytes2(46, 21, sheet);

			// Prot
			putBytes2(48, 22, sheet);

			// STR
			putBytes2(50, 25, sheet);

			// ENC
			putBytes2(52, 29, sheet);

			// Prec
			putBytes2(54, 28, sheet);

			// ATT
			putBytes2(56, 26, sheet);

			// Def
			putBytes2(58, 27, sheet);

			// MR
			putBytes2(60, 23, sheet);

			// Mor
			putBytes2(62, 24, sheet);

			// wpn1
			putBytes2(208, 2, sheet, new CallbackAdapter(){
				@Override
				public String found(String value) {
					if (Integer.parseInt(value) == 0) {
						return "";
					}
					return value;
				}
			});

			// wpn2
			putBytes2(210, 3, sheet, new CallbackAdapter(){
				@Override
				public String found(String value) {
					if (Integer.parseInt(value) == 0) {
						return "";
					}
					return value;
				}
			});

			// wpn3
			putBytes2(212, 4, sheet, new CallbackAdapter(){
				@Override
				public String found(String value) {
					if (Integer.parseInt(value) == 0) {
						return "";
					}
					return value;
				}
			});

			// wpn4
			putBytes2(214, 5, sheet, new CallbackAdapter(){
				@Override
				public String found(String value) {
					if (Integer.parseInt(value) == 0) {
						return "";
					}
					return value;
				}
			});

			// wpn5
			putBytes2(216, 6, sheet, new CallbackAdapter(){
				@Override
				public String found(String value) {
					if (Integer.parseInt(value) == 0) {
						return "";
					}
					return value;
				}
			});

			// wpn6
			putBytes2(218, 7, sheet, new CallbackAdapter(){
				@Override
				public String found(String value) {
					if (Integer.parseInt(value) == 0) {
						return "";
					}
					return value;
				}
			});

			// wpn7
			putBytes2(220, 8, sheet, new CallbackAdapter(){
				@Override
				public String found(String value) {
					if (Integer.parseInt(value) == 0) {
						return "";
					}
					return value;
				}
			});

			// armor1
			putBytes2(228, 9, sheet, new CallbackAdapter(){
				@Override
				public String found(String value) {
					if (Integer.parseInt(value) == 0) {
						return "";
					}
					return value;
				}
			});

			// armor2
			putBytes2(230, 10, sheet, new CallbackAdapter(){
				@Override
				public String found(String value) {
					if (Integer.parseInt(value) == 0) {
						return "";
					}
					return value;
				}
			});

			// armor3
			putBytes2(232, 11, sheet, new CallbackAdapter(){
				@Override
				public String found(String value) {
					if (Integer.parseInt(value) == 0) {
						return "";
					}
					return value;
				}
			});

			// basecost
			putBytes2(234, 15, sheet);

			stream = new FileInputStream(EXE_NAME);			
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
			putAttribute(sheet, "9D00", 35, new Callback() {
				
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
			putAttribute(sheet, "B600", 40, new CallbackAdapter(){
				// hand
				@Override
				public String notFound() {
					return "2";
				}
			});
			putAttribute(sheet, "B600", 41, new CallbackAdapter(){
				// head
				@Override
				public String notFound() {
					return "1";
				}
			});
			putAttribute(sheet, "B600", 42, new CallbackAdapter(){
				// body
				@Override
				public String notFound() {
					return "1";
				}
			});
			putAttribute(sheet, "B600", 43, new CallbackAdapter(){
				// foot
				@Override
				public String notFound() {
					return "1";
				}
			});
			putAttribute(sheet, "B600", 44, new CallbackAdapter(){
				// misc
				@Override
				public String notFound() {
					return "2";
				}
			});
			
			// Large bitmap
			for (String[] pair : DOTHESE) {
				rowNumber = 1;
				boolean[] boolArray = largeBitmap(pair[0], 248, Starts.MONSTER, Starts.MONSTER_SIZE, Starts.MONSTER_COUNT, values);
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
							if (largeBitmap("illusion", 248, Starts.MONSTER, Starts.MONSTER_SIZE, Starts.MONSTER_COUNT, values)[rowNumber-2]) {
								glamour = true;
							}
							cell.setCellValue(40 + Integer.parseInt(additional.equals("")?"0":additional) + (glamour?25:0));
						} else if (pair[0].equals("coldres15")) {
							String additional = getAttr("C900", rowNumber-1);
							boolean cold = false;
							if (largeBitmap("cold", 248, Starts.MONSTER, Starts.MONSTER_SIZE, Starts.MONSTER_COUNT, values)[rowNumber-2] || getAttr("DC00", rowNumber-1).length() > 0) {
								cold = true;
							}
							cell.setCellValue(15 + Integer.parseInt(additional.equals("")?"0":additional) + (cold?10:0));
						} else if (pair[0].equals("fireres15")) {
							String additional = getAttr("C600", rowNumber-1);
							boolean heat = false;
							if (largeBitmap("heat", 248, Starts.MONSTER, Starts.MONSTER_SIZE, Starts.MONSTER_COUNT, values)[rowNumber-2] || getAttr("3C01", rowNumber-1).length() > 0) {
								heat = true;
							}
							cell.setCellValue(15 + Integer.parseInt(additional.equals("")?"0":additional) + (heat?10:0));
						} else if (pair[0].equals("poisonres15")) {
							String additional = getAttr("C800", rowNumber-1);
							boolean poisoncloud = false;
							if (largeBitmap("undead", 248, Starts.MONSTER, Starts.MONSTER_SIZE, Starts.MONSTER_COUNT, values)[rowNumber-2] || largeBitmap("inanimate", 248, Starts.MONSTER, Starts.MONSTER_SIZE, Starts.MONSTER_COUNT, values)[rowNumber-2] || getAttr("6A00", rowNumber-1).length() > 0) {
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
							if (largeBitmap("cold", 248, Starts.MONSTER, Starts.MONSTER_SIZE, Starts.MONSTER_COUNT, values)[rowNumber-2] || getAttr("DC00", rowNumber-1).length() > 0) {
								cold = true;
							}
							String additional = getAttr("C900", rowNumber-1);
							int coldres = Integer.parseInt(additional.equals("")?"0":additional) + (cold?10:0);
							cell.setCellValue(coldres==0?"":Integer.toString(coldres));
						} else if (pair[0].equals("fireres15")) {
							boolean heat = false;
							if (largeBitmap("heat", 248, Starts.MONSTER, Starts.MONSTER_SIZE, Starts.MONSTER_COUNT, values)[rowNumber-2] || getAttr("3C01", rowNumber-1).length() > 0) {
								heat = true;
							}
							String additional = getAttr("C600", rowNumber-1);
							int coldres = Integer.parseInt(additional.equals("")?"0":additional) + (heat?10:0);
							cell.setCellValue(coldres==0?"":Integer.toString(coldres));
						} else if (pair[0].equals("poisonres15")) {
							boolean poisoncloud = false;
							if (largeBitmap("undead", 248, Starts.MONSTER, Starts.MONSTER_SIZE, Starts.MONSTER_COUNT, values)[rowNumber-2] || largeBitmap("inanimate", 248, Starts.MONSTER, Starts.MONSTER_SIZE, Starts.MONSTER_COUNT, values)[rowNumber-2] || getAttr("6A00", rowNumber-1).length() > 0) {
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
			
			/*stream = new FileInputStream(EXE_NAME);			
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
			stream = new FileInputStream(EXE_NAME);			
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
			putAttribute(sheet, "CD01", 104);

			// stormimmune
			putAttribute(sheet, "AF00", 94);
			
			// regeneration
			putAttribute(sheet, "BD00", 151);

			// secondshape
			putAttribute(sheet, "C200", 210);

			// firstshape
			putAttribute(sheet, "C300", 209);

			// shapechange
			putAttribute(sheet, "C100", 208);

			// secondtmpshape
			putAttribute(sheet, "C400", 211);
			
			// landshape
			putAttribute(sheet, "F500", 212);
			
			// watershape
			putAttribute(sheet, "F600", 213);
			
			// forestshape
			putAttribute(sheet, "4201", 214);
			
			// plainshape
			putAttribute(sheet, "4301", 215);
			
			// damagerev
			putAttribute(sheet, "CA00", 144, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return Integer.toString(Integer.parseInt(value)-1);
				}
			});

			// bloodvengeance
			putAttribute(sheet, "9E01", 231);
			
			// nobadevents
			putAttribute(sheet, "EB00", 178);
			
			// bringeroffortune
			putAttribute(sheet, "E400", 232);
			
			// darkvision
			putAttribute(sheet, "1901", 133);
			
			// fear
			putAttribute(sheet, "B700", 138);
			
			// voidsanity
			putAttribute(sheet, "1501", 132);
			
			// standard
			putAttribute(sheet, "6700", 117);
			
			// formationfighter
			putAttribute(sheet, "6E01", 115);
			
			// undisciplined
			putAttribute(sheet, "6F01", 114);
			
			// bodyguard
			putAttribute(sheet, "9801", 121);
			
			// summon
			putAttribute(sheet, "A400,A500,A600", 223);
			putAttribute(sheet, "A400", 224, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return "1";
				}
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "A500", 224, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return "2";
				}
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "A600", 224, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return "3";
				}
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "A400,A500,A600", 224, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return null;
				}
			});

			// inspirational
			putAttribute(sheet, "7001", 118);
			
			// pillagebonus
			putAttribute(sheet, "8300", 187);
			
			// berserk
			putAttribute(sheet, "BE00", 139);
			
			// pathcost
			putAttribute(sheet, "F300", 45);
			
			// default startdom
			putAttribute(sheet, "F300", 46, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return "1";
				}
			});
			// startdom
			putAttribute(sheet, "F200", 46, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			
			// waterbreathing
			putAttribute(sheet, "6F00", 122);

			// batstartsum1
			putAttribute(sheet, "B401", 227);
			// batstartsum2
			putAttribute(sheet, "B501", 228);
			// batstartsum3
			putAttribute(sheet, "B601", 236);
			// batstartsum4
			putAttribute(sheet, "B701", 237);
			// batstartsum5
			putAttribute(sheet, "B801", 238);

			// batstartsum1d6
			putAttribute(sheet, "B901", 239);
			// batstartsum2d6
			putAttribute(sheet, "BA01", 240);
			// batstartsum3d6
			putAttribute(sheet, "BB01", 241);
			// batstartsum4d6
			putAttribute(sheet, "BC01", 242);
			// batstartsum5d6
			putAttribute(sheet, "BD01", 243);
			// batstartsum6d6
			putAttribute(sheet, "BE01", 244);

			// autosummon (#summon1-5)
			putAttribute(sheet, "6B00,8F00", 225);
			putAttribute(sheet, "6B00", 226, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return "?";
				}
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "8F00", 226, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return "?";
				}
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "6B00,8F00", 226, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return null;
				}
			});
			
			// domsummon
			putMultipleAttributes(sheet, "A101,DB00,F100", 64l, 46l, 229, Starts.MONSTER, Starts.MONSTER_SIZE, Starts.MONSTER_COUNT, null);
			
			// turmoil summon
			putAttribute(sheet, "AD00", 245);
			
			// cold summon
			putAttribute(sheet, "9200", 246);
			
			// stormpower
			putAttribute(sheet, "AE00", 161);
			
			// firepower
			putAttribute(sheet, "B100", 162);

			// coldpower
			putAttribute(sheet, "B000", 163);

			// darkpower
			putAttribute(sheet, "2501", 164);
			
			// chaospower
			putAttribute(sheet, "A001", 165);
			
			// magicpower
			putAttribute(sheet, "4401", 166);
			
			// winterpower
			putAttribute(sheet, "EA00", 167);
			
			// springpower
			putAttribute(sheet, "E700", 168);
			
			// summerpower
			putAttribute(sheet, "E800", 169);
			
			// fallpower
			putAttribute(sheet, "E900", 170);
			
			// nametype
			putAttribute(sheet, "FB00", 219);
			
			// blind
			putAttribute(sheet, "AB00", 134);
			
			// eyes
			putAttribute(sheet, "B200", 279, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return Integer.toString((Integer.parseInt(value) + 2));
				}
			});
			
			// supplybonus
			putAttribute(sheet, "7A00", 192);
			
			// slave
			putAttribute(sheet, "7C01", 116);
			
			// awe
			putAttribute(sheet, "6900", 136);
			
			// siegebonus
			putAttribute(sheet, "7D00", 190);
			
			// researchbonus
			putAttribute(sheet, "7900", 195);
			
			// chaosrec
			putAttribute(sheet, "CA01", 186);
			
			// invulnerability
			putAttribute(sheet, "7E01", 124);
			
			// iceprot
			putAttribute(sheet, "BF00", 123);
			
			// reinvigoration
			putAttribute(sheet, "7500", 34);
			
			// ambidextrous
			putAttribute(sheet, "D900", 32);
			
			// spy
			putAttribute(sheet, "D600", 102);
			
			// scale walls
			putAttribute(sheet, "E301", 247);
			
			// dream seducer (succubus)
			putAttribute(sheet, "D200", 106);
			
			// seduction
			putAttribute(sheet, "2A01", 105);
			
			// assassin
			putAttribute(sheet, "D500", 103);
			
			// explode on death
			putAttribute(sheet, "2901", 249);
			
			// taskmaster
			putAttribute(sheet, "7B01", 119);
			
			// unique
			putAttribute(sheet, "1301", 216);
			
			// poisoncloud
			putAttribute(sheet, "6A00", 145);
			
			// startaff
			putAttribute(sheet, "B900", 250);
			
			// uwregen
			putAttribute(sheet, "DD00", 251);
			
			// patrolbonus
			putAttribute(sheet, "AA00", 188);
			
			// castledef
			putAttribute(sheet, "D700", 189);
			
			// sailingshipsize
			putAttribute(sheet, "7000", 98);
			
			// sailingmaxunitsize
			putAttribute(sheet, "9A01", 99);
			
			// incunrest
			putAttribute(sheet, "DF00", 193, new CallbackAdapter() {
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
			putAttribute(sheet, "BC00", 153, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return "1";
				}
			});
			
			// inn
			putAttribute(sheet, "4E01", 47);
			
			// stonebeing
			putAttribute(sheet, "1801", 80);
			
			// shrinkhp
			putAttribute(sheet, "5001", 252);
			
			// growhp
			putAttribute(sheet, "4F01", 253);
			
			// transformation
			putAttribute(sheet, "FD01", 254);
			
			// heretic
			putAttribute(sheet, "B800", 206);
			
			// popkill
			putAttribute(sheet, "2001", 194, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return Integer.toString(Integer.parseInt(value)*10);
				}
			});
			
			// autohealer
			putAttribute(sheet, "6201", 175);
			
			// fireshield
			putAttribute(sheet, "A300", 142);
			
			// startingaff
			putAttribute(sheet, "E200", 255);
			
			// fixedresearch
			putAttribute(sheet, "F800", 256);
			
			// divineins
			putAttribute(sheet, "C201", 257);
			
			// halt
			putAttribute(sheet, "4701", 137);
			
			// crossbreeder
			putAttribute(sheet, "AF01", 201);

			// reclimit
			putAttribute(sheet, "7D01", 14);
			
			// fixforgebonus
			putAttribute(sheet, "C501", 172);

			// mastersmith
			putAttribute(sheet, "6B01", 173);

			// lamiabonus
			putAttribute(sheet, "A900", 258);

			// homesick
			putAttribute(sheet, "FD00", 113);

			// banefireshield
			putAttribute(sheet, "DE00", 143);

			// animalawe
			putAttribute(sheet, "A200", 135);

			// autodishealer
			putAttribute(sheet, "6301", 176);

			// shatteredsoul
			putAttribute(sheet, "4801", 181);

			// voidsum
			putAttribute(sheet, "CE00", 205);

			// makepearls
			putAttribute(sheet, "AE01", 202);

			// inspiringres
			putAttribute(sheet, "5501", 197);

			// drainimmune
			putAttribute(sheet, "1101", 196);
			
			// diseasecloud
			putAttribute(sheet, "AC00", 146);

			// inquisitor
			putAttribute(sheet, "D300", 74);

			// beastmaster
			putAttribute(sheet, "7101", 120);

			// douse
			putAttribute(sheet, "7400", 198);

			// preanimator
			putAttribute(sheet, "6C01", 259);

			// dreanimator
			putAttribute(sheet, "6D01", 260);

			// mummify
			putAttribute(sheet, "FF01", 261);
			
			// onebattlespell
			putAttribute(sheet, "D100", 262, new CallbackAdapter() {
				// Hack to get Pazuzu Natural Storm 
				@Override
				public String notFound() {
					return null;
				}
			});

			// fireattuned
			putAttribute(sheet, "F501", 263);									

			// airattuned
			putAttribute(sheet, "F601", 264);

			// waterattuned
			putAttribute(sheet, "F701", 265);

			// earthattuned
			putAttribute(sheet, "F801", 266);

			// astralattuned
			putAttribute(sheet, "F901", 267);

			// deathattuned
			putAttribute(sheet, "FA01", 268);

			// natureattuned
			putAttribute(sheet, "FB01", 269);

			// bloodattuned
			putAttribute(sheet, "FC01", 270);

			// magicboost F
			putAttribute(sheet, "0A00", 271);

			// magicboost A
			putAttribute(sheet, "0B00", 272);

			// magicboost W
			putAttribute(sheet, "0C00", 273);

			// magicboost E
			putAttribute(sheet, "0D00", 274);

			// magicboost S
			putAttribute(sheet, "0E00", 275);

			// magicboost D
			putAttribute(sheet, "0F00", 276);

			// magicboost N
			putAttribute(sheet, "1000", 277);

			// magicboost ALL
			putAttribute(sheet, "1600", 278);

			// heatrec
			putAttribute(sheet, "EA01", 280, new CallbackAdapter() {
				@Override
				public String found(String value) {
					if (value.equals("10")) {
						return "0";
					}
					return "1";
				}
			});

			// coldrec
			putAttribute(sheet, "EB01", 281, new CallbackAdapter() {
				@Override
				public String found(String value) {
					if (value.equals("10")) {
						return "0";
					}
					return "1";
				}
			});

			// spread chaos
			putAttribute(sheet, "D801", 282, new CallbackAdapter() {
				@Override
				public String found(String value) {
					if (value.equals("100")) {
						return "1";
					}
					return "";
				}
			});

			// spread death
			putAttribute(sheet, "D801", 283, new CallbackAdapter() {
				@Override
				public String found(String value) {
					if (value.equals("103")) {
						return "1";
					}
					return "";
				}
			});
			
			// spread order
			putAttribute(sheet, "D901", 288, new CallbackAdapter() {
				@Override
				public String found(String value) {
					if (value.equals("100")) {
						return "1";
					}
					return "";
				}
			});

			// spread growth
			putAttribute(sheet, "D901", 289, new CallbackAdapter() {
				@Override
				public String found(String value) {
					if (value.equals("103")) {
						return "1";
					}
					return "";
				}
			});
			
			// corpseeater
			putAttribute(sheet, "EC00", 284);

			// poisonskin
			putAttribute(sheet, "0602", 285);

			// bug
			putAttribute(sheet, "AB01", 286);

			// uwbug
			putAttribute(sheet, "AC01", 287);

			// spreaddom
			putAttribute(sheet, "9C00", 291);

			// reform
			putAttribute(sheet, "8A01", 292);

			// battlesum5
			putAttribute(sheet, "8E01", 293);
			
			// acidsplash
			putAttribute(sheet, "F101", 294);
			
			// drake
			putAttribute(sheet, "F201", 295);

			// prophetshape
			putAttribute(sheet, "0402", 296);

			// horror
			putAttribute(sheet, "8701", 297);

			// insane
			putAttribute(sheet, "3501", 180);
			
			// sacr
			putAttribute(sheet, "D800", 199);
			
			// enchrebate50
			putAttribute(sheet, "1002", 298);
			
			// leper
			putAttribute(sheet, "7C00", 182);
			
			// resources
			putAttribute(sheet, "9701", 174);

			// slimer
			putAttribute(sheet, "9C01", 147);

			// mindslime
			putAttribute(sheet, "9201", 148);

			// corrupt
			putAttribute(sheet, "E000", 107, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return "1";
				}
			});
			
			// petrify
			putAttribute(sheet, "6800", 154, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return "1";
				}
			});
			
			// eyeloss
			putAttribute(sheet, "F400", 155);
			
			// ethtrue
			putAttribute(sheet, "0401", 157, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return "1";
				}
			});			

			// gemprod fire
			putAttribute(sheet, "1E00", 185, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return value + "F";
				}
			});
			
			// gemprod air
			putAttribute(sheet, "1F00", 185, new CallbackAdapter() {
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
			putAttribute(sheet, "2000", 185, new CallbackAdapter() {
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
			putAttribute(sheet, "2100", 185, new CallbackAdapter() {
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
			putAttribute(sheet, "2200", 185, new CallbackAdapter() {
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
			putAttribute(sheet, "2300", 185, new CallbackAdapter() {
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
			putAttribute(sheet, "2400", 185, new CallbackAdapter() {
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
			putAttribute(sheet, "2500", 185, new CallbackAdapter() {
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
			putAttribute(sheet, "B600", 40, new CallbackAdapter(){
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
			putAttribute(sheet, "B600", 41, new CallbackAdapter(){
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
			putAttribute(sheet, "B600", 42, new CallbackAdapter(){
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
			putAttribute(sheet, "B600", 43, new CallbackAdapter(){
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
			putAttribute(sheet, "B600", 44, new CallbackAdapter(){
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
			putAttribute(sheet, "1D01", 38, new CallbackAdapter(){
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
			putAttribute(sheet, "1C01", 39, new CallbackAdapter(){
				@Override
				public String notFound() {
					return null;
				}
			});
			
			// heroarrivallimit
			putAttribute(sheet, "1602", 299);
			
			// sailsize
			putAttribute(sheet, "1702", 300);
			
			// uwdamage
			putAttribute(sheet, "0201", 301);
			
			// landdamage
			putAttribute(sheet, "0E02", 302);
			
			// rpcost
			putBytes2(240, 303, sheet);

			// fixedname
			File heroesFile = new File("heroes.txt");
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
			}

			stream = new FileInputStream(EXE_NAME);			
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
							if (count == 4) {
								count = 62;
							}
							//System.out.print(ranMag.rand + "\t" + ranMag.nbr + "\t" + ranMag.link + "\t" + ranMag.mask + "\t");
						}
					}
				}
				//System.out.println("");
			}
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

}
