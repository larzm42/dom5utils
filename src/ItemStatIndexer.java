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
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ItemStatIndexer {
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

	private static int MASK[] = {0x0001, 0x0002, 0x0004, 0x0008, 0x0010, 0x0020, 0x0040, 0x0080,
		0x0100, 0x0200, 0x0400, 0x0800, 0x1000, 0x2000, 0x4000, 0x8000};

	private static String value1[] = {"bless", "luck", "", "airshield", "barkskin", "", "", "", "bers", "", "", "", "", "", "", "" };
	private static String value2[] = {"stoneskin", "fly", "quick", "", "", "", "", "", "", "", "", "eth", "ironskin", "", "", "" };
	private static String value3[] = {"", "", "", "float", "", "", "", "", "", "", "", "", "", "", "", "" };
	private static String value4[] = {"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
	private static String value5[] = {"", "", "", "", "", "", "trample", "", "", "", "", "", "", "", "", "fireshield?" };
	private static String value6[] = {"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
	private static String value7[] = {"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
	private static String value8[] = {"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
	private static String value9[] = {"disease", "curse", "", "", "", "", "", "", "", "", "", "", "", "", "", "cursed" };
	private static String value10[] = {"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
	private static String value11[] = {"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
	private static String value12[] = {"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };

	private static String DOTHESE[][] = {
		{"bless", "54"}, 
		{"luck", "55"}, 
		{"airshield", "15"}, 
		//{"barkskin", "15"}, 
		{"bers", "107"}, 
		//{"stoneskin", "15"}, 
		{"fly", "43"}, 
		{"float", "44"}, 
		{"quick", "51"}, 
		{"eth", "52"}, 
		//{"ironskin", "15"}, 
		{"trample", "53"}, 
		//{"fireshield?", "52"}, 
		{"disease", "58"}, 
		{"curse", "57"}, 
		{"cursed", "59"}, 
		};

	private static Map<Integer, String> columnsUsed = new HashMap<Integer, String>();
	
	private static String SkipColumns[] = {
		"id",
		"name",
		"type",
		"constlevel",
		"mainpath",
		"mainlevel",
		"secondarypath",
		"secondarylevel",
		"weapon",
		"armor",
		"test",
		"special",
		"autocombatspell",
		"startbattlespell",
		"itemspell",
		"restricted1",
		"restricted2",
		"restricted3",
		"restrictions",
		"ritual",
		"#sumbat",
		"sumbat",
		"affliction",
		"#sumrit",
		"sumrit",
		"#sumauto"
	};

	private static XSSFWorkbook readFile(String filename) throws IOException {
		return new XSSFWorkbook(new FileInputStream(filename));
	}
	
	private static void doit(XSSFSheet sheet, String attr, int column) throws IOException {
		doit(sheet, attr, column, null);
	}
	
	private static void doit(XSSFSheet sheet, String attr, int column, Callback callback) throws IOException {
        columnsUsed.remove(column);

        FileInputStream stream = new FileInputStream("Dominions5.exe");			
		stream.skip(Starts.ITEM);
		int rowNumber = 1;
		int i = 0;
		int k = 0;
		int pos = -1;
		long numFound = 0;
		byte[] c = new byte[2];
		stream.skip(120);
		while ((stream.read(c, 0, 2)) != -1) {
			String high = String.format("%02X", c[1]);
			String low = String.format("%02X", c[0]);
			int weapon = Integer.decode("0X" + high + low);
			if (weapon == 0) {
				boolean found = false;
				int value = 0;
				stream.skip(26l - numFound*2l);
				// Values
				for (int x = 0; x < numFound; x++) {
					byte[] d = new byte[4];
					stream.read(d, 0, 4);
					String high1 = String.format("%02X", d[3]);
					String low1 = String.format("%02X", d[2]);
					high = String.format("%02X", d[1]);
					low = String.format("%02X", d[0]);
					if (x == pos) {
						value = new BigInteger(high1 + low1 + high + low, 16).intValue();
						//System.out.print(value);
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
						cell.setCellValue(value);
					} else {
						if (callback.found(Integer.toString(value)) != null) {
							cell.setCellValue(callback.found(Integer.toString(value)));
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
				stream.skip(230l - 26l - numFound*4l);
				numFound = 0;
				pos = -1;
				k = 0;
				i++;
			} else {
				//System.out.print(low + high + " ");
				if ((low + high).equals(attr)) {
					pos = k;
				}
				k++;
				numFound++;
			}				
			if (i >= Starts.ITEM_COUNT) {
				break;
			}
		}
		stream.close();

	}
	
	private static boolean hasAttr(String attr, long position) throws IOException {
        FileInputStream stream = new FileInputStream("Dominions5.exe");	
        try {
    		stream.skip(position);
    		int i = 0;
    		byte[] c = new byte[2];
    		while ((stream.read(c, 0, 2)) != -1) {
    			String high = String.format("%02X", c[1]);
    			String low = String.format("%02X", c[0]);
    			int weapon = Integer.decode("0X" + high + low);
    			if (weapon == 0) {
    				return false;
    			} else {
    				if ((low + high).equals(attr)) {
    					return true;
    				}
    			}				
    			if (i >= Starts.ITEM_COUNT) {
    				break;
    			}
    		}
    		return false;
        } finally {
    		stream.close();
        }
	}

	public static void main(String[] args) {
		run();
	}
	
	public static void run() {
		FileInputStream stream = null;
		try {
	        long startIndex = Starts.ITEM;
	        int ch;

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(startIndex);
			
			XSSFWorkbook wb = ItemStatIndexer.readFile("BaseI.xlsx");
			FileOutputStream fos = new FileOutputStream("NewBaseI.xlsx");
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
				startIndex = startIndex + 232l;
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

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.ITEM);
			rowNumber = 1;
			// Const level
			int i = 0;
			byte[] c = new byte[1];
			stream.skip(36);
			while ((stream.read(c, 0, 1)) != -1) {
				XSSFRow row = sheet.getRow(rowNumber);
				rowNumber++;
				XSSFCell cell = row.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				if (c[0] == -1) {
					//System.out.println("");
					cell.setCellValue("");
				} else {
					//System.out.println(c[0]*2);
					cell.setCellValue(c[0]*2);
				}
				stream.skip(231l);
				i++;
				if (i >= Starts.ITEM_COUNT) {
					break;
				}
			}
			stream.close();

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.ITEM);
			rowNumber = 1;
			// Mainpath
			String[] paths = {"F", "A", "W", "E", "S", "D", "N", "B"};
			i = 0;
			c = new byte[1];
			stream.skip(37);
			while ((stream.read(c, 0, 1)) != -1) {
				XSSFRow row = sheet.getRow(rowNumber);
				rowNumber++;
				XSSFCell cell = row.getCell(4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				if (c[0] == -1) {
					//System.out.println("");
					cell.setCellValue("");
				} else {
					//System.out.println(paths[c[0]]);
					cell.setCellValue(paths[c[0]]);
				}
				stream.skip(231l);
				i++;
				if (i >= Starts.ITEM_COUNT) {
					break;
				}
			}
			stream.close();
			
			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.ITEM);
			rowNumber = 1;
			// main level
			i = 0;
			c = new byte[1];
			stream.skip(39);
			while ((stream.read(c, 0, 1)) != -1) {
				XSSFRow row = sheet.getRow(rowNumber);
				rowNumber++;
				XSSFCell cell = row.getCell(5, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				if (c[0] == -1) {
					//System.out.println("");
					cell.setCellValue("");
				} else {
					//System.out.println(c[0]);
					cell.setCellValue(c[0]);
				}
				stream.skip(231l);
				i++;
				if (i >= Starts.ITEM_COUNT) {
					break;
				}
			}
			stream.close();

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.ITEM);
			rowNumber = 1;
			// Secondary path
			i = 0;
			c = new byte[1];
			stream.skip(38);
			while ((stream.read(c, 0, 1)) != -1) {
				XSSFRow row = sheet.getRow(rowNumber);
				rowNumber++;
				XSSFCell cell = row.getCell(6, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				if (c[0] == -1) {
					//System.out.println("");
					cell.setCellValue("");
				} else {
					//System.out.println(paths[c[0]]);
					cell.setCellValue(paths[c[0]]);
				}
				stream.skip(231l);
				i++;
				if (i >= Starts.ITEM_COUNT) {
					break;
				}
			}
			stream.close();

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.ITEM);
			rowNumber = 1;
			// Secondary level
			i = 0;
			c = new byte[1];
			stream.skip(40);
			while ((stream.read(c, 0, 1)) != -1) {
				XSSFRow row = sheet.getRow(rowNumber);
				rowNumber++;
				XSSFCell cell = row.getCell(7, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				if (c[0] == -1 || c[0] == 0) {
					//System.out.println("");
					cell.setCellValue("");
				} else {
					//System.out.println(c[0]);
					cell.setCellValue(c[0]);
				}
				stream.skip(231l);
				i++;
				if (i >= Starts.ITEM_COUNT) {
					break;
				}
			}
			stream.close();

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.ITEM);
			rowNumber = 1;
			// type
			i = 0;
			c = new byte[1];
			stream.skip(41);
			while ((stream.read(c, 0, 1)) != -1) {
				XSSFRow row = sheet.getRow(rowNumber);
				rowNumber++;
				XSSFCell cell = row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				if (c[0] == -1 || c[0] == 0) {
					//System.out.println("");
					cell.setCellValue("");
				} else {
					//System.out.println(c[0]);
					if (c[0] == 1) {
						cell.setCellValue("1-h wpn");
					} else if (c[0] == 2) {
						cell.setCellValue("2-h wpn");
					} else if (c[0] == 3) {
						cell.setCellValue("missile");
					} else if (c[0] == 4) {
						cell.setCellValue("shield");
					} else if (c[0] == 5) {
						cell.setCellValue("armor");
					} else if (c[0] == 6) {
						cell.setCellValue("helm");
					} else if (c[0] == 7) {
						cell.setCellValue("boots");
					} else if (c[0] == 8) {
						cell.setCellValue("misc");
					}
				}
				stream.skip(231l);
				i++;
				if (i >= Starts.ITEM_COUNT) {
					break;
				}
			}
			stream.close();

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.ITEM);
			rowNumber = 1;
			// Weapon
			i = 0;
			c = new byte[2];
			stream.skip(44);
			while ((stream.read(c, 0, 2)) != -1) {
				XSSFRow row = sheet.getRow(rowNumber);
				rowNumber++;
				XSSFCell cell = row.getCell(8, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				String high = String.format("%02X", c[1]);
				String low = String.format("%02X", c[0]);
				int weapon = Integer.decode("0X" + high + low);
				if (weapon == 0) {
					//System.out.println("");
					cell.setCellValue("");
				} else {
					//System.out.println(Integer.decode("0X" + high + low));
					cell.setCellValue(Integer.decode("0X" + high + low));
				}
				stream.skip(230l);
				i++;
				if (i >= Starts.ITEM_COUNT) {
					break;
				}
			}
			stream.close();

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.ITEM);
			rowNumber = 1;
			// Armor
			i = 0;
			c = new byte[2];
			stream.skip(46);
			while ((stream.read(c, 0, 2)) != -1) {
				XSSFRow row = sheet.getRow(rowNumber);
				rowNumber++;
				XSSFCell cell = row.getCell(9, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				String high = String.format("%02X", c[1]);
				String low = String.format("%02X", c[0]);
				int weapon = Integer.decode("0X" + high + low);
				if (weapon == 0) {
					//System.out.println("");
					cell.setCellValue("");
				} else {
					//System.out.println(Integer.decode("0X" + high + low));
					cell.setCellValue(Integer.decode("0X" + high + low));
				}
				stream.skip(230l);
				i++;
				if (i >= Starts.ITEM_COUNT) {
					break;
				}
			}
			stream.close();

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.ITEM);
			rowNumber = 1;
	        startIndex = Starts.ITEM;
			// Itemspell
			stream.skip(48);
			startIndex = startIndex + 48l;
			i = 0;
			isr = new InputStreamReader(stream, "ISO-8859-1");
	        in = new BufferedReader(isr);
			while ((ch = in.read()) > -1) {
				StringBuffer name = new StringBuffer();
				while (ch != 0) {
					name.append((char)ch);
					ch = in.read();
				}
				in.close();

				stream = new FileInputStream("Dominions5.exe");		
				startIndex = startIndex + 232l;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);

				//System.out.println(name);
				XSSFRow row = sheet.getRow(rowNumber);
				rowNumber++;
				XSSFCell cell = row.getCell(130, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(name.toString());
				i++;
				if (i >= Starts.ITEM_COUNT) {
					break;
				}
			}
			in.close();
			stream.close();

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.ITEM);
			rowNumber = 1;
	        startIndex = Starts.ITEM;
			// Startbattlespell and Autocombatspell
			stream.skip(84);
			startIndex = startIndex + 84l;
			i = 0;
			isr = new InputStreamReader(stream, "ISO-8859-1");
	        in = new BufferedReader(isr);
			while ((ch = in.read()) > -1) {
				StringBuffer name = new StringBuffer();
				while (ch != 0) {
					name.append((char)ch);
					ch = in.read();
				}
				in.close();

				int blankCol = 129;
				int column = 128;
				if (hasAttr("8500", startIndex+36l)) {
					column = 129;
					blankCol = 128;
				}
				stream = new FileInputStream("Dominions5.exe");		
				startIndex = startIndex + 232l;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);

				//System.out.println(name);
				XSSFRow row = sheet.getRow(rowNumber);
				rowNumber++;
				XSSFCell cell = row.getCell(column, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(name.toString());
				XSSFCell blankcell = row.getCell(blankCol, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				blankcell.setCellValue("");
				i++;
				if (i >= Starts.ITEM_COUNT) {
					break;
				}
			}
			in.close();
			stream.close();

			// Fireres
			doit(sheet, "C600", 11);

			// Coldres
			doit(sheet, "C900", 12);

			// Poisonres
			doit(sheet, "C800", 13);

			// Shockres
			doit(sheet, "C700", 10);

			// Leadership
			doit(sheet, "9D00", 61);

			// str
			doit(sheet, "9700", 28);

			// fixforge
			doit(sheet, "C501", 18);

			// magic leadership
			doit(sheet, "9E00", 63);

			// undead leadership
			doit(sheet, "9F00", 62);

			// inspirational leadership
			doit(sheet, "7001", 64);

			// morale
			doit(sheet, "3401", 27);

			// penetration
			doit(sheet, "A100", 36);

			// pillage
			doit(sheet, "8300", 113);

			// fear
			doit(sheet, "B700", 66);

			// mr
			doit(sheet, "A000", 26);

			// taint
			doit(sheet, "0601", 60);

			// reinvigoration
			doit(sheet, "7500", 35);

			// awe
			doit(sheet, "6900", 67);

			// F
			doit(sheet, "0A00", 73);

			// A
			doit(sheet, "0B00", 74);

			// W
			doit(sheet, "0C00", 75);

			// E
			doit(sheet, "0D00", 76);

			// S
			doit(sheet, "0E00", 77);

			// D
			doit(sheet, "0F00", 78);

			// N
			doit(sheet, "1000", 79);

			// B
			doit(sheet, "1100", 80);

			// H
			doit(sheet, "1200", 81);

			// elemental
			doit(sheet, "1400", 73, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1400", 74, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1400", 75, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1400", 76, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});

			// sorcery
			doit(sheet, "1500", 77, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1500", 78, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1500", 79, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1500", 80, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			
			// all paths
			doit(sheet, "1600", 73, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1600", 74, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1600", 75, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1600", 76, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1600", 77, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1600", 78, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1600", 79, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1600", 80, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});

			// fire ritual range
			doit(sheet, "2800", 82);

			// air ritual range
			doit(sheet, "2900", 83);

			// water ritual range
			doit(sheet, "2A00", 84);

			// earth ritual range
			doit(sheet, "2B00", 85);

			// astral ritual range
			doit(sheet, "2C00", 86);

			// death ritual range
			doit(sheet, "2D00", 87);

			// nature ritual range
			doit(sheet, "2E00", 88);

			// blood ritual range
			doit(sheet, "2F00", 89);

			// elemental range
			doit(sheet, "1700", 82, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1700", 83, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1700", 84, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1700", 85, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});

			// sorcery range
			doit(sheet, "1800", 86, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1800", 87, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1800", 88, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1800", 89, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});

			// all range
			doit(sheet, "1900", 82, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1900", 83, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1900", 84, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1900", 85, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1900", 86, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1900", 87, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1900", 88, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "1900", 89, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});

			// darkvision
			doit(sheet, "1901", 14);

			// limited regeneration
			doit(sheet, "CE01", 16);

			// regeneration
			doit(sheet, "BD00", 141);

			// waterbreathing
			doit(sheet, "6E00", 47);

			// airbreathing
			doit(sheet, "8100", 49);

			// stealthb
			doit(sheet, "8601", 38);

			// stealth
			doit(sheet, "6C00", 37);

			// att
			doit(sheet, "9600", 29);

			// def
			doit(sheet, "7901", 30);

			// woundfend
			doit(sheet, "9601", 17);

			// berserk
			doit(sheet, "BE00", 106);

			// aging
			doit(sheet, "2F01", 145);

			// crossbreeder
			doit(sheet, "AF01", 125);

			// ivylord
			doit(sheet, "6500", 126);
			
			// forest
			doit(sheet, "A601", 39);

			// waste
			doit(sheet, "A701", 41);

			// mount
			doit(sheet, "A501", 40);

			// swamp
			doit(sheet, "A801", 42);

			// researchbonus
			doit(sheet, "7900", 118);
			
			// gitfofwater
			doit(sheet, "6F00", 48);

			// corpselord
			doit(sheet, "9A00", 146);

			// lictorlord
			doit(sheet, "6600", 147);

			// sumauto
			doit(sheet, "8B00", 134);

			// bloodsac
			doit(sheet, "D800", 148);

			// mastersmith
			doit(sheet, "6B01", 149);

			// alch
			doit(sheet, "8400", 21);

			// eyeloss
			doit(sheet, "7E00", 150);

			// armysize
			doit(sheet, "A301", 151);

			// defender
			doit(sheet, "8900", 152);

			// Hack for Forbidden Light
			doit(sheet, "7700", 73, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
				@Override
				public String found(String value) {
					return "2";
				}
			});
			doit(sheet, "7700", 77, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
				@Override
				public String found(String value) {
					return "2";
				}
			});

			// cannotwear
			doit(sheet, "C701", 153, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			doit(sheet, "C801", 153, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});

			// sailingshipsize
			doit(sheet, "7000", 45);

			// sailingmaxunitsize
			doit(sheet, "9A01", 46);

			// flytr
			doit(sheet, "7100", 50);

			// protf
			doit(sheet, "7F01", 24);

			// heretic
			doit(sheet, "B800", 119);

			// autodishealer
			doit(sheet, "6301", 19);

			// patrolbonus
			doit(sheet, "AA00", 114);

			// prec
			doit(sheet, "B500", 31);

			// tmpfiregems
			doit(sheet, "5801", 90);
			// tmpairgems
			doit(sheet, "5901", 91);
			// tmpwatergems
			doit(sheet, "5A01", 92);
			// tmpearthgems
			doit(sheet, "5B01", 93);
			// tmpastralgems
			doit(sheet, "5C01", 94);
			// tmpdeathgems
			doit(sheet, "5D01", 95);
			// tmpnaturegems
			doit(sheet, "5E01", 96);
			// tmpbloodgems
			doit(sheet, "5F01", 97);

			// healer
			doit(sheet, "6201", 20);

			// supplybonus
			doit(sheet, "7A00", 117);

			// mapspeed
			doit(sheet, "A901", 33);

			// gf
			doit(sheet, "1E00", 98);

			// ga
			doit(sheet, "1F00", 99);

			// gw
			doit(sheet, "2000", 100);

			// ge
			doit(sheet, "2100", 101);

			// gs
			doit(sheet, "2200", 102);

			// gd
			doit(sheet, "2300", 103);

			// gn
			doit(sheet, "2400", 104);

			// gb
			doit(sheet, "2500", 105);

			// reanimation bonus priest
			doit(sheet, "6C01", 154);
			
			// reanimation bonus death
			doit(sheet, "6D01", 155);
			
			// dragon mastery
			doit(sheet, "0E01", 156);
			
			// patience
			doit(sheet, "CD01", 157);
			
			// restricted
			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.ITEM);
			i = 0;
			int k = 0;
			Set<Integer> posSet = new HashSet<Integer>();
			long numFound = 0;
			c = new byte[2];
			stream.skip(120);
			while ((stream.read(c, 0, 2)) != -1) {
				String high = String.format("%02X", c[1]);
				String low = String.format("%02X", c[0]);
				int weapon = Integer.decode("0X" + high + low);
				if (weapon == 0) {
					stream.skip(26l - numFound*2l);
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
							XSSFCell cell = row.getCell(142+numRealms, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);							
							cell.setCellValue(fire-100);
							numRealms++;
						}
						//stream.skip(2);
					}
					
//					System.out.println("");
					stream.skip(230l - 26l - numFound*4l);
					numFound = 0;
					posSet.clear();
					k = 0;
					i++;
				} else {
					//System.out.print(low + high + " ");
					if ((low + high).equals("1601")) {
						posSet.add(k);
					}
					k++;
					numFound++;
				}				
				if (i >= Starts.ITEM_COUNT) {
					break;
				}
			}
			stream.close();

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
						if (pair[0].equals("airshield")) {
							cell.setCellValue(80);
						} else {
							cell.setCellValue(1);
						}
					} else {
							cell.setCellValue("");
					}
				}
			}

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.ITEM);
			i = 0;
			c = new byte[24];
			stream.skip(208);
			while ((stream.read(c, 0, 24)) != -1) {
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
							System.out.print((found?",":"") + (value6[j].equals("")?("*****"+(j+1)+"*****"):value6[j]));
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
							System.out.print((found?",":"") + (value7[j].equals("")?("*****"+(j+1)+"*****"):value7[j]));
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
							System.out.print((found?",":"") + (value8[j].equals("")?("*****"+(j+1)+"*****"):value8[j]));
							found = true;
						}
					}
					System.out.print("}");
				}
				high = String.format("%02X", c[17]);
				low = String.format("%02X", c[16]);
				//System.out.print(high + low + " ");
				val = Integer.decode("0X" + high + low);
				if (val > 0) {
					System.out.print(" 9:{");
					found = false;
					for (int j=0; j < 16; j++) {
						if ((val & MASK[j]) != 0) {
							System.out.print((found?",":"") + (value9[j].equals("")?("*****"+(j+1)+"*****"):value9[j]));
							found = true;
						}
					}
					System.out.print("}");
				}
				high = String.format("%02X", c[19]);
				low = String.format("%02X", c[18]);
				//System.out.print(high + low + " ");
				val = Integer.decode("0X" + high + low);
				if (val > 0) {
					System.out.print(" 10:{");
					found = false;
					for (int j=0; j < 16; j++) {
						if ((val & MASK[j]) != 0) {
							System.out.print((found?",":"") + (value10[j].equals("")?("*****"+(j+1)+"*****"):value10[j]));
							found = true;
						}
					}
					System.out.print("}");
				}
				high = String.format("%02X", c[21]);
				low = String.format("%02X", c[20]);
				//System.out.print(high + low + " ");
				val = Integer.decode("0X" + high + low);
				if (val > 0) {
					System.out.print(" 11:{");
					found = false;
					for (int j=0; j < 16; j++) {
						if ((val & MASK[j]) != 0) {
							System.out.print((found?",":"") + (value11[j].equals("")?("*****"+(j+1)+"*****"):value11[j]));
							found = true;
						}
					}
					System.out.print("}");
				}
				high = String.format("%02X", c[23]);
				low = String.format("%02X", c[22]);
				//System.out.print(high + low + " ");
				val = Integer.decode("0X" + high + low);
				if (val > 0) {
					System.out.print(" 12:{");
					found = false;
					for (int j=0; j < 16; j++) {
						if ((val & MASK[j]) != 0) {
							System.out.print((found?",":"") + (value12[j].equals("")?("*****"+(j+1)+"*****"):value12[j]));
							found = true;
						}
					}
					System.out.print("}");
				}
				
				if (!found) {
					//System.out.println("");
				}
				System.out.println(" ");
				stream.skip(208l);
				i++;
				if (i >= Starts.ITEM_COUNT) {
					break;
				}
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
	
	private static boolean[] largeBitmap(String fieldName) throws IOException {
		boolean[] boolArray = largeBitmapCache.get(fieldName);
		if (boolArray != null) {
			return boolArray;
		}
		FileInputStream stream = new FileInputStream("Dominions5.exe");			
		stream.skip(Starts.ITEM);
		int i = 0;
		byte[] c = new byte[24];
		boolArray = new boolean[Starts.ITEM_COUNT];
		stream.skip(208);
		while ((stream.read(c, 0, 24)) != -1) {
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
			high = String.format("%02X", c[17]);
			low = String.format("%02X", c[16]);
			val = Integer.decode("0X" + high + low);
			if (val > 0) {
				for (int j=0; j < 16; j++) {
					if ((val & MASK[j]) != 0) {
						if (value9[j].equals(fieldName)) {
							found = true;
						}
					}
				}
			}
			high = String.format("%02X", c[19]);
			low = String.format("%02X", c[18]);
			val = Integer.decode("0X" + high + low);
			if (val > 0) {
				for (int j=0; j < 16; j++) {
					if ((val & MASK[j]) != 0) {
						if (value10[j].equals(fieldName)) {
							found = true;
						}
					}
				}
			}
			high = String.format("%02X", c[21]);
			low = String.format("%02X", c[20]);
			val = Integer.decode("0X" + high + low);
			if (val > 0) {
				for (int j=0; j < 16; j++) {
					if ((val & MASK[j]) != 0) {
						if (value11[j].equals(fieldName)) {
							found = true;
						}
					}
				}
			}
			high = String.format("%02X", c[23]);
			low = String.format("%02X", c[22]);
			val = Integer.decode("0X" + high + low);
			if (val > 0) {
				for (int j=0; j < 16; j++) {
					if ((val & MASK[j]) != 0) {
						if (value12[j].equals(fieldName)) {
							found = true;
						}
					}
				}
			}
			boolArray[i] = found;
			stream.skip(208l);
			i++;
			if (i >= Starts.ITEM_COUNT) {
				break;
			}
		}
		stream.close();
		largeBitmapCache.put(fieldName, boolArray);
		return boolArray;
	}

}
