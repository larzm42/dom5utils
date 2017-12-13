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

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ItemStatIndexer extends AbstractStatIndexer {
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

	private static void putAttribute(XSSFSheet sheet, String attr, int column) throws IOException {
		putAttribute(sheet, attr, column, null);
	}
	
	private static void putAttribute(XSSFSheet sheet, String attr, int column, Callback callback) throws IOException {
		putAttribute(sheet, attr, Starts.ITEM_ATTRIBUTE_OFFSET, Starts.ITEM_ATTRIBUTE_GAP, column, Starts.ITEM, Starts.ITEM_SIZE, Starts.ITEM_COUNT, callback);
	}
	
	private static void putMultipleAttributes(XSSFSheet sheet, String attr, int column, Callback callback) throws IOException {
		putMultipleAttributes(sheet, attr, Starts.ITEM_ATTRIBUTE_OFFSET, Starts.ITEM_ATTRIBUTE_GAP, column, Starts.ITEM, Starts.ITEM_SIZE, Starts.ITEM_COUNT, callback);
	}
	
	private static boolean hasAttr(String attr, long position) throws IOException {
        FileInputStream stream = new FileInputStream(EXE_NAME);	
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

			stream = new FileInputStream(EXE_NAME);			
			stream.skip(startIndex);
			
			XSSFWorkbook wb = ItemStatIndexer.readFile("BaseI_Template.xlsx");
			FileOutputStream fos = new FileOutputStream("BaseI.xlsx");
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
				startIndex = startIndex + 232l;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);

				XSSFRow row = sheet.createRow(rowNumber);
				XSSFCell cell1 = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell1.setCellValue(rowNumber);
				rowNumber++;
				XSSFCell cell = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(name.toString());
			}
			in.close();
			stream.close();

			// Const level
			putBytes1(sheet, 36, 3, Starts.ITEM, Starts.ITEM_SIZE, Starts.ITEM_COUNT, new CallbackAdapter() {
				@Override
				public String found(String value) {
					if (value.equals("-1")) {
						return "";
					} else {
						return Integer.toString(Integer.parseInt(value)*2);
					}
				}
			});

			// Mainpath
			putBytes1(sheet, 37, 4, Starts.ITEM, Starts.ITEM_SIZE, Starts.ITEM_COUNT, new CallbackAdapter() {
				String[] paths = {"F", "A", "W", "E", "S", "D", "N", "B"};
				@Override
				public String found(String value) {
					if (value.equals("-1")) {
						return "";
					} else {
						return paths[Integer.parseInt(value)];
					}
				}
			});
			
			// main level
			putBytes1(sheet, 39, 5, Starts.ITEM, Starts.ITEM_SIZE, Starts.ITEM_COUNT);

			// Secondary path
			putBytes1(sheet, 38, 6, Starts.ITEM, Starts.ITEM_SIZE, Starts.ITEM_COUNT, new CallbackAdapter() {
				String[] paths = {"F", "A", "W", "E", "S", "D", "N", "B"};
				@Override
				public String found(String value) {
					if (value.equals("-1")) {
						return "";
					} else {
						return paths[Integer.parseInt(value)];
					}
				}
			});

			// Secondary level
			putBytes1(sheet, 40, 7, Starts.ITEM, Starts.ITEM_SIZE, Starts.ITEM_COUNT, new CallbackAdapter() {
				@Override
				public String found(String value) {
					if (value.equals("0")) {
						return "";
					} else {
						return value;
					}
				}
			});

			// type
			putBytes1(sheet, 41, 2, Starts.ITEM, Starts.ITEM_SIZE, Starts.ITEM_COUNT, new CallbackAdapter() {
				String[] types = {"1-h wpn", "2-h wpn", "missile", "shield", "armor", "helm", "boots", "misc"};
				@Override
				public String found(String value) {
					if (value.equals("-1") || value.equals("0")) {
						return "";
					} else {
						return types[Integer.parseInt(value)-1];
					}
				}
			});

			// Weapon
			putBytes2(sheet, 44, 8, Starts.ITEM, Starts.ITEM_SIZE, Starts.ITEM_COUNT, new CallbackAdapter() {
				@Override
				public String found(String value) {
					if (value.equals("0")) {
						return "";
					} else {
						return value;
					}
				}
			});

			// Armor
			putBytes2(sheet, 46, 9, Starts.ITEM, Starts.ITEM_SIZE, Starts.ITEM_COUNT, new CallbackAdapter() {
				@Override
				public String found(String value) {
					if (value.equals("0")) {
						return "";
					} else {
						return value;
					}
				}
			});

			// Itemspell
			stream = new FileInputStream(EXE_NAME);			
			stream.skip(Starts.ITEM);
			rowNumber = 1;
	        startIndex = Starts.ITEM;
			stream.skip(48);
			startIndex = startIndex + 48l;
			int i = 0;
			isr = new InputStreamReader(stream, "ISO-8859-1");
	        in = new BufferedReader(isr);
			while ((ch = in.read()) > -1) {
				StringBuffer name = new StringBuffer();
				while (ch != 0) {
					name.append((char)ch);
					ch = in.read();
				}
				in.close();

				stream = new FileInputStream(EXE_NAME);		
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

			// Startbattlespell and Autocombatspell
			stream = new FileInputStream(EXE_NAME);			
			stream.skip(Starts.ITEM);
			rowNumber = 1;
	        startIndex = Starts.ITEM;
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
				stream = new FileInputStream(EXE_NAME);		
				startIndex = startIndex + 232l;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);

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
			putAttribute(sheet, "C600", 11);

			// Coldres
			putAttribute(sheet, "C900", 12);

			// Poisonres
			putAttribute(sheet, "C800", 13);

			// Shockres
			putAttribute(sheet, "C700", 10);

			// Leadership
			putAttribute(sheet, "9D00", 61);

			// str
			putAttribute(sheet, "9700", 28);

			// fixforge
			putAttribute(sheet, "C501", 18);

			// magic leadership
			putAttribute(sheet, "9E00", 63);

			// undead leadership
			putAttribute(sheet, "9F00", 62);

			// inspirational leadership
			putAttribute(sheet, "7001", 64);

			// morale
			putAttribute(sheet, "3401", 27);

			// penetration
			putAttribute(sheet, "A100", 36);

			// pillage
			putAttribute(sheet, "8300", 113);

			// fear
			putAttribute(sheet, "B700", 66);

			// mr
			putAttribute(sheet, "A000", 26);

			// taint
			putAttribute(sheet, "0601", 60);

			// reinvigoration
			putAttribute(sheet, "7500", 35);

			// awe
			putAttribute(sheet, "6900", 67);

			// F
			putAttribute(sheet, "0A00", 73);

			// A
			putAttribute(sheet, "0B00", 74);

			// W
			putAttribute(sheet, "0C00", 75);

			// E
			putAttribute(sheet, "0D00", 76);

			// S
			putAttribute(sheet, "0E00", 77);

			// D
			putAttribute(sheet, "0F00", 78);

			// N
			putAttribute(sheet, "1000", 79);

			// B
			putAttribute(sheet, "1100", 80);

			// H
			putAttribute(sheet, "1200", 81);

			// elemental
			putAttribute(sheet, "1400", 73, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1400", 74, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1400", 75, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1400", 76, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});

			// sorcery
			putAttribute(sheet, "1500", 77, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1500", 78, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1500", 79, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1500", 80, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			
			// all paths
			putAttribute(sheet, "1600", 73, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1600", 74, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1600", 75, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1600", 76, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1600", 77, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1600", 78, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1600", 79, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1600", 80, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});

			// fire ritual range
			putAttribute(sheet, "2800", 82);

			// air ritual range
			putAttribute(sheet, "2900", 83);

			// water ritual range
			putAttribute(sheet, "2A00", 84);

			// earth ritual range
			putAttribute(sheet, "2B00", 85);

			// astral ritual range
			putAttribute(sheet, "2C00", 86);

			// death ritual range
			putAttribute(sheet, "2D00", 87);

			// nature ritual range
			putAttribute(sheet, "2E00", 88);

			// blood ritual range
			putAttribute(sheet, "2F00", 89);

			// elemental range
			putAttribute(sheet, "1700", 82, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1700", 83, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1700", 84, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1700", 85, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});

			// sorcery range
			putAttribute(sheet, "1800", 86, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1800", 87, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1800", 88, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1800", 89, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});

			// all range
			putAttribute(sheet, "1900", 82, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1900", 83, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1900", 84, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1900", 85, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1900", 86, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1900", 87, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1900", 88, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "1900", 89, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});

			// darkvision
			putAttribute(sheet, "1901", 14);

			// limited regeneration
			putAttribute(sheet, "CE01", 16);

			// regeneration
			putAttribute(sheet, "BD00", 141);

			// waterbreathing
			putAttribute(sheet, "6E00", 47);

			// airbreathing
			putAttribute(sheet, "8100", 49);

			// stealthb
			putAttribute(sheet, "8601", 38);

			// stealth
			putAttribute(sheet, "6C00", 37);

			// att
			putAttribute(sheet, "9600", 29);

			// def
			putAttribute(sheet, "7901", 30);

			// woundfend
			putAttribute(sheet, "9601", 17);

			// berserk
			putAttribute(sheet, "BE00", 106);

			// aging
			putAttribute(sheet, "2F01", 148);

			// crossbreeder
			putAttribute(sheet, "AF01", 125);

			// ivylord
			putAttribute(sheet, "6500", 126);
			
			// forest
			putAttribute(sheet, "A601", 39);

			// waste
			putAttribute(sheet, "A701", 41);

			// mount
			putAttribute(sheet, "A501", 40);

			// swamp
			putAttribute(sheet, "A801", 42);

			// researchbonus
			putAttribute(sheet, "7900", 118);
			
			// gitfofwater
			putAttribute(sheet, "6F00", 48);

			// corpselord
			putAttribute(sheet, "9A00", 149);

			// lictorlord
			putAttribute(sheet, "6600", 150);

			// sumauto
			putAttribute(sheet, "8B00", 134);

			// bloodsac
			putAttribute(sheet, "D800", 151);

			// mastersmith
			putAttribute(sheet, "6B01", 152);

			// alch
			putAttribute(sheet, "8400", 21);

			// eyeloss
			putAttribute(sheet, "7E00", 153);

			// armysize
			putAttribute(sheet, "A301", 154);

			// defender
			putAttribute(sheet, "8900", 155);

			// Hack for Forbidden Light
			putAttribute(sheet, "7700", 73, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
				@Override
				public String found(String value) {
					return "2";
				}
			});
			putAttribute(sheet, "7700", 77, new CallbackAdapter() {
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
			putAttribute(sheet, "C701", 156, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});
			putAttribute(sheet, "C801", 156, new CallbackAdapter() {
				@Override
				public String notFound() {
					return null;
				}
			});

			// sailingshipsize
			putAttribute(sheet, "7000", 45);

			// sailingmaxunitsize
			putAttribute(sheet, "9A01", 46);

			// flytr
			putAttribute(sheet, "7100", 50);

			// protf
			putAttribute(sheet, "7F01", 24);

			// heretic
			putAttribute(sheet, "B800", 119);

			// autodishealer
			putAttribute(sheet, "6301", 19);

			// patrolbonus
			putAttribute(sheet, "AA00", 114);

			// prec
			putAttribute(sheet, "B500", 31);

			// tmpfiregems
			putAttribute(sheet, "5801", 90);
			// tmpairgems
			putAttribute(sheet, "5901", 91);
			// tmpwatergems
			putAttribute(sheet, "5A01", 92);
			// tmpearthgems
			putAttribute(sheet, "5B01", 93);
			// tmpastralgems
			putAttribute(sheet, "5C01", 94);
			// tmpdeathgems
			putAttribute(sheet, "5D01", 95);
			// tmpnaturegems
			putAttribute(sheet, "5E01", 96);
			// tmpbloodgems
			putAttribute(sheet, "5F01", 97);

			// healer
			putAttribute(sheet, "6201", 20);

			// supplybonus
			putAttribute(sheet, "7A00", 117);

			// mapspeed
			putAttribute(sheet, "A901", 33);

			// gf
			putAttribute(sheet, "1E00", 98);

			// ga
			putAttribute(sheet, "1F00", 99);

			// gw
			putAttribute(sheet, "2000", 100);

			// ge
			putAttribute(sheet, "2100", 101);

			// gs
			putAttribute(sheet, "2200", 102);

			// gd
			putAttribute(sheet, "2300", 103);

			// gn
			putAttribute(sheet, "2400", 104);

			// gb
			putAttribute(sheet, "2500", 105);

			// reanimation bonus priest
			putAttribute(sheet, "6C01", 157);
			
			// reanimation bonus death
			putAttribute(sheet, "6D01", 158);
			
			// dragon mastery
			putAttribute(sheet, "0E01", 159);
			
			// patience
			putAttribute(sheet, "CD01", 160);
			
			// retinue
			putAttribute(sheet, "B401", 161);
			
			// restricted
			putMultipleAttributes(sheet, "1601", 142, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return Integer.toString(Integer.parseInt(value) - 100);
				}
			});

			// Large bitmap
			for (String[] pair : DOTHESE) {
				rowNumber = 1;
				boolean[] boolArray = largeBitmap(pair[0], 208l, Starts.ITEM, Starts.ITEM_SIZE, Starts.ITEM_COUNT, values);
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
