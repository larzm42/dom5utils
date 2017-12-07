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
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SiteStatIndexer {
	
	private static XSSFWorkbook readFile(String filename) throws IOException {
		return new XSSFWorkbook(new FileInputStream(filename));
	}

	public static void doit(XSSFSheet sheet, String attr, int column) throws IOException {
		doit(sheet, attr, column, null);
	}
	
	public static void doit2(XSSFSheet sheet, String attr, int column, int column2) throws IOException {
		FileInputStream stream = new FileInputStream("Dominions5.exe");			
		stream.skip(Starts.SITE);
		int rowNumber = 1;
		int i = 0;
		int k = 0;
		int numFound = 0;
		Set<Integer> posSet = new HashSet<Integer>();
		byte[] c = new byte[2];
		stream.skip(44);
		while ((stream.read(c, 0, 2)) != -1) {
			String high = String.format("%02X", c[1]);
			String low = String.format("%02X", c[0]);
			int weapon = Integer.decode("0X" + high + low);
			if (weapon == 0) {
				stream.skip(34l - numFound*2l);
				// Values
				boolean found = false;
				List<Integer> values = new ArrayList<Integer>();
				for (int x = 0; x < numFound; x++) {
					stream.read(c, 0, 2);
					high = String.format("%02X", c[1]);
					low = String.format("%02X", c[0]);
					if (posSet.contains(x)) {
						int fire = new BigInteger(new byte[]{c[1], c[0]}).intValue();
						if (!found) {
							found = true;
						} else {
							//System.out.print("\t");
						}
						//System.out.print(fire);
						values.add(fire);
					}
					stream.skip(6);
				}
				
				//System.out.println("");
				XSSFRow row = sheet.getRow(rowNumber);
				rowNumber++;
				int ind = 0;
				for (Integer mon : values) {
					if (ind == 0) {
						XSSFCell cell = row.getCell(column, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						cell.setCellValue(mon);
					} else {
						XSSFCell cell = row.getCell(column2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						cell.setCellValue(mon);
					}
					ind++;
				}
				stream.skip(214l - 34l - numFound*8l);
				numFound = 0;
				posSet.clear();
				k = 0;
				i++;
			} else {
				//System.out.print(low + high + " ");
				if ((low + high).equals(attr)) {
					posSet.add(k);
				}
				k++;
				numFound++;
			}				
			if (i >= Starts.SITE_COUNT) {
				break;
			}
		}
		stream.close();

	}
	
	public static void doit(XSSFSheet sheet, String attr, int column, Callback callback) throws IOException {
		FileInputStream stream = new FileInputStream("Dominions5.exe");			
		stream.skip(Starts.SITE);
		int rowNumber = 1;
		
		int i = 0;
		int k = 0;
		int pos = -1;
		long numFound = 0;
		byte[] c = new byte[2];
		stream.skip(44);
		while ((stream.read(c, 0, 2)) != -1) {
			String high = String.format("%02X", c[1]);
			String low = String.format("%02X", c[0]);
			int weapon = Integer.decode("0X" + high + low);
			if (weapon == 0) {
				boolean found = false;
				int value = 0;
				stream.skip(34l - numFound*2l);
				// Values
				for (int x = 0; x < numFound; x++) {
					stream.read(c, 0, 2);
					high = String.format("%02X", c[1]);
					low = String.format("%02X", c[0]);
					//System.out.print(low + high + " ");
					if (x == pos) {
						int tmp = new BigInteger(high + low, 16).intValue();
						if (tmp < 1000) {
							value = Integer.decode("0X" + high + low);
						} else {
							value = new BigInteger("FFFF" + high + low, 16).intValue();
						}
						//System.out.print(fire);
						found = true;
					}
					stream.skip(6);
				}
				
				//System.out.println("");
				XSSFRow row = sheet.getRow(rowNumber);
				rowNumber++;
				XSSFCell cell = row.getCell(column, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				if (found) {
					if (callback == null) {
						cell.setCellValue(value);
					} else {
						cell.setCellValue(callback.found(Integer.toString(value)));
					}
				} else {
					if (callback == null) {
						cell.setCellValue("");
					} else {
						cell.setCellValue(callback.notFound());
					}
				}
				stream.skip(214l - 34l - numFound*8l);
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
			if (i >= Starts.SITE_COUNT) {
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
	        long startIndex = Starts.SITE;
	        int ch;

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.SITE);
			
			XSSFWorkbook wb = SiteStatIndexer.readFile("MagicSites.xlsx");
			FileOutputStream fos = new FileOutputStream("NewMagicSites.xlsx");
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

				stream = new FileInputStream("Dominions5.exe");		
				startIndex = startIndex + 216l;
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
			stream.skip(Starts.SITE);
			rowNumber = 1;
			// rarity
			int i = 0;
			byte[] c = new byte[1];
			stream.skip(42);
			while ((stream.read(c, 0, 1)) != -1) {
				XSSFRow row = sheet.getRow(rowNumber);
				rowNumber++;
				XSSFCell cell = row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				if (c[0] == -1 || c[0] == 0) {
					//System.out.println("0");
					cell.setCellValue("0");
				} else {
					//System.out.println(c[0]);
					cell.setCellValue(c[0]);
				}
				stream.skip(215l);
				i++;
				if (i >= Starts.SITE_COUNT) {
					break;
				}
			}
			stream.close();

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.SITE);
			rowNumber = 1;
			// loc
			i = 0;
			c = new byte[2];
			stream.skip(208);
			while ((stream.read(c, 0, 2)) != -1) {
				String high = String.format("%02X", c[1]);
				String low = String.format("%02X", c[0]);
				//System.out.println(Integer.decode("0X" + high + low));
				XSSFRow row = sheet.getRow(rowNumber);
				rowNumber++;
				XSSFCell cell = row.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(Integer.decode("0X" + high + low));
				stream.skip(214l);
				i++;
				if (i >= Starts.SITE_COUNT) {
					break;
				}
			}
			stream.close();

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.SITE);
			rowNumber = 1;
			// level
			i = 0;
			c = new byte[1];
			stream.skip(40);
			while ((stream.read(c, 0, 1)) != -1) {
				String high = String.format("%02X", c[0]);
				//System.out.println(Integer.decode("0X" + high));
				XSSFRow row = sheet.getRow(rowNumber);
				rowNumber++;
				XSSFCell cell = row.getCell(4, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(Integer.decode("0X" + high));
				stream.skip(214l);
				i++;
				if (i >= Starts.SITE_COUNT) {
					break;
				}
			}
			stream.close();

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.SITE);
			rowNumber = 1;
			// path
			String[] paths = {"Fire", "Air", "Water", "Earth", "Astral", "Death", "Nature", "Blood", "Holy"};
			int[] spriteOffset = {1, 9, 18, 26, 35, 42, 50, 59, 68};
			i = 0;
			c = new byte[1];
			byte[] d = new byte[1];
			stream.skip(36);
			stream.read(d, 0, 1);
			stream.skip(1);
			while ((stream.read(c, 0, 1)) != -1) {
				XSSFRow row = sheet.getRow(rowNumber);
				rowNumber++;
				XSSFCell cell = row.getCell(5, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				XSSFCell cell2 = row.getCell(105, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				if (c[0] == -1) {
					//System.out.println("");
					cell.setCellValue("");
				} else {
					//System.out.println(paths[c[0]]);
					cell.setCellValue(paths[c[0]]);
					cell2.setCellValue(spriteOffset[c[0]] + d[0]);
				}
				stream.skip(213l);
				stream.read(d, 0, 1);
				stream.skip(1);
				i++;
				if (i >= Starts.SITE_COUNT) {
					break;
				}
			}
			stream.close();

			// F
			doit2(sheet, "0100", 6, 92);

			// A
			doit2(sheet, "0200", 7, 93);

			// W
			doit2(sheet, "0300", 8, 94);

			// E
			doit2(sheet, "0400", 9, 95);

			// S
			doit2(sheet, "0500", 10, 96);

			// D
			doit2(sheet, "0600", 11, 97);

			// N
			doit2(sheet, "0700", 12, 98);

			// B
			doit2(sheet, "0800", 13, 99);

			// gold
			doit(sheet, "0D00", 14);

			// res
			doit(sheet, "0E00", 15);

			// sup
			doit(sheet, "1400", 16);

			// unr
			doit(sheet, "1300", 17, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return Integer.toString(-Integer.parseInt(value));
				}
			});

			// exp
			doit(sheet, "1600", 18);

			// lab
			doit(sheet, "0F00", 19, new CallbackAdapter() {
				@Override
				public String found(String value) {
					return "lab";
				}
			});

			// fort
			doit(sheet, "1100", 20);

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.SITE);
			rowNumber = 1;
			// scales
			String[] scales = {"Turmoil", "Sloth", "Cold", "Death", "Misfortune", "Drain"};
			String[] opposite = {"Order", "Productivity", "Heat", "Growth", "Luck", "Magic"};
			i = 0;
			int k = 0;
			Set<Integer> scalesSet = new HashSet<Integer>();
			Set<Integer> oppositeSet = new HashSet<Integer>();
			long numFound = 0;
			c = new byte[2];
			stream.skip(44);
			while ((stream.read(c, 0, 2)) != -1) {
				String high = String.format("%02X", c[1]);
				String low = String.format("%02X", c[0]);
				int weapon = Integer.decode("0X" + high + low);
				if (weapon == 0) {
					stream.skip(34l - numFound*2l);
					// Values
					boolean found = false;
					String value[] = {"", ""};
					int index = 0;
					for (int x = 0; x < numFound; x++) {
						stream.read(c, 0, 2);
						high = String.format("%02X", c[1]);
						low = String.format("%02X", c[0]);
						//System.out.print(low + high + " ");
						if (oppositeSet.contains(x)) {
							//int fire = Integer.decode("0X" + high + low);
							int fire = new BigInteger(new byte[]{c[1], c[0]}).intValue();
							if (!found) {
								found = true;
							} else {
								//System.out.print("\t");
							}
							//System.out.print(opposite[fire]);
							value[index++] = opposite[fire];
						}
						if (scalesSet.contains(x)) {
							//int fire = Integer.decode("0X" + high + low);
							int fire = new BigInteger(new byte[]{c[1], c[0]}).intValue();
							if (!found) {
								found = true;
							} else {
								//System.out.print("\t");
							}
							//System.out.print(scales[fire]);
							value[index++] = scales[fire];
						}
						stream.skip(6);
					}
					
					//System.out.println("");
					XSSFRow row = sheet.getRow(rowNumber);
					rowNumber++;
					if (value[0].length() > 0) {
						XSSFCell cell = row.getCell(21, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						cell.setCellValue(value[0]);
					}
					if (value[1].length() > 0) {
						XSSFCell cell = row.getCell(22, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						cell.setCellValue(value[1]);
					}
					stream.skip(214l - 34l - numFound*8l);
					numFound = 0;
					scalesSet.clear();
					oppositeSet.clear();
					k = 0;
					i++;
				} else {
					//System.out.print(low + high + " ");
					if ((low + high).equals("1F00")) {
						oppositeSet.add(k);
					}
					if ((low + high).equals("2000")) {
						scalesSet.add(k);
					}
					k++;
					numFound++;
				}				
				if (i >= Starts.SITE_COUNT) {
					break;
				}
			}
			stream.close();

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.SITE);
			rowNumber = 1;
			// rit/ritr
			i = 0;
			k = 0;
			String rit = "";
			numFound = 0;
			int pos = -1;
			c = new byte[2];
			stream.skip(44);
			while ((stream.read(c, 0, 2)) != -1) {
				String high = String.format("%02X", c[1]);
				String low = String.format("%02X", c[0]);
				int weapon = Integer.decode("0X" + high + low);
				if (weapon == 0) {
					stream.skip(34l - numFound*2l);
					boolean found = false;
					int value = 0;
					// Values
					for (int x = 0; x < numFound; x++) {
						stream.read(c, 0, 2);
						high = String.format("%02X", c[1]);
						low = String.format("%02X", c[0]);
						//System.out.print(low + high + " ");
						if (x == pos) {
							//int fire = Integer.decode("0X" + high + low);
							//int fire = new BigInteger(new byte[]{c[1], c[0]}).intValue();
							//System.out.print(rit + "\t" + fire);
							found = true;
							value = new BigInteger(new byte[]{c[1], c[0]}).intValue();
						}
						stream.skip(6);
					}
					
					//System.out.println("");
					XSSFRow row = sheet.getRow(rowNumber);
					rowNumber++;
					XSSFCell cell1 = row.getCell(41, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					XSSFCell cell2 = row.getCell(42, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					if (found) {
						cell1.setCellValue(rit);
						cell2.setCellValue(value);
					} else {
						cell1.setCellValue("");
						cell2.setCellValue("");
					}
					
					stream.skip(214l - 34l - numFound*8l);
					numFound = 0;
					pos = -1;
					rit = "";
					k = 0;
					i++;
				} else {
					//System.out.print(low + high + " ");
					if ((low + high).equals("FA00")) {
						rit += "F";
						pos = k;
					}
					if ((low + high).equals("FB00")) {
						rit += "A";
						pos = k;
					}
					if ((low + high).equals("FC00")) {
						rit += "W";
						pos = k;
					}
					if ((low + high).equals("FD00")) {
						rit += "E";
						pos = k;
					}
					if ((low + high).equals("FE00")) {
						rit += "S";
						pos = k;
					}
					if ((low + high).equals("FF00")) {
						rit += "D";
						pos = k;
					}
					if ((low + high).equals("0001")) {
						rit += "N";
						pos = k;
					}
					if ((low + high).equals("0101")) {
						rit += "B";
						pos = k;
					}
					if ((low + high).equals("0401")) {
						rit += "FAWESDNB";
						pos = k;
					}
					k++;
					numFound++;
				}				
				if (i >= Starts.SITE_COUNT) {
					break;
				}
			}
			stream.close();

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.SITE);
			rowNumber = 1;
			// hmon
			i = 0;
			k = 0;
			numFound = 0;
			Set<Integer> posSet = new HashSet<Integer>();
			c = new byte[2];
			stream.skip(44);
			while ((stream.read(c, 0, 2)) != -1) {
				String high = String.format("%02X", c[1]);
				String low = String.format("%02X", c[0]);
				int weapon = Integer.decode("0X" + high + low);
				if (weapon == 0) {
					stream.skip(34l - numFound*2l);
					// Values
					boolean found = false;
					List<Integer> values = new ArrayList<Integer>();
					for (int x = 0; x < numFound; x++) {
						stream.read(c, 0, 2);
						high = String.format("%02X", c[1]);
						low = String.format("%02X", c[0]);
						if (posSet.contains(x)) {
							int fire = new BigInteger(new byte[]{c[1], c[0]}).intValue();
							if (!found) {
								found = true;
							} else {
								//System.out.print("\t");
							}
							//System.out.print(fire);
							values.add(fire);
						}
						stream.skip(6);
					}
					
					//System.out.println("");
					XSSFRow row = sheet.getRow(rowNumber);
					rowNumber++;
					int ind = 0;
					for (Integer mon : values) {
						XSSFCell cell = row.getCell(43+ind, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						cell.setCellValue(mon);
						ind++;
					}
					stream.skip(214l - 34l - numFound*8l);
					numFound = 0;
					posSet.clear();
					k = 0;
					i++;
				} else {
					//System.out.print(low + high + " ");
					if ((low + high).equals("1D00")) {
						posSet.add(k);
					}
					k++;
					numFound++;
				}				
				if (i >= Starts.SITE_COUNT) {
					break;
				}
			}
			stream.close();

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.SITE);
			rowNumber = 1;
			// hcom
			i = 0;
			k = 0;
			numFound = 0;
			posSet = new HashSet<Integer>();
			c = new byte[2];
			stream.skip(44);
			while ((stream.read(c, 0, 2)) != -1) {
				String high = String.format("%02X", c[1]);
				String low = String.format("%02X", c[0]);
				int weapon = Integer.decode("0X" + high + low);
				if (weapon == 0) {
					stream.skip(34l - numFound*2l);
					// Values
					boolean found = false;
					List<Integer> values = new ArrayList<Integer>();
					for (int x = 0; x < numFound; x++) {
						stream.read(c, 0, 2);
						high = String.format("%02X", c[1]);
						low = String.format("%02X", c[0]);
						if (posSet.contains(x)) {
							int fire = new BigInteger(new byte[]{c[1], c[0]}).intValue();
							if (!found) {
								found = true;
							} else {
								//System.out.print("\t");
							}
							//System.out.print(fire);
							values.add(fire);
						}
						stream.skip(6);
					}
					
					//System.out.println("");
					XSSFRow row = sheet.getRow(rowNumber);
					rowNumber++;
					int ind = 0;
					for (Integer mon : values) {
						XSSFCell cell = row.getCell(73+ind, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						cell.setCellValue(mon);
						ind++;
					}
					stream.skip(214l - 34l - numFound*8l);
					numFound = 0;
					posSet.clear();
					k = 0;
					i++;
				} else {
					//System.out.print(low + high + " ");
					if ((low + high).equals("1E00")) {
						posSet.add(k);
					}
					k++;
					numFound++;
				}				
				if (i >= Starts.SITE_COUNT) {
					break;
				}
			}
			stream.close();

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.SITE);
			rowNumber = 1;
			// mon
			i = 0;
			k = 0;
			numFound = 0;
			posSet = new HashSet<Integer>();
			c = new byte[2];
			stream.skip(44);
			while ((stream.read(c, 0, 2)) != -1) {
				String high = String.format("%02X", c[1]);
				String low = String.format("%02X", c[0]);
				int weapon = Integer.decode("0X" + high + low);
				if (weapon == 0) {
					stream.skip(34l - numFound*2l);
					// Values
					boolean found = false;
					List<Integer> values = new ArrayList<Integer>();
					for (int x = 0; x < numFound; x++) {
						stream.read(c, 0, 2);
						high = String.format("%02X", c[1]);
						low = String.format("%02X", c[0]);
						if (posSet.contains(x)) {
							int fire = new BigInteger(new byte[]{c[1], c[0]}).intValue();
							if (!found) {
								found = true;
							} else {
								//System.out.print("\t");
							}
							//System.out.print(fire);
							values.add(fire);
						}
						stream.skip(6);
					}
					
					//System.out.println("");
					XSSFRow row = sheet.getRow(rowNumber);
					rowNumber++;
					int ind = 0;
					for (Integer mon : values) {
						XSSFCell cell = row.getCell(78+ind, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						cell.setCellValue(mon);
						ind++;
					}
					stream.skip(214l - 34l - numFound*4l);
					numFound = 0;
					posSet.clear();
					k = 0;
					i++;
				} else {
					//System.out.print(low + high + " ");
					if ((low + high).equals("0B00")) {
						posSet.add(k);
					}
					k++;
					numFound++;
				}				
				if (i >= Starts.SITE_COUNT) {
					break;
				}
			}
			stream.close();

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.SITE);
			rowNumber = 1;
			// com
			i = 0;
			k = 0;
			numFound = 0;
			posSet = new HashSet<Integer>();
			c = new byte[2];
			stream.skip(44);
			while ((stream.read(c, 0, 2)) != -1) {
				String high = String.format("%02X", c[1]);
				String low = String.format("%02X", c[0]);
				int weapon = Integer.decode("0X" + high + low);
				if (weapon == 0) {
					stream.skip(34l - numFound*2l);
					// Values
					boolean found = false;
					List<Integer> values = new ArrayList<Integer>();
					for (int x = 0; x < numFound; x++) {
						stream.read(c, 0, 2);
						high = String.format("%02X", c[1]);
						low = String.format("%02X", c[0]);
						if (posSet.contains(x)) {
							int fire = new BigInteger(new byte[]{c[1], c[0]}).intValue();
							if (!found) {
								found = true;
							} else {
								//System.out.print("\t");
							}
							//System.out.print(fire);
							values.add(fire);
						}
						stream.skip(6);
					}
					
					//System.out.println("");
					XSSFRow row = sheet.getRow(rowNumber);
					rowNumber++;
					int ind = 0;
					for (Integer mon : values) {
						XSSFCell cell = row.getCell(83+ind, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						cell.setCellValue(mon);
						ind++;
					}
					stream.skip(214l - 34l - numFound*8l);
					numFound = 0;
					posSet.clear();
					k = 0;
					i++;
				} else {
					//System.out.print(low + high + " ");
					if ((low + high).equals("0C00")) {
						posSet.add(k);
					}
					k++;
					numFound++;
				}				
				if (i >= Starts.SITE_COUNT) {
					break;
				}
			}
			stream.close();

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.SITE);
			rowNumber = 1;
			// provdef
			i = 0;
			k = 0;
			numFound = 0;
			posSet = new HashSet<Integer>();
			c = new byte[2];
			stream.skip(44);
			while ((stream.read(c, 0, 2)) != -1) {
				String high = String.format("%02X", c[1]);
				String low = String.format("%02X", c[0]);
				int weapon = Integer.decode("0X" + high + low);
				if (weapon == 0) {
					stream.skip(34l - numFound*2l);
					// Values
					boolean found = false;
					List<Integer> values = new ArrayList<Integer>();
					for (int x = 0; x < numFound; x++) {
						stream.read(c, 0, 2);
						high = String.format("%02X", c[1]);
						low = String.format("%02X", c[0]);
						if (posSet.contains(x)) {
							int fire = new BigInteger(new byte[]{c[1], c[0]}).intValue();
							if (!found) {
								found = true;
							} else {
								//System.out.print("\t");
							}
							//System.out.print(fire);
							values.add(fire);
						}
						stream.skip(6);
					}
					
					//System.out.println("");
					XSSFRow row = sheet.getRow(rowNumber);
					rowNumber++;
					int ind = 0;
					for (Integer mon : values) {
						XSSFCell cell = row.getCell(89+ind, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						cell.setCellValue(mon);
						ind++;
					}
					stream.skip(214l - 34l - numFound*4l);
					numFound = 0;
					posSet.clear();
					k = 0;
					i++;
				} else {
					//System.out.print(low + high + " ");
					if ((low + high).equals("E000")) {
						posSet.add(k);
					}
					k++;
					numFound++;
				}				
				if (i >= Starts.SITE_COUNT) {
					break;
				}
			}
			stream.close();

			// conj
			doit(sheet, "3C00", 55, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return value + "%";
				}
			});

			// alter
			doit(sheet, "3D00", 56, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return value + "%";
				}
			});

			// evo
			doit(sheet, "3E00", 57, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return value + "%";
				}
			});

			// const
			doit(sheet, "3F00", 58, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return value + "%";
				}
			});

			// ench
			doit(sheet, "4000", 59, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return value + "%";
				}
			});

			// thau
			doit(sheet, "4100", 60, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return value + "%";
				}
			});

			// blood
			doit(sheet, "4200", 61, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return value + "%";
				}
			});

			// heal
			doit(sheet, "4600", 62, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return value + "%";
				}
			});

			// disease
			doit(sheet, "1500", 63, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return value + "%";
				}
			});

			// curse
			doit(sheet, "4700", 64, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return value + "%";
				}
			});

			// horror
			doit(sheet, "1800", 65, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return value + "%";
				}
			});

			// holyfire
			doit(sheet, "4400", 66, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return value + "%";
				}
			});

			// holypower
			doit(sheet, "4300", 67, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return value + "%";
				}
			});

			// scry
			doit(sheet, "4800", 68);

			// adventure
			doit(sheet, "C000", 69);

			// voidgate
			doit(sheet, "3900", 48, new CallbackAdapter(){
				@Override
				public String found(String value) {
					return value + "%";
				}
			});

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.SITE);
			rowNumber = 1;
			// summoning
			i = 0;
			k = 0;
			posSet = new HashSet<Integer>();
			int sum1 = 0;
			int sum1count = 0;
			int sum2 = 0;
			int sum2count = 0;
			int sum3 = 0;
			int sum3count = 0;
			int sum4 = 0;
			int sum4count = 0;
			numFound = 0;
			c = new byte[2];
			stream.skip(44);
			while ((stream.read(c, 0, 2)) != -1) {
				String high = String.format("%02X", c[1]);
				String low = String.format("%02X", c[0]);
				int weapon = Integer.decode("0X" + high + low);
				if (weapon == 0) {
					stream.skip(34l - numFound*2l);
					// Values
					for (int x = 0; x < numFound; x++) {
						stream.read(c, 0, 2);
						high = String.format("%02X", c[1]);
						low = String.format("%02X", c[0]);
						if (posSet.contains(x)) {
							int fire = new BigInteger(new byte[]{c[1], c[0]}).intValue();
							if (sum1 == 0 || sum1 == fire) {
								sum1 = fire;
								sum1count++;
							} else if (sum2 == 0 || sum2 == fire) {
								sum2 = fire;
								sum2count++;
							} else if (sum3 == 0 || sum3 == fire) {
								sum3 = fire;
								sum3count++;
							} else if (sum4 == 0 || sum4 == fire) {
								sum4 = fire;
								sum4count++;
							}
						}
						stream.skip(6);
					}

					XSSFRow row = sheet.getRow(rowNumber);
					rowNumber++;
					//String sum = "";
					if (sum1 > 0) {
						//sum += sum1 + "\t" + sum1count;
						XSSFCell cell1 = row.getCell(49, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						cell1.setCellValue(sum1);
						XSSFCell cell2 = row.getCell(50, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						cell2.setCellValue(sum1count);

					}
					if (sum2 > 0) {
						//sum += "\t" + sum2 + "\t" + sum2count;
						XSSFCell cell1 = row.getCell(51, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						cell1.setCellValue(sum2);
						XSSFCell cell2 = row.getCell(52, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						cell2.setCellValue(sum2count);
					}
					if (sum3 > 0) {
						//sum += "\t" + sum3 + "\t" + sum3count;
						XSSFCell cell1 = row.getCell(53, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						cell1.setCellValue(sum3);
						XSSFCell cell2 = row.getCell(54, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						cell2.setCellValue(sum3count);
					}
					if (sum4 > 0) {
						//sum += "\t" + sum3 + "\t" + sum3count;
						XSSFCell cell1 = row.getCell(71, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						cell1.setCellValue(sum4);
						XSSFCell cell2 = row.getCell(72, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
						cell2.setCellValue(sum4count);
					}
					//System.out.println(sum);
					stream.skip(214l - 34l - numFound*8l);
					numFound = 0;
					posSet.clear();
					k = 0;
					sum1 = 0;
					sum1count = 0;
					sum2 = 0;
					sum2count = 0;
					sum3 = 0;
					sum3count = 0;
					sum4 = 0;
					sum4count = 0;
					i++;
				} else {
					//System.out.print(low + high + " ");
					if ((low + high).equals("1200")) {
						posSet.add(k);
					}
					k++;
					numFound++;
				}				
				if (i >= Starts.SITE_COUNT) {
					break;
				}
			}
			stream.close();

			// domspread
			doit(sheet, "1501", 23);

			// turmoil/order
			doit(sheet, "1901", 24);

			// sloth/prod
			doit(sheet, "1A01", 25);

			// cold/heat
			doit(sheet, "1B01", 26);

			// death/growth
			doit(sheet, "1C01", 27);

			// misfortune/luck
			doit(sheet, "1D01", 28);

			// drain/magic
			doit(sheet, "1E01", 29);

			// fire resistence
			doit(sheet, "FB01", 30);

			// cold resistence
			doit(sheet, "FC01", 31);

			// str
			doit(sheet, "FA01", 34);

			// prec
			doit(sheet, "0402", 35);

			// mor
			doit(sheet, "F401", 36);

			// shock resistence
			doit(sheet, "FD01", 32);

			// undying
			doit(sheet, "F801", 37);

			// att
			doit(sheet, "F501", 38);

			// poison resistence
			doit(sheet, "FE01", 33);

			// darkvision
			doit(sheet, "0302", 39);

			// animal awe
			doit(sheet, "0102", 40);

			// reveal
			doit(sheet, "0601", 88);
			
			// def
			doit(sheet, "F601", 91);

			// awe
			doit(sheet, "0202", 100);
			
			// reinvigoration
			doit(sheet, "FF01", 101);
			
			// airshield
			doit(sheet, "0002", 102);
			
			// provdefcom
			doit(sheet, "4A00", 103);
			
			// domconflict
			doit(sheet, "2A00", 104);
			
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
