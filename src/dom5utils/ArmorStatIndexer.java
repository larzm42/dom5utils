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

public class ArmorStatIndexer {
	
	private static XSSFWorkbook readFile(String filename) throws IOException {
		return new XSSFWorkbook(new FileInputStream(filename));
	}
	
	private static void doit(XSSFSheet sheet, int skip, int column) throws IOException {
		FileInputStream stream = new FileInputStream("Dominions5.exe");			
		stream.skip(Starts.ARMOR);
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

			if (value == 0) {
				//System.out.println("0");
				cell.setCellValue(0);
			} else {
				//System.out.println(Integer.decode("0X" + high + low));
				cell.setCellValue(value);
			}
			stream.skip(102l);
			i++;
			if (i >= Starts.ARMOR_COUNT) {
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

	public static void main(String[] args) {
		run();
	}
	
	public static void run() {
		FileInputStream stream = null;
		try {
	        long startIndex = Starts.ARMOR;
	        int ch;

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(startIndex);
			
			XSSFWorkbook wb = ArmorStatIndexer.readFile("BaseA.xlsx");
			
			FileOutputStream fos = new FileOutputStream("NewBaseA.xlsx");
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
				startIndex = startIndex + 104l;
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

			// type
			doit(sheet, 66, 2);

			// def
			doit(sheet, 62, 3);

			// enc		
			doit(sheet, 64, 4);

			// rcost
			doit(sheet, 68, 5);

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
		
		// protections_by_armor
		try {
	        long startIndex = Starts.ARMOR;
	        int ch;

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(startIndex);
			
			XSSFWorkbook wb = ArmorStatIndexer.readFile("protections_by_armor.xlsx");
			
			FileOutputStream fos = new FileOutputStream("Newprotections_by_armor.xlsx");
			XSSFSheet sheet = wb.getSheetAt(0);

			// name
			InputStreamReader isr = new InputStreamReader(stream, "ISO-8859-1");
	        Reader in = new BufferedReader(isr);
	        int armorNumber = 1;
	        List<Protections> protections = new ArrayList<Protections>();
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
				
				long newIndex = startIndex+36;
				
				short zone = doit2(newIndex);
				while (zone != 0) {
					newIndex+=2;
					short prot = doit2(newIndex);
					protections.add(new Protections(zone, prot, armorNumber));

					newIndex+=2;					
					zone = doit2(newIndex);
				}
				armorNumber++;
				//System.out.println(name);

				stream = new FileInputStream("Dominions5.exe");		
				startIndex = startIndex + 104l;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);
				
			}
			
			int rowNum = 1;
			for (Protections prot : protections) {
				XSSFRow row = sheet.getRow(rowNum);
				XSSFCell cell1 = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell1.setCellValue(prot.zone_number);
				XSSFCell cell2 = row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell2.setCellValue(prot.protection);
				XSSFCell cell3 = row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell3.setCellValue(prot.armor_number);
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
		// attributes_by_armor
		try {
	        long startIndex = Starts.ARMOR;
	        int ch;

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(startIndex);
			
			XSSFWorkbook wb = ArmorStatIndexer.readFile("attributes_by_armor.xlsx");
			
			FileOutputStream fos = new FileOutputStream("Newattributes_by_armor.xlsx");
			XSSFSheet sheet = wb.getSheetAt(0);

			// name
			InputStreamReader isr = new InputStreamReader(stream, "ISO-8859-1");
	        Reader in = new BufferedReader(isr);
	        int armorNumber = 1;
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
				
				long newIndex = startIndex+72;
				
				int attrib = doit3(newIndex);
				while (attrib != 0) {
					attributes.add(new Attributes(attrib, armorNumber));
					newIndex+=4;					
					attrib = doit3(newIndex);
				}
				armorNumber++;
				//System.out.println(name);

				stream = new FileInputStream("Dominions5.exe");		
				startIndex = startIndex + 104l;
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

	}	
	
	private static class Protections {
		short zone_number;
		short protection;
		int armor_number;
		public Protections(short zone_number, short protection, int armor_number) {
			super();
			this.zone_number = zone_number;
			this.protection = protection;
			this.armor_number = armor_number;
		}
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
