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

public class MercenaryStatIndexer {
	
	private static XSSFWorkbook readFile(String filename) throws IOException {
		return new XSSFWorkbook(new FileInputStream(filename));
	}
	
	private static void doit(XSSFSheet sheet, int skip, int column) throws IOException {
		FileInputStream stream = new FileInputStream("Dominions5.exe");			
		stream.skip(Starts.MERCENARY);
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
			int value = Integer.decode("0X" + high + low);
			if (value == 0) {
				//System.out.println("0");
				cell.setCellValue(0);
			} else {
				//System.out.println(Integer.decode("0X" + high + low));
				cell.setCellValue(Integer.decode("0X" + high + low));
			}
			stream.skip(310l);
			i++;
			if (i >= Starts.MERCENARY_COUNT) {
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
	        long startIndex = Starts.MERCENARY;
	        int ch;

			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(startIndex);
			
			XSSFWorkbook wb = MercenaryStatIndexer.readFile("Mercenary.xlsx");
			
			FileOutputStream fos = new FileOutputStream("NewMercenary.xlsx");
			XSSFSheet sheet = wb.getSheetAt(0);

//			XSSFRow titleRow = sheet.getRow(0);
//			int cellNum = 0;
//			XSSFCell titleCell = titleRow.getCell(cellNum);
//			while (titleCell != null) {
//				cellNum++;
//				titleCell = titleRow.getCell(cellNum);
//			}
			
			// name
			InputStreamReader isr = new InputStreamReader(stream, "ISO-8859-1");
	        Reader in = new BufferedReader(isr);
	        int rowNumber = 1;
			while ((ch = in.read()) > -1) {
				StringBuffer name = new StringBuffer();
				while (ch != 0) {
//					if (ch == 0xc3) {
//						ch = in.read();
//						ch += 0x40;
//					}
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
				startIndex = startIndex + 312l;
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

			// bossname
			stream = new FileInputStream("Dominions5.exe");			
			startIndex = Starts.MERCENARY + 36;
			stream.skip(startIndex);
			isr = new InputStreamReader(stream, "ISO-8859-1");
	        in = new BufferedReader(isr);
	        rowNumber = 1;
			while ((ch = in.read()) > -1) {
				StringBuffer name = new StringBuffer();
				while (ch != 0) {
//					if (ch == 0xc3) {
//						ch = in.read();
//						ch += 0x40;
//					}
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
				startIndex = startIndex + 312l;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);

				//System.out.println(name);
				
				XSSFRow row = sheet.getRow(rowNumber);
				rowNumber++;
				XSSFCell cell = row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				cell.setCellValue(name.toString());
				
				if (rowNumber > Starts.MERCENARY_COUNT) {
					break;
				}
			}
			in.close();
			stream.close();

			// com
			doit(sheet, 74, 3);

			// unit
			doit(sheet, 78, 4);

			// nrunits
			doit(sheet, 82, 5);

			// minmen
			doit(sheet, 86, 7);

			// minpay
			doit(sheet, 90, 8);

			// xp
			doit(sheet, 94, 9);

			// randequip
			doit(sheet, 158, 10);

			// recrate
			doit(sheet, 306, 11);

			// item1
			stream = new FileInputStream("Dominions5.exe");			
			startIndex = Starts.MERCENARY + 162;
			stream.skip(startIndex);
			isr = new InputStreamReader(stream, "ISO-8859-1");
	        in = new BufferedReader(isr);
	        rowNumber = 1;
			while ((ch = in.read()) > -1) {
				StringBuffer name = new StringBuffer();
				while (ch != 0) {
					if (ch == 0xc3) {
						ch = in.read();
						ch += 0x40;
					}
					name.append((char)ch);
					ch = in.read();
				}
				in.close();

				stream = new FileInputStream("Dominions5.exe");		
				startIndex = startIndex + 312l;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);

				//System.out.println(name);
				
				XSSFRow row = sheet.getRow(rowNumber);
				rowNumber++;
				if (name.length() != 0) {
					XSSFCell cell = row.getCell(12, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					cell.setCellValue(name.toString());
				}
				
				if (rowNumber > Starts.MERCENARY_COUNT) {
					break;
				}
			}
			in.close();
			stream.close();

			// item2
			stream = new FileInputStream("Dominions5.exe");			
			startIndex = Starts.MERCENARY + 198;
			stream.skip(startIndex);
			isr = new InputStreamReader(stream, "ISO-8859-1");
	        in = new BufferedReader(isr);
	        rowNumber = 1;
			while ((ch = in.read()) > -1) {
				StringBuffer name = new StringBuffer();
				while (ch != 0) {
					if (ch == 0xc3) {
						ch = in.read();
						ch += 0x40;
					}
					name.append((char)ch);
					ch = in.read();
				}
				in.close();

				stream = new FileInputStream("Dominions5.exe");		
				startIndex = startIndex + 312l;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);

				//System.out.println(name);
				
				XSSFRow row = sheet.getRow(rowNumber);
				rowNumber++;
				if (name.length() != 0) {
					XSSFCell cell = row.getCell(13, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
					cell.setCellValue(name.toString());
				}
				
				if (rowNumber > Starts.MERCENARY_COUNT) {
					break;
				}
			}
			in.close();
			stream.close();
			
			// eramask
			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.MERCENARY-2);
			rowNumber = 1;
			int i = 0;
			byte[] c = new byte[1];
			while ((stream.read(c, 0, 1)) != -1) {
				XSSFRow row = sheet.getRow(rowNumber);
				rowNumber++;
				XSSFCell cell = row.getCell(14, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				String low = String.format("%02X", c[0]);
				int weapon = Integer.decode("0X00" + low);
				if (weapon == 0) {
					//System.out.println("");
					cell.setCellValue("");
				} else {
					//System.out.println(Integer.decode("0X00" + low));
					cell.setCellValue(Integer.decode("0X00" + low));
				}
				stream.skip(311l);
				i++;
				if (i >= Starts.MERCENARY_COUNT) {
					break;
				}
			}
			stream.close();

			// level
			stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.MERCENARY-1);
			rowNumber = 1;
			i = 0;
			c = new byte[1];
			while ((stream.read(c, 0, 1)) != -1) {
				XSSFRow row = sheet.getRow(rowNumber);
				rowNumber++;
				XSSFCell cell = row.getCell(6, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
				String low = String.format("%02X", c[0]);
				int weapon = Integer.decode("0X00" + low);
				if (weapon == 0) {
					//System.out.println("0");
					cell.setCellValue(0);
				} else {
					//System.out.println(Integer.decode("0X" + low));
					cell.setCellValue(Integer.decode("0X00" + low));
				}
				stream.skip(311l);
				i++;
				if (i >= Starts.MERCENARY_COUNT) {
					break;
				}
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
