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

public class MercenaryStatIndexer extends AbstractStatIndexer {
	
	private static void putBytes2(XSSFSheet sheet, int skip, int column) throws IOException {
		putBytes2(sheet, skip, column, Starts.MERCENARY, Starts.MERCENARY_SIZE, Starts.MERCENARY_COUNT);
	}
	
	public static void main(String[] args) {
		run();
	}
	
	public static void run() {
		FileInputStream stream = null;
		try {
	        long startIndex = Starts.MERCENARY;
	        int ch;

			stream = new FileInputStream(EXE_NAME);			
			stream.skip(startIndex);
			
			XSSFWorkbook wb = MercenaryStatIndexer.readFile("Mercenary_Template.xlsx");
			
			FileOutputStream fos = new FileOutputStream("Mercenary.xlsx");
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

				stream = new FileInputStream(EXE_NAME);		
				startIndex = startIndex + 312l;
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

			// bossname
			putString(sheet, 36, 2, Starts.MERCENARY, Starts.MERCENARY_SIZE);

			// com
			putBytes2(sheet, 74, 3);

			// unit
			putBytes2(sheet, 78, 4);

			// nrunits
			putBytes2(sheet, 82, 5);

			// minmen
			putBytes2(sheet, 86, 7);

			// minpay
			putBytes2(sheet, 90, 8);

			// xp
			putBytes2(sheet, 94, 9);

			// randequip
			putBytes2(sheet, 158, 10);

			// recrate
			putBytes2(sheet, 306, 11);

			// item1
			putString(sheet, 162, 12, Starts.MERCENARY, Starts.MERCENARY_SIZE);

			// item2
			putString(sheet, 198, 13, Starts.MERCENARY, Starts.MERCENARY_SIZE);
			
			// eramask
			stream = new FileInputStream(EXE_NAME);			
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
			stream = new FileInputStream(EXE_NAME);			
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
