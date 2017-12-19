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

public abstract class AbstractStatIndexer {
	public static String EXE_NAME = "Dominions5.exe";
	
	private static Map<String, boolean[]> largeBitmapCache = new HashMap<String, boolean[]>();

	protected static class CacheKey {
		String key;
		int id;
		public CacheKey(String key, int id) {
			super();
			this.key = key;
			this.id = id;
		}
		
	}
	protected static Map<CacheKey, String> attrCache = new HashMap<CacheKey, String>();
	
	private static int MASK[] = {0x0001, 0x0002, 0x0004, 0x0008, 0x0010, 0x0020, 0x0040, 0x0080,
			0x0100, 0x0200, 0x0400, 0x0800, 0x1000, 0x2000, 0x4000, 0x8000};

	protected static XSSFWorkbook readFile(String filename) throws IOException {
		return new XSSFWorkbook(new FileInputStream(filename));
	}
	
	protected static short getBytes1(long skip) throws IOException {
		FileInputStream stream = new FileInputStream(EXE_NAME);			
		short value = 0;
		byte[] c = new byte[1];
		stream.skip(skip);
		while ((stream.read(c, 0, 1)) != -1) {
			String low = String.format("%02X", c[0]);
			value = new BigInteger(low, 16).shortValue();
			if (value > 100) {
				value = new BigInteger("FF" + low, 16).shortValue();
			}
			break;
		}
		stream.close();
		return value;
	}
	
	protected static short getBytes2(long skip) throws IOException {
		FileInputStream stream = new FileInputStream(EXE_NAME);			
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

	protected static int getBytes4(long skip) throws IOException {
		FileInputStream stream = new FileInputStream(EXE_NAME);			
		int value = 0;
		byte[] c = new byte[4];
		stream.skip(skip);
		while ((stream.read(c, 0, 4)) != -1) {
			String high1 = String.format("%02X", c[3]);
			String low1 = String.format("%02X", c[2]);
			String high2 = String.format("%02X", c[1]);
			String low2 = String.format("%02X", c[0]);
			value = new BigInteger(high1 + low1 + high2 + low2, 16).intValue();
			break;
		}
		stream.close();
		return value;
	}
	
	protected static long getBytes6(long skip) throws IOException {
		FileInputStream stream = new FileInputStream(EXE_NAME);			
		long value = 0;
		byte[] c = new byte[6];
		stream.skip(skip);
		while ((stream.read(c, 0, 6)) != -1) {
			String high1 = String.format("%02X", c[5]);
			String low1 = String.format("%02X", c[4]);
			String high2 = String.format("%02X", c[3]);
			String low2 = String.format("%02X", c[2]);
			String high3 = String.format("%02X", c[1]);
			String low3 = String.format("%02X", c[0]);
			
			value = new BigInteger(high1 + low1 + high2 + low2 + high3 + low3, 16).longValue();
			break;
		}
		stream.close();
		return value;
	}
	
	protected static String getString(long skip) throws IOException {
		FileInputStream stream = new FileInputStream(EXE_NAME);			
		InputStreamReader isr = new InputStreamReader(stream, "ISO-8859-1");
        Reader in = new BufferedReader(isr);
        int ch;
		StringBuffer name = new StringBuffer();
		stream.skip(skip);
		ch = in.read();
		while (ch != 0) {
			name.append((char)ch);
			ch = in.read();
		}
		in.close();
		stream.close();
		return name.toString();
	}
	
	protected static void putBytes2(XSSFSheet sheet, int skip, int column, long start, long size, int count, Callback callback) throws IOException {
		FileInputStream stream = new FileInputStream(EXE_NAME);			
		stream.skip(start);
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
			if (value > 60000) {
				value = new BigInteger("FFFF" + high + low, 16).intValue();
			}

			if (callback != null) {
				cell.setCellValue(callback.found(Integer.toString(value)));
			} else {
				cell.setCellValue(value);
			}

			stream.skip(size-2);
			i++;
			if (i >= count) {
				break;
			}
		}
		stream.close();
	}
	
	protected static List<AttributeValue> getAttributes(long skip, long attrGap) throws IOException {
		return getAttributes(skip, attrGap, 4l);
	}
	protected static List<AttributeValue> getAttributes(long skip, long attrGap, long attrValueLength) throws IOException {
		FileInputStream stream = new FileInputStream(EXE_NAME);	
		List<AttributeValue> attrList = new ArrayList<AttributeValue>();
		stream.skip(skip);
		byte[] c = new byte[2];
		while ((stream.read(c, 0, 2)) != -1) {
			String high = String.format("%02X", c[1]);
			String low = String.format("%02X", c[0]);
			if (Integer.decode("0X" + high + low) != 0) {
				AttributeValue val = new AttributeValue(low + high);
				attrList.add(val);
			} else {
				stream.skip(attrGap - attrList.size()*2l);
				// Values
				for (AttributeValue attrVal : attrList) {
					byte[] d = new byte[4];
					stream.read(d, 0, 4);
					String high1 = String.format("%02X", d[3]);
					String low1 = String.format("%02X", d[2]);
					high = String.format("%02X", d[1]);
					low = String.format("%02X", d[0]);
					attrVal.values.add(Integer.toString(new BigInteger(high1 + low1 + high + low, 16).intValue()));
					stream.skip(attrValueLength-4);
				}
				break;
			}				
		}
		List<AttributeValue> newList = new ArrayList<AttributeValue>();
		for (AttributeValue attrVal : attrList) {
			if (!newList.contains(attrVal)) {
				newList.add(attrVal);
			} else {
				AttributeValue oldAttr = newList.get(newList.indexOf(attrVal));
				oldAttr.values.add(attrVal.values.get(0));
			}
		}
		stream.close();
		return newList;
	}
	
	protected static void putMultipleAttributes(XSSFSheet sheet, String attr, long attrStart, long attrGap, int column, long start, long size, int count, Callback callback) throws IOException {
		FileInputStream stream = new FileInputStream(EXE_NAME);			
		stream.skip(start);
		int i = 0;
		int k = 0;
		Set<Integer> posSet = new HashSet<Integer>();
		long numFound = 0;
		byte[] c = new byte[2];
		stream.skip(attrStart);
		while ((stream.read(c, 0, 2)) != -1) {
			String high = String.format("%02X", c[1]);
			String low = String.format("%02X", c[0]);
			int attribute = Integer.decode("0X" + high + low);
			if (attribute == 0) {
				stream.skip(attrGap - numFound*2l);
				int numColumns = 0;
				// Values
				for (int x = 0; x < numFound; x++) {
					byte[] d = new byte[4];
					stream.read(d, 0, 4);
					String high1 = String.format("%02X", d[3]);
					String low1 = String.format("%02X", d[2]);
					high = String.format("%02X", d[1]);
					low = String.format("%02X", d[0]);
					if (posSet.contains(x)) {
						int value = new BigInteger(high1 + low1 + high + low, 16).intValue();
						XSSFRow row = sheet.getRow(i+1);
						XSSFCell cell = row.getCell(column+numColumns, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);	
						if (callback == null) {
							cell.setCellValue(value);
						} else {
							if (callback.found(Integer.toString(value)) != null) {
								cell.setCellValue(callback.found(Integer.toString(value)));
							}
						}
						numColumns++;
					}
				}

				stream.skip(size - 2 - attrGap - numFound*4l);
				numFound = 0;
				posSet.clear();
				k = 0;
				i++;
			} else {
				if (attr.indexOf(low + high) != -1) {
					posSet.add(k);
				}
				k++;
				numFound++;
			}				
			if (i >= count) {
				break;
			}
		}
		stream.close();
	}

	protected static void putAttribute(XSSFSheet sheet, String attr, long attrStart, long attrGap, int column, long start, long size, int count, boolean append, Callback callback) throws IOException {
        FileInputStream stream = new FileInputStream(EXE_NAME);			
		stream.skip(start);
		int rowNumber = 1;
		int i = 0;
		int k = 0;
		int pos = -1;
		long numFound = 0;
		byte[] c = new byte[2];
		stream.skip(attrStart);
		while ((stream.read(c, 0, 2)) != -1) {
			String high = String.format("%02X", c[1]);
			String low = String.format("%02X", c[0]);
			int attribute = Integer.decode("0X" + high + low);
			if (attribute == 0) {
				boolean found = false;
				int value = 0;
				stream.skip(attrGap - numFound*2l);
				
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
						found = true;
					}
				}
				
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
				stream.skip(size - 2 - attrGap - numFound*4l);
				numFound = 0;
				pos = -1;
				k = 0;
				i++;
			} else {
				if (attr.indexOf(low + high) != -1) {
					if (pos == -1) {
						pos = k;
					}
				}
				k++;
				numFound++;
			}				
			if (i >= count) {
				break;
			}
		}
		stream.close();
	}

	protected static boolean[] largeBitmap(String fieldName, long bitmapOffset, long start, long size, int count, String[][] values) throws IOException {
		return largeBitmap(fieldName, bitmapOffset, start, size, count, values, false);
	}
	
	protected static boolean[] largeBitmap(String fieldName, long bitmapOffset, long start, long size, int count, String[][] values, boolean debug) throws IOException {
		boolean[] boolArray = largeBitmapCache.get(fieldName);
		if (boolArray != null) {
			return boolArray;
		}
		FileInputStream stream = new FileInputStream(EXE_NAME);			
		stream.skip(start);
		int i = 0;
		byte[] c = new byte[2];
		boolArray = new boolean[count];
		stream.skip(bitmapOffset);
		while ((stream.read(c, 0, 2)) != -1) {
			boolean found = false;
			if (debug) { System.out.print("(" + (i+1) + ") "); }
			for (int k = 0; k < values.length; k++) {
				String high = String.format("%02X", c[1]);
				String low = String.format("%02X", c[0]);
				int val = Integer.decode("0X" + high + low);
				if (val > 0) {
					if (debug) { System.out.print((k+1)+":{"); }
					for (int j=0; j < 16; j++) {
						if ((val & MASK[j]) != 0) {
							if (debug) { System.out.print((found?",":"") + (values[k][j].equals("")?("*****"+(j+1)+"*****"):values[k][j])); }
							if (values[k][j].equals(fieldName)) {
								found = true;
							}
						}
					}
					if (debug) { System.out.print("}"); }
				}
				if (k < values.length-1) {
					stream.read(c, 0, 2);
				}
			}
			boolArray[i] = found;
			if (debug) { System.out.println(" "); }
			stream.skip(bitmapOffset);
			i++;
			if (i >= count) {
				break;
			}
		}
		stream.close();
		largeBitmapCache.put(fieldName, boolArray);
		return boolArray;
	}

	protected static List<String> largeBitmap(long start, String[][] values) throws IOException {
		FileInputStream stream = new FileInputStream(EXE_NAME);		
		List<String> boolList = new ArrayList<String>();
		stream.skip(start);
		byte[] c = new byte[2];
		stream.read(c, 0, 2);
		for (int k = 0; k < values.length; k++) {
			String high = String.format("%02X", c[1]);
			String low = String.format("%02X", c[0]);
			int val = Integer.decode("0X" + high + low);
			if (val > 0) {
				for (int j=0; j < 16; j++) {
					if ((val & MASK[j]) != 0) {
						if (!values[k][j].equals("")) {
							boolList.add(values[k][j]);
						}
					}
				}
			}
			if (k < values.length-1) {
				stream.read(c, 0, 2);
			}
		}
		stream.close();
		return boolList;
	}

	protected static class Attribute {
		int object_number;
		int attribute;
		long raw_value;
		public Attribute(int object_number, int attribute, long raw_value) {
			super();
			this.object_number = object_number;
			this.attribute = attribute;
			this.raw_value = raw_value;
		}
	}
	
	protected static class AttributeValue {
		String attribute;
		List<String> values = new ArrayList<String>();
		
		public AttributeValue(String attribute) {
			super();
			this.attribute = attribute;
		}
		@Override
		public int hashCode() {
			final int prime = 31;
			int result = 1;
			result = prime * result + ((attribute == null) ? 0 : attribute.hashCode());
			return result;
		}
		@Override
		public boolean equals(Object obj) {
			if (this == obj)
				return true;
			if (obj == null)
				return false;
			if (getClass() != obj.getClass())
				return false;
			AttributeValue other = (AttributeValue) obj;
			if (attribute == null) {
				if (other.attribute != null)
					return false;
			} else if (!attribute.equals(other.attribute))
				return false;
			return true;
		}
	}
	
	protected static class Effect {
		int record_number;
		int effect_number;
		int duration;
		int ritual;
		String object_type;
		int raw_argument;
		long modifiers_mask;
		int range_base;
		int range_per_level;
		int range_strength_divisor;
		int area_base;
		int area_per_level;
		int area_battlefield_pct;
		int sound_number;
		int flight_sprite_number;
		int flight_sprite_length;
		int explosion_sprite_number;
		int explosion_sprite_length;

	}
	

}


