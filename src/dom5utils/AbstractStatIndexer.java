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
import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.Reader;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.StringTokenizer;

public abstract class AbstractStatIndexer {
	protected static String EXE_NAME = "Dominions5.exe";
	
	private static int MASK[] = {0x0001, 0x0002, 0x0004, 0x0008, 0x0010, 0x0020, 0x0040, 0x0080,
			0x0100, 0x0200, 0x0400, 0x0800, 0x1000, 0x2000, 0x4000, 0x8000};

	static private Map<Integer, String> attrKeys = new HashMap<Integer, String>();
	static private Map<Integer, String> effectInfoKeys = new HashMap<Integer, String>();
	static private Map<Long, String> effectModifierKeys = new HashMap<Long, String>();

	static {
		try {
			File attributesFile = new File("attribute_keys.csv");
			FileReader attributesFileReader = new FileReader(attributesFile);
			BufferedReader bufferedReader = new BufferedReader(attributesFileReader);
			String line;
			while ((line = bufferedReader.readLine()) != null) {
				StringTokenizer tok = new StringTokenizer(line, "\t");
				Integer attrNum = Integer.parseInt(tok.nextToken());
				String attrVal = tok.nextToken();
				attrKeys.put(attrNum, attrVal);
			}
			bufferedReader.close();
			
			File effectsInfoFile = new File("effects_info.csv");
			FileReader effectsInfoFileReader = new FileReader(effectsInfoFile);
			bufferedReader = new BufferedReader(effectsInfoFileReader);
			while ((line = bufferedReader.readLine()) != null) {
				StringTokenizer tok = new StringTokenizer(line, "\t");
				Integer attrNum = Integer.parseInt(tok.nextToken());
				String attrVal = tok.nextToken();
				effectInfoKeys.put(attrNum, attrVal);
			}
			bufferedReader.close();
			
			File effectModBitsFile = new File("effect_modifier_bits.csv");
			FileReader effectModBitsFileReader = new FileReader(effectModBitsFile);
			bufferedReader = new BufferedReader(effectModBitsFileReader);
			while ((line = bufferedReader.readLine()) != null) {
				StringTokenizer tok = new StringTokenizer(line, "\t");
				Long attrNum = Long.parseLong(tok.nextToken());
				String attrVal = tok.nextToken();
				effectModifierKeys.put(attrNum, attrVal);
			}
			bufferedReader.close();
			
		} catch (Exception e) {
			System.err.println("Problem reading file");
			e.printStackTrace();
		}

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
	
	protected static int getBytes2(long skip) throws IOException {
		FileInputStream stream = new FileInputStream(EXE_NAME);			
		int value = 0;
		byte[] c = new byte[2];
		stream.skip(skip);
		while ((stream.read(c, 0, 2)) != -1) {
			String high = String.format("%02X", c[1]);
			String low = String.format("%02X", c[0]);
			value = Integer.decode("0X" + high + low).intValue();
			if (value > 60000) {
				value = new BigInteger("FFFF" + high + low, 16).intValue();
			}
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
	
	protected static long getBytes8(long skip) throws IOException {
		FileInputStream stream = new FileInputStream(EXE_NAME);			
		long value = 0;
		byte[] c = new byte[8];
		stream.skip(skip);
		while ((stream.read(c, 0, 8)) != -1) {
			String high0 = String.format("%02X", c[7]);
			String low0 = String.format("%02X", c[6]);
			String high1 = String.format("%02X", c[5]);
			String low1 = String.format("%02X", c[4]);
			String high2 = String.format("%02X", c[3]);
			String low2 = String.format("%02X", c[2]);
			String high3 = String.format("%02X", c[1]);
			String low3 = String.format("%02X", c[0]);
			
			value = new BigInteger(high0 + low0 + high1 + low1 + high2 + low2 + high3 + low3, 16).longValue();
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
	
	protected static void printAttributes(List<Attribute> attributes, BufferedWriter writer) throws IOException {
		for (Attribute attr : attributes) {
			String string = attrKeys.get(attr.attribute);
			if (string == null) {
				string = "Unknown Attribute:<" + attr.attribute+">";
			} else {
				string = string + "<" + attr.attribute + ">";
			}
			writer.write("\t" + string + ": " + attr.raw_value);
			writer.newLine();
		}
	}

	protected static void printEffect(Effect effect, BufferedWriter writer) throws IOException {
		String string = effectInfoKeys.get(effect.effect_number);
		if (string == null) {
			string = "Unknown Effect:<" + effect.effect_number+">";
		} else {
			string = string + "<" + effect.effect_number + ">";
		}
		writer.write("\teffect:" + string);
		writer.newLine();
		writer.write("\tduration: " + effect.duration);
		writer.newLine();
		writer.write("\tritual: " + effect.ritual);
		writer.newLine();
		writer.write("\traw_argument: " + effect.raw_argument);
		writer.newLine();
		writer.write("\trange_base: " + effect.range_base);
		writer.newLine();
		writer.write("\trange_per_level: " + effect.range_per_level);
		writer.newLine();
		writer.write("\trange_strength_divisor: " + effect.range_strength_divisor);
		writer.newLine();
		writer.write("\tarea_base: " + effect.area_base);
		writer.newLine();
		writer.write("\tarea_per_level: " + effect.area_per_level);
		writer.newLine();
		writer.write("\tarea_battlefield_pct: " + effect.area_battlefield_pct);
		writer.newLine();
		writer.write("\tmodifiers_mask: " + effect.modifiers_mask);
		writer.newLine();
		printModiferMask(effect.modifiers_mask, writer);
	}
	
	protected static void printModiferMask(long bitmap, BufferedWriter writer) throws IOException {
		for (Map.Entry<Long, String> mask : effectModifierKeys.entrySet()) {
			if ((mask.getKey().longValue() & bitmap) != 0) {
				writer.write("\t\t" + mask.getValue());
				writer.newLine();
			}
		}
	}
	
	protected static void dumpUnknown(Set<String> attr, BufferedWriter writer) throws IOException {
		for (String att : attr) {
			writer.write(att);
			writer.newLine();
		}
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
		long raw_argument;
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


