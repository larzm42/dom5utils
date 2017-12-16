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
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

public class AttributeDumperSite {
	private static String[] KNOWN_ITEM_ATTRS = {
		"0100", // F
		"0200", // A
		"0300", // W
		"0400", // E
		"0500", // S
		"0600", // D
		"0700", // N
		"0800", // B
		"0D00", // gold
		"0E00", // res
		"1400", // sup
		"1300", // unr
		"1600", // exp
		"0F00", // lab
		"1100", // fort
		"1F00", // scales
		"2000", // scales
		"FA00", // rit/ritr
		"FB00", // rit/ritr
		"FC00", // rit/ritr
		"FD00", // rit/ritr
		"FE00", // rit/ritr
		"FF00", // rit/ritr
		"0001", // rit/ritr
		"0101", // rit/ritr
		"0401", // rit/ritr
		"1E00", // hcom
		"1D00", // hmon
		"0C00", // com
		"0B00", // mon
		"3C00", // conj
		"3D00", // alter
		"3E00", // evo
		"3F00", // const
		"4000", // ench
		"4100", // thau
		"4200", // blood
		"4600", // heal
		"1500", // disease
		"4700", // curse
		"1800", // horror
		"4400", // holyfire
		"4300", // holypower
		"4800", // scry
		"C000", // adventure
		"3900", // voidgate
		"1200", // summoning
		"1501", // domspread
		"1901", // turmoil/order
		"1A01", // sloth/prod
		"1B01", // cold/heat
		"1C01", // death/growth
		"1D01", // misfortune/luck
		"1E01", // drain/magic
		"FB01", // fire resistence
		"FC01", // cold resistence
		"FA01", // str
		"0402", // prec
		"F401", // mor
		"FD01", // shock resistence
		"F801", // undying
		"F501", // att
		"FE01", // poison resistence
		"0302", // darkvision
		"0102", // animal awe
		"1401", // throne?
		"0A01", // fortparts
		"0601", // reveal
		"E000", // provdef
		"F601", // def
		"0202", // awe
		"FF01", // reinvigoration
		"0002", // airshield
		"4A00", // provdefcom
};

	private static List<String> attrList = new ArrayList<String>();
	
	private static Map<String, Integer> Summaries = new HashMap<String, Integer>();
	
	public static void main(String[] args) {
		try {
			FileInputStream stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.SITE);
			int i = 0;
			long numFound = 0;
			byte[] c = new byte[2];
			stream.skip(44);
			while ((stream.read(c, 0, 2)) != -1) {
				String high = String.format("%02X", c[1]);
				String low = String.format("%02X", c[0]);
				int weapon = Integer.decode("0X" + high + low);
				if (weapon == 0) {
					stream.skip(34l - numFound*2l);
					System.out.print("id:" + (i+1));
					// Values
					for (int x = 0; x < numFound; x++) {
						byte[] d = new byte[4];
						stream.read(d, 0, 4);
						String high1 = String.format("%02X", d[3]);
						String low1 = String.format("%02X", d[2]);
						high = String.format("%02X", d[1]);
						low = String.format("%02X", d[0]);
						if (!Arrays.asList(KNOWN_ITEM_ATTRS).contains(attrList.get(x))) {
							System.out.print("\n\t" + attrList.get(x) + ": ");
							System.out.print(new BigInteger(high1 + low1 + high + low, 16).intValue() + " ");
						}
					}

					System.out.println("");
					stream.skip(214l - 34l - numFound*8l);
					numFound = 0;
					i++;
					
					for (String ids : attrList) {
						Integer count = Summaries.get(ids);
						if (count == null) {
							Summaries.put(ids, Integer.valueOf(1));
						} else {
							Summaries.put(ids, Integer.valueOf(count.intValue()+1));
						}
					}
					attrList = new ArrayList<String>();
				} else {
					attrList.add(low + high);
					numFound++;
				}				
				if (i >= Starts.SITE_COUNT) {
					break;
				}
			}
			
			Set<Entry<String, Integer>> set = Summaries.entrySet();
	        List<Entry<String, Integer>> list = new ArrayList<Entry<String, Integer>>(set);
	        Collections.sort( list, new Comparator<Map.Entry<String, Integer>>()
	        {
	            public int compare( Map.Entry<String, Integer> o1, Map.Entry<String, Integer> o2 )
	            {
	                return (o2.getValue()).compareTo( o1.getValue() );
	            }
	        } );

	        int indexes = 0;
	        int sites = 0;
			System.out.println("Summary:");
			for (Entry<String, Integer> entry : list) {
				if (!Arrays.asList(KNOWN_ITEM_ATTRS).contains(entry.getKey())) {
					System.out.println("id: " + entry.getKey() + " " + entry.getValue());
					indexes++;
					sites += entry.getValue();
				}
			}
			System.out.println("---------------------------");
			System.out.println("indexes: " + indexes + " sites: " + sites);
			stream.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
