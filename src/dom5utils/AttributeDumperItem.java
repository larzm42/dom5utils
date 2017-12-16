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

public class AttributeDumperItem {
	private static String[] KNOWN_ITEM_ATTRS = {
		"C600", // Fireres
		"C900", // Coldres
		"C800", // Poisonres
		"C700", // Shockres
		"9D00", // Leadership
		"9700", // str
		"C501", // fixforge
		"9E00", // magic leadership
		"9F00", // undead leadership
		"7001", // inspirational leadership
		"3401", // morale
		"A100", // penetration
		"8300", // pillage
		"B700", // fear
		"A000", // mr
		"0601", // taint
		"7500", // reinvigoration
		"6900", // awe
		"8500", // autospell
		"0A00", // F
		"0B00", // A
		"0C00", // W
		"0D00", // E
		"0E00", // S
		"0F00", // D
		"1000", // N
		"1100", // B
		"1200", // H
		"1400", // elemental
		"1500", // sorcery
		"1600", // all
		"1700", // elemental range
		"1800", // sorcery range
		"1900", // all range
		"2800", // fire ritual range
		"2900", // air ritual range
		"2A00", // water ritual range
		"2B00", // earth ritual range
		"2C00", // astral ritual range
		"2D00", // death ritual range
		"2E00", // nature ritual range
		"2F00", // blood ritual range
		"1901", // darkvision
		"CE01", // limited regeneration
		"BD00", // regeneration
		"6E00", // waterbreathing
		"8601", // stealthb
		"6C00", // stealth
		"9600", // att
		"7901", // def
		"9601", // woundfend
		"1601", // restricted
		"BE00", // berserk
		"2F01", // aging
		"6500", // ivylord
		"A501", // mount
		"A601", // forest
		"A701", // waste
		"A801", // swamp
		"7900", // researchbonus
		"6F00", // gitfofwater
		"9A00", // corpselord
		"6600", // lictorlord
		"8B00", // sumauto
		"D800", // bloodsac
		"6B01", // mastersmith
		"8400", // alch
		"7E00", // eyeloss
		"A301", // armysize
		"8900", // defender
		"7700", // forbidden light
		"C701", // cannotwear
		"C801", // cannotwear
		"7000", // sailingshipsize
		"9A01", // sailingmaxunitsize
		"7100", // flytr
		"7F01", // protf
		"B800", // heretic
		"6301", // autodishealer
		"AA00", // patrolbonus
		"B500", // prec
		"5801", // tmpfiregems
		"5901", // tmpairgems
		"5A01", // tmpwatergems
		"5B01", // tmpearthgems
		"5C01", // tmpastralgems
		"5D01", // tmpdeathgems
		"5E01", // tmpnaturegems
		"5F01", // tmpbloodgems
		"6201", // healer
		"7A00", // supplybonus
		"8100", // airbreathing
		"A901", // mapspeed
		"1E00", // gf
		"1F00", // ga
		"2000", // gw
		"2100", // ge
		"2200", // gs
		"2300", // gd
		"2400", // gn
		"2500", // gb
		"6C01", // reanimation bonus priest
		"6D01", // reanimation bonus death
		"0E01", // dragon mastery
		"AF01", // crossbreeder
		"CD01", // patience
};

	private static List<String> attrList = new ArrayList<String>();
	
	private static Map<String, Integer> Summaries = new HashMap<String, Integer>();
	
	public static void main(String[] args) {
		try {
			FileInputStream stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.ITEM);
			int i = 0;
			long numFound = 0;
			byte[] c = new byte[2];
			stream.skip(Starts.ITEM_ATTRIBUTE_OFFSET);
			while ((stream.read(c, 0, 2)) != -1) {
				String high = String.format("%02X", c[1]);
				String low = String.format("%02X", c[0]);
				int weapon = Integer.decode("0X" + high + low);
				if (weapon == 0) {
					stream.skip(Starts.ITEM_ATTRIBUTE_GAP - numFound*2l);
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
					stream.skip(Starts.ITEM_SIZE - 2 - Starts.ITEM_ATTRIBUTE_GAP - numFound*4l);
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
				if (i >= Starts.ITEM_COUNT) {
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
	        int items = 0;
			System.out.println("Summary:");
			for (Entry<String, Integer> entry : list) {
				if (!Arrays.asList(KNOWN_ITEM_ATTRS).contains(entry.getKey())) {
					System.out.println("id: " + entry.getKey() + " " + entry.getValue());
					indexes++;
					items += entry.getValue();
				}
			}
			System.out.println("---------------------------");
			System.out.println("indexes: " + indexes + " items: " + items);
			stream.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
