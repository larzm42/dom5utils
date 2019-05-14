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
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.NoSuchFileException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Comparator;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.SortedMap;
import java.util.StringTokenizer;
import java.util.TreeMap;
import java.util.TreeSet;

public class ItemSpriteIndexer {
	
	static Map<String, List<String>> indexToInt = new HashMap<String, List<String>>();

	public static void main(String[] args) {
		run();
	}
	
	public static void run() {
		indexToInt.put("first", new ArrayList<String>(Arrays.asList(new String[]{"00", "01"})));
		indexToInt.put("1h", new ArrayList<String>(Arrays.asList(new String[]{"03"})));
		indexToInt.put("2h", new ArrayList<String>(Arrays.asList(new String[]{"07"})));
		indexToInt.put("bow", new ArrayList<String>(Arrays.asList(new String[]{"0B"})));
		indexToInt.put("shield", new ArrayList<String>(Arrays.asList(new String[]{"0F"})));
		indexToInt.put("armor", new ArrayList<String>(Arrays.asList(new String[]{"13"})));
		indexToInt.put("helm", new ArrayList<String>(Arrays.asList(new String[]{"17"})));
		indexToInt.put("boot", new ArrayList<String>(Arrays.asList(new String[]{"1B"})));
		indexToInt.put("misc", new ArrayList<String>(Arrays.asList(new String[]{"1F"})));
		indexToInt.put("potion", new ArrayList<String>(Arrays.asList(new String[]{"23"})));
		indexToInt.put("prost", new ArrayList<String>(Arrays.asList(new String[]{"27"})));
		
		FileInputStream stream = null;
		try {
			Path monstersPath = Files.createDirectories(Paths.get("items", "output"));
			Files.walkFileTree(monstersPath, new DirCleaner());

			stream = new FileInputStream("item.trs");
			stream.skip(Starts.ITEM_TRS_INDEX);

			SortedMap<Integer, String> indexes1 = new TreeMap<Integer, String>();
			int index1 = 0;
			while (index1 != -1) {
				byte[] d = new byte[4];
				stream.read(d, 0, 4);
				String high1 = String.format("%02X", d[3]);
				String low1 = String.format("%02X", d[2]);
				String high = String.format("%02X", d[1]);
				String low = String.format("%02X", d[0]);
				index1 = new BigInteger(low + high + low1 + high1, 16).intValue();
				if (index1 == -1) break;
				
				StringBuffer buffer = new StringBuffer();
				byte[] b = new byte[1];
				while (stream.read(b) != -1) {
					if (b[0] != 0) {
						buffer.append(new String(new byte[] {b[0]}));
					} else {
						break;
					}
				}
				indexes1.put(index1, buffer.toString());
			}
			stream.close();
			
			indexes1.put(0, "first");
			for (Map.Entry<Integer, String> entry : indexes1.entrySet()) {
				System.out.println(entry.getValue() + ": " + entry.getKey());
			}
			
			Map<String, List<String>> map = new HashMap<String, List<String>>();
			stream = new FileInputStream("Dominions5.exe");

			byte[] b = new byte[32];
			byte[] c = new byte[2];

			stream.skip(Starts.ITEM);
			int id = 1;
			Set<String> indexes = new HashSet<String>();
			while (stream.read(b, 0, 32) != -1) {
				stream.skip(10);
				stream.read(c, 0, 2);

				StringBuffer name = new StringBuffer();
				for (int i = 0; i < 32; i++) {
					if (b[i] != 0) {
						name.append(new String(new byte[] {b[i]}));
					}
				}
				if (name.toString().equals("end")) {
					break;
				}
				String index = String.format("%02X", c[1]);
				String offset = String.format("%02X", c[0]);
				indexes.add(index);
				List<String> list = map.get(index);
				if (list == null) {
					list = new ArrayList<String>();
					map.put(index, list);
				}
				list.add(id + ": " + name + ": " + offset + " " + index);
				System.out.println(id + ":" + name + ": " + offset + " " + index);

				id++;
				stream.skip(196);
			}
			TreeSet<String> sorted = new TreeSet<String>(new Comparator<String>() {
				@Override
				public int compare(String o1, String o2) {
					return Integer.decode("0X" + o1).compareTo(Integer.decode("0X" + o2));
				}
			});
			sorted.addAll(indexes);
			int things = 0;
			Iterator<String> iter = sorted.iterator();
			while (iter.hasNext()) {
				String ind = iter.next();
				System.out.println(ind);
				List<String> list = map.get(ind);
				List<SortedByOffset> sortedSet = new ArrayList<SortedByOffset>();
				for (String myList : list) {
					sortedSet.add(new SortedByOffset(myList));
				}
				Collections.sort(sortedSet);
				for (SortedByOffset thing : sortedSet) {
					System.out.println("  " + thing.value);
					things++;
				}
			}
			System.out.println("------------------------");
			System.out.println("Indexes:" + indexes.size() + " Items:" + things);
				
			for (Map.Entry<Integer, String> entry : indexes1.entrySet()) {
				List<String> mappings = indexToInt.get(entry.getValue());
				if (mappings != null) {
					boolean first = true;
					int groupNegativeOffset = 0;
					int groupPositiveOffset = 0;
					for (String group : mappings) {
						indexes.remove(group);
						List<String> list = map.get(group);
						List<SortedByOffset> sortedSet = new ArrayList<SortedByOffset>();
						for (String myList : list) {
							sortedSet.add(new SortedByOffset(myList));
						}
						Collections.sort(sortedSet);
						if (first) {
							if (sortedSet.size() == 0) {
								System.err.println("Empty Set: " + group);
								continue;
							}
							SortedByOffset sortedByOffset = sortedSet.get(0);
							groupNegativeOffset = sortedByOffset.getIntValue();
							groupPositiveOffset = entry.getKey();
							first = false;
						}
						int tweak = -1;
						if (entry.getValue().equals("first")) {
							tweak = 3;
						}
						if (entry.getValue().equals("armor")) {
							tweak = 50;
						}
						if (entry.getValue().equals("prost")) {
							tweak = 1;
						}
						if (entry.getValue().equals("shield")) {
							tweak = 3;
						}
						if (entry.getValue().equals("potion")) {
							tweak = 0;
						}
						for (SortedByOffset ugh : sortedSet) {
							int val = groupPositiveOffset - groupNegativeOffset + ugh.getIntValue()+2+tweak;
							System.out.println(ugh.value + ": " + val);
							if (val > 0) {
								StringTokenizer tok = new StringTokenizer(ugh.value);
								String idStr = tok.nextToken();
								idStr = idStr.substring(0, idStr.length()-1);
								String oldFileName1 = "item_" + String.format("%04d", val) + ".tga";
								String newFileName1 = "item" + idStr + ".tga";
								
								System.out.println(oldFileName1 + "->" + newFileName1);

								Path old1 = Paths.get("items", oldFileName1);
								Path new1 = Paths.get("items", "output", newFileName1);
								try {
									Files.copy(old1, new1);
								} catch (NoSuchFileException e) {
									e.printStackTrace();
								} catch (Exception e) {
									e.printStackTrace();
								}
							} else {
								System.err.println("FAILED");
							}

						}
					}
				}
			}

			for (String output : indexes) {
				System.out.println(output);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}

