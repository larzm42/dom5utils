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
import java.io.OutputStream;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

public class ItemMonsterDescDumper {
	enum State {
		ITEM, NOTHING, UNIT
	};
	static Map<Integer, String> monsters = new HashMap<Integer, String>();
	static String LAST_ITEM = "Jinn Bottle";
	static String FIRST_UNIT = "Minister of Magic";
	static String LAST_UNIT = "Lord of Fertility";
	static Set<Integer> blockedMonsterIds = new HashSet<Integer>();
	
	public static void main(String[] args) {
		run();
	}
	
	public static void run() {
		State state = State.ITEM;
		
		FileInputStream stream = null;
		try {
			byte[] b32 = new byte[32];
			byte[] c = new byte[2];

			stream = new FileInputStream("Dominions5.exe");
			stream.skip(Starts.MONSTER);
			int id = 1;
			while (stream.read(b32, 0, 32) != -1) {
				stream.skip(4);
				stream.read(c, 0, 2);
				
				StringBuffer name = new StringBuffer();
				for (int i = 0; i < 32; i++) {
					if (b32[i] != 0) {
						name.append(new String(new byte[] {b32[i]}));
					}
				}
				if (name.toString().equals("end")) {
					break;
				}
				monsters.put(id, name.toString().toUpperCase());
				id++;
				stream.skip(226 + 24);
			}
			stream.close();

			
			stream = new FileInputStream("Dominions5.exe");
			stream.skip(Starts.ITEM_AND_MONSTER_DESC_INDEX);

			List<Integer> indexes = new ArrayList<Integer>();
			int index = -1;
			while (index != 0) {
				byte[] d = new byte[4];
				stream.read(d, 0, 4);
				String high1 = String.format("%02X", d[3]);
				String low1 = String.format("%02X", d[2]);
				String high = String.format("%02X", d[1]);
				String low = String.format("%02X", d[0]);
				index = new BigInteger(high1 + low1 + high + low, 16).intValue();
				if (index != 0) {
					indexes.add(index);
				}
			}
			stream.close();
			
			Path itemDescPath = Files.createDirectories(Paths.get("items", "desc"));
			Files.walkFileTree(itemDescPath, new DirCleaner());
			Path descPath = Files.createDirectories(Paths.get("monsters", "desc"));
			Files.walkFileTree(descPath, new DirCleaner());

			stream = new FileInputStream("Dominions5.exe");
			byte[] b = new byte[1];
			int firstIndex = indexes.get(0);
			stream.skip(Starts.ITEM_AND_MONSTER_DESC);
			List<String> names = new ArrayList<String>();
			String desc = null;

			for (Integer offset : indexes) {
				StringBuffer buffer = new StringBuffer();
				stream = new FileInputStream("Dominions5.exe");
				stream.skip(Starts.ITEM_AND_MONSTER_DESC+offset-firstIndex);
				while (stream.read(b) != -1) {
					if (b[0] != 0) {
						buffer.append(new String(new byte[] {b[0]}));
					} else {
						break;
					}
				}
				//System.out.println(buffer.toString());
				stream.close();

				if (buffer.toString().startsWith(":")) {
					if (buffer.toString().substring(1).equals(FIRST_UNIT)) {
						state = State.UNIT;
					}
					names.add(buffer.toString().substring(1));
					desc = null;
				} else {
					desc = buffer.toString();
				}

				if (names.size() > 0 && desc != null) {
					for (String name : names) {
						if (state == State.ITEM) {
							Path path = Paths.get("items", "desc", name.replaceAll("[^a-zA-Z0-9\\-]", "") + ".txt");
							OutputStream os = Files.newOutputStream(path);
							os.write(desc.getBytes());
							os.close();
							if (name.equals(LAST_ITEM)) {
								state = State.NOTHING;
							}
						}
						if (state == State.UNIT) {
							List<Integer> idsInt = new ArrayList<Integer>();
							if (name.toString().startsWith("mon ")) {
								idsInt.add(Integer.valueOf(name.substring(4)));
								blockedMonsterIds.add(Integer.valueOf(name.substring(4)));
							} else if (name.toString().startsWith("mon")) {
								idsInt.add(Integer.valueOf(name.substring(3)));
								blockedMonsterIds.add(Integer.valueOf(name.substring(3)));
							} else {
								Set<Entry<Integer, String>> entrySet = monsters.entrySet();
								for (Map.Entry<Integer, String> entry : entrySet) {
									if (entry.getValue().equals(name.toString().toUpperCase())) {
										if (!blockedMonsterIds.contains(entry.getKey())) {
											idsInt.add(entry.getKey());
										}
									}
								}
							}
							for (Integer idInt : idsInt) {
								Path path = Paths.get("monsters", "desc", String.format("%04d", idInt) + ".txt");
								OutputStream os = Files.newOutputStream(path);
								os.write(desc.getBytes());
								os.close();
							}
							if (name.equals(LAST_UNIT)) {
								state = State.NOTHING;
							}
						}
					}
					names.clear();
					desc = null;
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
