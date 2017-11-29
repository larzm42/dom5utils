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
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.NoSuchFileException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class ItemSpriteIndexer {
	
	public static void main(String[] args) {
		run();
	}
	
	public static void run() {
		Map<String, List<String>> map = new HashMap<String, List<String>>();
		FileInputStream stream = null;
		try {
			Path itemsPath = Files.createDirectories(Paths.get("items", "output"));
			Files.walkFileTree(itemsPath, new DirCleaner());

			stream = new FileInputStream("Dominions5.exe");
			
			byte[] b = new byte[32];
			byte[] c = new byte[2];

			stream.skip(Starts.ITEM);
			int id = 1;
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
				List<String> list = map.get(index);
				if (list == null) {
					list = new ArrayList<String>();
					map.put(index, list);
				}
				list.add(name + ": " + offset + " " + index);
				System.out.println(id + ":" + name + ": " + offset + " " + index + "=" + Integer.decode("0X" + index + offset));
				
				int val = Integer.decode("0X" + index + offset);
				if (val > 10000) {
					val -= 9268;
				} else if (val > 9000) {
					val -= 8276;
				} else if (val > 8000) {
					val -= 7337;
				} else if (val > 7000) {
					val -= 6339;
				} else if (val > 6000) {
					val -= 5383;
				} else if (val > 5000) {
					val -= 4444;
				} else if (val > 4000) {
					val -= 3475;
				} else if (val > 3000) {
					val -= 2505;
				} else
				if (val > 2000) {
					val -= 1519;
				} else
				if (val > 1000) {
					val -= 538;
				}
				String.format("%04d", val);
				if (val > 0) {
					String oldFileName1 = "item_" + String.format("%04d", val) + ".tga";
					String newFileName1 = "item" + id + ".tga";

					System.out.println(oldFileName1 + "->" + newFileName1);

					Path old1 = Paths.get("items", oldFileName1);
					Path new1 = Paths.get("items", "output", newFileName1);
					try {
						Files.copy(old1, new1);
					} catch (NoSuchFileException e) {

					}
				} else {
					System.err.println("FAILED");
				}
				
				id++;
				stream.skip(188);
			}
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
