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
import java.util.List;

public class SpellDescDumper {
	
	public static void main(String[] args) {
		run();
	}
	
	public static void run() {		
		FileInputStream stream = null;
		try {
			stream = new FileInputStream("Dominions5.exe");
			stream.skip(Starts.SPELL_DESC_INDEX);

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
			
			Path spellsDescPath = Files.createDirectories(Paths.get("spells"));
			Files.walkFileTree(spellsDescPath, new DirCleaner());
			
			stream = new FileInputStream("Dominions5.exe");
			byte[] b = new byte[1];
			int firstIndex = indexes.get(0);
			stream.skip(Starts.SPELL_DESC);
			List<String> names = new ArrayList<String>();
			String desc = null;
			for (Integer offset : indexes) {
				StringBuffer buffer = new StringBuffer();
				stream = new FileInputStream("Dominions5.exe");
				stream.skip(Starts.SPELL_DESC+offset-firstIndex);
				while (stream.read(b) != -1) {
					if (b[0] != 0) {
						buffer.append(new String(new byte[] {b[0]}));
					} else {
						break;
					}
				}
				System.out.println(buffer.toString());
				stream.close();
				
				if (buffer.toString().startsWith(":")) {
					names.add(buffer.toString().substring(1));
					desc = null;
				} else {
					desc = buffer.toString();
				}

				if (names.size() > 0 && desc != null) {
					for (String name : names) {
						Path path = Paths.get("spells", name.replaceAll("[^a-zA-Z0-9\\-]", "") + ".txt");
						OutputStream os = Files.newOutputStream(path);
						os.write(desc.getBytes());
						os.close();
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
