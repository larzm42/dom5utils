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
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.Reader;

public class NametypeIndexer {
	
	public static void main(String[] args) {
		run();
	}
	
	public static void run() {
		FileInputStream stream = null;
		try {
			stream = new FileInputStream("Dominions5.exe");
			
			InputStreamReader isr = new InputStreamReader(stream, "ISO-8859-1");
	        Reader in = new BufferedReader(isr);
	        int ch;

			stream.skip(Starts.NAMES);
			int id = 100;
			System.out.println("100");
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
					id ++;
					if (id > Starts.NAMES_COUNT) { break;}
					System.out.println(id +"");
					continue;
				}
				
				System.out.println(name);
			}
			
			in.close();
			
			stream = new FileInputStream("Dominions5.exe");
			
			isr = new InputStreamReader(stream, "ISO-8859-1");
	        in = new BufferedReader(isr);

			System.out.println("**********Fixed Names**********");
			stream.skip(Starts.FIXED_NAMES);
			while ((ch = in.read()) > -1) {
				
				StringBuffer name = new StringBuffer();
				while (ch != 0) {
					name.append((char)ch);
					ch = in.read();
				}
				if (name.length() == 0) {
					continue;
				}
				if (name.toString().equals("Active Mods")) {
					break;
				}
				
				System.out.println(name);
			}
			
			in.close();

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
