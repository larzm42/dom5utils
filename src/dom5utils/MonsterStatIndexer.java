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
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.Reader;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MonsterStatIndexer extends AbstractStatIndexer {
	public static String[] unit_columns = {"id", "name", "wpn1", "wpn2", "wpn3", "wpn4", "wpn5", "wpn6", "wpn7", "armor1", "armor2", "armor3", "armor4", 
			"rt", "reclimit", "basecost", "gmon", "gcom", "rcost", "size", "ressize", "hp", "prot", "mr", "mor", "str", "att", "def", "prec", "enc", 
			"mapmove", "ap", "ambidextrous", "mounted", "reinvigoration", "leader", "undeadleader", "magicleader", "startage", "maxage", "hand", "head", 
			"body", "foot", "misc", "pathcost", "startdom", "inn", "F", "A", "W", "E", "S", "D", "N", "B", "H", "rand1", "nbr1", "link1", "mask1", "rand2", 
			"nbr2", "link2", "mask2", "rand3", "nbr3", "link3", "mask3", "rand4", "nbr4", "link4", "mask4", "holy", "inquisitor", "mind", "inanimate", 
			"undead", "demon", "magicbeing", "stonebeing", "animal", "coldblood", "female", "forestsurvival", "mountainsurvival", "wastesurvival", 
			"swampsurvival", "cavesurvival", "aquatic", "amphibian", "pooramphibian", "float", "flying", "stormimmune", "teleport", "immobile", 
			"noriverpass", "sailingshipsize", "sailingmaxunitsize", "stealthy", "illusion", "spy", "assassin", "patience", "seduce", "succubus", 
			"corrupt", "heal", "immortal", "reinc", "noheal", "neednoteat", "homesick", "undisciplined", "formationfighter", "slave", "standard", 
			"inspirational", "taskmaster", "beastmaster", "bodyguard", "waterbreathing", "iceprot", "invulnerable", "slashres", "bluntres", "pierceres", 
			"shockres", "fireres", "coldres", "poisonres", "voidsanity", "darkvision", "blind", "animalawe", "awe", "halt", "fear", "berserk", "cold", 
			"heat", "fireshield", "banefireshield", "damagerev", "poisoncloud", "diseasecloud", "slimer", "mindslime", "disbel", "reform", "regeneration", 
			"reanimator", "barbs", "petrify", "eyeloss", "ethereal", "ethtrue", "deathcurse", "trample", "trampswallow", "stormpower", "firepower", 
			"coldpower", "darkpower", "chaospower", "magicpower", "winterpower", "springpower", "summerpower", "fallpower", "forgebonus", "fixforgebonus", 
			"mastersmith", "resources", "autohealer", "autodishealer", "alch", "nobadevents", "event", "insane", "shatteredsoul", "leper", "taxcollector", 
			"gem", "gemprod", "chaosrec", "pillagebonus", "patrolbonus", "castledef", "siegebonus", "incprovdef", "supplybonus", "incunrest", "popkill", 
			"researchbonus", "drainimmune", "inspiringres", "douse", "sacr", "ivylord", "crossbreeder", "makepearls", "pathboost", "allrange", "voidsum", 
			"heretic", "elegist", "shapechange", "firstshape", "secondshape", "secondtmpshape", "landshape", "watershape", "forestshape", "plainshape", 
			"unique", "fixedname", "special", "nametype", "type", "typeclass", "from", "summon", "n_summon", "autosum", "n_autosum", "batstartsum1", 
			"batstartsum2", "domsummon1", "domsummon2", "bloodvengeance", "bringeroffortune", "realm1", "realm2", "realm3", "batstartsum3", "batstartsum4", 
			"batstartsum5", "batstartsum1d6", "batstartsum2d6", "batstartsum3d6", "batstartsum4d6", "batstartsum5d6", "batstartsum6d6", "turmoilsummon", 
			"coldsummon", "scalewalls", "baseleadership", "explodeondeath", "startaff", "uwregen", "shrinkhp", "growhp", "transformation", "startingaff", 
			"fixedresearch", "divineins", "lamiabonus", "preanimator", "dreanimator", "mummify", "onebattlespell", "fireattuned", "airattuned", 
			"waterattuned", "earthattuned", "astralattuned", "deathattuned", "natureattuned", "bloodattuned", "magicboostF", "magicboostA", "magicboostW", 
			"magicboostE", "magicboostS", "magicboostD", "magicboostN", "magicboostALL", "eyes", "heatrec", "coldrec", "spreadchaos", "spreaddeath", 
			"corpseeater", "poisonskin", "bug", "uwbug", "spreadorder", "spreadgrowth", "startitem", "spreaddom", "reform", "battlesum5", "acidsplash", 
			"drake", "prophetshape", "horror", "enchrebate50", "heroarrivallimit", "sailsize", "uwdamage", "landdamage", "rpcost", "buffer", 
			"rand5", "nbr5", "link5", "mask5", "rand6", "nbr6", "link6", "mask6", "end"}; 
			
	private static String values[][] = {{"heal", "mounted", "animal", "amphibian", "wastesurvival", "undead", "coldres15", "heat", "neednoteat", "fireres15", "poisonres15", "aquatic", "flying", "trample", "immobile", "immortal" },
										{"cold", "forestsurvival", "shockres15", "swampsurvival", "demon", "holy", "mountainsurvival", "illusion", "noheal", "ethereal", "pooramphibian", "stealthy40", "misc2", "coldblood", "inanimate", "female" },
										{"bluntres", "slashres", "pierceres", "slow_to_recruit", "float", "", "teleport", "", "", "", "", "", "", "", "", "" },
										{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" },
										{"magicbeing", "", "", "poormagicleader", "okmagicleader", "goodmagicleader", "expertmagicleader", "superiormagicleader", "poorundeadleader", "okundeadleader", "goodundeadleader", "expertundeadleader", "superiorundeadleader", "", "", "" },
										{"", "", "", "", "", "", "", "", "noleader", "poorleader", "goodleader", "expertleader", "superiorleader", "", "", "" },
										{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" },
										{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" }};
	
	private static String[][] KNOWN_MONSTER_ATTRS = {
		{"6C00", "stealthy"},
		{"C900", "coldres"},
		{"DC00", "cold"},
		{"C600", "fireres"},
		{"3C01", "heat"},
		{"C800", "poisonres"},
		{"C700", "shockres"},
		{"6A00", "poisoncloud"},
		{"CD01", "patience"},
		{"AF00", "stormimmune"},
		{"BD00", "regeneration"},
		{"C200", "secondshape"},
		{"C300", "firstshape"},
		{"C100", "shapechange"},
		{"C400", "secondtmpshape"},
		{"1C01", "maxage"},
		{"CA00", "damagerev"},
		{"9E01", "bloodvengeance"},
		{"EB00", "nobadevents"},
		{"E400", "bringeroffortune"},
		{"1901", "darkvision"},
		{"B700", "fear"},
		{"1501", "voidsanity"},
		{"6700", "standard"},
		{"6E01", "formationfighter"},
		{"6F01", "undisciplined"},
		{"9801", "bodyguard"},
		{"A400", "summon"},
		{"A500", "summon"},
		{"A600", "summon"},
		{"7001", "inspirational"},
		{"8300", "pillagebonus"},
		{"BE00", "berserk"},
		{"F200", "startdom"},
		{"F300", "pathcost"},
		{"6F00", "waterbreathing"},
		{"5A02", "realm#"},
		{"B401", "batstartsum1"},
		{"B501", "batstartsum2"},
		{"B601", "batstartsum3"},
		{"B701", "batstartsum4"},
		{"B801", "batstartsum5"},
		{"B901", "batstartsum1d6"},
		{"BA01", "batstartsum2d6"},
		{"BB01", "batstartsum3d6"},
		{"BC01", "batstartsum4d6"},
		{"BD01", "batstartsum5d6"},
		{"BE01", "batstartsum6d6"},
		{"2501", "darkpower"},
		{"AE00", "stormpower"},
		{"B100", "firepower"},
		{"B000", "coldpower"},
		{"2501", "darkpower"},
		{"A001", "chaospower"},
		{"4401", "magicpower"},
		{"EA00", "winterpower"},
		{"E700", "springpower"},
		{"E800", "summerpower"},
		{"E900", "fallpower"},
		{"FB00", "nametype"},
		{"B600", "itemslots"},
		{"0901", "ressize"},
		{"1D01", "startage"},
		{"AB00", "blind"},
		{"B200", "eyes"},
		{"7A00", "supplybonus"},
		{"7C01", "slave"},
		{"7900", "researchbonus"},
		{"CA01", "chaosrec"},
		{"7D00", "siegebonus"},
		{"D900", "ambidextrous"},
		{"7E01", "invulnerable"},
		{"BF00", "iceprot"},
		{"7500", "reinvigoration"},
		{"D600", "spy"},
		{"E301", "scalewalls"},
		{"D500", "assassin"},
		{"D200", "succubus"},
		{"2A01", "seduce"},
		{"2901", "explodeondeath"},
		{"7B01", "taskmaster"},
		{"1301", "unique"},
		{"6900", "awe"},
		{"9D00", "additional leadership"},
		{"B900", "startaff"},
		{"F500", "landshape"},
		{"F600", "watershape"},
		{"4201", "forestshape"},
		{"4301", "plainshape"},
		{"DD00", "uwregen"},
		{"AA00", "patrolbonus"},
		{"D700", "castledef"},
		{"7000", "sailingshipsize"},
		{"9A01", "sailingmaxunitsize"},
		{"DF00", "incunrest"},
		{"BC00", "barbs"},
		{"4E01", "inn"},
		{"1801", "stonebeing"},
		{"5001", "shrinkhp"},
		{"4F01", "growhp"},
		{"FD01", "transformation"},
		{"A101", "domsummon#"},
		{"DB00", "domsummon#"},
		{"F100", "domsummon#"},
		{"6B00", "autosum"},
		{"8F00", "autosum"},
		{"AD00", "turmoilsummon"},
		{"9200", "coldsummon"},
		{"B800", "heretic"},
		{"2001", "popkill"},
		{"6201", "autohealer"},
		{"A300", "fireshield"},
		{"E200", "startingaff"},
		{"1E00", "gemprod fire"},
		{"1F00", "gemprod air"},
		{"2000", "gemprod water"},
		{"2100", "gemprod earth"},
		{"2200", "gemprod astral"},
		{"2300", "gemprod death"},
		{"2400", "gemprod nature"},
		{"2500", "gemprod blood"},
		{"F800", "fixedresearch"},
		{"C201", "divineins"},
		{"4701", "halt"},
		{"AF01", "crossbreeder"},
		{"7D01", "reclimit"},
		{"C501", "fixforgebonus"},
		{"6B01", "mastersmith"},
		{"A900", "lamiabonus"},
		{"FD00", "homesick"},
		{"DE00", "banefireshield"},
		{"A200", "animalawe"},
		{"6301", "autodishealer"},
		{"4801", "shatteredsoul"},
		{"CE00", "voidsum"},
		{"AE01", "makepearls"},
		{"5501", "inspiringres"},
		{"1101", "drainimmune"},
		{"AC00", "diseasecloud"},
		{"D300", "inquisitor"},
		{"7101", "beastmaster"},
		{"7400", "douse"},
		{"6C01", "preanimator"},
		{"6D01", "dreanimator"},
		{"FF01", "mummify"},
		{"D100", "onebattlespell"},
		{"F501", "fireattuned"},
		{"F601", "airattuned"},
		{"F701", "waterattuned"},
		{"F801", "earthattuned"},
		{"F901", "astralattuned"},
		{"FA01", "deathattuned"},
		{"FB01", "natureattuned"},
		{"FC01", "bloodattuned"},
		{"0A00", "magicboostF"},
		{"0B00", "magicboostA"},
		{"0C00", "magicboostW"},
		{"0D00", "magicboostE"},
		{"0E00", "magicboostS"},
		{"0F00", "magicboostD"},
		{"1000", "magicboostN"},
		{"1600", "magicboostALL"},
		{"EA01", "heatrec"},
		{"EB01", "coldrec"},
		{"D801", "spreadchaos/death"},
		{"D901", "spreadorder/growth"},
		{"EC00", "corpseeater"},
		{"0602", "poisonskin"},
		{"AB01", "bug"},
		{"AC01", "uwbug"},
		{"9C00", "spreaddom"},
		{"8A01", "reform"},
		{"8E01", "battlesum5"},
		{"F101", "acidsplash"},
		{"F201", "drake"},
		{"0402", "prophetshape"},
		{"8701", "horror"},
		{"3501", "insane"},
		{"D800", "sacr"},
		{"1002", "enchrebate50"},
		{"7C00", "leper"},
		{"9C01", "slimer"},
		{"9201", "mindslime"},
		{"9701", "resources"},
		{"E000", "corrupt"},
		{"6800", "petrify"},
		{"F400", "eyeloss"},
		{"0401", "ethtrue"},
		{"1602", "heroarrivallimit"},
		{"1702", "sailsize"},
		{"0201", "uwdamage"},
		{"0E02", "landdamage"}
	};

	
	private static class Magic {
		public int F;
		public int A;
		public int W;
		public int E;
		public int S;
		public int D;
		public int N;
		public int B;
		public int H;
		public List<RandomMagic> rand;
	}
	
	private static class RandomMagic {
		public int rand;
		public int nbr;
		public int link;
		public int mask;
	}
	
	private static String magicStrip(int mag) {
		return mag > 0 ? Integer.toString(mag) : "";
	}
	
	private static Map<Integer, Magic> monsterMagic = new HashMap<Integer, Magic>();

	public static void main(String[] args) {
		run();
	}
	
	public static void run() {
		FileInputStream stream = null;
        List<Monster> monsterList = new ArrayList<Monster>();

		try {
			BufferedWriter writer = new BufferedWriter(new FileWriter("monsters.txt"));
			BufferedWriter writerUnknown = new BufferedWriter(new FileWriter("monstersUnknown.txt"));
	        long startIndex = Starts.MONSTER;
	        int ch;
			stream = new FileInputStream(EXE_NAME);			
			stream.skip(Starts.MONSTER_MAGIC);
			Set<String> unknown = new HashSet<String>();

			// magic
			byte[] c = new byte[4];
			while ((stream.read(c, 0, 4)) != -1) {
				String high = String.format("%02X", c[1]);
				String low = String.format("%02X", c[0]);
				if ((high + low).equals("FFFF")) {
					break;
				}
				int id = Integer.decode("0X" + high + low);
				Magic monMagic = monsterMagic.get(id);
				if (monMagic == null) {
					monMagic = new Magic();
					monsterMagic.put(id, monMagic);
				}
				
				stream.read(c, 0, 4);
				high = String.format("%02X", c[1]);
				low = String.format("%02X", c[0]);
				int path = Integer.decode("0X" + high + low);
				stream.read(c, 0, 4);
				high = String.format("%02X", c[1]);
				low = String.format("%02X", c[0]);
				int value = Integer.decode("0X" + high + low);
				switch (path) {
				case 0:
					monMagic.F = value;
					break;
				case 1:
					monMagic.A = value;
					break;
				case 2:
					monMagic.W = value;
					break;
				case 3:
					monMagic.E = value;
					break;
				case 4:
					monMagic.S = value;
					break;
				case 5:
					monMagic.D = value;
					break;
				case 6:
					monMagic.N = value;
					break;
				case 7:
					monMagic.B = value;
					break;
				case 8:
					monMagic.H = value;
					break;
				default:
					RandomMagic monRandomMagic = null;
					List<RandomMagic> randomMagicList = monMagic.rand;
					if (randomMagicList == null) {
						randomMagicList = new ArrayList<RandomMagic>();
						monRandomMagic = new RandomMagic();
						if (path == 50) {
							monRandomMagic.mask = 32640;
							monRandomMagic.nbr = 1;
							monRandomMagic.link = value;
							monRandomMagic.rand = 100;
						} else if (path == 51) {
							monRandomMagic.mask = 1920;
							monRandomMagic.nbr = 1;
							monRandomMagic.link = value;
							monRandomMagic.rand = 100;
						} else {
							monRandomMagic.mask = path;
							monRandomMagic.nbr = 1;
							if (value > 100) {
								monRandomMagic.link = value / 100;
								monRandomMagic.rand = 100;
							} else {
								monRandomMagic.link = 1;
								monRandomMagic.rand = value;
							}
						}
						randomMagicList.add(monRandomMagic);
						monMagic.rand = randomMagicList;
					} else {
						boolean found = false;
						for (RandomMagic ranMagic : randomMagicList) {
							if (ranMagic.mask == path && ranMagic.rand == value) {
								ranMagic.nbr++;
								found = true;
							}
							if (ranMagic.mask == 32640 && path == 50 && ranMagic.link == value) {
								ranMagic.nbr++;
								found = true;
							}
							if (ranMagic.mask == 1920 && path == 51 && ranMagic.link == value) {
								ranMagic.nbr++;
								found = true;
							}
						}
						if (!found) {
							monRandomMagic = new RandomMagic();
							if (path == 50) {
								monRandomMagic.mask = 32640;
								monRandomMagic.nbr = 1;
								monRandomMagic.link = value;
								monRandomMagic.rand = 100;
							} else if (path == 51) {
								monRandomMagic.mask = 1920;
								monRandomMagic.nbr = 1;
								monRandomMagic.link = value;
								monRandomMagic.rand = 100;
							} else {
								monRandomMagic.mask = path;
								monRandomMagic.nbr = 1;
								if (value > 100) {
									monRandomMagic.link = value / 100;
									monRandomMagic.rand = 100;
								} else {
									monRandomMagic.link = 1;
									monRandomMagic.rand = value;
								}
							}
							randomMagicList.add(monRandomMagic);
						}
					}
				}
			}
			stream.close();
	        
			stream = new FileInputStream(EXE_NAME);			
			stream.skip(startIndex);
			
			// Name
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
				
				Monster monster = new Monster();
				monster.parameters = new TreeMap<String, Object>();
				monster.parameters.put("id", rowNumber);
				monster.parameters.put("name", name.toString());
				monster.parameters.put("ap", getBytes2(startIndex + 40));
				monster.parameters.put("mapmove", getBytes2(startIndex + 42));
				monster.parameters.put("size", getBytes2(startIndex + 44));
				monster.parameters.put("ressize", getBytes2(startIndex + 44));
				monster.parameters.put("hp", getBytes2(startIndex + 46));
				monster.parameters.put("prot", getBytes2(startIndex + 48));
				monster.parameters.put("str", getBytes2(startIndex + 50));
				monster.parameters.put("enc", getBytes2(startIndex + 52));
				monster.parameters.put("prec", getBytes2(startIndex + 54));
				monster.parameters.put("att", getBytes2(startIndex + 56));
				monster.parameters.put("def", getBytes2(startIndex + 58));
				monster.parameters.put("mr", getBytes2(startIndex + 60));
				monster.parameters.put("mor", getBytes2(startIndex + 62));
				monster.parameters.put("rcost", getBytes2(startIndex + 236));
				monster.parameters.put("wpn1", getBytes2(startIndex + 208) == 0 ? "" : getBytes2(startIndex + 208));
				monster.parameters.put("wpn2", getBytes2(startIndex + 210) == 0 ? "" : getBytes2(startIndex + 210));
				monster.parameters.put("wpn3", getBytes2(startIndex + 212) == 0 ? "" : getBytes2(startIndex + 212));
				monster.parameters.put("wpn4", getBytes2(startIndex + 214) == 0 ? "" : getBytes2(startIndex + 214));
				monster.parameters.put("wpn5", getBytes2(startIndex + 216) == 0 ? "" : getBytes2(startIndex + 216));
				monster.parameters.put("wpn6", getBytes2(startIndex + 218) == 0 ? "" : getBytes2(startIndex + 218));
				monster.parameters.put("wpn7", getBytes2(startIndex + 220) == 0 ? "" : getBytes2(startIndex + 220));
				monster.parameters.put("armor1", getBytes2(startIndex + 228) == 0 ? "" : getBytes2(startIndex + 228));
				monster.parameters.put("armor2", getBytes2(startIndex + 230) == 0 ? "" : getBytes2(startIndex + 230));
				monster.parameters.put("armor3", getBytes2(startIndex + 232) == 0 ? "" : getBytes2(startIndex + 232));
				monster.parameters.put("basecost", getBytes2(startIndex + 234));
				monster.parameters.put("rpcost", getBytes2(startIndex + 240));

				List<AttributeValue> attributes = getAttributes(startIndex + Starts.MONSTER_ATTRIBUTE_OFFSET, Starts.MONSTER_ATTRIBUTE_GAP);
				for (AttributeValue attr : attributes) {
					for (int x = 0; x < KNOWN_MONSTER_ATTRS.length; x++) {
						if (KNOWN_MONSTER_ATTRS[x][0].equals(attr.attribute)) {
							if (KNOWN_MONSTER_ATTRS[x][1].endsWith("#")) {
								int i = 1;
								for (String value : attr.values) {
									monster.parameters.put(KNOWN_MONSTER_ATTRS[x][1].replace("#", i+""), Integer.parseInt(value));
									i++;
								}
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("startage")) {
								int age = Integer.parseInt(attr.values.get(0));
								if (age == -1) {
									age = 0;
								}
								monster.parameters.put(KNOWN_MONSTER_ATTRS[x][1], Integer.toString((int)(age+age*.1)));
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("heatrec")) {
								int val = Integer.parseInt(attr.values.get(0));
								if (val == 10) {
									val = 0;
								} else {
									val = 1;
								}
								monster.parameters.put(KNOWN_MONSTER_ATTRS[x][1], Integer.toString(val));
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("coldrec")) {
								int val = Integer.parseInt(attr.values.get(0));
								if (val == 10) {
									val = 0;
								} else {
									val = 1;
								}
								monster.parameters.put(KNOWN_MONSTER_ATTRS[x][1], Integer.toString(val));
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("barbs")) {
	 							monster.parameters.put(KNOWN_MONSTER_ATTRS[x][1], "1");
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("corrupt")) {
	 							monster.parameters.put(KNOWN_MONSTER_ATTRS[x][1], "1");
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("ethtrue")) {
	 							monster.parameters.put(KNOWN_MONSTER_ATTRS[x][1], "1");
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("incunrest")) {
								double val = Double.parseDouble(attr.values.get(0))/10d;
								if (val < 1 && val > -1) {
									monster.parameters.put(KNOWN_MONSTER_ATTRS[x][1], Double.toString(val));
								} else {
									monster.parameters.put(KNOWN_MONSTER_ATTRS[x][1], Integer.toString((int)val));
								}
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("petrify")) {
	 							monster.parameters.put(KNOWN_MONSTER_ATTRS[x][1], "1");
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("eyes")) {
								int age = Integer.parseInt(attr.values.get(0));
								monster.parameters.put(KNOWN_MONSTER_ATTRS[x][1], Integer.toString(age+2));
							} else if (KNOWN_MONSTER_ATTRS[x][0].equals("A400")) {
	 							monster.parameters.put("n_summon", "1");
	 							monster.parameters.put(KNOWN_MONSTER_ATTRS[x][1], attr.values.get(0));
							} else if (KNOWN_MONSTER_ATTRS[x][0].equals("A500")) {
	 							monster.parameters.put("n_summon", "2");
	 							monster.parameters.put(KNOWN_MONSTER_ATTRS[x][1], attr.values.get(0));
							} else if (KNOWN_MONSTER_ATTRS[x][0].equals("A600")) {
	 							monster.parameters.put("n_summon", "3");
	 							monster.parameters.put(KNOWN_MONSTER_ATTRS[x][1], attr.values.get(0));
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("gemprod fire")) {
								String gemprod = "";
								if (monster.parameters.get("gemprod") != null) {
									gemprod = monster.parameters.get("gemprod").toString();
								}
								gemprod += attr.values.get(0) + "F";
	 							monster.parameters.put("gemprod", gemprod);
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("gemprod air")) {
								String gemprod = "";
								if (monster.parameters.get("gemprod") != null) {
									gemprod = monster.parameters.get("gemprod").toString();
								}
								gemprod += attr.values.get(0) + "A";
	 							monster.parameters.put("gemprod", gemprod);
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("gemprod water")) {
								String gemprod = "";
								if (monster.parameters.get("gemprod") != null) {
									gemprod = monster.parameters.get("gemprod").toString();
								}
								gemprod += attr.values.get(0) + "W";
	 							monster.parameters.put("gemprod", gemprod);
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("gemprod earth")) {
								String gemprod = "";
								if (monster.parameters.get("gemprod") != null) {
									gemprod = monster.parameters.get("gemprod").toString();
								}
								gemprod += attr.values.get(0) + "E";
	 							monster.parameters.put("gemprod", gemprod);
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("gemprod astral")) {
								String gemprod = "";
								if (monster.parameters.get("gemprod") != null) {
									gemprod = monster.parameters.get("gemprod").toString();
								}
								gemprod += attr.values.get(0) + "S";
	 							monster.parameters.put("gemprod", gemprod);
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("gemprod death")) {
								String gemprod = "";
								if (monster.parameters.get("gemprod") != null) {
									gemprod = monster.parameters.get("gemprod").toString();
								}
								gemprod += attr.values.get(0) + "D";
	 							monster.parameters.put("gemprod", gemprod);
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("gemprod nature")) {
								String gemprod = "";
								if (monster.parameters.get("gemprod") != null) {
									gemprod = monster.parameters.get("gemprod").toString();
								}
								gemprod += attr.values.get(0) + "N";
	 							monster.parameters.put("gemprod", gemprod);
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("gemprod blood")) {
								String gemprod = "";
								if (monster.parameters.get("gemprod") != null) {
									gemprod = monster.parameters.get("gemprod").toString();
								}
								gemprod += attr.values.get(0) + "B";
	 							monster.parameters.put("gemprod", gemprod);
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("pathcost")) {
								if (monster.parameters.get("startdom") == null) {
									monster.parameters.put("startdom", 1);
								}
	 							monster.parameters.put(KNOWN_MONSTER_ATTRS[x][1], attr.values.get(0));
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("spreadchaos/death")) {
								if (attr.values.get(0).equals("100")) {
									monster.parameters.put("spreadchaos", 1);
								} else if (attr.values.get(0).equals("103")) {
									monster.parameters.put("spreaddeath", 1);
								}
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("spreadorder/growth")) {
								if (attr.values.get(0).equals("100")) {
									monster.parameters.put("spreadorder", 1);
								} else if (attr.values.get(0).equals("103")) {
									monster.parameters.put("spreadgrowth", 1);
								}
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("damagerev")) {
								int val = Integer.parseInt(attr.values.get(0));
								monster.parameters.put(KNOWN_MONSTER_ATTRS[x][1], Integer.toString(val-1));
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("popkill")) {
								int val = Integer.parseInt(attr.values.get(0));
								monster.parameters.put(KNOWN_MONSTER_ATTRS[x][1], Integer.toString(val*10));
							} else {
	 							monster.parameters.put(KNOWN_MONSTER_ATTRS[x][1], attr.values.get(0));
							}
						} else {
 							monster.parameters.put("\tUnknown Attribute<" + attr.attribute + ">", attr.values.get(0));							
 							unknown.add(attr.attribute);
						}
					}
				}
				
				List<String> largeBitmap = largeBitmap(startIndex + 248, values);
				for (String bit : largeBitmap) {
					monster.parameters.put(bit, 1);
				}

				if (monster.parameters.containsKey("slow_to_recruit")) {
					monster.parameters.put("rt", "2");
				} else {
					monster.parameters.put("rt", "1");
				}
				
				if (largeBitmap.contains("heat")) {
					monster.parameters.put("heat", "3");
				}
				
				if (largeBitmap.contains("cold")) {
					monster.parameters.put("cold", "3");
				}
				
				String additionalLeader = "0";
				if (monster.parameters.get("additional leadership") != null) {
					additionalLeader = monster.parameters.get("additional leadership").toString();
				}
				monster.parameters.put("leader", 40+Integer.parseInt(additionalLeader));
				if (largeBitmap.contains("noleader")) {
					monster.parameters.put("baseleadership", 0);
					if (!"".equals(additionalLeader)) {
						monster.parameters.put("leader", additionalLeader);
					} else {
						monster.parameters.put("leader", 0);
					}
				}
				if (largeBitmap.contains("poorleader")) {
					monster.parameters.put("baseleadership", 10);
					if (!"".equals(additionalLeader)) {
						monster.parameters.put("leader", 10+Integer.parseInt(additionalLeader));
					} else {
						monster.parameters.put("leader", 10);
					}
				}
				if (largeBitmap.contains("goodleader")) {
					monster.parameters.put("baseleadership", 80);
					if (!"".equals(additionalLeader)) {
						monster.parameters.put("leader", 80+Integer.parseInt(additionalLeader));
					} else {
						monster.parameters.put("leader", 80);
					}
				}
				if (largeBitmap.contains("expertleader")) {
					monster.parameters.put("baseleadership", 120);
					if (!"".equals(additionalLeader)) {
						monster.parameters.put("leader", 120+Integer.parseInt(additionalLeader));
					} else {
						monster.parameters.put("leader", 120);
					}
				}
				if (largeBitmap.contains("superiorleader")) {
					monster.parameters.put("baseleadership", 160);
					if (!"".equals(additionalLeader)) {
						monster.parameters.put("leader", 160+Integer.parseInt(additionalLeader));
					} else {
						monster.parameters.put("leader", 160);
					}
				}
				if (largeBitmap.contains("poormagicleader")) {
					monster.parameters.put("magicleader", 10);
				}
				if (largeBitmap.contains("okmagicleader")) {
					monster.parameters.put("magicleader", 40);
				}
				if (largeBitmap.contains("goodmagicleader")) {
					monster.parameters.put("magicleader", 80);
				}
				if (largeBitmap.contains("expertmagicleader")) {
					monster.parameters.put("magicleader", 120);
				}
				if (largeBitmap.contains("superiormagicleader")) {
					monster.parameters.put("magicleader", 160);
				}
				if (largeBitmap.contains("poorundeadleader")) {
					monster.parameters.put("undeadleader", 10);
				}
				if (largeBitmap.contains("okundeadleader")) {
					monster.parameters.put("undeadleader", 40);
				}
				if (largeBitmap.contains("goodundeadleader")) {
					monster.parameters.put("undeadleader", 80);
				}
				if (largeBitmap.contains("expertundeadleader")) {
					monster.parameters.put("undeadleader", 120);
				}
				if (largeBitmap.contains("superiorundeadleader")) {
					monster.parameters.put("undeadleader", 160);
				}
				
				if (largeBitmap.contains("coldres15")) {
					String additionalCold = "0";
					if (monster.parameters.get("coldres") != null) {
						additionalCold = monster.parameters.get("coldres").toString();
					}
					boolean cold = false;
					if (largeBitmap.contains("cold") || monster.parameters.get("cold") != null) {
						cold = true;
					}
					monster.parameters.put("coldres", 15 + Integer.parseInt(additionalCold.equals("")?"0":additionalCold) + (cold?10:0));
				} else {
					String additionalCold = "0";
					if (monster.parameters.get("coldres") != null) {
						additionalCold = monster.parameters.get("coldres").toString();
					}
					boolean cold = false;
					if (largeBitmap.contains("cold") || monster.parameters.get("cold") != null) {
						cold = true;
					}
					int coldres = Integer.parseInt(additionalCold.equals("")?"0":additionalCold) + (cold?10:0);
					monster.parameters.put("coldres", coldres==0?"":Integer.toString(coldres));
				}
				if (largeBitmap.contains("fireres15")) {
					String additionalFire = "0";
					if (monster.parameters.get("fireres") != null) {
						additionalFire = monster.parameters.get("fireres").toString();
					}
					boolean heat = false;
					if (largeBitmap.contains("heat") || monster.parameters.get("heat") != null) {
						heat = true;
					}
					monster.parameters.put("fireres", 15 + Integer.parseInt(additionalFire.equals("")?"0":additionalFire) + (heat?10:0));
				} else {
					String additionalFire = "0";
					if (monster.parameters.get("fireres") != null) {
						additionalFire = monster.parameters.get("fireres").toString();
					}
					boolean heat = false;
					if (largeBitmap.contains("heat") || monster.parameters.get("heat") != null) {
						heat = true;
					}
					int fireres = Integer.parseInt(additionalFire.equals("")?"0":additionalFire) + (heat?10:0);
					monster.parameters.put("fireres", fireres==0?"":Integer.toString(fireres));
				}
				if (largeBitmap.contains("poisonres15")) {
					String additionalPoisin = "0";
					if (monster.parameters.get("poisonres") != null) {
						additionalPoisin = monster.parameters.get("poisonres").toString();
					}
					boolean poisoncloud = false;
					if (largeBitmap.contains("undead")|| largeBitmap.contains("inanimate") || monster.parameters.get("poisoncloud") != null) {
						poisoncloud = true;
					}
					monster.parameters.put("poisonres", 15 + Integer.parseInt(additionalPoisin.equals("")?"0":additionalPoisin) + (poisoncloud?10:0));
				} else {
					String additionalPoison = "0";
					if (monster.parameters.get("poisonres") != null) {
						additionalPoison = monster.parameters.get("poisonres").toString();
					}
					boolean poisoncloud = false;
					if (largeBitmap.contains("undead")|| largeBitmap.contains("inanimate") || monster.parameters.get("poisoncloud") != null) {
						poisoncloud = true;
					}
					int poisonres = Integer.parseInt(additionalPoison.equals("")?"0":additionalPoison) + (poisoncloud?10:0);
					monster.parameters.put("poisonres", poisonres==0?"":Integer.toString(poisonres));
				}
				if (largeBitmap.contains("shockres15")) {
					String additionalShock = "0";
					if (monster.parameters.get("shockres") != null) {
						additionalShock = monster.parameters.get("shockres").toString();
					}
					monster.parameters.put("shockres", 15 + Integer.parseInt(additionalShock.equals("")?"0":additionalShock));
				} else {
					String additionalShock = "0";
					if (monster.parameters.get("shockres") != null) {
						additionalShock = monster.parameters.get("shockres").toString();
					}
					int shockres = Integer.parseInt(additionalShock.equals("")?"0":additionalShock);
					monster.parameters.put("shockres", shockres==0?"":Integer.toString(shockres));
				}
				if (largeBitmap.contains("stealthy40")) {
					String additionalStealth = "0";
					if (monster.parameters.get("stealthy") != null) {
						additionalStealth = monster.parameters.get("stealthy").toString();
					}
					boolean glamour = false;
					if (monster.parameters.get("illusion") != null) {
						glamour = true;
					}
					monster.parameters.put("stealthy", 40 + Integer.parseInt(additionalStealth.equals("")?"0":additionalStealth) + (glamour?25:0));
				} else {
					String additionalStealth = "0";
					if (monster.parameters.get("stealthy") != null) {
						additionalStealth = monster.parameters.get("stealthy").toString();
					}
					monster.parameters.put("stealthy", additionalStealth == null || additionalStealth.equals("0") ? "" : additionalStealth);
				}
				
				// magic
				Magic monMagic = monsterMagic.get(rowNumber);
				if (monMagic != null) {
					monster.parameters.put("F", magicStrip(monMagic.F));
					monster.parameters.put("A", magicStrip(monMagic.A));
					monster.parameters.put("W", magicStrip(monMagic.W));
					monster.parameters.put("E", magicStrip(monMagic.E));
					monster.parameters.put("S", magicStrip(monMagic.S));
					monster.parameters.put("D", magicStrip(monMagic.D));
					monster.parameters.put("N", magicStrip(monMagic.N));
					monster.parameters.put("B", magicStrip(monMagic.B));
					monster.parameters.put("H", magicStrip(monMagic.H));
					
					if (monMagic.rand != null) {
						int count = 1;
						for (RandomMagic ranMag : monMagic.rand) {
							monster.parameters.put("rand"+count, magicStrip(ranMag.rand));
							monster.parameters.put("nbr"+count, magicStrip(ranMag.nbr));
							monster.parameters.put("link"+count, magicStrip(ranMag.link));
							monster.parameters.put("mask"+count, magicStrip(ranMag.mask));

							count++;
						}
					}
				}

				if (largeBitmap.contains("misc2")) {
					monster.parameters.put("hand", 0);
					monster.parameters.put("head", 0);
					monster.parameters.put("body", 0);
					monster.parameters.put("foot", 0);
					monster.parameters.put("misc", 2);
				}
				if (monster.parameters.containsKey("itemslots")) {
					String slots = monster.parameters.get("itemslots").toString();
					int numHands = 0;
					int numHeads = 0;
					int numBody = 0;
					int numFoot = 0;
					int numMisc = 0;
					long val = Long.parseLong(slots);
					if ((val & 0x0002) != 0) {
						numHands++;
					}
					if ((val & 0x0004) != 0) {
						numHands++;
					}
					if ((val & 0x0008) != 0) {
						numHands++;
					}
					if ((val & 0x0010) != 0) {
						numHands++;
					}
					if ((val & 0x0080) != 0) {
						numHeads++;
					}
					if ((val & 0x0100) != 0) {
						numHeads++;
					}
					if ((val & 0x0200) != 0) {
						numHeads++;
					}
					if ((val & 0x0400) != 0) {
						numBody++;
					}
					if ((val & 0x0800) != 0) {
						numFoot++;
					}
					if ((val & 0x1000) != 0) {
						numMisc++;
					}
					if ((val & 0x2000) != 0) {
						numMisc++;
					}
					if ((val & 0x4000) != 0) {
						numMisc++;
					}
					if ((val & 0x8000) != 0) {
						numMisc++;
					}
					if ((val & 0x10000) != 0) {
						numMisc++;
					}
					monster.parameters.put("hand", numHands);
					monster.parameters.put("head", numHeads);
					monster.parameters.put("body", numBody);
					monster.parameters.put("foot", numFoot);
					monster.parameters.put("misc", numMisc);
				} else if (!largeBitmap.contains("misc2")) {
					monster.parameters.put("hand", 2);
					monster.parameters.put("head", 1);
					monster.parameters.put("body", 1);
					monster.parameters.put("foot", 1);
					monster.parameters.put("misc", 2);
				}
				if (largeBitmap.contains("mounted")) {
					monster.parameters.put("foot", 0);
				}

				monsterList.add(monster);

				stream = new FileInputStream(EXE_NAME);		
				startIndex = startIndex + Starts.MONSTER_SIZE;
				stream.skip(startIndex);
				isr = new InputStreamReader(stream, "ISO-8859-1");
		        in = new BufferedReader(isr);

				rowNumber++;
			}
			in.close();
			stream.close();

			XSSFWorkbook wb = new XSSFWorkbook();
			FileOutputStream fos = new FileOutputStream("BaseU.xlsx");
			XSSFSheet sheet = wb.createSheet();

			int rowNum = 0;
			for (Monster monster : monsterList) {
				if (rowNum == 0) {
					XSSFRow row = sheet.createRow(rowNum);
					for (int i = 0; i < unit_columns.length; i++) {
						row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(unit_columns[i]);
					}
					rowNum++;
				}
				XSSFRow row = sheet.createRow(rowNum);
				for (int i = 0; i < unit_columns.length; i++) {
					Object object = monster.parameters.get(unit_columns[i]);
					if (object != null) {
						row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(object.toString());
					}
				}
				
				dumpTextFile(monster, writer);

				rowNum++;
			}
			
			dumpUnknown(unknown, writerUnknown);
			
			writer.close();
			writerUnknown.close();

			wb.write(fos);
			fos.close();
			wb.close();

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
	
	private static void dumpTextFile(Monster monster, BufferedWriter writer) throws IOException {
		Object name = monster.parameters.get("name");
		Object id = monster.parameters.get("id");
		writer.write(name.toString() + "(" + id + ")");
		writer.newLine();
		for (Map.Entry<String, Object> entry : monster.parameters.entrySet()) {
			if (!entry.getKey().equals("name") && !entry.getKey().equals("id") && entry.getValue() != null && !entry.getValue().equals("")) {
				writer.write("\t" + entry.getKey() + ": " + entry.getValue());
				writer.newLine();
			}
		}
		writer.newLine();
	}

	private static class Monster {
		Map<String, Object> parameters;
	}

}
