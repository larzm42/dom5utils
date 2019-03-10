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
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
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
import java.util.TreeSet;

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
			"corrupt", "heal", "immortal", "domimmortal", "reinc", "noheal", "neednoteat", "homesick", "undisciplined", "formationfighter", "slave", "standard", 
			"inspirational", "taskmaster", "beastmaster", "bodyguard", "waterbreathing", "iceprot", "invulnerable", "slashres", "bluntres", "pierceres", 
			"shockres", "fireres", "coldres", "poisonres", "voidsanity", "darkvision", "blind", "animalawe", "awe", "halt", "fear", "berserk", "cold", 
			"heat", "fireshield", "banefireshield", "damagerev", "poisoncloud", "diseasecloud", "slimer", "mindslime", "disbel", "reform", "regeneration", 
			"reanimator", "barbs", "petrify", "eyeloss", "ethereal", "ethtrue", "deathcurse", "trample", "trampswallow", "stormpower", "firepower", 
			"coldpower", "darkpower", "chaospower", "magicpower", "winterpower", "springpower", "summerpower", "fallpower", "forgebonus", "fixforgebonus", 
			"mastersmith", "resources", "autohealer", "autodishealer", "alch", "nobadevents", "event", "insane", "shatteredsoul", "leper", "taxcollector", 
			"gem", "gemprod", "chaosrec", "pillagebonus", "patrolbonus", "castledef", "siegebonus", "incprovdef", "supplybonus", "incunrest", "popkill", 
			"researchbonus", "drainimmune", "inspiringres", "douse", "sacr", "crossbreeder", "makepearls", "pathboost", "allrange", "voidsum", 
			"heretic", "elegist", "shapechange", "firstshape", "secondshape", "secondtmpshape", "landshape", "watershape", "forestshape", "plainshape", "xpshape",
			"unique", "fixedname", "special", "nametype", "summon", "n_summon", "autosum", "n_autosum", "batstartsum1", 
			"batstartsum2", "domsummon1", "domsummon2", "bloodvengeance", "bringeroffortune", "realm1", "realm2", "realm3", "batstartsum3", "batstartsum4", 
			"batstartsum5", "batstartsum1d6", "batstartsum2d6", "batstartsum3d6", "batstartsum4d6", "batstartsum5d6", "batstartsum6d6", "turmoilsummon", 
			"coldsummon", "scalewalls", "explodeondeath", "startaff", "uwregen", "shrinkhp", "growhp", "transformation", "startingaff", 
			"fixedresearch", "divineins", "lamiabonus", "preanimator", "dreanimator", "mummify", "onebattlespell", "fireattuned", "airattuned", 
			"waterattuned", "earthattuned", "astralattuned", "deathattuned", "natureattuned", "bloodattuned", "magicboostF", "magicboostA", "magicboostW", 
			"magicboostE", "magicboostS", "magicboostD", "magicboostN", "magicboostALL", "eyes", "heatrec", "coldrec", "spreadchaos", "spreaddeath", 
			"corpseeater", "poisonskin", "bug", "uwbug", "spreadorder", "spreadgrowth", "startitem", "spreaddom", "battlesum5", "acidsplash", 
			"drake", "prophetshape", "horror", "enchrebate50", "heroarrivallimit", "sailsize", "uwdamage", "landdamage", "rpcost", "buffer", 
			"rand5", "nbr5", "link5", "mask5", "rand6", "nbr6", "link6", "mask6", 
			"mummification", "disres", "raiseonkill", "raiseshape", "sendlesserhorrormult", "xploss", "theftofthesunawe", "incorporate", "hpoverslow", "berserkwhenblessed",
			"dragonmastery", "curseattacker", "uwheataura", "slothresearch", "horrorsonly", "mindvessel", "elementrange", "sorceryrange", "startagemodifier",
			"disbelieveillusions", "firerange", "astralrange", "landreinvigoration", "naturerange", "beartattoo", "horsetattoo", "reincarnation", "wolftattoo", "boartattoo",
			"sleepaura", "snaketattoo", "supplysize", "astralfetters", "foreignmagicboost", "templesummon", "infernoret", "kokytosret", "infernalcrossbreedingmult", "unsurroundable",
			"combatcaster", "homeshape", "speciallook", "aisinglerec", "nowish", "swarmbody", "mason", "onisummoner", "sunawe", "spiritsight", "defenceorganiser",
			"invisibility", "startaffliction", "ivylord", "spellsinging", "magicallyattunedresearcher", "triplegod", "triplegodmag", "unify", "triple3mon",
			"poweroftheturningyear", "fortkill", "thronekill", "digest", "indepmove", "unteleportable", "reanimpriest", "stunimmunity",
			"vineshield", "alchemy", "afflictionresistance", "leavespostbattleifwoundedorhaskilled", "makesarmylooksmallerorlarger",
			"summon5", "ainorecruit", "autocomslave", "researchwithoutmagic", "captureslaves", "mustfightinarena", "deathwail", "adventurers", "cleanshape", "requireslabtorecruit",
			"requirestempletorecruit", "horrormarked", "changetargetgenderforseductionandseductionimmune", "corpseconstruct", "guardianspiritmodifier", "isashah", "iceforging",
			"isayazad", "isadaeva", "flieswhenblessed", "plant", "clockworklord", "commaster", "comslave", "minsizeleader", "snowmove", "swimming", "stupid",
			"end"}; 
			
	private static String values[][] = {{"heal", "mounted", "animal", "amphibian", "wastesurvival", "undead", "coldres15", "heat", "neednoteat", "fireres15", "poisonres15", "aquatic", "flying", "trample", "immobile", "immortal" },
										{"cold", "forestsurvival", "shockres15", "swampsurvival", "demon", "holy", "mountainsurvival", "illusion", "noheal", "ethereal", "pooramphibian", "stealthy40", "misc2", "coldblood", "inanimate", "female" },
										{"bluntres", "slashres", "pierceres", "slow_to_recruit", "float", "stormimmune", "teleport", "snowmove", "swimming", "domimmortal", "", "", "", "", "", "" },
										{"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" },
										{"magicbeing", "", "darkvision100", "poormagicleader", "okmagicleader", "goodmagicleader", "expertmagicleader", "superiormagicleader", "poorundeadleader", "okundeadleader", "goodundeadleader", "expertundeadleader", "superiorundeadleader", "", "", "" },
										{"", "", "", "", "stupid", "", "", "", "noleader", "poorleader", "goodleader", "expertleader", "superiorleader", "", "", "" },
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
		{"FE01", "xpshape"},
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
		//{"A101", "domsummon#"},
		//{"DB00", "domsummon#"},
		//{"F100", "domsummon#"},
		{"6B00", "autosum"},
		{"8F00", "autosum"},
		{"AD00", "turmoilsummon"},
		{"9200", "coldsummon"},
		{"B800", "heretic"},
		{"2001", "popkill"},
		{"6201", "autohealer"},
		{"8C02", "fireshield"},
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
		{"0E02", "landdamage"},
		
		{"0002", "mummification"},
		{"0102", "disres"},
		{"0202", "raiseonkill"},
		{"0302", "raiseshape"},
		{"0502", "sendlesserhorrormult"},
		{"0702", "xploss"},
		{"0802", "theftofthesunawe"},
		{"0A02", "incorporate"},
		{"0B02", "hpoverslow"},
		{"0D02", "berserkwhenblessed"},
		{"0E01", "dragonmastery"},
		{"0F01", "curseattacker"},
		{"0F02", "uwheataura"},
		{"1001", "slothresearch"},
		{"1201", "horrorsonly"},
		{"1302", "mindvessel"},
		{"1700", "elementrange"},
		{"1800", "sorceryrange"},
		{"1E01", "startagemodifier"},
		{"2601", "illusion"},
		{"2701", "disbelieveillusions"},
		{"2800", "firerange"},
		{"2C00", "astralrange"},
		{"2C01", "landreinvigoration"},
		{"2E00", "naturerange"},
		{"2E02", "beartattoo"},
		{"2F02", "horsetattoo"},
		{"3001", "reincarnation"},
		{"3002", "wolftattoo"},
		{"3102", "boartattoo"},
		{"3201", "sleepaura"},
		{"3202", "snaketattoo"},
		{"3301", "supplysize"},
		{"3302", "astralfetters"},
		{"3400", "foreignmagicboost"},
		{"3402", "templesummon"},
		{"3901", "infernoret"},
		{"3A01", "kokytosret"},
		{"3D01", "infernalcrossbreedingmult"},
		{"3D02", "unsurroundable"},
		{"4502", "combatcaster"},
		{"4902", "homeshape"},
		{"4A01", "speciallook"},
		{"4B01", "aisinglerec"},
		{"4C02", "nowish"},
		{"5201", "swarmbody"},
		{"5202", "mason"},
		{"5601", "taxcollector"},
		{"5602", "onisummoner"},
		{"5D02", "sunawe"},
		{"5F02", "spiritsight"},
		{"6001", "defenceorganiser"},
		{"6002", "invisibility"},
		{"6102", "startaffliction"},
		{"6500", "ivylord"},
		//{"6601", "formationfighter"},
		{"6802", "spellsinging"},
		{"6A02", "magicallyattunedresearcher"},
		{"6E02", "triplegod"},
		{"6F02", "triplegodmag"},
		{"7002", "unify"},
		{"7102", "triple3mon"},
		{"7202", "poweroftheturningyear"},
		{"7302", "fortkill"},
		{"7402", "thronekill"},
		{"7501", "trampswallow"},
		{"7600", "forgebonus"},
		{"7601", "digest"},
		{"7602", "indepmove"},
		{"7801", "unteleportable"},
		{"7802", "reanimpriest"},
		{"7E02", "stunimmunity"},
		{"8000", "vineshield"},
		{"8400", "alchemy"},
		//{"8B00", "domsummon2"},
		{"9601", "afflictionresistance"},
		{"9901", "leavespostbattleifwoundedorhaskilled"},
		{"9B01", "elegist"},
		{"9D01", "deathcurse"},
		{"A201", "noriverpass"},
		{"A301", "makesarmylooksmallerorlarger"},
		{"A800", "summon5"},
		{"AD01", "ainorecruit"},
		{"B001", "autocomslave"},
		{"C301", "researchwithoutmagic"},
		{"C500", "captureslaves"},
		{"CB01", "mustfightinarena"},
		{"CC00", "deathwail"},
		{"CC01", "adventurers"},
		{"D101", "cleanshape"},
		{"D501", "requireslabtorecruit"},
		{"D601", "requirestempletorecruit"},
		{"D701", "horrormarked"},
		{"DA00", "changetargetgenderforseductionandseductionimmune"},
		{"DE01", "corpseconstruct"},
		{"E601", "guardianspiritmodifier"},
		{"E801", "isashah"},
		{"E901", "iceforging"},
		{"EC01", "isayazad"},
		{"ED00", "corpseeater"},
		{"ED01", "isadaeva"},
		{"F000", "flieswhenblessed"},
		{"F401", "plant"},
		{"F900", "clockworklord"},
		{"8802", "commaster"},
		{"8402", "minsizeleader"},
		{"8702", "comslave"}		

	};
	
//	0701	Feeblemind chance in province? Kurgi Only
//	1202	Aboleth - Can Cast Mind Vessel?
//	1A01	? Angels and Celestial Beings only
//	1B01	? Rudra & Devata only
//	3300	land all magic penalty? Kaijin Only
//	3602	Startitem - looks to refer to item attribute 3702
//	3B01	? Void Spectre Only (Voidret/Insanity?)
//	3F01	? Delgnat Only
//	4001	? Defiler of Dreams only
//	4501	? Leshy Plainshape only
//	4702	? Chariots only
//	5801	? Buer Only. Spread Heat?
//	5802	Minprison? Uttervast Only
//	5902	? Scabiel Only
//	5C01	? Kurgi Only
//	6701	? Barbarians & Bakemono-Sho
//	6801	? Knight & Barbarian commanders
//	7401	Spritesize? Elementals and Draugr
//	7502	FarThroneKill? Doom Horrors only
//	7701	Aciddigest? Eater of the Dead & Ancient Presence only
//	7702	Indepstay? Maker of Ruins & Eater of Gods only
//	7A02	? Eater of Dreams only
//	7B02	Polymorph Immunity? Doom Horrors & God Vessels only
//	7F00	? Eater of Gods only
//	8401	? Horrors and God Vessels only
//	8501	? Slave to Unreason only
//	8801	Resummon as this shape? Eater of the Dead only
//	8A00	? Cockatrice only
//	9101	Turns before reform? Worm mage only
//	9300	Immortality delay modifier? Vampires and Wraith Lords only
//	9500	? Vampires only
//	9F01	? Soultorn has this at 10
//	B301	?Doom Horrors have this?
//	BB00	?Spiders & ghost cats have this. Shapechange related?
//	D201	?God Vessel has this?
//	D301	?God vessel & legion of gods? (% of moving around?)
//	D401	? God vessel & Legion of Gods
//	E401	? Unused unit 10 only
//	E501	? Unused unit 10 only
//	F001	?Acid cube has this?
//	F700	? Siren landshape only (Lure?)

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
			Set<String> unknown = new TreeSet<String>();

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
				monster.setId(rowNumber);
				monster.setName(name.toString());
				monster.setAp(getBytes2(startIndex + 40));
				monster.setMapmove(getBytes2(startIndex + 42));
				monster.setSize(getBytes2(startIndex + 44));
				monster.setRessize(getBytes2(startIndex + 44));
				monster.setHp(getBytes2(startIndex + 46));
				monster.setProt(getBytes2(startIndex + 48));
				monster.setStr(getBytes2(startIndex + 50));
				monster.setEnc(getBytes2(startIndex + 52));
				monster.setPrec(getBytes2(startIndex + 54));
				monster.setAtt(getBytes2(startIndex + 56));
				monster.setDef(getBytes2(startIndex + 58));
				monster.setMr(getBytes2(startIndex + 60));
				monster.setMor(getBytes2(startIndex + 62));
				monster.setWpn1(getBytes2(startIndex + 208) == 0 ? null : getBytes2(startIndex + 208));
				monster.setWpn2(getBytes2(startIndex + 210) == 0 ? null : getBytes2(startIndex + 210));
				monster.setWpn3(getBytes2(startIndex + 212) == 0 ? null : getBytes2(startIndex + 212));
				monster.setWpn4(getBytes2(startIndex + 214) == 0 ? null : getBytes2(startIndex + 214));
				monster.setWpn5(getBytes2(startIndex + 216) == 0 ? null : getBytes2(startIndex + 216));
				monster.setWpn6(getBytes2(startIndex + 218) == 0 ? null : getBytes2(startIndex + 218));
				monster.setWpn7(getBytes2(startIndex + 220) == 0 ? null : getBytes2(startIndex + 220));
				monster.setArmor1(getBytes2(startIndex + 228) == 0 ? null : getBytes2(startIndex + 228));
				monster.setArmor2(getBytes2(startIndex + 230) == 0 ? null : getBytes2(startIndex + 230));
				monster.setArmor3(getBytes2(startIndex + 232) == 0 ? null : getBytes2(startIndex + 232));
				monster.setBasecost(getBytes2(startIndex + 234));
				monster.setRcost(getBytes2(startIndex + 236));
				monster.setRpcost(getBytes2(startIndex + 240));

				List<AttributeValue> attributes = getAttributes(startIndex + Starts.MONSTER_ATTRIBUTE_OFFSET, Starts.MONSTER_ATTRIBUTE_GAP);
				for (AttributeValue attr : attributes) {
					boolean found = false;
					for (int x = 0; x < KNOWN_MONSTER_ATTRS.length; x++) {
						if (KNOWN_MONSTER_ATTRS[x][0].equals(attr.attribute)) {
							found = true;
							if (KNOWN_MONSTER_ATTRS[x][1].endsWith("#")) {
								int i = 1;
								for (String value : attr.values) {
									monster.addAttribute(new Attr(KNOWN_MONSTER_ATTRS[x][1].replace("#", i+""), Integer.parseInt(value)));
									i++;
								}
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("startage")) {
								int age = Integer.parseInt(attr.values.get(0));
								if (age == -1) {
									age = 0;
								}
								monster.addAttribute(new Attr(KNOWN_MONSTER_ATTRS[x][1], Integer.toString((int)(age+age*.1))));
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("heatrec")) {
								int val = Integer.parseInt(attr.values.get(0));
								if (val == 10) {
									val = 0;
								} else {
									val = 1;
								}
								monster.addAttribute(new Attr(KNOWN_MONSTER_ATTRS[x][1], Integer.toString(val)));
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("coldrec")) {
								int val = Integer.parseInt(attr.values.get(0));
								if (val == 10) {
									val = 0;
								} else {
									val = 1;
								}
								monster.addAttribute(new Attr(KNOWN_MONSTER_ATTRS[x][1], Integer.toString(val)));
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("barbs")) {
	 							monster.addAttribute(new Attr(KNOWN_MONSTER_ATTRS[x][1], "1"));
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("corrupt")) {
	 							monster.addAttribute(new Attr(KNOWN_MONSTER_ATTRS[x][1], "1"));
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("ethtrue")) {
	 							monster.addAttribute(new Attr(KNOWN_MONSTER_ATTRS[x][1], "1"));
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("incunrest")) {
								double val = Double.parseDouble(attr.values.get(0))/10d;
								if (val < 1 && val > -1) {
									monster.addAttribute(new Attr(KNOWN_MONSTER_ATTRS[x][1], Double.toString(val)));
								} else {
									monster.addAttribute(new Attr(KNOWN_MONSTER_ATTRS[x][1], Integer.toString((int)val)));
								}
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("petrify")) {
	 							monster.addAttribute(new Attr(KNOWN_MONSTER_ATTRS[x][1], "1"));
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("eyes")) {
								int age = Integer.parseInt(attr.values.get(0));
								monster.addAttribute(new Attr(KNOWN_MONSTER_ATTRS[x][1], Integer.toString(age+2)));
							} else if (KNOWN_MONSTER_ATTRS[x][0].equals("A400")) {
	 							monster.addAttribute(new Attr("n_summon", "1"));
	 							monster.addAttribute(new Attr(KNOWN_MONSTER_ATTRS[x][1], attr.values.get(0)));
							} else if (KNOWN_MONSTER_ATTRS[x][0].equals("A500")) {
	 							monster.addAttribute(new Attr("n_summon", "2"));
	 							monster.addAttribute(new Attr(KNOWN_MONSTER_ATTRS[x][1], attr.values.get(0)));
							} else if (KNOWN_MONSTER_ATTRS[x][0].equals("A600")) {
	 							monster.addAttribute(new Attr("n_summon", "3"));
	 							monster.addAttribute(new Attr(KNOWN_MONSTER_ATTRS[x][1], attr.values.get(0)));
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("gemprod fire")) {
								String gemprod = "";
								if (monster.getAttribute("gemprod") != null) {
									gemprod = monster.getAttribute("gemprod").toString();
								}
								gemprod += attr.values.get(0) + "F";
	 							monster.addAttribute(new Attr("gemprod", gemprod));
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("gemprod air")) {
								String gemprod = "";
								if (monster.getAttribute("gemprod") != null) {
									gemprod = monster.getAttribute("gemprod").toString();
								}
								gemprod += attr.values.get(0) + "A";
	 							monster.addAttribute(new Attr("gemprod", gemprod));
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("gemprod water")) {
								String gemprod = "";
								if (monster.getAttribute("gemprod") != null) {
									gemprod = monster.getAttribute("gemprod").toString();
								}
								gemprod += attr.values.get(0) + "W";
	 							monster.addAttribute(new Attr("gemprod", gemprod));
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("gemprod earth")) {
								String gemprod = "";
								if (monster.getAttribute("gemprod") != null) {
									gemprod = monster.getAttribute("gemprod").toString();
								}
								gemprod += attr.values.get(0) + "E";
	 							monster.addAttribute(new Attr("gemprod", gemprod));
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("gemprod astral")) {
								String gemprod = "";
								if (monster.getAttribute("gemprod") != null) {
									gemprod = monster.getAttribute("gemprod").toString();
								}
								gemprod += attr.values.get(0) + "S";
	 							monster.addAttribute(new Attr("gemprod", gemprod));
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("gemprod death")) {
								String gemprod = "";
								if (monster.getAttribute("gemprod") != null) {
									gemprod = monster.getAttribute("gemprod").toString();
								}
								gemprod += attr.values.get(0) + "D";
	 							monster.addAttribute(new Attr("gemprod", gemprod));
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("gemprod nature")) {
								String gemprod = "";
								if (monster.getAttribute("gemprod") != null) {
									gemprod = monster.getAttribute("gemprod").toString();
								}
								gemprod += attr.values.get(0) + "N";
	 							monster.addAttribute(new Attr("gemprod", gemprod));
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("gemprod blood")) {
								String gemprod = "";
								if (monster.getAttribute("gemprod") != null) {
									gemprod = monster.getAttribute("gemprod").toString();
								}
								gemprod += attr.values.get(0) + "B";
	 							monster.addAttribute(new Attr("gemprod", gemprod));
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("pathcost")) {
								if (monster.getAttribute("startdom") == null) {
									monster.addAttribute(new Attr("startdom", 1));
								}
	 							monster.addAttribute(new Attr(KNOWN_MONSTER_ATTRS[x][1], attr.values.get(0)));
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("spreadchaos/death")) {
								if (attr.values.get(0).equals("100")) {
									monster.addAttribute(new Attr("spreadchaos", 1));
								} else if (attr.values.get(0).equals("103")) {
									monster.addAttribute(new Attr("spreaddeath", 1));
								}
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("spreadorder/growth")) {
								if (attr.values.get(0).equals("100")) {
									monster.addAttribute(new Attr("spreadorder", 1));
								} else if (attr.values.get(0).equals("103")) {
									monster.addAttribute(new Attr("spreadgrowth", 1));
								}
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("damagerev")) {
								int val = Integer.parseInt(attr.values.get(0));
								monster.addAttribute(new Attr(KNOWN_MONSTER_ATTRS[x][1], Integer.toString(val-1)));
							} else if (KNOWN_MONSTER_ATTRS[x][1].equals("popkill")) {
								int val = Integer.parseInt(attr.values.get(0));
								monster.addAttribute(new Attr(KNOWN_MONSTER_ATTRS[x][1], Integer.toString(val*10)));
							} else {
	 							monster.addAttribute(new Attr(KNOWN_MONSTER_ATTRS[x][1], attr.values.get(0)));
							}
						}
					}
					if (!found) {
						monster.addAttribute(new Attr("\tUnknown Attribute<" + attr.attribute + ">", attr.values.get(0)));							
						unknown.add(attr.attribute);
					}
				}
				
				List<String> largeBitmap = largeBitmap(startIndex + 248, values);
				for (String bit : largeBitmap) {
					monster.addAttribute(new Attr(bit, 1));
				}

				if (monster.getAttribute("slow_to_recruit") != null) {
					monster.addAttribute(new Attr("rt", "2"));
				} else {
					monster.addAttribute(new Attr("rt", "1"));
				}
				
				if (largeBitmap.contains("heat")) {
					monster.addAttribute(new Attr("heat", "3"));
				}
				
				if (largeBitmap.contains("cold")) {
					monster.addAttribute(new Attr("cold", "3"));
				}
				
				String additionalLeader = "0";
				if (monster.getAttribute("additional leadership") != null) {
					additionalLeader = monster.getAttribute("additional leadership").toString();
				}
				monster.addAttribute(new Attr("leader", 40+Integer.parseInt(additionalLeader)));
				if (largeBitmap.contains("noleader")) {
					if (!"".equals(additionalLeader)) {
						monster.addAttribute(new Attr("leader", additionalLeader));
					} else {
						monster.addAttribute(new Attr("leader", 0));
					}
				}
				if (largeBitmap.contains("poorleader")) {
					if (!"".equals(additionalLeader)) {
						monster.addAttribute(new Attr("leader", 10+Integer.parseInt(additionalLeader)));
					} else {
						monster.addAttribute(new Attr("leader", 10));
					}
				}
				if (largeBitmap.contains("goodleader")) {
					if (!"".equals(additionalLeader)) {
						monster.addAttribute(new Attr("leader", 80+Integer.parseInt(additionalLeader)));
					} else {
						monster.addAttribute(new Attr("leader", 80));
					}
				}
				if (largeBitmap.contains("expertleader")) {
					if (!"".equals(additionalLeader)) {
						monster.addAttribute(new Attr("leader", 120+Integer.parseInt(additionalLeader)));
					} else {
						monster.addAttribute(new Attr("leader", 120));
					}
				}
				if (largeBitmap.contains("superiorleader")) {
					if (!"".equals(additionalLeader)) {
						monster.addAttribute(new Attr("leader", 160+Integer.parseInt(additionalLeader)));
					} else {
						monster.addAttribute(new Attr("leader", 160));
					}
				}
				if (largeBitmap.contains("poormagicleader")) {
					monster.addAttribute(new Attr("magicleader", 10));
				}
				if (largeBitmap.contains("okmagicleader")) {
					monster.addAttribute(new Attr("magicleader", 40));
				}
				if (largeBitmap.contains("goodmagicleader")) {
					monster.addAttribute(new Attr("magicleader", 80));
				}
				if (largeBitmap.contains("expertmagicleader")) {
					monster.addAttribute(new Attr("magicleader", 120));
				}
				if (largeBitmap.contains("superiormagicleader")) {
					monster.addAttribute(new Attr("magicleader", 160));
				}
				if (largeBitmap.contains("poorundeadleader")) {
					monster.addAttribute(new Attr("undeadleader", 10));
				}
				if (largeBitmap.contains("okundeadleader")) {
					monster.addAttribute(new Attr("undeadleader", 40));
				}
				if (largeBitmap.contains("goodundeadleader")) {
					monster.addAttribute(new Attr("undeadleader", 80));
				}
				if (largeBitmap.contains("expertundeadleader")) {
					monster.addAttribute(new Attr("undeadleader", 120));
				}
				if (largeBitmap.contains("superiorundeadleader")) {
					monster.addAttribute(new Attr("undeadleader", 160));
				}
				
				if (largeBitmap.contains("coldres15")) {
					String additionalCold = "0";
					if (monster.getAttribute("coldres") != null) {
						additionalCold = monster.getAttribute("coldres").toString();
					}
					boolean cold = false;
					if (largeBitmap.contains("cold") || monster.getAttribute("cold") != null) {
						cold = true;
					}
					monster.addAttribute(new Attr("coldres", 15 + Integer.parseInt(additionalCold.equals("")?"0":additionalCold) + (cold?10:0)));
				} else {
					String additionalCold = "0";
					if (monster.getAttribute("coldres") != null) {
						additionalCold = monster.getAttribute("coldres").toString();
					}
					boolean cold = false;
					if (largeBitmap.contains("cold") || monster.getAttribute("cold") != null) {
						cold = true;
					}
					int coldres = Integer.parseInt(additionalCold.equals("")?"0":additionalCold) + (cold?10:0);
					monster.addAttribute(new Attr("coldres", coldres==0?"":Integer.toString(coldres)));
				}
				if (largeBitmap.contains("fireres15")) {
					String additionalFire = "0";
					if (monster.getAttribute("fireres") != null) {
						additionalFire = monster.getAttribute("fireres").toString();
					}
					boolean heat = false;
					if (largeBitmap.contains("heat") || monster.getAttribute("heat") != null) {
						heat = true;
					}
					monster.addAttribute(new Attr("fireres", 15 + Integer.parseInt(additionalFire.equals("")?"0":additionalFire) + (heat?10:0)));
				} else {
					String additionalFire = "0";
					if (monster.getAttribute("fireres") != null) {
						additionalFire = monster.getAttribute("fireres").toString();
					}
					boolean heat = false;
					if (largeBitmap.contains("heat") || monster.getAttribute("heat") != null) {
						heat = true;
					}
					int fireres = Integer.parseInt(additionalFire.equals("")?"0":additionalFire) + (heat?10:0);
					monster.addAttribute(new Attr("fireres", fireres==0?"":Integer.toString(fireres)));
				}
				if (largeBitmap.contains("poisonres15")) {
					String additionalPoisin = "0";
					if (monster.getAttribute("poisonres") != null) {
						additionalPoisin = monster.getAttribute("poisonres").toString();
					}
					boolean poisoncloud = false;
					if (largeBitmap.contains("undead")|| largeBitmap.contains("inanimate") || monster.getAttribute("poisoncloud") != null) {
						poisoncloud = true;
					}
					monster.addAttribute(new Attr("poisonres", 15 + Integer.parseInt(additionalPoisin.equals("")?"0":additionalPoisin) + (poisoncloud?10:0)));
				} else {
					String additionalPoison = "0";
					if (monster.getAttribute("poisonres") != null) {
						additionalPoison = monster.getAttribute("poisonres").toString();
					}
					boolean poisoncloud = false;
					if (largeBitmap.contains("undead")|| largeBitmap.contains("inanimate") || monster.getAttribute("poisoncloud") != null) {
						poisoncloud = true;
					}
					int poisonres = Integer.parseInt(additionalPoison.equals("")?"0":additionalPoison) + (poisoncloud?10:0);
					monster.addAttribute(new Attr("poisonres", poisonres==0?"":Integer.toString(poisonres)));
				}
				if (largeBitmap.contains("shockres15")) {
					String additionalShock = "0";
					if (monster.getAttribute("shockres") != null) {
						additionalShock = monster.getAttribute("shockres").toString();
					}
					monster.addAttribute(new Attr("shockres", 15 + Integer.parseInt(additionalShock.equals("")?"0":additionalShock)));
				} else {
					String additionalShock = "0";
					if (monster.getAttribute("shockres") != null) {
						additionalShock = monster.getAttribute("shockres").toString();
					}
					int shockres = Integer.parseInt(additionalShock.equals("")?"0":additionalShock);
					monster.addAttribute(new Attr("shockres", shockres==0?"":Integer.toString(shockres)));
				}
				if (largeBitmap.contains("stealthy40")) {
					String additionalStealth = "0";
					if (monster.getAttribute("stealthy") != null) {
						additionalStealth = monster.getAttribute("stealthy").toString();
					}
					boolean glamour = false;
					if (monster.getAttribute("illusion") != null) {
						glamour = true;
					}
					monster.addAttribute(new Attr("stealthy", 40 + Integer.parseInt(additionalStealth.equals("")?"0":additionalStealth) + (glamour?25:0)));
				} else {
					String additionalStealth = "0";
					if (monster.getAttribute("stealthy") != null) {
						additionalStealth = monster.getAttribute("stealthy").toString();
					}
					monster.addAttribute(new Attr("stealthy", additionalStealth == null || additionalStealth.equals("0") ? "" : additionalStealth));
				}
				if (largeBitmap.contains("darkvision100")) {
					if (monster.getAttribute("spiritsight") == null) {
						monster.addAttribute(new Attr("darkvision", 100));
					}
				}

				// magic
				Magic monMagic = monsterMagic.get(rowNumber);
				if (monMagic != null) {
					monster.addAttribute(new Attr("F", magicStrip(monMagic.F)));
					monster.addAttribute(new Attr("A", magicStrip(monMagic.A)));
					monster.addAttribute(new Attr("W", magicStrip(monMagic.W)));
					monster.addAttribute(new Attr("E", magicStrip(monMagic.E)));
					monster.addAttribute(new Attr("S", magicStrip(monMagic.S)));
					monster.addAttribute(new Attr("D", magicStrip(monMagic.D)));
					monster.addAttribute(new Attr("N", magicStrip(monMagic.N)));
					monster.addAttribute(new Attr("B", magicStrip(monMagic.B)));
					monster.addAttribute(new Attr("H", magicStrip(monMagic.H)));
					
					if (monMagic.rand != null) {
						int count = 1;
						for (RandomMagic ranMag : monMagic.rand) {
							monster.addAttribute(new Attr("rand"+count, magicStrip(ranMag.rand)));
							monster.addAttribute(new Attr("nbr"+count, magicStrip(ranMag.nbr)));
							monster.addAttribute(new Attr("link"+count, magicStrip(ranMag.link)));
							monster.addAttribute(new Attr("mask"+count, magicStrip(ranMag.mask)));

							count++;
						}
					}
				}

				if (largeBitmap.contains("misc2")) {
					monster.addAttribute(new Attr("hand", 0));
					monster.addAttribute(new Attr("head", 0));
					monster.addAttribute(new Attr("body", 0));
					monster.addAttribute(new Attr("foot", 0));
					monster.addAttribute(new Attr("misc", 2));
				}
				if (monster.getAttribute("itemslots") != null) {
					String slots = monster.getAttribute("itemslots").toString();
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
					monster.addAttribute(new Attr("hand", numHands));
					monster.addAttribute(new Attr("head", numHeads));
					monster.addAttribute(new Attr("body", numBody));
					monster.addAttribute(new Attr("foot", numFoot));
					monster.addAttribute(new Attr("misc", numMisc));
				} else if (!largeBitmap.contains("misc2")) {
					monster.addAttribute(new Attr("hand", 2));
					monster.addAttribute(new Attr("head", 1));
					monster.addAttribute(new Attr("body", 1));
					monster.addAttribute(new Attr("foot", 1));
					monster.addAttribute(new Attr("misc", 2));
				}
				if (largeBitmap.contains("mounted")) {
					monster.addAttribute(new Attr("foot", 0));
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
			
			// fixedname
			File heroesFile = new File("heroes.txt");
			Set<Integer> heroes = new HashSet<Integer>();
			File namesFile = new File("names.txt");
			List<String> names = new ArrayList<String>();
			FileReader herosFileReader = new FileReader(heroesFile);
			FileReader namesFileReader = new FileReader(namesFile);
			BufferedReader bufferedReader = new BufferedReader(herosFileReader);
			String line;
			while ((line = bufferedReader.readLine()) != null) {
				heroes.add(Integer.parseInt(line));
			}
			bufferedReader.close();
			bufferedReader = new BufferedReader(namesFileReader);
			while ((line = bufferedReader.readLine()) != null) {
				names.add(line);
			}
			bufferedReader.close();
			int nameIndex = 0;
			for (Monster monster : monsterList) {
				Object unique = monster.getAttribute("unique");
				int id = monster.getId();
				if (id == 621 || id == 980 ||id == 981||id==994||id==995||id==996||id==997||
					id==1484 || id==1485|| id==1486|| id==1487 || (id >= 2765 && id <=2781)) {
					unique = "0";
				}
				if (heroes.contains(monster.getId()) || (unique != null && unique.equals("1"))/* || monster.getAttribute("id").equals("641")*/) {
					monster.addAttribute(new Attr("fixedname", names.get(nameIndex++)));
				}
			}

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
					Object object = monster.getAttribute(unit_columns[i]);
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
			
			XSSFWorkbook wb2 = new XSSFWorkbook();
			FileOutputStream fos2 = new FileOutputStream("monster.xlsx");
			XSSFSheet sheet2 = wb2.createSheet();

			rowNum = 0;
			for (Monster monster : monsterList) {
				Set<Attr> attributes = monster.getAttributes();
				if (attributes != null) {
					for (Attr attr : attributes) {
						if (attr.getKey().trim().startsWith("Unknown") || attr.getValue() == null || attr.getValue().equals("")) {
							continue;
						}
						XSSFRow row = sheet2.createRow(rowNum);
						row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(rowNum);
						row.getCell(1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(monster.getId());
						row.getCell(2, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(attr.getKey());
						row.getCell(3, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK).setCellValue(attr.getValue());
						rowNum++;
					}
				}
			}
			
			wb2.write(fos2);
			fos2.close();
			wb2.close();
			
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
		Object name = monster.getAttribute("name");
		Object id = monster.getAttribute("id");
		writer.write(name.toString() + "(" + id + ")");
		writer.newLine();
		for (Attr entry : monster.getAllAttributes()) {
			if (!entry.getKey().equals("name") && !entry.getKey().equals("id") && entry.getValue() != null && !entry.getValue().equals("")) {
				writer.write("\t" + entry.getKey() + ": " + entry.getValue());
				writer.newLine();
			}
		}
		writer.newLine();
	}

	
}
