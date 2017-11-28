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

public class AttributeDumperMonster {
	private static String[] KNOWN_MONSTER_ATTRS = {
		"6C00", // stealthy
		"C900", // coldres
		"DC00", // cold
		"C600", // fireres
		"3C01", // heat
		"C800", // poisonres
		"C700", // shockres
		"6A00", // poisoncloud
		"CD01", // patience
		"AF00", // stormimmune
		"BD00", // regeneration
		"C200", // secondshape
		"C300", // firstshape
		"C100", // shapechange
		"C400", // secondtmpshape
		"1C01", // maxage
		"CA00", // damagerev
		"9E01", // bloodvengeance
		"EB00", // nobadevents
		"E400", // bringeroffortune
		"1901", // darkvision
		"B700", // fear
		"1501", // voidsanity
		"6700", // standard
		"6E01", // formationfighter
		"6F01", // undisciplined
		"9801", // bodyguard
		"A400", // summon1
		"A500", // summon2
		"A600", // summon3
		"7001", // inspirational
		"8300", // pillagebonus
		"BE00", // berserk
		"F200", // startdom
		"F300", // pathcost
		"6F00", // waterbreathing
		"AA01", // realm
		"B401", // batstartsum1
		"B501", // batstartsum2
		"B601", // batstartsum3
		"B701", // batstartsum4
		"B801", // batstartsum5
		"B901", // batstartsum1d6
		"BA01", // batstartsum2d6
		"BB01", // batstartsum3d6
		"BC01", // batstartsum4d6
		"BD01", // batstartsum5d6
		"BE01", // batstartsum6d6
		"2501", // darkpower
		"AE00", // stormpower
		"B100", // firepower
		"B000", // coldpower
		"2501", // darkpower
		"A001", // chaospower
		"4401", // magicpower
		"EA00", // winterpower
		"E700", // springpower
		"E800", // summerpower
		"E900",	// fallpower
		"FB00", // nametype
		"B600",	// itemslots
		"0901", // ressize
		"1D01", // startage
		"AB00", // blind
		"B200", // eyes
		"7A00", // supplybonus
		"7C01", // slave
		"7900", // researchbonus
		"CA01", // chaosrec
		"7D00", // siegebonus
		"D900", // ambidextrous
		"7E01", // invulnerability
		"BF00", // iceprot
		"7500", // reinvigoration
		"D600", // spy
		"E301", // scale walls
		"D500", // assassin
		"D200", // dream seducer
		"2A01", // seduction
		"2901", // explode on death
		"7B01", // taskmaster
		"1301", // unique
		"6900", // awe
		"9D00", // additional leadership
		"B900", // startaff
		"F500", // landshape
		"F600", // watershape
		"4201", // forestshape
		"4301", // plainshape
		"DD00", // uwregen
		"AA00", // patrolbonus
		"D700", // castledef
		"7000", // sailingshipsize
		"9A01", // sailingmaxunitsize
		"DF00", // incunrest
		"BC00", // barbs
		"4E01", // inn
		"1801", // stonebeing
		"5001", // shrinkhp
		"4F01", // growhp
		"FD01", // transformation
		"A101", // domsummon
		"DB00", // domsummon
		"F100", // domsummon
		"6B00", // autosummon
		"8F00", // autosummon
		"AD00", // turmoil summon
		"9200", // cold summon
		"B800", // heretic
		"2001", // popkill
		"6201", // autohealer
		"A300", // fireshield
		"E200", // startingaff
		"1E00", // gemprod fire
		"1F00", // gemprod air
		"2000", // gemprod water
		"2100", // gemprod earth
		"2200", // gemprod astral
		"2300", // gemprod death
		"2400", // gemprod nature
		"2500", // gemprod blood
		"F800", // fixdresearch
		"C201", // divineins
		"4701", // halt
		"AF01", // crossbreeder
		"7D01", // reclimit
		"C501", // fixforgebonus
		"6B01", // mastersmith
		"A900", // lamiabonus
		"FD00", // homesick
		"DE00", // banefireshield
		"A200", // animalawe
		"6301", // autodishealer
		"4801", // shatteredsoul
		"CE00", // voidsum
		"AE01", // makepearls
		"5501", // inspiringres
		"1101", // drainimmune
		"AC00", // diseasecloud
		"D300", // inquisitor
		"7101", // beastmaster
		"7400", // douse
		"6C01", // preanimator
		"6D01", // dreanimator
		"FF01", // mummify
		"D100", // onebattlespell
		"F501", // fireattuned
		"F601", // airattuned
		"F701", // waterattuned
		"F801", // earthattuned
		"F901", // astralattuned
		"FA01", // deathattuned
		"FB01", // natureattuned
		"FC01", // bloodattuned
		"0A00", // magicboost F
		"0B00", // magicboost A
		"0C00", // magicboost W
		"0D00", // magicboost E
		"0E00", // magicboost S
		"0F00", // magicboost D
		"1000", // magicboost N
		"1600", // magicboost ALL
		"EA01", // heatrec
		"EB01", // coldrec
		"D801", // spreadchaos/death
		"D901", // spreadorder/growth
		"EC00", // corpseeater
		"0602", // poisonskin
		"AB01", // bug
		"AC01", // uwbug
		"9C00", // spreaddom
		"8A01", // reform
		"8E01", // battlesum5
		"F101", // acidsplash
		"F201", // drake
		"0402", // prophetshape
		"8701", // horror
		"3501", // insane
		"D800", // sacr
		"1002", // enchrebate50
		"7C00", // leper
		"9C01", // slimer
		"9201", // mindslime
		"9701", // resources
		"E000", // corrupt
		"6800", // petrify
		"F400", // eyeloss
		"0401", // ethtrue
		"1602", // heroarrivallimit
		"1702", // sailsize
		"0201", // uwdamage
		"0E02", // landdamage

};

	private static List<String> attrList = new ArrayList<String>();
	
	private static Map<String, Integer> Summaries = new HashMap<String, Integer>();
	
	public static void main(String[] args) {
		try {
			FileInputStream stream = new FileInputStream("Dominions5.exe");			
			stream.skip(Starts.MONSTER);
			int i = 0;
			long numFound = 0;
			byte[] c = new byte[2];
			stream.skip(64);
			while ((stream.read(c, 0, 2)) != -1) {
				String high = String.format("%02X", c[1]);
				String low = String.format("%02X", c[0]);
				int weapon = Integer.decode("0X" + high + low);
				if (weapon == 0) {
					stream.skip(46l - numFound*2l);
					System.out.print("id:" + (i+1));
					// Values
					for (int x = 0; x < numFound; x++) {
						byte[] d = new byte[4];
						stream.read(d, 0, 4);
						String high1 = String.format("%02X", d[3]);
						String low1 = String.format("%02X", d[2]);
						high = String.format("%02X", d[1]);
						low = String.format("%02X", d[0]);
						if (!Arrays.asList(KNOWN_MONSTER_ATTRS).contains(attrList.get(x))) {
							System.out.print("\n\t" + attrList.get(x) + ": ");
							System.out.print(new BigInteger(high1 + low1 + high + low, 16).intValue() + " ");
						}
					}

					System.out.println("");
					stream.skip(254l - 46l - numFound*4l);
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
				if (i >= Starts.MONSTER_COUNT) {
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
	        int monsters = 0;
			System.out.println("Summary:");
			for (Entry<String, Integer> entry : list) {
				if (!Arrays.asList(KNOWN_MONSTER_ATTRS).contains(entry.getKey())) {
					System.out.println("id: " + entry.getKey() + " " + entry.getValue());
					indexes++;
					monsters += entry.getValue();
				}
			}
			System.out.println("---------------------------");
			System.out.println("indexes: " + indexes + " monsters: " + monsters);
			stream.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
