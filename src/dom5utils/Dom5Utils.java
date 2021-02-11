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

public class Dom5Utils {
	public static void main(String[] args) {
		// Stats
		System.out.println("Beginning item stats"); 
		ItemStatIndexer.run();
		System.out.println("Beginning site stats"); 
		SiteStatIndexer.run();
		System.out.println("Beginning monster stats"); 
		MonsterStatIndexer.run();
		System.out.println("Beginning event stats"); 
		EventStatIndexer.run();
		System.out.println("Beginning merc stats"); 
		MercenaryStatIndexer.run();
		System.out.println("Beginning armor stats"); 
		ArmorStatIndexer.run();
		System.out.println("Beginning wpn stats"); 
		WeaponStatIndexer.run();
		System.out.println("Beginning spell stats"); 
		SpellStatIndexer.run();
		System.out.println("Beginning nation stats"); 
		NationStatIndexer.run();
		
		// Descriptions
		System.out.println("Beginning item/monster descr"); 
		ItemMonsterDescDumper.run();
		System.out.println("Beginning spell descr"); 
		SpellDescDumper.run();
		
		// Events
		System.out.println("Beginning event stats"); 
		EventStatIndexer.run();
		
		// Sprites
		System.out.println("Beginning item sprites"); 
		ItemSpriteIndexer.run();
		System.out.println("Beginning monster sprites"); 
		MonsterSpriteIndexer.run();
		
		// Names
		System.out.println("Beginning names"); 
		NametypeIndexer.run();
	}
}
