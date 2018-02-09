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
		ItemStatIndexer.run();
		SiteStatIndexer.run();
		MonsterStatIndexer.run();
		EventStatIndexer.run();
		MercenaryStatIndexer.run();
		ArmorStatIndexer.run();
		WeaponStatIndexer.run();
		SpellStatIndexer.run();
		NationStatIndexer.run();
		
		// Descriptions
		ItemMonsterDescDumper.run();
		SpellDescDumper.run();
		
		// Events
		EventStatIndexer.run();
		
		// Sprites
		ItemSpriteIndexer.run();
		MonsterSpriteIndexer.run();
		
		// Names
		NametypeIndexer.run();
	}
}
