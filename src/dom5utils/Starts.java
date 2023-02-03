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

public class Starts {
	public static final long ITEM = 0x01168d18l;
	public static final long ITEM_SIZE = 240l;
	public static final long ITEM_ATTRIBUTE_OFFSET = 120l;
	public static final long ITEM_ATTRIBUTE_GAP = 30l;
	public static final long ITEM_BITMAP_START = 216l;
	
	public static final long MONSTER = 0x011f56b8l;
	public static final long MONSTER_SIZE = 288l;
	public static final long MONSTER_ATTRIBUTE_OFFSET = 64l;
	public static final long MONSTER_ATTRIBUTE_GAP = 54l;

	public static final long MONSTER_MAGIC = 0x015b6550l;
	public static final long ITEM_AND_MONSTER_DESC = 0x00233b00l;
	public static final long ITEM_AND_MONSTER_DESC_INDEX = 0x0038ee70l;
	
	public static final long MONSTER_TRS_INDEX = 0x0001ac1c;
	
	public static final long ITEM_TRS_INDEX = 0x0000243cl;

	public static final long SITE = 0x01479640l;
	public static final long SITE_SIZE = 216l;
	public static final long SITE_ATTRIBUTE_OFFSET = 44;
	public static final long SITE_ATTRIBUTE_GAP = 34l;
	
	public static final long NAMES = 0x00ff00bcl;
	public static final long FIXED_NAMES = 0x01137e58l;
	public static final int NAMES_COUNT = 161;
	
	public static final long SPELL = 0x014e3650l;
	public static final long SPELL_SIZE = 216l;

	public static final long SPELL_DESC = 0x003a26f0l;
	public static final long SPELL_DESC_INDEX = 0x003ee158l;

	public static final long EVENT = 0x003f5fa8l;
	public static final long EVENT_SIZE = 1728l;
	
	public static final long MERCENARY = 0x011c28e2l;
	public static final long MERCENARY_SIZE = 312l;

	public static final long ARMOR = 0x0020c480l;
	public static final long ARMOR_SIZE = 104l;

	public static final long WEAPON = 0x015dd270l;
	public static final long WEAPON_SIZE = 112l;
	
	public static final long NATION = 0x00eb9e00l;
	public static final long NATION_SIZE = 2400l;
}