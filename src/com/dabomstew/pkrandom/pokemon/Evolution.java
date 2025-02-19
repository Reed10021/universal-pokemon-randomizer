package com.dabomstew.pkrandom.pokemon;

/*----------------------------------------------------------------------------*/
/*--  Evolution.java - represents an evolution between 2 Pokemon.           --*/
/*--                                                                        --*/
/*--  Part of "Universal Pokemon Randomizer" by Dabomstew                   --*/
/*--  Pokemon and any associated names and the like are                     --*/
/*--  trademark and (C) Nintendo 1996-2012.                                 --*/
/*--                                                                        --*/
/*--  The custom code written here is licensed under the terms of the GPL:  --*/
/*--                                                                        --*/
/*--  This program is free software: you can redistribute it and/or modify  --*/
/*--  it under the terms of the GNU General Public License as published by  --*/
/*--  the Free Software Foundation, either version 3 of the License, or     --*/
/*--  (at your option) any later version.                                   --*/
/*--                                                                        --*/
/*--  This program is distributed in the hope that it will be useful,       --*/
/*--  but WITHOUT ANY WARRANTY; without even the implied warranty of        --*/
/*--  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the          --*/
/*--  GNU General Public License for more details.                          --*/
/*--                                                                        --*/
/*--  You should have received a copy of the GNU General Public License     --*/
/*--  along with this program. If not, see <http://www.gnu.org/licenses/>.  --*/
/*----------------------------------------------------------------------------*/

public class Evolution implements Comparable<Evolution> {

    public Pokemon from;
    public Pokemon to;
    public boolean carryStats;
    public EvolutionType type;
    public int extraInfo;

    public Evolution(Pokemon from, Pokemon to, boolean carryStats, EvolutionType type, int extra) {
        this.from = from;
        this.to = to;
        this.carryStats = carryStats;
        this.type = type;
        this.extraInfo = extra;
    }

    public Evolution(Evolution evo) {
        this.from = evo.from;
        this.to = evo.to;
        this.carryStats = evo.carryStats;
        this.type = evo.type;
        this.extraInfo = evo.extraInfo;
    }

    @Override
    public int hashCode() {
        final int prime = 31;
        int result = 1;
        result = prime * result + from.number;
        result = prime * result + to.number;
        result = prime * result + type.ordinal();
        return result;
    }

    @Override
    public boolean equals(Object obj) {
        if (this == obj)
            return true;
        if (obj == null)
            return false;
        if (getClass() != obj.getClass())
            return false;
        Evolution other = (Evolution) obj;
        if (from != other.from)
            return false;
        if (to != other.to)
            return false;
        if (type != other.type)
            return false;
        return true;
    }

    @Override
    public int compareTo(Evolution o) {
        if (this.from.number < o.from.number) {
            return -1;
        } else if (this.from.number > o.from.number) {
            return 1;
        } else if (this.to.number < o.to.number) {
            return -1;
        } else if (this.to.number > o.to.number) {
            return 1;
        } else if (this.type.ordinal() < o.type.ordinal()) {
            return -1;
        } else if (this.type.ordinal() > o.type.ordinal()) {
            return 1;
        } else {
            return 0;
        }
    }

    @Override
    public String toString() {
        return "\nfrom = " + from.name + "\n to = " + to.name + "\ncarryStats = " + carryStats
                + "\ntype = " + type.toString() + "\nextraInfo = " + extraInfo;
    }

}
