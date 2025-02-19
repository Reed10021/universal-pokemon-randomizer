package com.dabomstew.pkrandom;

/*----------------------------------------------------------------------------*/
/*--  Randomizer.java - Can randomize a file based on settings.             --*/
/*--                    Output varies by seed.                              --*/
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


import java.io.*;
import java.util.*;
import java.util.stream.Collectors;

import com.dabomstew.pkrandom.pokemon.*;
import com.dabomstew.pkrandom.romhandlers.Gen1RomHandler;
import com.dabomstew.pkrandom.romhandlers.Gen4RomHandler;
import com.dabomstew.pkrandom.romhandlers.Gen5RomHandler;
import com.dabomstew.pkrandom.romhandlers.RomHandler;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

// Can randomize a file based on settings. Output varies by seed.
public class Randomizer {

    private static final String NEWLINE = System.getProperty("line.separator");

    private final Settings settings;
    private final RomHandler romHandler;

    public Randomizer(Settings settings, RomHandler romHandler) {
        this.settings = settings;
        this.romHandler = romHandler;
    }

    public int randomize(final String filename, final PrintStream log, final WorkbookHandler workbookHandler, long seed) {
        final long startTime = System.currentTimeMillis();
        RandomSource.seed(seed);
        final boolean raceMode = settings.isRaceMode();

        int checkValue = 0;

        // Deep copy the evolutions
        Map<Pokemon, List<Evolution>> originalEvos = new HashMap<>();
        for (Pokemon pk : romHandler.getPokemon()) {
            if(pk != null) {
                List<Evolution> t = pk.evolutionsFrom.stream()
                        .map(Evolution::new)
                        .collect(Collectors.toList());
                originalEvos.put(pk, t);
            }
        }

        // limit pokemon?
        if (settings.isLimitPokemon()) {
            romHandler.setPokemonPool(settings.getCurrentRestrictions());
            romHandler.removeEvosForPokemonPool();
        } else {
            romHandler.setPokemonPool(null);
        }

        List<Move> oldMoves = romHandler.getMoves();

        // Move updates & data changes
        if (settings.isUpdateMoves()) {
            romHandler.initMoveUpdates();
            if (!(romHandler instanceof Gen5RomHandler)) {
                romHandler.updateMovesToGen5();
            }
            if (!settings.isUpdateMovesLegacy()) {
                romHandler.updateMovesToGen6();
            }
            romHandler.printMoveUpdates();
        }

        if (settings.isRandomizeMovePowers()) {
            romHandler.randomizeMovePowers();
        }

        if (settings.isRandomizeMoveAccuracies()) {
            romHandler.randomizeMoveAccuracies();
        }

        if (settings.isRandomizeMovePPs()) {
            romHandler.randomizeMovePPs();
        }

        if (settings.isRandomizeMoveTypes()) {
            romHandler.randomizeMoveTypes();
        }

        if (settings.isRandomizeMoveCategory() && romHandler.hasPhysicalSpecialSplit()) {
            romHandler.randomizeMoveCategory();
        }

        List<Move> moves = romHandler.getMoves();

        // Misc Tweaks?
        int currentMiscTweaks = settings.getCurrentMiscTweaks();
        if (romHandler.miscTweaksAvailable() != 0) {
            int codeTweaksAvailable = romHandler.miscTweaksAvailable();
            List<MiscTweak> tweaksToApply = new ArrayList<MiscTweak>();

            for (MiscTweak mt : MiscTweak.allTweaks) {
                if ((codeTweaksAvailable & mt.getValue()) > 0 && (currentMiscTweaks & mt.getValue()) > 0) {
                    tweaksToApply.add(mt);
                }
            }

            // Sort so priority is respected in tweak ordering.
            Collections.sort(tweaksToApply);

            // Now apply in order.
            for (MiscTweak mt : tweaksToApply) {
                romHandler.applyMiscTweak(mt);
            }
        }

        if (settings.isUpdateBaseStats()) {
            romHandler.updatePokemonStats();
        }

        // Base stats changing
        switch (settings.getBaseStatisticsMod()) {
        case SHUFFLE:
            romHandler.shufflePokemonStats(settings.isBaseStatsFollowEvolutions());
            break;
        case RANDOM:
            romHandler.randomizePokemonStats(settings.isBaseStatsFollowEvolutions());
            break;
        case TRUERANDOM:
            romHandler.truerandomizePokemonStats(settings.isBaseStatsFollowEvolutions());
            break;
        default:
            break;
        }

        if (settings.isStandardizeEXPCurves()) {
            romHandler.standardizeEXPCurves();
        }

        // Abilities? (new 1.0.2)
        if (romHandler.abilitiesPerPokemon() > 0 && settings.getAbilitiesMod() == Settings.AbilitiesMod.RANDOMIZE) {
            romHandler.randomizeAbilities(settings.isAbilitiesFollowEvolutions(), settings.isAllowWonderGuard(),
                    settings.isBanTrappingAbilities(), settings.isBanNegativeAbilities(), settings.isForcingTwoAbilities() );
        }

        // Pokemon Types
        switch (settings.getTypesMod()) {
        case RANDOM_FOLLOW_EVOLUTIONS:
            romHandler.randomizePokemonTypes(true);
            break;
        case COMPLETELY_RANDOM:
            romHandler.randomizePokemonTypes(false);
            break;
        default:
            break;
        }

        // Wild Held Items?
        if (settings.isRandomizeWildPokemonHeldItems()) {
            romHandler.randomizeWildHeldItems(settings.isBanBadRandomWildPokemonHeldItems(), settings.getForceHeldItemMode() );
        }

        maybeLogBaseStatAndTypeChanges(log, romHandler);
        if (settings.getBaseStatisticsMod() != Settings.BaseStatisticsMod.UNCHANGED
                || settings.getTypesMod() != Settings.TypesMod.UNCHANGED
                || settings.getAbilitiesMod() != Settings.AbilitiesMod.UNCHANGED
                || settings.isRandomizeWildPokemonHeldItems()) {
            workbookHandler.logToWorkbookBaseStatAndTypeChanges(romHandler);
        }

        for (Pokemon pkmn : romHandler.getPokemon()) {
            if (pkmn != null) {
                checkValue = addToCV(checkValue, pkmn.hp, pkmn.attack, pkmn.defense, pkmn.speed, pkmn.spatk,
                        pkmn.spdef, pkmn.ability1, pkmn.ability2, pkmn.ability3);
            }
        }

        // Random Evos
        // Applied after type to pick new evos based on new types.
        if (settings.getEvolutionsMod() == Settings.EvolutionsMod.RANDOM) {
            romHandler.randomizeEvolutions(settings.isEvosSimilarStrength(), settings.isEvosSameTyping(),
                    settings.isEvosMaxThreeStages(), settings.isEvosForceChange());

            logRandomizedEvolutions(log, romHandler);
            if (!settings.isChangeImpossibleEvolutions() && !settings.isMakeEvolutionsEasier()) {
                // Only output evolutions to workbook once. So if we're not done making changes, don't log it.
                workbookHandler.logToWorkbookRandomizedEvolutions(romHandler, originalEvos);
            }
        }

        // Trade evolutions removal
        if (settings.isChangeImpossibleEvolutions()) {
            romHandler.removeTradeEvolutions(!(settings.getMovesetsMod() == Settings.MovesetsMod.UNCHANGED));
            // Again, if we're not done making changes to evolutions yet, keep going and don't log it to the workbook
            if(!settings.isMakeEvolutionsEasier()) {
                workbookHandler.logToWorkbookRandomizedEvolutions(romHandler, originalEvos);
            }
        }

        // Easier evolutions
        if (settings.isMakeEvolutionsEasier()) {
            romHandler.condenseLevelEvolutions(40, 30);
            workbookHandler.logToWorkbookRandomizedEvolutions(romHandler, originalEvos);
        }

        // Starter Pokemon
        // Applied after type to update the strings correctly based on new types
        List<Pokemon> oldStarters = romHandler.getStarters();
        maybeChangeAndLogStarters(log, romHandler);
        // If starters changed, log it to the workbook
        if(!oldStarters.containsAll(romHandler.getStarters())) {
            workbookHandler.logToWorkbookStarters(romHandler, oldStarters);
        }

        // Move Data Log
        // Placed here so it matches its position in the randomizer interface
        maybeLogMoveChanges(log, romHandler);
        if (settings.isRandomizeMoveAccuracies() || settings.isRandomizeMovePowers() || settings.isRandomizeMovePPs()
                || settings.isRandomizeMoveCategory() || settings.isRandomizeMoveTypes() || settings.isUpdateMoves()) {
            workbookHandler.logToWorkbookMoveChanges(romHandler);
        }

        // Movesets
        boolean noBrokenMoves = settings.doBlockBrokenMoves();
        boolean forceFourLv1s = romHandler.supportsFourStartingMoves() && settings.isStartWithFourMoves();
        double msGoodDamagingProb = settings.isMovesetsForceGoodDamaging() ? settings.getMovesetsGoodDamagingPercent() / 100.0
                : 0;
        if (settings.getMovesetsMod() == Settings.MovesetsMod.RANDOM_PREFER_SAME_TYPE) {
            romHandler.randomizeMovesLearnt(1, noBrokenMoves, forceFourLv1s, msGoodDamagingProb);
        }else if (settings.getMovesetsMod() == Settings.MovesetsMod.RANDOM_STRICT_TYPE_NORMAL) {
            romHandler.randomizeMovesLearnt(2, noBrokenMoves, forceFourLv1s, msGoodDamagingProb);
        }else if (settings.getMovesetsMod() == Settings.MovesetsMod.RANDOM_STRICT_TYPE) {
            romHandler.randomizeMovesLearnt(3, noBrokenMoves, forceFourLv1s, msGoodDamagingProb);
        } else if (settings.getMovesetsMod() == Settings.MovesetsMod.COMPLETELY_RANDOM) {
            romHandler.randomizeMovesLearnt(0, noBrokenMoves, forceFourLv1s, msGoodDamagingProb);
        } else {
            if (noBrokenMoves) {
                romHandler.removeBrokenMoves();
            }
            if( forceFourLv1s ) {
                romHandler.forceFourStartingMovesOnly();
            }
        }
        

        if (settings.isReorderDamagingMoves()) {
            romHandler.orderDamagingMovesByDamage();
        }

        // Trainer Pokemon
        if (settings.getTrainersMod() == Settings.TrainersMod.UNCHANGED && settings.isTrainersLevelModified() ){
            romHandler.levelUpTrainerPokes(settings.getTrainersLevelModifier(),
                    settings.getMinimumDifficulty());
        } else if (settings.getTrainersMod() == Settings.TrainersMod.RANDOM) {
            romHandler.randomizeTrainerPokes(settings.isTrainersUsePokemonOfSimilarStrength(),
                    settings.isTrainersBlockLegendaries(), settings.isTrainersBlockEarlyWonderGuard(),
                    settings.isTrainersLevelModified() ? settings.getTrainersLevelModifier() : 0,
                    settings.getMinimumDifficulty());
        } else if (settings.getTrainersMod() == Settings.TrainersMod.TYPE_THEMED) {
            romHandler.typeThemeTrainerPokes(settings.isTrainersUsePokemonOfSimilarStrength(),
                    settings.isTrainersMatchTypingDistribution(), settings.isTrainersBlockLegendaries(),
                    settings.isTrainersBlockEarlyWonderGuard(),
                    settings.isTrainersLevelModified() ? settings.getTrainersLevelModifier() : 0,
                    settings.getMinimumDifficulty());
        }

        if ((settings.getTrainersMod() != Settings.TrainersMod.UNCHANGED || settings.getStartersMod() != Settings.StartersMod.UNCHANGED)
                && settings.isRivalCarriesStarterThroughout()) {
            romHandler.rivalCarriesStarter();
        }

        if (settings.isTrainersForceFullyEvolved()) {
            romHandler.forceFullyEvolvedTrainerPokes(settings.getTrainersForceFullyEvolvedLevel());
        }

        // Trainer names & class names randomization
        // done before trainer log to add proper names

        if (romHandler.canChangeTrainerText()) {
            if (settings.isRandomizeTrainerClassNames()) {
                romHandler.randomizeTrainerClassNames(settings.getCustomNames());
            }

            if (settings.isRandomizeTrainerNames()) {
                romHandler.randomizeTrainerNames(settings.getCustomNames());
            }
        }

        // Apply metronome only mode now that trainers have been dealt with
        if (settings.getMovesetsMod() == Settings.MovesetsMod.METRONOME_ONLY) {
            romHandler.metronomeOnlyMode();
        }

        List<Trainer> trainers = romHandler.getTrainers();
        for (Trainer t : trainers) {
            for (TrainerPokemon tpk : t.pokemon) {
                checkValue = addToCV(checkValue, tpk.level, tpk.pokemon.number);
            }
        }


        maybeLogMovesetChanges(log, romHandler, forceFourLv1s);
        if (settings.getMovesetsMod() != Settings.MovesetsMod.UNCHANGED || settings.doBlockBrokenMoves() || forceFourLv1s) {
            workbookHandler.logToWorkbookMovesetChanges(romHandler);
        }
        maybeLogTrainerChanges(log, romHandler);
        if(settings.getTrainersMod() != Settings.TrainersMod.UNCHANGED
                || settings.isRivalCarriesStarterThroughout()
                || settings.isTrainersLevelModified()) {
            workbookHandler.logToWorkbookTrainerChanges(romHandler);
        }

        // Static Pokemon
        List<Pokemon> oldStatics = romHandler.getStaticPokemon();
        checkValue = maybeChangeAndLogStaticPokemon(log, romHandler, raceMode, checkValue);
        if(!oldStatics.containsAll(romHandler.getStaticPokemon())) {
            workbookHandler.logToWorkbookStaticPokemon(romHandler, oldStatics);
        }

        // Wild Pokemon
        if (settings.isUseMinimumCatchRate()) {
            boolean gen5 = romHandler instanceof Gen5RomHandler;
            int normalMin, legendaryMin;
            switch (settings.getMinimumCatchRateLevel()) {
            case 1:
            default:
                normalMin = gen5 ? 50 : 75;
                legendaryMin = gen5 ? 25 : 37;
                break;
            case 2:
                normalMin = gen5 ? 100 : 128;
                legendaryMin = gen5 ? 45 : 64;
                break;
            case 3:
                normalMin = gen5 ? 180 : 200;
                legendaryMin = gen5 ? 75 : 100;
                break;
            case 4:
                normalMin = legendaryMin = 255;
                break;
            }
            romHandler.minimumCatchRate(normalMin, legendaryMin);
        }

        switch (settings.getWildPokemonMod()) {
        case RANDOM:
            romHandler.randomEncounters(settings.isUseTimeBasedEncounters(),
                    settings.getWildPokemonRestrictionMod() == Settings.WildPokemonRestrictionMod.CATCH_EM_ALL,
                    settings.getWildPokemonRestrictionMod() == Settings.WildPokemonRestrictionMod.TYPE_THEME_AREAS,
                    settings.getWildPokemonRestrictionMod() == Settings.WildPokemonRestrictionMod.SIMILAR_STRENGTH,
                    settings.isBlockWildLegendaries(),
                    settings.isWildLevelModifiedHigh() ? settings.getWildLevelHighModifier() : -1, 
                    settings.isWildLevelModifiedLow() ? settings.getWildLevelLowModifier() : -1);
            break;
        case AREA_MAPPING:
            romHandler.area1to1Encounters(settings.isUseTimeBasedEncounters(),
                    settings.getWildPokemonRestrictionMod() == Settings.WildPokemonRestrictionMod.CATCH_EM_ALL,
                    settings.getWildPokemonRestrictionMod() == Settings.WildPokemonRestrictionMod.TYPE_THEME_AREAS,
                    settings.getWildPokemonRestrictionMod() == Settings.WildPokemonRestrictionMod.SIMILAR_STRENGTH,
                    settings.isBlockWildLegendaries(),
                    settings.isWildLevelModifiedHigh() ? settings.getWildLevelHighModifier() : -1, 
                    settings.isWildLevelModifiedLow() ? settings.getWildLevelLowModifier() : -1);
            break;
        case GLOBAL_MAPPING:
            romHandler.game1to1Encounters(settings.isUseTimeBasedEncounters(),
                    settings.getWildPokemonRestrictionMod() == Settings.WildPokemonRestrictionMod.SIMILAR_STRENGTH,
                    settings.isBlockWildLegendaries(),
                    settings.isWildLevelModifiedHigh() ? settings.getWildLevelHighModifier() : -1, 
                    settings.isWildLevelModifiedLow() ? settings.getWildLevelLowModifier() : -1);
            break;
        default:
            if( settings.isWildLevelModifiedHigh() )
            {
                List<EncounterSet> encounters = romHandler.getEncounters(settings.isUseTimeBasedEncounters());
                for (EncounterSet es : encounters) 
                {
                    for (Encounter enc : es.encounters) 
                    {
                        enc = romHandler.levelUpEncounterPub(enc, settings.getWildLevelHighModifier(),
                                settings.isWildLevelModifiedLow() ? settings.getWildLevelLowModifier() : -1);
                    }
                }
                romHandler.setEncounters(settings.isUseTimeBasedEncounters(), encounters);
            }
            break;
        }

        maybeLogWildPokemonChanges(log, romHandler);
        if(settings.getWildPokemonMod() != Settings.WildPokemonMod.UNCHANGED || settings.isWildLevelModifiedHigh()) {
            workbookHandler.logToWorkbookWildPokemonChanges(romHandler, settings.isUseTimeBasedEncounters());
        }

        List<EncounterSet> encounters = romHandler.getEncounters(settings.isUseTimeBasedEncounters());
        for (EncounterSet es : encounters) {
            for (Encounter e : es.encounters) {
                checkValue = addToCV(checkValue, e.level, e.pokemon.number);
            }
        }

        // TMs
        List<Integer> oldTms = romHandler.getTMMoves();
        if (!(settings.getMovesetsMod() == Settings.MovesetsMod.METRONOME_ONLY)
                && settings.getTmsMod() == Settings.TMsMod.RANDOM) {
            double goodDamagingProb = settings.isTmsForceGoodDamaging() ? settings.getTmsGoodDamagingPercent() / 100.0
                    : 0;
            romHandler.randomizeTMMoves(noBrokenMoves, settings.isKeepFieldMoveTMs(), goodDamagingProb);
            log.println("--TM Moves--");
            List<Integer> tmMoves = romHandler.getTMMoves();
            for (int i = 0; i < tmMoves.size(); i++) {
                log.printf("TM%02d %s" + NEWLINE, i + 1, moves.get(tmMoves.get(i)).name);
                checkValue = addToCV(checkValue, tmMoves.get(i));
            }
            log.println();
        } else if (settings.getMovesetsMod() == Settings.MovesetsMod.METRONOME_ONLY) {
            log.println("TM Moves: Metronome Only." + NEWLINE);
        } else {
            log.println("TM Moves: Unchanged." + NEWLINE);
        }

        if(!oldTms.containsAll(romHandler.getTMMoves())) {
            workbookHandler.logToWorkbookRandomizedTmMoves(romHandler, oldTms);
        }

        // TM/HM compatibility
        switch (settings.getTmsHmsCompatibilityMod()) {
        case RANDOM_PREFER_TYPE:
        case RANDOM_PREFER_TYPE_AND_NORMAL:
        case COMPLETELY_RANDOM:
            romHandler.randomizeTMHMCompatibility(settings.getTmsHmsCompatibilityMod());
            break;
        case FULL:
            romHandler.fullTMHMCompatibility();
            break;
        default:
            break;
        }

        if (settings.isTmLevelUpMoveSanity()) {
            romHandler.ensureTMCompatSanity();
        }

        if (settings.isFullHMCompat()) {
            romHandler.fullHMCompatibility();
        }

        if(settings.getTmsHmsCompatibilityMod() != Settings.TMsHMsCompatibilityMod.FULL &&
                settings.getTmsHmsCompatibilityMod() != Settings.TMsHMsCompatibilityMod.UNCHANGED) {
            Map<Pokemon, boolean[]> compatMap = romHandler.getTMHMCompatibility();
            List<String> movesets = new ArrayList<String>();
            log.println("--TM Compatibility--");
            for (Pokemon pkmn : compatMap.keySet()) {
                StringBuilder sb = new StringBuilder();
                sb.append(String.format("%03d %-10s : ", pkmn.number, pkmn.name));
                boolean[] data = compatMap.get(pkmn);
                boolean first = true;
                for (int i = 1; i < data.length; i++) {
                    if(data[i]) {
                        if (!first) {
                            sb.append(", ");
                        }
                        if(romHandler.getTMMoves().size() < i) {
                            sb.append(String.format("HM%02d ", (i - romHandler.getTMMoves().size()))).append(moves.get(romHandler.getHMMoves().get(i - romHandler.getTMMoves().size() - 1)).name);
                        } else {
                            sb.append(String.format("TM%02d ", i)).append(moves.get(romHandler.getTMMoves().get(i - 1)).name);
                        }
                        first = false;
                    }
                }
                movesets.add(sb.toString());
            }
            Collections.sort(movesets);
            for (String moveset : movesets) {
                log.println(moveset);
            }
            log.println();
            workbookHandler.logtoWorkbookTmHmCompatability(romHandler);
        }

        // Move Tutors (new 1.0.3)
        if (romHandler.hasMoveTutors()) {
            if (!(settings.getMovesetsMod() == Settings.MovesetsMod.METRONOME_ONLY)
                    && settings.getMoveTutorMovesMod() == Settings.MoveTutorMovesMod.RANDOM) {
                List<Integer> oldMtMoves = romHandler.getMoveTutorMoves();
                double goodDamagingProb = settings.isTutorsForceGoodDamaging() ? settings
                        .getTutorsGoodDamagingPercent() / 100.0 : 0;
                romHandler.randomizeMoveTutorMoves(noBrokenMoves, settings.isKeepFieldMoveTutors(), goodDamagingProb);
                log.println("--Move Tutor Moves--");
                List<Integer> newMtMoves = romHandler.getMoveTutorMoves();
                for (int i = 0; i < newMtMoves.size(); i++) {
                    log.printf("%s => %s" + NEWLINE, moves.get(oldMtMoves.get(i)).name,
                            moves.get(newMtMoves.get(i)).name);
                    checkValue = addToCV(checkValue, newMtMoves.get(i));
                }
                log.println();
                workbookHandler.logToWorkbookRandomizedMoveTutors(romHandler, oldMtMoves);
            } else if (settings.getMovesetsMod() == Settings.MovesetsMod.METRONOME_ONLY) {
                log.println("Move Tutor Moves: Metronome Only." + NEWLINE);
            } else {
                log.println("Move Tutor Moves: Unchanged." + NEWLINE);
            }

            // Compatibility
            switch (settings.getMoveTutorsCompatibilityMod()) {
            case RANDOM_PREFER_TYPE:
            case RANDOM_PREFER_TYPE_AND_NORMAL:
            case COMPLETELY_RANDOM:
                romHandler.randomizeMoveTutorCompatibility(settings.getMoveTutorsCompatibilityMod());
                break;
            case FULL:
                romHandler.fullMoveTutorCompatibility();
                break;
            default:
                break;
            }

            if (settings.isTutorLevelUpMoveSanity()) {
                romHandler.ensureMoveTutorCompatSanity();
            }

            if(settings.getMoveTutorsCompatibilityMod() != Settings.MoveTutorsCompatibilityMod.FULL &&
                    settings.getMoveTutorsCompatibilityMod() != Settings.MoveTutorsCompatibilityMod.UNCHANGED) {
                Map<Pokemon, boolean[]> compatMap = romHandler.getMoveTutorCompatibility();
                List<String> movesets = new ArrayList<>();
                log.println("--Move Tutor Compatibility--");
                for (Pokemon pkmn : compatMap.keySet()) {
                    StringBuilder sb = new StringBuilder();
                    sb.append(String.format("%03d %-10s : ", pkmn.number, pkmn.name));
                    boolean[] data = compatMap.get(pkmn);
                    boolean first = true;
                    for (int i = 1; i < data.length; i++) {
                        if(data[i]) {
                            if (!first) {
                                sb.append(", ");
                            }
                            sb.append(moves.get(romHandler.getMoveTutorMoves().get(i - 1)).name);
                            first = false;
                        }
                    }
                    movesets.add(sb.toString());
                }
                Collections.sort(movesets);
                for (String moveset : movesets) {
                    log.println(moveset);
                }
                log.println();
                workbookHandler.logToWorkbookRandomizedMoveTutorCompat(romHandler);
            }
        }

        // In-game trades
        List<IngameTrade> oldTrades = romHandler.getIngameTrades();
        if (settings.getInGameTradesMod() == Settings.InGameTradesMod.RANDOMIZE_GIVEN) {
            romHandler.randomizeIngameTrades(false, settings.isRandomizeInGameTradesNicknames(),
                    settings.isRandomizeInGameTradesOTs(), settings.isRandomizeInGameTradesIVs(),
                    settings.isRandomizeInGameTradesItems(), settings.getCustomNames());
        } else if (settings.getInGameTradesMod() == Settings.InGameTradesMod.RANDOMIZE_GIVEN_AND_REQUESTED) {
            romHandler.randomizeIngameTrades(true, settings.isRandomizeInGameTradesNicknames(),
                    settings.isRandomizeInGameTradesOTs(), settings.isRandomizeInGameTradesIVs(),
                    settings.isRandomizeInGameTradesItems(), settings.getCustomNames());
        }

        if (!(settings.getInGameTradesMod() == Settings.InGameTradesMod.UNCHANGED)) {
            log.println("--In-Game Trades--");
            List<IngameTrade> newTrades = romHandler.getIngameTrades();
            int size = oldTrades.size();
            for (int i = 0; i < size; i++) {
                IngameTrade oldT = oldTrades.get(i);
                IngameTrade newT = newTrades.get(i);
                log.printf("Trading %s for %s the %s has become trading %s for %s the %s" + NEWLINE,
                        oldT.requestedPokemon.name, oldT.nickname, oldT.givenPokemon.name, newT.requestedPokemon.name,
                        newT.nickname, newT.givenPokemon.name);
            }
            log.println();
            workbookHandler.logToWorkbookRandomizedTrades(romHandler, oldTrades);
        }

        // Field Items
        List<Integer> oldItems = romHandler.getRegularFieldItems();
        List<Integer> oldTMs = romHandler.getCurrentFieldTMs();
        if (settings.getFieldItemsMod() == Settings.FieldItemsMod.SHUFFLE) {
            romHandler.shuffleFieldItems();
        } else if (settings.getFieldItemsMod() == Settings.FieldItemsMod.RANDOM) {
            romHandler.randomizeFieldItems(settings.isBanBadRandomFieldItems());
        }

        if(settings.getFieldItemsMod() != Settings.FieldItemsMod.UNCHANGED) {
            workbookHandler.logToWorkbookRandomizedItems(romHandler, oldItems, oldTMs);
        }

        // Signature...
        romHandler.applySignature();

        // Record check value?
        romHandler.writeCheckValueToROM(checkValue);

        // Save
        romHandler.saveRom(filename);

        // Log tail
        log.println("------------------------------------------------------------------");
        log.println("Randomization of " + romHandler.getROMName() + " completed.");
        log.println("Time elapsed: " + (System.currentTimeMillis() - startTime) + "ms");
        log.println("RNG Calls: " + RandomSource.callsSinceSeed());
        log.println("Seed: " + seed);
        log.println("Config string: " + Settings.VERSION + "" + settings.toString());
        log.println("------------------------------------------------------------------");

        return checkValue;
    }

    private void maybeLogBaseStatAndTypeChanges(final PrintStream log, final RomHandler romHandler) {
        List<Pokemon> allPokes = romHandler.getPokemon();
        String[] itemNames = romHandler.getItemNames();
        // Log base stats & types if changed at all
        if (settings.getBaseStatisticsMod() == Settings.BaseStatisticsMod.UNCHANGED
                && settings.getTypesMod() == Settings.TypesMod.UNCHANGED
                && settings.getAbilitiesMod() == Settings.AbilitiesMod.UNCHANGED
                && !settings.isRandomizeWildPokemonHeldItems()) {
            log.println("Pokemon base stats & type: unchanged" + NEWLINE);
        } else {
            log.println("--Pokemon Base Stats & Types--");
            if (romHandler instanceof Gen1RomHandler) {
                log.println("NUM|NAME      |TYPE             |  HP| ATK| DEF| SPE|SPEC");
                for (Pokemon pkmn : allPokes) {
                    if (pkmn != null) {
                        String typeString = pkmn.primaryType == null ? "???" : pkmn.primaryType.toString();
                        if (pkmn.secondaryType != null) {
                            typeString += "/" + pkmn.secondaryType.toString();
                        }
                        log.printf("%3d|%-10s|%-17s|%4d|%4d|%4d|%4d|%4d" + NEWLINE, pkmn.number, pkmn.name, typeString,
                                pkmn.hp, pkmn.attack, pkmn.defense, pkmn.speed, pkmn.special);
                    }
                }
            } else {
                log.print("NUM|NAME      |TYPE             |  HP| ATK| DEF| SPE|SATK|SDEF");
                int abils = romHandler.abilitiesPerPokemon();
                for (int i = 0; i < abils; i++) {
                    log.print("|ABILITY" + (i + 1) + "    ");
                }
                log.print("|ITEM");
                log.println();
                for (Pokemon pkmn : allPokes) {
                    if (pkmn != null) {
                        String typeString = pkmn.primaryType == null ? "???" : pkmn.primaryType.toString();
                        if (pkmn.secondaryType != null) {
                            typeString += "/" + pkmn.secondaryType.toString();
                        }
                        log.printf("%3d|%-10s|%-17s|%4d|%4d|%4d|%4d|%4d|%4d", pkmn.number, pkmn.name, typeString,
                                pkmn.hp, pkmn.attack, pkmn.defense, pkmn.speed, pkmn.spatk, pkmn.spdef);
                        if (abils > 0) {
                            log.printf("|%-12s|%-12s", romHandler.abilityName(pkmn.ability1),
                                    romHandler.abilityName(pkmn.ability2));
                            if (abils > 2) {
                                log.printf("|%-12s", romHandler.abilityName(pkmn.ability3));
                            }
                        }
                        log.print("|");
                        if (pkmn.guaranteedHeldItem > 0) {
                            log.print(itemNames[pkmn.guaranteedHeldItem] + " (100%)");
                        } else {
                            int itemCount = 0;
                            if (pkmn.commonHeldItem > 0) {
                                itemCount++;
                                log.print(itemNames[pkmn.commonHeldItem] + " (common)");
                            }
                            if (pkmn.rareHeldItem > 0) {
                                if (itemCount > 0) {
                                    log.print(", ");
                                }
                                itemCount++;
                                log.print(itemNames[pkmn.rareHeldItem] + " (rare)");
                            }
                            if (pkmn.darkGrassHeldItem > 0) {
                                if (itemCount > 0) {
                                    log.print(", ");
                                }
                                itemCount++;
                                log.print(itemNames[pkmn.darkGrassHeldItem] + " (dark grass only)");
                            }
                        }
                        log.println();
                    }
                }
            }
            log.println();
        }
    }

    private void logRandomizedEvolutions(PrintStream log, RomHandler romHandler) {
        log.println("--Randomized Evolutions--");
        List<Pokemon> allPokes = romHandler.getPokemon();
        for (Pokemon pk : allPokes) {
            if (pk != null) {
                int numEvos = pk.evolutionsFrom.size();
                if (numEvos > 0) {
                    StringBuilder evoStr = new StringBuilder(pk.evolutionsFrom.get(0).to.name);
                    for (int i = 1; i < numEvos; i++) {
                        if (i == numEvos - 1) {
                            evoStr.append(" and " + pk.evolutionsFrom.get(i).to.name);
                        } else {
                            evoStr.append(", " + pk.evolutionsFrom.get(i).to.name);
                        }
                    }
                    log.println(pk.name + " now evolves into " + evoStr.toString());
                }
            }
        }
        log.println();
    }

    private void maybeChangeAndLogStarters(final PrintStream log, final RomHandler romHandler) {
        if (romHandler.canChangeStarters()) {
            if (settings.getStartersMod() == Settings.StartersMod.CUSTOM) {
                log.println("--Custom Starters--");
                List<Pokemon> romPokemon = romHandler.getPokemon();
                int[] customStarters = settings.getCustomStarters();
                Pokemon pkmn1 = romPokemon.get(customStarters[0]);
                log.println("Set starter 1 to " + pkmn1.name);
                Pokemon pkmn2 = romPokemon.get(customStarters[1]);
                log.println("Set starter 2 to " + pkmn2.name);
                if (romHandler.isYellow()) {
                    romHandler.setStarters(Arrays.asList(pkmn1, pkmn2));
                } else {
                    Pokemon pkmn3 = romPokemon.get(customStarters[2]);
                    log.println("Set starter 3 to " + pkmn3.name);
                    romHandler.setStarters(Arrays.asList(pkmn1, pkmn2, pkmn3));
                }
                log.println();

            } else if (settings.getStartersMod() == Settings.StartersMod.COMPLETELY_RANDOM) {
                // Randomise
                log.println("--Random Starters--");
                int starterCount = 3;
                if (romHandler.isYellow()) {
                    starterCount = 2;
                }
                List<Pokemon> starters = new ArrayList<Pokemon>();
                for (int i = 0; i < starterCount; i++) {
                    Pokemon pkmn = romHandler.randomPokemon();
                    while (starters.contains(pkmn)) {
                        pkmn = romHandler.randomPokemon();
                    }
                    log.println("Set starter " + (i + 1) + " to " + pkmn.name);
                    starters.add(pkmn);
                }
                romHandler.setStarters(starters);
                log.println();
                
            } else if (settings.getStartersMod() == Settings.StartersMod.RANDOM_WITH_ONE_OR_TWO_EVOLUTIONS) {
                // Randomise
                log.println("--Random 1/2-Evolution Starters--");
                int starterCount = 3;
                if (romHandler.isYellow()) {
                    starterCount = 2;
                }
                List<Pokemon> starters = new ArrayList<Pokemon>();
                for (int i = 0; i < starterCount; i++) {
                    Pokemon pkmn = romHandler.random1or2EvosPokemon();
                    while (starters.contains(pkmn)) {
                        pkmn = romHandler.random1or2EvosPokemon();
                    }
                    log.println("Set starter " + (i + 1) + " to " + pkmn.name);
                    starters.add(pkmn);
                }
                romHandler.setStarters(starters);
                log.println();
            } else if (settings.getStartersMod() == Settings.StartersMod.RANDOM_WITH_TWO_EVOLUTIONS) {
                // Randomise
                log.println("--Random 2-Evolution Starters--");
                int starterCount = 3;
                if (romHandler.isYellow()) {
                    starterCount = 2;
                }
                List<Pokemon> starters = new ArrayList<Pokemon>();
                for (int i = 0; i < starterCount; i++) {
                    Pokemon pkmn = romHandler.random2EvosPokemon();
                    while (starters.contains(pkmn)) {
                        pkmn = romHandler.random2EvosPokemon();
                    }
                    log.println("Set starter " + (i + 1) + " to " + pkmn.name);
                    starters.add(pkmn);
                }
                romHandler.setStarters(starters);
                log.println();
            }
            if (settings.isRandomizeStartersHeldItems() && !(romHandler instanceof Gen1RomHandler)) {
                romHandler.randomizeStarterHeldItems(settings.isBanBadRandomStarterHeldItems());
            }
        }
    }

    private void maybeLogMovesetChanges(PrintStream log, RomHandler romHandler, boolean forceFourLv1s) {
        // Show the new movesets if applicable
        List<Move> moves = romHandler.getMoves();
        if (settings.getMovesetsMod() == Settings.MovesetsMod.UNCHANGED) {
            if(!settings.doBlockBrokenMoves() && !forceFourLv1s ) {
                log.println("Pokemon Movesets: Unchanged." + NEWLINE);
            }
            else
            {
                if (settings.doBlockBrokenMoves()) {
                    log.print("Pokemon Movesets: Removed Game-Breaking Moves (");
                    List<Integer> gameBreakingMoves = romHandler.getGameBreakingMoves();
                    int numberPrinted = 0;
                    for (Move move : moves) {
                        if (move == null) {
                            continue;
                        }
                        if (gameBreakingMoves.contains(move.number)) {
                            numberPrinted++;
                            log.print(move.name);
                            if (numberPrinted < gameBreakingMoves.size()) {
                                log.print(", ");
                            }
                        }
                    }
                    log.println(")" + NEWLINE);
                }

                if( forceFourLv1s ) {
                    log.println("--Pokemon Movesets--");
                    List<String> movesets = new ArrayList<String>();
                    Map<Pokemon, List<MoveLearnt>> moveData = romHandler.getMovesLearnt();
                    for (Pokemon pkmn : moveData.keySet()) {
                        StringBuilder sb = new StringBuilder();
                        sb.append(String.format("%03d %-10s : ", pkmn.number, pkmn.name));
                        List<MoveLearnt> data = moveData.get(pkmn);
                        boolean first = true;
                        for (MoveLearnt ml : data) {
                            if (!first) {
                                sb.append(", ");
                            }
                            try {
                                sb.append(moves.get(ml.move).name).append(" at level ").append(ml.level);
                            } catch (NullPointerException ex) {
                                sb.append("invalid move at level" + ml.level);
                            }
                            first = false;
                        }
                        movesets.add(sb.toString());
                    }
                    Collections.sort(movesets);
                    for (String moveset : movesets) {
                        log.println(moveset);
                    }
                    log.println();
                }
            }
        } else if (settings.getMovesetsMod() == Settings.MovesetsMod.METRONOME_ONLY) {
            log.println("Pokemon Movesets: Metronome Only." + NEWLINE);
        } else {
            log.println("--Pokemon Movesets--");
            List<String> movesets = new ArrayList<String>();
            Map<Pokemon, List<MoveLearnt>> moveData = romHandler.getMovesLearnt();
            for (Pokemon pkmn : moveData.keySet()) {
                StringBuilder sb = new StringBuilder();
                sb.append(String.format("%03d %-10s : ", pkmn.number, pkmn.name));
                List<MoveLearnt> data = moveData.get(pkmn);
                boolean first = true;
                for (MoveLearnt ml : data) {
                    if (!first) {
                        sb.append(", ");
                    }
                    try {
                        sb.append(moves.get(ml.move).name).append(" at level ").append(ml.level);
                    } catch (NullPointerException ex) {
                        sb.append("invalid move at level" + ml.level);
                    }
                    first = false;
                }
                movesets.add(sb.toString());
            }
            Collections.sort(movesets);
            for (String moveset : movesets) {
                log.println(moveset);
            }
            log.println();
        }
    }

    private void maybeLogWildPokemonChanges(final PrintStream log, final RomHandler romHandler) {
        if (settings.getWildPokemonMod() == Settings.WildPokemonMod.UNCHANGED && !settings.isWildLevelModifiedHigh() ) {
            log.println("Wild Pokemon: Unchanged." + NEWLINE);
        } else {
            log.println("--Wild Pokemon--");
            List<EncounterSet> encounters = romHandler.getEncounters(settings.isUseTimeBasedEncounters());
            int idx = 0;
            for (EncounterSet es : encounters) {
                //skip unused EncounterSets in DPPT
                if(romHandler instanceof Gen4RomHandler) {
                    if(es.displayName.contains("? Unknown ?")) {
                        continue;
                    }
                }
                idx++;
                log.print("Set #" + idx + " ");
                if (es.displayName != null) {
                    log.print("- " + es.displayName + " ");
                }
                log.print("(rate=" + es.rate + ")");
                log.print(" - ");
                for (int i = 0; i < es.encounters.size(); i++) {
                    Encounter e = es.encounters.get(i);
                    if (i > 0) {
                        log.print(", ");
                    }
                    if(romHandler instanceof Gen4RomHandler && es.displayName.contains("Swarm/Radar/GBA")) {
                        if(i == 0) {
                            log.print("Swarm: ");
                        } else if(i == 2) {
                            log.print("PokeRadar: ");
                        } else if(i == 6) {
                            log.print("Ruby GBA: ");
                        } else if(i == 8) {
                            log.print("Sapphire GBA: ");
                        } else if(i == 10) {
                            log.print("Emerald GBA: ");
                        } else if(i == 12) {
                            log.print("Fire Red GBA: ");
                        } else if(i == 14) {
                            log.print("Leaf Green GBA: ");
                        }
                    }
                    log.print(e.pokemon.name + " Lv");
                    if (e.maxLevel > 0 && e.maxLevel != e.level) {
                        log.print("s " + e.level + "-" + e.maxLevel);
                    } else {
                        log.print(e.level);
                    }
                }
                log.println();
            }
            log.println();
        }
    }

    private void maybeLogTrainerChanges(final PrintStream log, final RomHandler romHandler) {
        if (settings.getTrainersMod() == Settings.TrainersMod.UNCHANGED && !settings.isRivalCarriesStarterThroughout() && !settings.isTrainersLevelModified()) {
            log.println("Trainers: Unchanged." + NEWLINE);
        } else {
            log.println("--Trainers Pokemon--");
            List<Trainer> trainers = romHandler.getTrainers();
            int idx = 0;
            for (Trainer t : trainers) {
                idx++;
                log.print("#" + idx + " ");
                if (t.fullDisplayName != null) {
                    log.print("(" + t.fullDisplayName + ")");
                } else if (t.name != null) {
                    log.print("(" + t.name + ")");
                }
                if (t.offset != idx && t.offset != 0) {
                    log.printf("@%X", t.offset);
                }
                log.print(" - ");
                boolean first = true;
                for (TrainerPokemon tpk : t.pokemon) {
                    if (!first) {
                        log.print(", ");
                    }
                    log.print(tpk.pokemon.name + " Lv" + tpk.level + "(Ivs: " + (int)Math.floor((tpk.difficulty*31)/255.0D) + ")");
                    first = false;
                }
                log.println();
            }
            log.println();
        }
    }

    private int maybeChangeAndLogStaticPokemon(final PrintStream log, final RomHandler romHandler, boolean raceMode,
            int checkValue) {
        if (romHandler.canChangeStaticPokemon()) {
            List<Pokemon> oldStatics = romHandler.getStaticPokemon();
            if (settings.getStaticPokemonMod() == Settings.StaticPokemonMod.RANDOM_MATCHING) {
                romHandler.randomizeStaticPokemon(true);
            } else if (settings.getStaticPokemonMod() == Settings.StaticPokemonMod.COMPLETELY_RANDOM) {
                romHandler.randomizeStaticPokemon(false);
            }
            List<Pokemon> newStatics = romHandler.getStaticPokemon();
            if (settings.getStaticPokemonMod() == Settings.StaticPokemonMod.UNCHANGED) {
                log.println("Static Pokemon: Unchanged." + NEWLINE);
            } else {
                log.println("--Static Pokemon--");
                Map<Pokemon, Integer> seenPokemon = new TreeMap<Pokemon, Integer>();
                for (int i = 0; i < oldStatics.size(); i++) {
                    Pokemon oldP = oldStatics.get(i);
                    Pokemon newP = newStatics.get(i);
                    checkValue = addToCV(checkValue, newP.number);
                    log.print(oldP.name);
                    if (seenPokemon.containsKey(oldP)) {
                        int amount = seenPokemon.get(oldP);
                        log.print("(" + (++amount) + ")");
                        seenPokemon.put(oldP, amount);
                    } else {
                        seenPokemon.put(oldP, 1);
                    }
                    log.println(" => " + newP.name);
                }
                log.println();
            }
        }
        return checkValue;
    }

    private void maybeLogMoveChanges(final PrintStream log, final RomHandler romHandler) {
        if (!settings.isRandomizeMoveAccuracies() && !settings.isRandomizeMovePowers()
                && !settings.isRandomizeMovePPs() && !settings.isRandomizeMoveCategory()
                && !settings.isRandomizeMoveTypes()) {
            if (!settings.isUpdateMoves()) {
                log.println("Move Data: Unchanged." + NEWLINE);
            }
        } else {
            log.println("--Move Data--");
            log.print("NUM|NAME           |TYPE    |POWER|ACC.|PP");
            if (romHandler.hasPhysicalSpecialSplit()) {
                log.print(" |CATEGORY");
            }
            log.println();
            List<Move> allMoves = romHandler.getMoves();
            for (Move mv : allMoves) {
                if (mv != null) {
                    String mvType = (mv.type == null) ? "???" : mv.type.toString();
                    log.printf("%3d|%-15s|%-8s|%5d|%4d|%3d", mv.internalId, mv.name, mvType, mv.power,
                            (int) mv.hitratio, mv.pp);
                    if (romHandler.hasPhysicalSpecialSplit()) {
                        log.printf("| %s", mv.category.toString());
                    }
                    log.println();
                }
            }
            log.println();
        }
    }

    private static int addToCV(int checkValue, int... values) {
        for (int value : values) {
            checkValue = Integer.rotateLeft(checkValue, 3);
            checkValue ^= value;
        }
        return checkValue;
    }
}