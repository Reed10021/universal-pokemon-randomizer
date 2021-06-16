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
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

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

    public int randomize(final String filename) {
        return randomize(filename, new PrintStream(new OutputStream() {
            @Override
            public void write(int b) {
            }
        }));
    }

    public int randomize(final String filename, final PrintStream log) {
        long seed = RandomSource.pickSeed();
        return randomize(filename, log, seed);
    }

    public int randomize(final String filename, final PrintStream log, long seed) {
        Workbook wb = new XSSFWorkbook();
        wb.createSheet("Stats, Type, Ability, Item");
        wb.createSheet("Evolutions");
        wb.createSheet("Starters, Static, Trades");
        wb.createSheet("Moves, TMs, Move Tutors");
        wb.createSheet("Poke Movesets");
        wb.createSheet("Poke TMHM Compat");
        wb.createSheet("Poke Move Tutor Compat");
        wb.createSheet("Wild Pokemon");
        wb.createSheet("Trainers");
        wb.createSheet("Randomized Items");

        final long startTime = System.currentTimeMillis();
        RandomSource.seed(seed);
        final boolean raceMode = settings.isRaceMode();

        int checkValue = 0;

        // limit pokemon?
        if (settings.isLimitPokemon()) {
            romHandler.setPokemonPool(settings.getCurrentRestrictions());
            romHandler.removeEvosForPokemonPool();
        } else {
            romHandler.setPokemonPool(null);
        }

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

        maybeLogBaseStatAndTypeChanges(wb, log, romHandler);
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

            // Only output evolutions to workbook once. So if we're not done making changes, don't log it.
            if(settings.isChangeImpossibleEvolutions() || settings.isMakeEvolutionsEasier()) {
                logRandomizedEvolutions(log, romHandler);
            } else {
                logRandomizedEvolutions(log, romHandler);
                logToWorkbookRandomizedEvolutions(wb, romHandler);
            }
        }

        // Trade evolutions removal
        if (settings.isChangeImpossibleEvolutions()) {
            romHandler.removeTradeEvolutions(!(settings.getMovesetsMod() == Settings.MovesetsMod.UNCHANGED));
            // Again, if we're not done making changes to evolutions yet, keep going and don't log it to the workbook
            if(!settings.isMakeEvolutionsEasier()) {
                logToWorkbookRandomizedEvolutions(wb, romHandler);
            }
        }

        // Easier evolutions
        if (settings.isMakeEvolutionsEasier()) {
            romHandler.condenseLevelEvolutions(40, 30);
            logToWorkbookRandomizedEvolutions(wb, romHandler);
        }

        // Starter Pokemon
        // Applied after type to update the strings correctly based on new types
        List<Pokemon> oldStarters = romHandler.getStarters();
        maybeChangeAndLogStarters(log, romHandler);
        // If starters changed, log it to the workbook
        if(!oldStarters.containsAll(romHandler.getStarters())) {
            logToWorkbookStarters(wb, romHandler, oldStarters);
        }

        // Move Data Log
        // Placed here so it matches its position in the randomizer interface
        maybeLogMoveChanges(log, romHandler);
        maybeLogToWorkbookMoveChanges(wb, romHandler);

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
        maybeLogToWorkbookMovesetChanges(wb, romHandler, forceFourLv1s);
        maybeLogTrainerChanges(log, romHandler);
        maybeLogToWorkbookTrainerChanges(wb,romHandler);

        // Static Pokemon
        List<Pokemon> oldStatics = romHandler.getStaticPokemon();
        checkValue = maybeChangeAndLogStaticPokemon(log, romHandler, raceMode, checkValue);
        if(!oldStatics.containsAll(romHandler.getStaticPokemon())) {
            logToWorkbookStaticPokemon(wb, romHandler, oldStatics);
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
        maybeLogToWorkbookWildPokemonChanges(wb, romHandler);

        List<EncounterSet> encounters = romHandler.getEncounters(settings.isUseTimeBasedEncounters());
        for (EncounterSet es : encounters) {
            for (Encounter e : es.encounters) {
                checkValue = addToCV(checkValue, e.level, e.pokemon.number);
            }
        }

        // TMs
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
                List<Integer> mts = romHandler.getMoveTutorMoves();
                List<Move> moveData = romHandler.getMoves();
                List<String> movesets = new ArrayList<String>();
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
        }

        // Field Items
        if (settings.getFieldItemsMod() == Settings.FieldItemsMod.SHUFFLE) {
            romHandler.shuffleFieldItems();
        } else if (settings.getFieldItemsMod() == Settings.FieldItemsMod.RANDOM) {
            romHandler.randomizeFieldItems(settings.isBanBadRandomFieldItems());
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

        try(OutputStream fileOut = new FileOutputStream(filename+".xlsx")) {
            wb.write(fileOut);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return checkValue;
    }



    private void maybeLogBaseStatAndTypeChanges(final Workbook wb, final PrintStream log, final RomHandler romHandler) {
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
            logToWorkbookBaseStatAndTypeChanges(wb, romHandler);
        }
    }

    private void logToWorkbookBaseStatAndTypeChanges(Workbook wb, RomHandler romHandler) {
        List<Pokemon> allPokes = romHandler.getPokemon();
        String[] itemNames = romHandler.getItemNames();
        int rowCounter = 0;
        int cellCounter = 0;
        Sheet poke = wb.getSheetAt(0);
        poke.createFreezePane(0,1);
        Row rowOne = poke.createRow(rowCounter++);
        // Handle Gen 1 since it's special.
        if (romHandler instanceof Gen1RomHandler) {
            rowOne.createCell(cellCounter++).setCellValue("NUM");
            rowOne.createCell(cellCounter++).setCellValue("NAME");
            rowOne.createCell(cellCounter++).setCellValue("TYPE");
            rowOne.createCell(cellCounter++).setCellValue("HP");
            rowOne.createCell(cellCounter++).setCellValue("ATK");
            rowOne.createCell(cellCounter++).setCellValue("DEF");
            rowOne.createCell(cellCounter++).setCellValue("SPD");
            rowOne.createCell(cellCounter++).setCellValue("SPEC");
            rowOne.createCell(cellCounter++).setCellValue("BST");
            CellStyle centerCells = wb.createCellStyle();
            centerCells.setAlignment(HorizontalAlignment.CENTER);
            for(int i = 0; i < cellCounter; i++) {
                rowOne.getCell(i).setCellStyle(centerCells);
            }

            for(Pokemon pkmn : allPokes) {
                if(pkmn != null) {
                    String typeString = pkmn.primaryType == null ? "???" : pkmn.primaryType.toString();
                    if (pkmn.secondaryType != null) {
                        typeString += "/" + pkmn.secondaryType.toString();
                    }
                    Row temp = poke.createRow(rowCounter++);
                    temp.createCell(0).setCellValue(pkmn.number);
                    temp.createCell(1).setCellValue(pkmn.name);
                    temp.createCell(2).setCellValue(typeString);
                    temp.createCell(3).setCellValue(pkmn.hp);
                    temp.createCell(4).setCellValue(pkmn.attack);
                    temp.createCell(5).setCellValue(pkmn.defense);
                    temp.createCell(6).setCellValue(pkmn.speed);
                    temp.createCell(7).setCellValue(pkmn.special);
                    temp.createCell(8).setCellValue(pkmn.hp + pkmn.attack + pkmn.defense + pkmn.speed + pkmn.special);
                }
            }
            for (int i = 0; i < 10; i++) {
                poke.autoSizeColumn(i);
            }
            // This is annoying. You cannot do CellRangeAddress.valueOf("$D:$D") to apply it to the whole column, so this is the workaround.
            SheetConditionalFormatting sheetConditionalFormatting = poke.getSheetConditionalFormatting();
            CellRangeAddress [] hpRegion = { CellRangeAddress.valueOf("D2:D" + poke.getLastRowNum()+1) };
            CellRangeAddress [] atkRegion = { CellRangeAddress.valueOf("E2:E" + poke.getLastRowNum()+1) };
            CellRangeAddress [] defRegion = { CellRangeAddress.valueOf("F2:F" + poke.getLastRowNum()+1) };
            CellRangeAddress [] spdRegion = { CellRangeAddress.valueOf("G2:G" + poke.getLastRowNum()+1) };
            CellRangeAddress [] specialRegion = { CellRangeAddress.valueOf("H2:H" + poke.getLastRowNum()+1) };
            CellRangeAddress [] bstRegion = { CellRangeAddress.valueOf("I2:I" + poke.getLastRowNum()+1) };

            ConditionalFormattingRule formattingRule = sheetConditionalFormatting.createConditionalFormattingColorScaleRule();
            ColorScaleFormatting colorScaleFormatting = formattingRule.getColorScaleFormatting();
            Color [] colors = {new XSSFColor(new java.awt.Color(248, 105, 107)), new XSSFColor(new java.awt.Color(255, 235, 132)), new XSSFColor(new java.awt.Color(99, 190, 123))};
            colorScaleFormatting.setColors(colors);
            colorScaleFormatting.getThresholds()[0].setRangeType(ConditionalFormattingThreshold.RangeType.MIN);
            colorScaleFormatting.getThresholds()[1].setRangeType(ConditionalFormattingThreshold.RangeType.PERCENTILE);
            colorScaleFormatting.getThresholds()[1].setValue(50d);
            colorScaleFormatting.getThresholds()[2].setRangeType(ConditionalFormattingThreshold.RangeType.MAX);

            sheetConditionalFormatting.addConditionalFormatting(hpRegion, formattingRule);
            sheetConditionalFormatting.addConditionalFormatting(atkRegion, formattingRule);
            sheetConditionalFormatting.addConditionalFormatting(defRegion, formattingRule);
            sheetConditionalFormatting.addConditionalFormatting(spdRegion, formattingRule);
            sheetConditionalFormatting.addConditionalFormatting(specialRegion, formattingRule);
            sheetConditionalFormatting.addConditionalFormatting(bstRegion, formattingRule);

        } else {
            // Handle the rest of the generations of pokemon.
            int abils = romHandler.abilitiesPerPokemon();
            rowOne.createCell(cellCounter++).setCellValue("NUM");
            rowOne.createCell(cellCounter++).setCellValue("NAME");
            rowOne.createCell(cellCounter++).setCellValue("TYPE");
            rowOne.createCell(cellCounter++).setCellValue("HP");
            rowOne.createCell(cellCounter++).setCellValue("ATK");
            rowOne.createCell(cellCounter++).setCellValue("DEF");
            rowOne.createCell(cellCounter++).setCellValue("SPD");
            rowOne.createCell(cellCounter++).setCellValue("SPATK");
            rowOne.createCell(cellCounter++).setCellValue("SPDEF");
            rowOne.createCell(cellCounter++).setCellValue("BST");
            for (int i = 0; i < abils; i++) {
                rowOne.createCell(cellCounter++).setCellValue("ABILITY " + (i+1));
            }
            rowOne.createCell(cellCounter++).setCellValue("ITEM 1");
            rowOne.createCell(cellCounter++).setCellValue("ITEM 2");
            rowOne.createCell(cellCounter++).setCellValue("ITEM 3");

            CellStyle centerCells = wb.createCellStyle();
            centerCells.setAlignment(HorizontalAlignment.CENTER);
            for(int i = 0; i < cellCounter; i++) {
                rowOne.getCell(i).setCellStyle(centerCells);
            }

            for (Pokemon pkmn : allPokes) {
                if (pkmn != null) {
                    String typeString = pkmn.primaryType == null ? "???" : pkmn.primaryType.toString();
                    if (pkmn.secondaryType != null) {
                        typeString += "/" + pkmn.secondaryType.toString();
                    }
                    Row temp = poke.createRow(rowCounter++);
                    int tempCellCounter = 0;
                    temp.createCell(tempCellCounter++).setCellValue(pkmn.number);
                    temp.createCell(tempCellCounter++).setCellValue(pkmn.name);
                    temp.createCell(tempCellCounter++).setCellValue(typeString);
                    temp.createCell(tempCellCounter++).setCellValue(pkmn.hp);
                    temp.createCell(tempCellCounter++).setCellValue(pkmn.attack);
                    temp.createCell(tempCellCounter++).setCellValue(pkmn.defense);
                    temp.createCell(tempCellCounter++).setCellValue(pkmn.speed);
                    temp.createCell(tempCellCounter++).setCellValue(pkmn.spatk);
                    temp.createCell(tempCellCounter++).setCellValue(pkmn.spdef);
                    temp.createCell(tempCellCounter++).setCellValue(pkmn.hp + pkmn.attack + pkmn.defense + pkmn.speed + pkmn.spatk + pkmn.spdef);

                    if (abils > 0) {
                        temp.createCell(tempCellCounter++).setCellValue(romHandler.abilityName(pkmn.ability1));
                        temp.createCell(tempCellCounter++).setCellValue(romHandler.abilityName(pkmn.ability2));
                        if (abils > 2) {
                            temp.createCell(tempCellCounter++).setCellValue(romHandler.abilityName(pkmn.ability3));
                        }
                    }

                    if (pkmn.guaranteedHeldItem > 0) {
                        temp.createCell(tempCellCounter++).setCellValue(itemNames[pkmn.guaranteedHeldItem] + " (100%)");
                    } else {
                        if (pkmn.commonHeldItem > 0) {
                            temp.createCell(tempCellCounter++).setCellValue(itemNames[pkmn.commonHeldItem] + " (common)");
                        }
                        if (pkmn.rareHeldItem > 0) {
                            temp.createCell(tempCellCounter++).setCellValue(itemNames[pkmn.rareHeldItem] + " (rare)");
                        }
                        if (pkmn.darkGrassHeldItem > 0) {
                            temp.createCell(tempCellCounter++).setCellValue(itemNames[pkmn.darkGrassHeldItem] + " (dark grass only)");
                        }
                    }
                }
            }
            for (int i = 0; i < 14; i++) {
                poke.autoSizeColumn(i);
            }
            SheetConditionalFormatting sheetConditionalFormatting = poke.getSheetConditionalFormatting();
            CellRangeAddress [] hpRegion = { CellRangeAddress.valueOf("D2:D" + poke.getLastRowNum()+1) };
            CellRangeAddress [] atkRegion = { CellRangeAddress.valueOf("E2:E" + poke.getLastRowNum()+1) };
            CellRangeAddress [] defRegion = { CellRangeAddress.valueOf("F2:F" + poke.getLastRowNum()+1) };
            CellRangeAddress [] spdRegion = { CellRangeAddress.valueOf("G2:G" + poke.getLastRowNum()+1) };
            CellRangeAddress [] spatkRegion = { CellRangeAddress.valueOf("H2:H" + poke.getLastRowNum()+1) };
            CellRangeAddress [] spdefRegion = { CellRangeAddress.valueOf("I2:I" + poke.getLastRowNum()+1) };
            CellRangeAddress [] bstRegion = { CellRangeAddress.valueOf("J2:J" + poke.getLastRowNum()+1) };

            ConditionalFormattingRule formattingRule = sheetConditionalFormatting.createConditionalFormattingColorScaleRule();
            ColorScaleFormatting colorScaleFormatting = formattingRule.getColorScaleFormatting();
            Color [] colors = {new XSSFColor(new java.awt.Color(248, 105, 107)), new XSSFColor(new java.awt.Color(255, 235, 132)), new XSSFColor(new java.awt.Color(99, 190, 123))};
            colorScaleFormatting.setColors(colors);
            colorScaleFormatting.getThresholds()[0].setRangeType(ConditionalFormattingThreshold.RangeType.MIN);
            colorScaleFormatting.getThresholds()[1].setRangeType(ConditionalFormattingThreshold.RangeType.PERCENTILE);
            colorScaleFormatting.getThresholds()[1].setValue(50d);
            colorScaleFormatting.getThresholds()[2].setRangeType(ConditionalFormattingThreshold.RangeType.MAX);

            sheetConditionalFormatting.addConditionalFormatting(hpRegion, formattingRule);
            sheetConditionalFormatting.addConditionalFormatting(atkRegion, formattingRule);
            sheetConditionalFormatting.addConditionalFormatting(defRegion, formattingRule);
            sheetConditionalFormatting.addConditionalFormatting(spdRegion, formattingRule);
            sheetConditionalFormatting.addConditionalFormatting(spatkRegion, formattingRule);
            sheetConditionalFormatting.addConditionalFormatting(spdefRegion, formattingRule);
            sheetConditionalFormatting.addConditionalFormatting(bstRegion, formattingRule);
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

    private void logToWorkbookRandomizedEvolutions(Workbook wb, RomHandler romHandler) {
        int rowCounter = 0;
        Sheet evos = wb.getSheetAt(1);

        List<Pokemon> allPokes = romHandler.getPokemon();
        String[] itemNames = romHandler.getItemNames();
        List<Move> moves = romHandler.getMoves();
        for (Pokemon pk : allPokes) {
            if (pk != null) {
                int numEvos = pk.evolutionsFrom.size();
                if (numEvos > 0) {
                    int tempCellCounter = 0;
                    Row temp = evos.createRow(rowCounter++);
                    temp.createCell(tempCellCounter++).setCellValue(pk.evolutionsFrom.get(0).from.name);
                    temp.createCell(tempCellCounter++).setCellValue("now evolves into");

                    StringBuilder evoStr = new StringBuilder();
                    StringBuilder evoTypeStr = new StringBuilder();
                    for (int i = 0; i < numEvos; i++) {
                        EvolutionType evoType = pk.evolutionsFrom.get(i).type;
                        evoTypeStr.append(evoType.toString());
                        if(evoType.usesLevel()) {
                            evoTypeStr.append(pk.evolutionsFrom.get(i).extraInfo);
                        } else if(evoType.isItem()) {
                            evoTypeStr.append(itemNames[pk.evolutionsFrom.get(i).extraInfo]);
                        } else if(evoType == EvolutionType.LEVEL_WITH_MOVE) {
                            for(Move move : moves) {
                                if(move == null)
                                    continue;
                                if(move.number == pk.evolutionsFrom.get(i).extraInfo) {
                                    evoTypeStr.append(move.name);
                                }
                            }
                        } else if(evoType == EvolutionType.LEVEL_WITH_OTHER || evoType == EvolutionType.TRADE_SPECIAL) {
                            evoTypeStr.append(allPokes.get(pk.evolutionsFrom.get(i).extraInfo).name);
                        }
                        evoStr.append( pk.evolutionsFrom.get(i).to.name);

                        if (i == numEvos-2) {
                            evoStr.append(" and ");
                            evoTypeStr.append("; ");
                        } else if( i < numEvos-1) {
                            evoStr.append(", ");
                            evoTypeStr.append("; ");
                        }
                    }
                    temp.createCell(tempCellCounter++).setCellValue(evoStr.toString());

                    temp.createCell(tempCellCounter++).setCellValue("by way of (respectively): " + evoTypeStr.toString());
                }
            }
        }
        for (int i = 0; i < 4; i++) {
            evos.autoSizeColumn(i);
        }
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

    private void logToWorkbookStarters(Workbook wb, RomHandler romHandler, List<Pokemon> oldStarters) {
        List<Pokemon> newStarters = romHandler.getStarters();
        int rowCounter = 0;
        int cellCounter = 0;
        Sheet starters = wb.getSheetAt(2);
        starters.createFreezePane(0,1);
        Row rowOne = starters.createRow(rowCounter++);

        rowOne.createCell(cellCounter++).setCellValue("STARTERS");
        rowOne.createCell(cellCounter++);
        rowOne.createCell(cellCounter++);
        starters.addMergedRegion(new CellRangeAddress(0, 0, 0, 2));
        CellStyle centerCells = wb.createCellStyle();
        centerCells.setAlignment(HorizontalAlignment.CENTER);
        rowOne.getCell(0).setCellStyle(centerCells);

        int oldStarterIndex = 0;
        for(Pokemon starter : newStarters) {
            int tempCellCounter = 0;
            Row tempRow = starters.createRow(rowCounter++);
            tempRow.createCell(tempCellCounter++).setCellValue(oldStarters.get(oldStarterIndex).name);
            tempRow.createCell(tempCellCounter++).setCellValue("changed to");
            tempRow.createCell(tempCellCounter).setCellValue(starter.name);

            oldStarterIndex++;
        }
        for (int i = 0; i < cellCounter; i++) {
            starters.autoSizeColumn(i);
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

    private void maybeLogToWorkbookMovesetChanges(Workbook wb, RomHandler romHandler, boolean forceFourLv1s) {
        if (settings.getMovesetsMod() != Settings.MovesetsMod.UNCHANGED || settings.doBlockBrokenMoves() || forceFourLv1s) {
            List<Move> moves = romHandler.getMoves();
            List<String> movesets = new ArrayList<String>();
            Map<Pokemon, List<MoveLearnt>> moveData = romHandler.getMovesLearnt();
            int cellCounter = 0;
            int rowCounter = 0;
            Sheet sheetMoveset = wb.getSheetAt(4);
            sheetMoveset.createFreezePane(2,0);
            Row rowOne = sheetMoveset.createRow(rowCounter++);
            rowOne.createCell(cellCounter++).setCellValue("NUM");
            rowOne.createCell(cellCounter++).setCellValue("NAME");

            CellStyle centerCells = wb.createCellStyle();
            centerCells.setAlignment(HorizontalAlignment.CENTER);
            for(int i = 0; i < cellCounter; i++) {
                rowOne.getCell(i).setCellStyle(centerCells);
            }

            for (Pokemon pkmn : moveData.keySet()) {
                int tempCellCounter = 0;
                Row tempRow = sheetMoveset.createRow(rowCounter++);
                tempRow.createCell(tempCellCounter++).setCellValue(pkmn.number);
                tempRow.createCell(tempCellCounter++).setCellValue(pkmn.name);

                List<MoveLearnt> data = moveData.get(pkmn);
                for (MoveLearnt ml : data) {
                    if(ml != null && moves.get(ml.move) != null) {
                        tempRow.createCell(tempCellCounter++).setCellValue(moves.get(ml.move).name + " at level " + ml.level);
                    } else {
                        if(ml != null) {
                            System.out.println("BAD MOVE ALERT: " + ml.move + " LV: " + ml.level + " FOR PKMN " + pkmn.name);
                        } else {
                            System.out.println("BAD MOVE ALERT: NULLPTR FOR PKMN " + pkmn.name);
                        }
                    }
                }

                if(tempCellCounter > cellCounter) {
                    cellCounter = tempCellCounter;
                }
            }

            for (int i = 0; i < cellCounter; i++) {
                sheetMoveset.autoSizeColumn(i);
            }
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

    private void maybeLogToWorkbookWildPokemonChanges(Workbook wb, RomHandler romHandler) {
        if(settings.getWildPokemonMod() != Settings.WildPokemonMod.UNCHANGED || settings.isWildLevelModifiedHigh()) {
            int rowCounter = 0;
            int cellCounter = 0;
            Sheet sheetTrainers = wb.getSheetAt(7);
            sheetTrainers.createFreezePane(2,0);
            Row rowOne = sheetTrainers.createRow(rowCounter++);
            rowOne.createCell(cellCounter++).setCellValue("NUM");
            rowOne.createCell(cellCounter++).setCellValue("LOCATION");

            CellStyle centerCells = wb.createCellStyle();
            centerCells.setAlignment(HorizontalAlignment.CENTER);
            for(int i = 0; i < cellCounter; i++) {
                rowOne.getCell(i).setCellStyle(centerCells);
            }

            List<EncounterSet> encounters = romHandler.getEncounters(settings.isUseTimeBasedEncounters());
            int idx = 0;
            for (EncounterSet es : encounters) {
                //skip unused EncounterSets in DPPT
                if(romHandler instanceof Gen4RomHandler) {
                    if(es.displayName.contains("? Unknown ?")) {
                        continue;
                    } else if(es.displayName.contains("Swarm/Radar/GBA")) {
                        // Skip swarm stuff until the end.
                        continue;
                    }
                }
                idx++;
                int tempCellCounter = 0;
                Row tempRow = sheetTrainers.createRow(rowCounter++);
                tempRow.createCell(tempCellCounter++).setCellValue("Set " + idx + " (Rate: " + es.rate + ")");
                if (es.displayName != null) {
                    tempRow.createCell(tempCellCounter++).setCellValue(es.displayName);
                }

                for (int i = 0; i < es.encounters.size(); i++) {
                    Encounter e = es.encounters.get(i);
                    String pkmnStr = e.pokemon.name + " Lv";
                    if (e.maxLevel > 0 && e.maxLevel != e.level) {
                        pkmnStr += "s " + e.level + "-" + e.maxLevel;
                    } else {
                        pkmnStr+= e.level;
                    }
                    tempRow.createCell(tempCellCounter++).setCellValue(pkmnStr);
                    if(tempCellCounter > cellCounter) {
                        cellCounter = tempCellCounter;
                    }
                }
            }

            // Now do swarm stuff.
            if(romHandler instanceof Gen4RomHandler) {
                for (EncounterSet es : encounters) {
                    //skip unused EncounterSets in DPPT
                    if (es.displayName.contains("? Unknown ?")) {
                        continue;
                    } else if (!es.displayName.contains("Swarm/Radar/GBA")) {
                        continue;
                    }
                    idx++;
                    int tempCellCounter = 0;
                    Row tempRow = sheetTrainers.createRow(rowCounter++);
                    tempRow.createCell(tempCellCounter++).setCellValue("Set " + idx + " (Rate: " + es.rate + ")");
                    if (es.displayName != null) {
                        tempRow.createCell(tempCellCounter++).setCellValue(es.displayName.replace("Swarm/Radar/GBA", "Swarm"));
                    }

                    for (int i = 0; i < es.encounters.size(); i++) {
                        if (es.displayName.contains("Swarm/Radar/GBA")) {
                            // i == 0 is Swarm; already taken care of.
                            if (i == 2) {
                                tempCellCounter = 0;
                                tempRow = sheetTrainers.createRow(rowCounter++);
                                tempRow.createCell(tempCellCounter++).setCellValue("Set " + idx + " (Rate: " + es.rate + ")");
                                tempRow.createCell(tempCellCounter++).setCellValue(es.displayName.replace("Swarm/Radar/GBA", "PokeRadar"));
                            } else if (i == 6) {
                                tempCellCounter = 0;
                                tempRow = sheetTrainers.createRow(rowCounter++);
                                tempRow.createCell(tempCellCounter++).setCellValue("Set " + idx + " (Rate: " + es.rate + ")");
                                tempRow.createCell(tempCellCounter++).setCellValue(es.displayName.replace("Swarm/Radar/GBA", "Ruby GBA"));
                            } else if (i == 8) {
                                tempCellCounter = 0;
                                tempRow = sheetTrainers.createRow(rowCounter++);
                                tempRow.createCell(tempCellCounter++).setCellValue("Set " + idx + " (Rate: " + es.rate + ")");
                                tempRow.createCell(tempCellCounter++).setCellValue(es.displayName.replace("Swarm/Radar/GBA", "Sapphire GBA"));
                            } else if (i == 10) {
                                tempCellCounter = 0;
                                tempRow = sheetTrainers.createRow(rowCounter++);
                                tempRow.createCell(tempCellCounter++).setCellValue("Set " + idx + " (Rate: " + es.rate + ")");
                                tempRow.createCell(tempCellCounter++).setCellValue(es.displayName.replace("Swarm/Radar/GBA", "Emerald GBA"));
                            } else if (i == 12) {
                                tempCellCounter = 0;
                                tempRow = sheetTrainers.createRow(rowCounter++);
                                tempRow.createCell(tempCellCounter++).setCellValue("Set " + idx + " (Rate: " + es.rate + ")");
                                tempRow.createCell(tempCellCounter++).setCellValue(es.displayName.replace("Swarm/Radar/GBA", "Fire Red GBA"));
                            } else if (i == 14) {
                                tempCellCounter = 0;
                                tempRow = sheetTrainers.createRow(rowCounter++);
                                tempRow.createCell(tempCellCounter++).setCellValue("Set " + idx + " (Rate: " + es.rate + ")");
                                tempRow.createCell(tempCellCounter++).setCellValue(es.displayName.replace("Swarm/Radar/GBA", "Leaf Green GBA"));
                            }
                        }
                        Encounter e = es.encounters.get(i);
                        String pkmnStr = e.pokemon.name + " Lv";
                        if (e.maxLevel > 0 && e.maxLevel != e.level) {
                            pkmnStr += "s " + e.level + "-" + e.maxLevel;
                        } else {
                            pkmnStr += e.level;
                        }
                        tempRow.createCell(tempCellCounter++).setCellValue(pkmnStr);
                    }
                }
            }
            for (int i = 0; i < cellCounter; i++) {
                sheetTrainers.autoSizeColumn(i);
            }
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
                    log.print(tpk.pokemon.name + " Lv" + tpk.level + "(Iv: " + (int)Math.floor((tpk.difficulty*31)/255.0D) + ")");
                    first = false;
                }
                log.println();
            }
            log.println();
        }
    }

    private void maybeLogToWorkbookTrainerChanges(Workbook wb, RomHandler romHandler) {
        if(settings.getTrainersMod() != Settings.TrainersMod.UNCHANGED
                || settings.isRivalCarriesStarterThroughout()
                || settings.isTrainersLevelModified()) {

            List<Trainer> trainers = romHandler.getTrainers();
            int rowCounter = 0;
            int cellCounter = 0;
            Sheet sheetTrainers = wb.getSheetAt(8);
            sheetTrainers.createFreezePane(2,1);
            Row rowOne = sheetTrainers.createRow(rowCounter++);
            rowOne.createCell(cellCounter++).setCellValue("NUM");
            rowOne.createCell(cellCounter++).setCellValue("NAME");
            rowOne.createCell(cellCounter++).setCellValue("PKMN 1");
            rowOne.createCell(cellCounter++).setCellValue("PKMN 2");
            rowOne.createCell(cellCounter++).setCellValue("PKMN 3");
            rowOne.createCell(cellCounter++).setCellValue("PKMN 4");
            rowOne.createCell(cellCounter++).setCellValue("PKMN 5");
            rowOne.createCell(cellCounter++).setCellValue("PKMN 6");

            CellStyle centerCells = wb.createCellStyle();
            centerCells.setAlignment(HorizontalAlignment.CENTER);
            for(int i = 0; i < cellCounter; i++) {
                rowOne.getCell(i).setCellStyle(centerCells);
            }

            int idx = 0;
            for (Trainer t : trainers) {
                idx++;
                int tempCellCounter = 0;
                Row tempRow = sheetTrainers.createRow(rowCounter++);
                tempRow.createCell(tempCellCounter++).setCellValue(idx);
                String name = "";
                if (t.fullDisplayName != null) {
                    name += t.fullDisplayName;
                } else if (t.name != null) {
                    name += t.name;
                }
                if (t.offset != idx && t.offset != 0) {
                    name += " - " + String.format("@%X", t.offset);
                }

                tempRow.createCell(tempCellCounter++).setCellValue(name);

                for (TrainerPokemon tpk : t.pokemon) {
                    tempRow.createCell(tempCellCounter++).setCellValue(tpk.pokemon.name + " Lv" + tpk.level + " (IVS: " + (int)Math.floor((tpk.difficulty*31)/255.0D) + ")");
                }
            }
            for (int i = 0; i < cellCounter; i++) {
                sheetTrainers.autoSizeColumn(i);
            }
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

    private void logToWorkbookStaticPokemon(Workbook wb, RomHandler romHandler, List<Pokemon> oldStatics) {
        int rowCounter = 0;
        int cellCounter = 0;
        Sheet sheetStatics = wb.getSheetAt(2);
        Row rowOne = sheetStatics.getRow(0);
        boolean doesRowOneExist = false;
        if(rowOne == null) {
            // rowOne is null, which means starters haven't been added.
            sheetStatics.createFreezePane(0,1);
            rowOne = sheetStatics.createRow(rowCounter++);
        } else {
            // rowOne is not null, starters have been added. Work from there.
            rowCounter++;
            cellCounter = 3;
            doesRowOneExist = true;
        }
        rowOne.createCell(cellCounter++);
        rowOne.createCell(cellCounter++).setCellValue("OLD");
        rowOne.createCell(cellCounter++).setCellValue("TO");
        rowOne.createCell(cellCounter++).setCellValue("NEW");

        CellStyle centerCells = wb.createCellStyle();
        centerCells.setAlignment(HorizontalAlignment.CENTER);
        for(int i = 0; i < cellCounter; i++) {
            rowOne.getCell(i).setCellStyle(centerCells);
        }

        List<Pokemon> newStatics = romHandler.getStaticPokemon();
        Map<Pokemon, Integer> seenPokemon = new TreeMap<Pokemon, Integer>();
        for (int i = 0; i < oldStatics.size(); i++) {
            Pokemon oldP = oldStatics.get(i);
            Pokemon newP = newStatics.get(i);
            int tempCellCounter = doesRowOneExist ? 3 : 0;
            Row tempRow = sheetStatics.getRow(rowCounter);
            if(tempRow == null) {
                tempRow = sheetStatics.createRow(rowCounter);
            }
            rowCounter++;
            tempRow.createCell(tempCellCounter++);

            if (seenPokemon.containsKey(oldP)) {
                int amount = seenPokemon.get(oldP);
                seenPokemon.put(oldP, amount);
                tempRow.createCell(tempCellCounter++).setCellValue(oldP.name + " (" + (++amount) + ")");
            } else {
                seenPokemon.put(oldP, 1);
                tempRow.createCell(tempCellCounter++).setCellValue(oldP.name);
            }
            tempRow.createCell(tempCellCounter++).setCellValue("=>");
            tempRow.createCell(tempCellCounter).setCellValue(newP.name);
        }
        for (int i = 0; i < cellCounter; i++) {
            sheetStatics.autoSizeColumn(i);
        }
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

    private void maybeLogToWorkbookMoveChanges(Workbook wb, RomHandler romHandler) {
        if (settings.isRandomizeMoveAccuracies() || settings.isRandomizeMovePowers() || settings.isRandomizeMovePPs()
                || settings.isRandomizeMoveCategory() || settings.isRandomizeMoveTypes() || settings.isUpdateMoves()) {
            int rowCounter = 0;
            int cellCounter = 0;
            Sheet moveSheet = wb.getSheetAt(3);
            moveSheet.createFreezePane(0,1);
            Row rowOne = moveSheet.createRow(rowCounter++);
            rowOne.createCell(cellCounter++).setCellValue("NUM");
            rowOne.createCell(cellCounter++).setCellValue("NAME");
            rowOne.createCell(cellCounter++).setCellValue("TYPE");
            rowOne.createCell(cellCounter++).setCellValue("POWER");
            rowOne.createCell(cellCounter++).setCellValue("ACCURACY");
            rowOne.createCell(cellCounter++).setCellValue("PP");
            if(romHandler.hasPhysicalSpecialSplit()) {
                rowOne.createCell(cellCounter++).setCellValue("CATEGORY");
            }
            CellStyle centerCells = wb.createCellStyle();
            centerCells.setAlignment(HorizontalAlignment.CENTER);
            for(int i = 0; i < cellCounter; i++) {
                rowOne.getCell(i).setCellStyle(centerCells);
            }
            List<Move> allMoves = romHandler.getMoves();
            for (Move mv : allMoves) {
                if (mv != null) {
                    int tempCellCounter = 0;
                    Row tempRow = moveSheet.createRow(rowCounter++);
                    String mvType = (mv.type == null) ? "???" : mv.type.toString();

                    tempRow.createCell(tempCellCounter++).setCellValue(mv.internalId);
                    tempRow.createCell(tempCellCounter++).setCellValue(mv.name);
                    tempRow.createCell(tempCellCounter++).setCellValue(mvType);
                    tempRow.createCell(tempCellCounter++).setCellValue(mv.power);
                    tempRow.createCell(tempCellCounter++).setCellValue((int)mv.hitratio);
                    tempRow.createCell(tempCellCounter++).setCellValue(mv.pp);
                    if (romHandler.hasPhysicalSpecialSplit()) {
                        tempRow.createCell(tempCellCounter).setCellValue(mv.category.toString());
                    }
                }
            }
            for (int i = 0; i < cellCounter; i++) {
                moveSheet.autoSizeColumn(i);
            }
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