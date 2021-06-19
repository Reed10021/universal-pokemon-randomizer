package com.dabomstew.pkrandom;

import com.dabomstew.pkrandom.constants.*;
import com.dabomstew.pkrandom.pokemon.*;
import com.dabomstew.pkrandom.romhandlers.Gen1RomHandler;
import com.dabomstew.pkrandom.romhandlers.Gen4RomHandler;
import com.dabomstew.pkrandom.romhandlers.RomHandler;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

public class WorkbookHandler {
    private Workbook workbook;

    public WorkbookHandler() {
        workbook = new XSSFWorkbook();
        workbook.createSheet("Stats, Type, Ability, Item");
        workbook.createSheet("Evolutions");
        workbook.createSheet("Starters, Static, Trades");
        workbook.createSheet("Moves, TMs, Move Tutors, Items");
        workbook.createSheet("Poke Movesets");
        workbook.createSheet("Poke TMHM Compat");
        workbook.createSheet("Poke Move Tutor Compat");
        workbook.createSheet("Wild Pokemon");
        workbook.createSheet("Trainers");
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public void logToWorkbookBaseStatAndTypeChanges(RomHandler romHandler) {
        int rowCounter = 0;
        int cellCounter = 0;
        Sheet poke = workbook.getSheetAt(0);
        poke.createFreezePane(0,1);
        Row rowOne = poke.createRow(rowCounter++);

        List<Pokemon> allPokes = romHandler.getPokemon();
        String[] itemNames = romHandler.getItemNames();
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
            CellStyle centerCells = workbook.createCellStyle();
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
            CellRangeAddress[] hpRegion = { CellRangeAddress.valueOf("D2:D" + poke.getLastRowNum()+1) };
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

            CellStyle centerCells = workbook.createCellStyle();
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

    public void logToWorkbookRandomizedEvolutions(RomHandler romHandler, Map<Pokemon, List<Evolution>> originalEvos) {
        int rowCounter = 0;
        int cellCounter = 0;
        Sheet evos = workbook.getSheetAt(1);

        List<Pokemon> allPokes = romHandler.getPokemon();
        String[] itemNames = romHandler.getItemNames();
        List<Move> moves = romHandler.getMoves();

        Map<Pokemon, List<Evolution>> newEvoList = new HashMap<>();
        for (Pokemon pk : allPokes) {
            if(pk != null) {
                newEvoList.put(pk, new ArrayList<>(pk.evolutionsFrom));
            }
        }

        for (Pokemon pk : allPokes) {
            if (pk != null) {
                int numEvos = pk.evolutionsFrom.size();
                if (numEvos > 0) {
                    if(originalEvos.get(pk).containsAll(newEvoList.get(pk))) {
                        continue;
                    }
                    int tempCellCounter = 0;
                    Row temp = evos.createRow(rowCounter++);
                    List<Evolution> evoFrom = pk.evolutionsFrom;

                    temp.createCell(tempCellCounter++).setCellValue(evoFrom.get(0).from.name);
                    temp.createCell(tempCellCounter++).setCellValue("now evolves into");
                    evoFrom.sort(Evolution::compareTo);

                    // Get Pokemon Name, get all Evo Types.
                    for(int i = 0; i < evoFrom.size(); i++) {
                        List<EvolutionType> evoTypes = new ArrayList<>(Collections.singletonList(evoFrom.get(i).type));
                        List<Evolution> evoList = new ArrayList<>(Collections.singletonList(evoFrom.get(i)));
                        Evolution evo = evoFrom.get(i);
                        if(temp.getCell(tempCellCounter-1).getStringCellValue().contains(evo.to.name)) {
                            Cell cell = temp.getCell(tempCellCounter-1);
                            evoTypes.add(evoFrom.get(i-1).type);
                            evoList.add(evoFrom.get(i-1));
                            String tempStr = outputEvoString(evoList, evoTypes, allPokes, itemNames, moves);
                            cell.setCellValue(tempStr);
                        } else {
                            String tempStr = outputEvoString(evoList, evoTypes, allPokes, itemNames, moves);
                            temp.createCell(tempCellCounter++).setCellValue(tempStr);
                        }
                        if (tempCellCounter > cellCounter) {
                            cellCounter = tempCellCounter;
                        }
                    }
                }
            }
        }
        for (int i = 0; i < cellCounter; i++) {
            evos.autoSizeColumn(i);
        }
    }

    private String outputEvoString(List<Evolution> evo, List<EvolutionType> evoTypes, List<Pokemon> allPokes, String[] itemNames, List<Move> moves) {
        StringBuilder evoTypeStr = new StringBuilder();
        for(int j = 0; j < evoTypes.size(); j++) {
            if(!evoTypeStr.toString().isEmpty()) {
                evoTypeStr.append(" AND ");
            } else {
                evoTypeStr.append(evo.get(j).to.name);
                evoTypeStr.append(" (");
            }
            EvolutionType evoType = evoTypes.get(j);
            evoTypeStr.append(evoType.toString());
            if (evoType.usesLevel()) {
                evoTypeStr.append(evo.get(j).extraInfo);
            } else if (evoType.isItem()) {
                evoTypeStr.append(itemNames[evo.get(j).extraInfo]);
            } else if (evoType == EvolutionType.LEVEL_WITH_MOVE) {
                for (Move move : moves) {
                    if (move == null)
                        continue;
                    if (move.number == evo.get(j).extraInfo) {
                        evoTypeStr.append(move.name);
                    }
                }
            } else if (evoType == EvolutionType.LEVEL_WITH_OTHER || evoType == EvolutionType.TRADE_SPECIAL) {
                evoTypeStr.append(allPokes.get(evo.get(j).extraInfo).name);
            }
        }
        evoTypeStr.append(")");
        return evoTypeStr.toString();
    }

    public void logToWorkbookStarters(RomHandler romHandler, List<Pokemon> oldStarters) {
        int rowCounter = 0;
        int cellCounter = 0;
        Sheet starters = workbook.getSheetAt(2);
        starters.createFreezePane(0,1);
        Row rowOne = starters.createRow(rowCounter++);

        rowOne.createCell(cellCounter++).setCellValue("STARTERS");
        rowOne.createCell(cellCounter++);
        rowOne.createCell(cellCounter++);
        starters.addMergedRegion(new CellRangeAddress(0, 0, 0, 2));
        CellStyle centerCells = workbook.createCellStyle();
        centerCells.setAlignment(HorizontalAlignment.CENTER);
        rowOne.getCell(0).setCellStyle(centerCells);

        int oldStarterIndex = 0;
        List<Pokemon> newStarters = romHandler.getStarters();
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

    public void logToWorkbookMoveChanges(RomHandler romHandler) {
        int rowCounter = 0;
        int cellCounter = 0;
        Sheet moveSheet = workbook.getSheetAt(3);
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
        CellStyle centerCells = workbook.createCellStyle();
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

    public void logToWorkbookMovesetChanges(RomHandler romHandler) {
        int cellCounter = 0;
        int rowCounter = 0;
        Sheet sheetMoveset = workbook.getSheetAt(4);
        sheetMoveset.createFreezePane(2,0);
        Row rowOne = sheetMoveset.createRow(rowCounter++);
        rowOne.createCell(cellCounter++).setCellValue("NUM");
        rowOne.createCell(cellCounter++).setCellValue("NAME");

        CellStyle centerCells = workbook.createCellStyle();
        centerCells.setAlignment(HorizontalAlignment.CENTER);
        for(int i = 0; i < cellCounter; i++) {
            rowOne.getCell(i).setCellStyle(centerCells);
        }

        List<Move> moves = romHandler.getMoves();
        Map<Pokemon, List<MoveLearnt>> moveData = romHandler.getMovesLearnt();
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

    public void logToWorkbookTrainerChanges(RomHandler romHandler) {
        int rowCounter = 0;
        int cellCounter = 0;
        Sheet sheetTrainers = workbook.getSheetAt(8);
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

        CellStyle centerCells = workbook.createCellStyle();
        centerCells.setAlignment(HorizontalAlignment.CENTER);
        for(int i = 0; i < cellCounter; i++) {
            rowOne.getCell(i).setCellStyle(centerCells);
        }

        int idx = 0;
        List<Trainer> trainers = romHandler.getTrainers();
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

    public void logToWorkbookStaticPokemon(RomHandler romHandler, List<Pokemon> oldStatics) {
        int rowCounter = 0;
        int cellCounter = 0;
        Sheet sheetStatics = workbook.getSheetAt(2);
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
        if(cellCounter != 0) {
            rowOne.createCell(cellCounter++);
        }
        rowOne.createCell(cellCounter++).setCellValue("OLD STATIC");
        rowOne.createCell(cellCounter++).setCellValue("TO");
        rowOne.createCell(cellCounter++).setCellValue("NEW STATIC");

        CellStyle centerCells = workbook.createCellStyle();
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
            if(tempCellCounter != 0) {
                tempRow.createCell(tempCellCounter++);
            }

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

    public void logToWorkbookWildPokemonChanges(RomHandler romHandler, boolean isUseTimeBasedEncounters) {
        int rowCounter = 0;
        int cellCounter = 0;
        Sheet sheetTrainers = workbook.getSheetAt(7);
        sheetTrainers.createFreezePane(2,0);
        Row rowOne = sheetTrainers.createRow(rowCounter++);
        rowOne.createCell(cellCounter++).setCellValue("NUM");
        rowOne.createCell(cellCounter++).setCellValue("LOCATION");

        CellStyle centerCells = workbook.createCellStyle();
        centerCells.setAlignment(HorizontalAlignment.CENTER);
        for(int i = 0; i < cellCounter; i++) {
            rowOne.getCell(i).setCellStyle(centerCells);
        }

        List<EncounterSet> encounters = romHandler.getEncounters(isUseTimeBasedEncounters);
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

    public void logToWorkbookRandomizedTmMoves(RomHandler romHandler, List<Integer> oldTmMoves) {
        int rowCounter = 0;
        int cellCounter = 0;
        int startingCell = 0;
        Sheet moveSheet = workbook.getSheetAt(3);
        Row rowOne = moveSheet.getRow(0);
        boolean doesRowOneExist = false;
        if(rowOne == null) {
            // rowOne is null, which means moves haven't been added.
            moveSheet.createFreezePane(0,1);
            rowOne = moveSheet.createRow(rowCounter++);
        } else {
            // rowOne is not null, moves have been added. Work from there.
            rowCounter++;
            doesRowOneExist = true;
            for(int i = 0; i < 16; i++) {
                if(rowOne.getCell(i) == null) {
                    startingCell = i;
                    cellCounter = i;
                    break;
                }
            }
        }

        if(cellCounter != 0) {
            rowOne.createCell(cellCounter++);
        }
        rowOne.createCell(cellCounter++).setCellValue("OLD TM");
        rowOne.createCell(cellCounter++).setCellValue("TO");
        rowOne.createCell(cellCounter++).setCellValue("NEW TM");

        CellStyle centerCells = workbook.createCellStyle();
        centerCells.setAlignment(HorizontalAlignment.CENTER);
        for(int i = 0; i < cellCounter; i++) {
            rowOne.getCell(i).setCellStyle(centerCells);
        }

        List<Move> moves = romHandler.getMoves();
        List<Integer> newTmMoves = romHandler.getTMMoves();
        for (int i = 0; i < newTmMoves.size(); i++) {
            int tempCellCounter = doesRowOneExist ? startingCell : 0;
            Row tempRow = moveSheet.getRow(rowCounter);
            if(tempRow == null) {
                tempRow = moveSheet.createRow(rowCounter);
            }
            rowCounter++;
            if(tempCellCounter != 0) {
                tempRow.createCell(tempCellCounter++);
            }
            String temp = String.format("TM%02d %s", i + 1, moves.get(oldTmMoves.get(i)).name);
            tempRow.createCell(tempCellCounter++).setCellValue(temp);
            tempRow.createCell(tempCellCounter++).setCellValue("=>");
            temp = String.format("TM%02d %s", i + 1, moves.get(newTmMoves.get(i)).name);
            tempRow.createCell(tempCellCounter++).setCellValue(temp);
            if(tempCellCounter > cellCounter) {
                cellCounter = tempCellCounter;
            }
        }

        for (int i = 0; i < cellCounter; i++) {
            moveSheet.autoSizeColumn(i);
        }
    }

    public void logtoWorkbookTmHmCompatability(RomHandler romHandler) {
        int cellCounter = 0;
        int rowCounter = 0;
        Sheet sheetTmHmCompat = workbook.getSheetAt(5);
        sheetTmHmCompat.createFreezePane(2,0);
        Row rowOne = sheetTmHmCompat.createRow(rowCounter++);
        rowOne.createCell(cellCounter++).setCellValue("NUM");
        rowOne.createCell(cellCounter++).setCellValue("NAME");

        CellStyle centerCells = workbook.createCellStyle();
        centerCells.setAlignment(HorizontalAlignment.CENTER);
        for(int i = 0; i < cellCounter; i++) {
            rowOne.getCell(i).setCellStyle(centerCells);
        }

        Map<Pokemon, boolean[]> compatMap = romHandler.getTMHMCompatibility();

        List<Move> moves = romHandler.getMoves();
        for (Pokemon pkmn : compatMap.keySet()) {
            int tempCellCounter = 0;
            Row tempRow = sheetTmHmCompat.createRow(rowCounter++);
            tempRow.createCell(tempCellCounter++).setCellValue(pkmn.number);
            tempRow.createCell(tempCellCounter++).setCellValue(pkmn.name);

            boolean[] data = compatMap.get(pkmn);
            for (int i = 1; i < data.length; i++) {
                if(data[i]) {
                    if(romHandler.getTMMoves().size() < i) {
                        tempRow.createCell(tempCellCounter++).setCellValue(
                                String.format("HM%02d ", (i - romHandler.getTMMoves().size()))
                                        + " "
                                        + moves.get(romHandler.getHMMoves().get(i - romHandler.getTMMoves().size() - 1)).name
                        );
                    } else {
                        tempRow.createCell(tempCellCounter++).setCellValue(String.format("TM%02d ", i)
                        + " "
                        + moves.get(romHandler.getTMMoves().get(i - 1)).name
                        );
                    }
                }
                if(tempCellCounter > cellCounter) {
                    cellCounter = tempCellCounter;
                }
            }
        }

        for (int i = 0; i < cellCounter; i++) {
            sheetTmHmCompat.autoSizeColumn(i);
        }
    }

    public void logToWorkbookRandomizedMoveTutors(RomHandler romHandler, List<Integer> oldMtMoves) {
        int rowCounter = 0;
        int cellCounter = 0;
        int startingCell = 0;
        Sheet moveSheet = workbook.getSheetAt(3);
        Row rowOne = moveSheet.getRow(0);
        boolean doesRowOneExist = false;
        if(rowOne == null) {
            // rowOne is null, which means moves haven't been added.
            moveSheet.createFreezePane(0,1);
            rowOne = moveSheet.createRow(rowCounter++);
        } else {
            // rowOne is not null, moves have been added. Work from there.
            rowCounter++;
            doesRowOneExist = true;
            for(int i = 0; i < 16; i++) {
                if(rowOne.getCell(i) == null) {
                    startingCell = i;
                    cellCounter = i;
                    break;
                }
            }
        }
        if(cellCounter != 0) {
            rowOne.createCell(cellCounter++);
        }
        rowOne.createCell(cellCounter++).setCellValue("OLD MOVE TUTOR");
        rowOne.createCell(cellCounter++).setCellValue("TO");
        rowOne.createCell(cellCounter++).setCellValue("NEW MOVE TUTOR");

        CellStyle centerCells = workbook.createCellStyle();
        centerCells.setAlignment(HorizontalAlignment.CENTER);
        for(int i = 0; i < cellCounter; i++) {
            rowOne.getCell(i).setCellStyle(centerCells);
        }

        List<Move> moves = romHandler.getMoves();
        List<Integer> newMtMoves = romHandler.getMoveTutorMoves();
        for (int i = 0; i < newMtMoves.size(); i++) {
            int tempCellCounter = doesRowOneExist ? startingCell : 0;
            Row tempRow = moveSheet.getRow(rowCounter);
            if(tempRow == null) {
                tempRow = moveSheet.createRow(rowCounter);
            }
            rowCounter++;
            if(tempCellCounter != 0) {
                tempRow.createCell(tempCellCounter++);
            }
            tempRow.createCell(tempCellCounter++).setCellValue(moves.get(oldMtMoves.get(i)).name);
            tempRow.createCell(tempCellCounter++).setCellValue("=>");
            tempRow.createCell(tempCellCounter++).setCellValue(moves.get(newMtMoves.get(i)).name);
            if(tempCellCounter > cellCounter) {
                cellCounter = tempCellCounter;
            }
        }

        for (int i = 0; i < cellCounter; i++) {
            moveSheet.autoSizeColumn(i);
        }
    }

    public void logToWorkbookRandomizedMoveTutorCompat(RomHandler romHandler) {
        int cellCounter = 0;
        int rowCounter = 0;
        Sheet sheetTutorCompat = workbook.getSheetAt(6);
        sheetTutorCompat.createFreezePane(2,0);
        Row rowOne = sheetTutorCompat.createRow(rowCounter++);
        rowOne.createCell(cellCounter++).setCellValue("NUM");
        rowOne.createCell(cellCounter++).setCellValue("NAME");

        CellStyle centerCells = workbook.createCellStyle();
        centerCells.setAlignment(HorizontalAlignment.CENTER);
        for(int i = 0; i < cellCounter; i++) {
            rowOne.getCell(i).setCellStyle(centerCells);
        }

        Map<Pokemon, boolean[]> compatMap = romHandler.getMoveTutorCompatibility();
        List<Move> moves = romHandler.getMoves();
        for (Pokemon pkmn : compatMap.keySet()) {
            int tempCellCounter = 0;
            Row tempRow = sheetTutorCompat.createRow(rowCounter++);
            tempRow.createCell(tempCellCounter++).setCellValue(pkmn.number);
            tempRow.createCell(tempCellCounter++).setCellValue(pkmn.name);

            boolean[] data = compatMap.get(pkmn);
            for (int i = 1; i < data.length; i++) {
                if(data[i]) {
                    tempRow.createCell(tempCellCounter++).setCellValue(moves.get(romHandler.getMoveTutorMoves().get(i - 1)).name);
                }
                if(tempCellCounter > cellCounter) {
                    cellCounter = tempCellCounter;
                }
            }
        }

        for (int i = 0; i < cellCounter; i++) {
            sheetTutorCompat.autoSizeColumn(i);
        }
    }

    public void logToWorkbookRandomizedTrades(RomHandler romHandler, List<IngameTrade> oldTrades) {
        int rowCounter = 0;
        int cellCounter = 0;
        Sheet sheetTrades = workbook.getSheetAt(2);
        Row rowOne = sheetTrades.getRow(0);
        boolean doesRowOneExist = false;
        boolean areStaticsRandom = false;
        if(rowOne == null) {
            // rowOne is null, which means starters haven't been added.
            sheetTrades.createFreezePane(0,1);
            rowOne = sheetTrades.createRow(rowCounter++);
        } else {
            // rowOne is not null, starters have been added. Work from there.
            doesRowOneExist = true;
            rowCounter++;
            cellCounter = 3;

            if(rowOne.getCell(cellCounter) != null) {
                areStaticsRandom = true;
                cellCounter += 4;
            }
        }
        if(cellCounter != 0) {
            rowOne.createCell(cellCounter++);
        }
        int col1 = cellCounter;
        rowOne.createCell(cellCounter++).setCellValue("OLD TRADE");
        rowOne.createCell(cellCounter++).setCellValue("OLD TRADE");
        rowOne.createCell(cellCounter++).setCellValue("OLD TRADE");
        sheetTrades.addMergedRegion(new CellRangeAddress(0, 0, col1, cellCounter-1));
        rowOne.createCell(cellCounter++).setCellValue("TO");
        col1 = cellCounter;
        rowOne.createCell(cellCounter++).setCellValue("NEW TRADE");
        rowOne.createCell(cellCounter++).setCellValue("NEW TRADE");
        rowOne.createCell(cellCounter++).setCellValue("NEW TRADE");
        sheetTrades.addMergedRegion(new CellRangeAddress(0, 0, col1, cellCounter-1));

        CellStyle centerCells = workbook.createCellStyle();
        centerCells.setAlignment(HorizontalAlignment.CENTER);
        for(int i = 0; i < cellCounter; i++) {
            rowOne.getCell(i).setCellStyle(centerCells);
        }

        List<IngameTrade> newTrades = romHandler.getIngameTrades();
        for (int i = 0; i < oldTrades.size(); i++) {
            int tempCellCounter = doesRowOneExist ? 3 : 0;
            if(areStaticsRandom) {
                tempCellCounter += 4;
            }
            Row tempRow = sheetTrades.getRow(rowCounter);
            if(tempRow == null) {
                tempRow = sheetTrades.createRow(rowCounter);
            }
            rowCounter++;
            if(tempCellCounter != 0) {
                tempRow.createCell(tempCellCounter++);
            }
            IngameTrade oldT = oldTrades.get(i);
            IngameTrade newT = newTrades.get(i);
            tempRow.createCell(tempCellCounter++).setCellValue(oldT.requestedPokemon.name);
            tempRow.createCell(tempCellCounter++).setCellValue("FOR");
            tempRow.createCell(tempCellCounter++).setCellValue(oldT.nickname + " the " + oldT.givenPokemon.name);
            tempRow.createCell(tempCellCounter++).setCellValue("=>");
            tempRow.createCell(tempCellCounter++).setCellValue(newT.requestedPokemon.name);
            tempRow.createCell(tempCellCounter++).setCellValue("FOR");
            tempRow.createCell(tempCellCounter++).setCellValue(newT.nickname + " the " + newT.givenPokemon.name);
            if(tempCellCounter > cellCounter) {
                cellCounter = tempCellCounter;
            }
        }

        for (int i = 0; i < cellCounter; i++) {
            sheetTrades.autoSizeColumn(i);
        }
    }

    public void logToWorkbookRandomizedItems(RomHandler romHandler, List<Integer> oldItems, List<Integer> oldTMs) {
        int rowCounter = 0;
        int cellCounter = 0;
        int startingCell = 0;
        Sheet sheetItems = workbook.getSheetAt(3);
        Row rowOne = sheetItems.getRow(0);
        boolean doesRowOneExist = false;
        if(rowOne == null) {
            // rowOne is null, which means moves haven't been added.
            sheetItems.createFreezePane(0,1);
            rowOne = sheetItems.createRow(rowCounter++);
        } else {
            // rowOne is not null, moves have been added. Work from there.
            rowCounter++;
            doesRowOneExist = true;
            for(int i = 0; i < 16; i++) {
                if(rowOne.getCell(i) == null) {
                    startingCell = i;
                    cellCounter = i;
                    break;
                }
            }
        }

        if(cellCounter != 0) {
            rowOne.createCell(cellCounter++);
        }
        rowOne.createCell(cellCounter++).setCellValue("OLD ITEM");
        rowOne.createCell(cellCounter++).setCellValue("TO");
        rowOne.createCell(cellCounter++).setCellValue("NEW ITEM");

        CellStyle centerCells = workbook.createCellStyle();
        centerCells.setAlignment(HorizontalAlignment.CENTER);
        for(int i = 0; i < cellCounter; i++) {
            rowOne.getCell(i).setCellStyle(centerCells);
        }

        List<Integer> newItems = romHandler.getRegularFieldItems();
        List<Integer> newTMs = romHandler.getCurrentFieldTMs();
        String[] itemNames = romHandler.getItemNames();

        for (int i = 0; i < newItems.size(); i++) {
            int tempCellCounter = doesRowOneExist ? startingCell : 0;

            Row tempRow = sheetItems.getRow(rowCounter);
            if(tempRow == null) {
                tempRow = sheetItems.createRow(rowCounter);
            }
            rowCounter++;
            if(tempCellCounter != 0) {
                tempRow.createCell(tempCellCounter++);
            }
            tempRow.createCell(tempCellCounter++).setCellValue(itemNames[oldItems.get(i)]);
            tempRow.createCell(tempCellCounter++).setCellValue("=>");
            tempRow.createCell(tempCellCounter++).setCellValue(itemNames[newItems.get(i)]);
            if(tempCellCounter > cellCounter) {
                cellCounter = tempCellCounter;
            }
        }

        int tmIndex = 0;
        if( romHandler.generationOfPokemon() == 1 ) {
            tmIndex = Gen1Constants.tmsStartIndex - 1;
        } else if( romHandler.generationOfPokemon() == 3 ) {
            tmIndex = Gen3Constants.tmItemOffset - 1;
        } else if( romHandler.generationOfPokemon() == 4 ) {
            tmIndex = Gen4Constants.tmItemOffset - 1;
        }

        if( tmIndex != 0 ) {
            for(int i = 0; i < newTMs.size(); i++) {
                int tempCellCounter = doesRowOneExist ? startingCell : 0;

                Row tempRow = sheetItems.getRow(rowCounter);
                if(tempRow == null) {
                    tempRow = sheetItems.createRow(rowCounter);
                }
                rowCounter++;
                if(tempCellCounter != 0) {
                    tempRow.createCell(tempCellCounter++);
                }
                tempRow.createCell(tempCellCounter++).setCellValue(itemNames[oldTMs.get(i) + tmIndex]);
                tempRow.createCell(tempCellCounter++).setCellValue("=>");
                tempRow.createCell(tempCellCounter++).setCellValue(itemNames[newTMs.get(i)+ tmIndex]);
                if(tempCellCounter > cellCounter) {
                    cellCounter = tempCellCounter;
                }
            }
        } else {
            if( romHandler.generationOfPokemon() == 2 ) {
                for(int i = 0; i < newTMs.size(); i++) {
                    int oldTmID = oldTMs.get(i);
                    int newTmID = newTMs.get(i);

                    if (oldTmID >= 1 && oldTmID <= Gen2Constants.tmBlockOneSize) {
                        oldTmID += Gen2Constants.tmBlockOneIndex - 1;
                    } else if (oldTmID >= Gen2Constants.tmBlockOneSize + 1
                            && oldTmID <= Gen2Constants.tmBlockOneSize + Gen2Constants.tmBlockTwoSize) {
                        oldTmID += Gen2Constants.tmBlockTwoIndex - 1 - Gen2Constants.tmBlockOneSize;
                    } else {
                        oldTmID += Gen2Constants.tmBlockThreeIndex - 1 - Gen2Constants.tmBlockOneSize
                                - Gen2Constants.tmBlockTwoSize;
                    }

                    if (newTmID >= 1 && newTmID <= Gen2Constants.tmBlockOneSize) {
                        newTmID += Gen2Constants.tmBlockOneIndex - 1;
                    } else if (newTmID >= Gen2Constants.tmBlockOneSize + 1
                            && newTmID <= Gen2Constants.tmBlockOneSize + Gen2Constants.tmBlockTwoSize) {
                        newTmID += Gen2Constants.tmBlockTwoIndex - 1 - Gen2Constants.tmBlockOneSize;
                    } else {
                        newTmID += Gen2Constants.tmBlockThreeIndex - 1 - Gen2Constants.tmBlockOneSize
                                - Gen2Constants.tmBlockTwoSize;
                    }

                    int tempCellCounter = doesRowOneExist ? startingCell : 0;

                    Row tempRow = sheetItems.getRow(rowCounter);
                    if(tempRow == null) {
                        tempRow = sheetItems.createRow(rowCounter);
                    }
                    rowCounter++;
                    if(tempCellCounter != 0) {
                        tempRow.createCell(tempCellCounter++);
                    }
                    tempRow.createCell(tempCellCounter++).setCellValue(itemNames[oldTMs.get(i) + tmIndex]);
                    tempRow.createCell(tempCellCounter++).setCellValue("=>");
                    tempRow.createCell(tempCellCounter++).setCellValue(itemNames[newTMs.get(i)+ tmIndex]);
                    if(tempCellCounter > cellCounter) {
                        cellCounter = tempCellCounter;
                    }
                }
            } else if( romHandler.generationOfPokemon() == 5 ) {
                for(int i = 0; i < newTMs.size(); i++) {
                    int oldTmID = oldTMs.get(i);
                    int newTmID = newTMs.get(i);
                    if (oldTmID >= 1 && oldTmID <= Gen5Constants.tmBlockOneCount) {
                        oldTmID = oldTmID + (Gen5Constants.tmBlockOneOffset - 1);
                    } else {
                        oldTmID = oldTmID + (Gen5Constants.tmBlockTwoOffset - 1 - Gen5Constants.tmBlockOneCount);
                    }

                    if (newTmID >= 1 && newTmID <= Gen5Constants.tmBlockOneCount) {
                        newTmID = newTmID + (Gen5Constants.tmBlockOneOffset - 1);
                    } else {
                        newTmID = newTmID + (Gen5Constants.tmBlockTwoOffset - 1 - Gen5Constants.tmBlockOneCount);
                    }

                    int tempCellCounter = doesRowOneExist ? startingCell : 0;

                    Row tempRow = sheetItems.getRow(rowCounter);
                    if(tempRow == null) {
                        tempRow = sheetItems.createRow(rowCounter);
                    }
                    rowCounter++;
                    if(tempCellCounter != 0) {
                        tempRow.createCell(tempCellCounter++);
                    }
                    tempRow.createCell(tempCellCounter++).setCellValue(itemNames[oldTMs.get(i) + tmIndex]);
                    tempRow.createCell(tempCellCounter++).setCellValue("=>");
                    tempRow.createCell(tempCellCounter++).setCellValue(itemNames[newTMs.get(i)+ tmIndex]);
                    if(tempCellCounter > cellCounter) {
                        cellCounter = tempCellCounter;
                    }
                }
            }
        }
        for (int i = 0; i < cellCounter; i++) {
            sheetItems.autoSizeColumn(i);
        }
    }
}
