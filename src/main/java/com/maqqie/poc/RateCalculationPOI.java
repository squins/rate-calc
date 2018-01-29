package com.maqqie.poc;

import static org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator.evaluateAllFormulaCells;

import java.io.IOException;
import java.math.RoundingMode;
import java.text.DecimalFormat;

import org.apache.commons.lang3.time.StopWatch;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class RateCalculationPOI {

    public static void main(String[] args) throws IOException {

        StopWatch duration = new StopWatch();
        duration.start();
        doCalculation();

        duration.stop();

        System.out.println("Total duration: " + duration);
    }

    private static void doCalculation() throws IOException {
        ClassLoader classloader = Thread.currentThread().getContextClassLoader();

        XSSFWorkbook excelWorkBook = new XSSFWorkbook(classloader.getResourceAsStream("Excel1.xlsm"));
        FormulaEvaluator formulaEvaluator = excelWorkBook.getCreationHelper().createFormulaEvaluator();

        DecimalFormat df = new DecimalFormat("#.##");
        df.setRoundingMode(RoundingMode.HALF_UP);

        // Calculate bruto uurloon for 100 hourly rates.
        for (double hourlyRate = 1.00; hourlyRate < 100; hourlyRate++) {
            double totalPivotGoalSeek;

            setWbInputHourlyRate(excelWorkBook, hourlyRate);

            double testBrutoUurloonValue = hourlyRate; //To bring values closer
            setWorkbookCalculationBrutoUurloon(excelWorkBook, testBrutoUurloonValue);
            formulaEvaluator.evaluateAll();

            totalPivotGoalSeek = hourlyRate - (readWbCalculationE8(excelWorkBook) + readWbCalculationF24(excelWorkBook) + readWbCalculationF30(excelWorkBook) + readWbCalculationF35(excelWorkBook));


            StopWatch uurloonCalculationStopwatch = new StopWatch();
            uurloonCalculationStopwatch.start();
            while (!((testBrutoUurloonValue - totalPivotGoalSeek < 0.001 && testBrutoUurloonValue - totalPivotGoalSeek > -0.001))) { //limit
                testBrutoUurloonValue = totalPivotGoalSeek; // reducing the difference between F37 and the pivotValue.
                setWorkbookCalculationBrutoUurloon(excelWorkBook, testBrutoUurloonValue);
                evaluateAllFormulaCells(excelWorkBook);
                totalPivotGoalSeek = hourlyRate -
                        (readWbCalculationE8(excelWorkBook) +
                        readWbCalculationF24(excelWorkBook) +
                        readWbCalculationF30(excelWorkBook) + readWbCalculationF35(excelWorkBook));
            }
            uurloonCalculationStopwatch.stop();

            System.out.println("Goal seek achieved; hourlyRate input: " + hourlyRate + " Bruto uurloon calculated: " + df.format(testBrutoUurloonValue) + ", duration: " + uurloonCalculationStopwatch);
        }
    }

    private static void setWbInputHourlyRate(XSSFWorkbook excelWorkBook, double hourlyRate) {

        XSSFSheet inputSheet = excelWorkBook.getSheetAt(0);

        XSSFRow inputRow = inputSheet.getRow(45);
        XSSFCell inputCell = inputRow.getCell(2);
        inputCell.setCellValue(hourlyRate);

    }

    private static void setWorkbookCalculationBrutoUurloon(XSSFWorkbook excelWorkBook, double input) {
        XSSFSheet calculationSheet = excelWorkBook.getSheetAt(2);
        XSSFRow brutoUurloonRow = calculationSheet.getRow(36);
        XSSFCell brutoUUrloonCell = brutoUurloonRow.getCell(5);
        brutoUUrloonCell.setCellValue(input);
    }

    private static double readWbCalculationE8(XSSFWorkbook excelWorkBook) {

        XSSFSheet calculationSheet = excelWorkBook.getSheetAt(2);
        XSSFRow E8Row = calculationSheet.getRow(7);
        XSSFCell E8Cell = E8Row.getCell(4);

        double E8 = Double.parseDouble(E8Cell.getRawValue());
        //System.out.println("E8 : "+E8);
        return (double) Math.round(E8 * 10000) / 10000;
        //return E8;
    }

    private static double readWbCalculationF24(XSSFWorkbook excelWorkBook) {

        XSSFSheet calculationSheet = excelWorkBook.getSheetAt(2);
        XSSFRow F24Row = calculationSheet.getRow(23);
        XSSFCell F24Cell = F24Row.getCell(5);

        double F24 = Double.parseDouble(F24Cell.getRawValue());
        //System.out.println("F24 : "+F24);
        return (double) Math.round(F24 * 10000) / 10000;

        //return F24;
    }

    private static double readWbCalculationF30(XSSFWorkbook excelWorkBook) {

        XSSFSheet calculationSheet = excelWorkBook.getSheetAt(2);
        XSSFRow F30Row = calculationSheet.getRow(29);
        XSSFCell F30Cell = F30Row.getCell(5);

        double F30 = Double.parseDouble(F30Cell.getRawValue());
        //System.out.println("F30 : "+F30);
        return (double) Math.round(F30 * 10000) / 10000;

        //return F30;
    }

    private static double readWbCalculationF35(XSSFWorkbook excelWorkBook) {

        XSSFSheet calculationSheet = excelWorkBook.getSheetAt(2);
        XSSFRow F35Row = calculationSheet.getRow(34);
        XSSFCell F35Cell = F35Row.getCell(5);

        double F35 = Double.parseDouble(F35Cell.getRawValue());
        //System.out.println("F35 : "+F35);
        return (double) Math.round(F35 * 10000) / 10000;

        //return F35;
    }

}
