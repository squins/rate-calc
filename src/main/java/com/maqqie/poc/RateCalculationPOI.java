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

    private static XSSFWorkbook excelWorkBook;
    private static FormulaEvaluator formulaEvaluator;
    private static DecimalFormat amountFormat;

    public static void main(String[] args) throws IOException {
        ClassLoader classloader = Thread.currentThread().getContextClassLoader();

        excelWorkBook = new XSSFWorkbook(classloader.getResourceAsStream("Excel1.xlsm"));
        formulaEvaluator = excelWorkBook.getCreationHelper().createFormulaEvaluator();
        amountFormat = new DecimalFormat("#.00");
        amountFormat.setRoundingMode(RoundingMode.HALF_UP);

        StopWatch duration = new StopWatch();
        duration.start();
        doCalculation();

        duration.stop();

        System.out.println("Total duration: " + duration);
    }

    private static void doCalculation() {
        // Calculate bruto uurloon for 100 hourly rates.
        for (double hourlyRateInput = 20; hourlyRateInput < 150; hourlyRateInput++) {
            setWbInputHourlyRate(hourlyRateInput);

            double testBrutoUurloon = hourlyRateInput -1; //To bring values closer
            setWbCalculationBrutoUurloon(excelWorkBook, testBrutoUurloon);
            formulaEvaluator.evaluateAll();

            // comment out one of them, don't run concurrently, as POI seems to cache evaluations so last one is always fastest.
            calculateWithOriginalAgorithm(hourlyRateInput, testBrutoUurloon);
            calculateWithExcelLikeAlgorithm(hourlyRateInput, testBrutoUurloon);

        }
    }

    /**
     * Original algorithm uses a (bit) different algorithm than the original Excel.
     */
    private static double calculateWithOriginalAgorithm(double hourlyRateInput, double testBrutoUurloon) {

        StopWatch duration = new StopWatch();
        duration.start();
        double totalPivotGoalSeek = hourlyRateInput - (readWbCalculationE8() + readWbCalculationF24() + readWbCalculationF30()
                + readWbCalculationF35());
        while (!((testBrutoUurloon - totalPivotGoalSeek < 0.001 && testBrutoUurloon - totalPivotGoalSeek > -0.001))) { //limit
            testBrutoUurloon = totalPivotGoalSeek; // reducing the difference between F37 and the pivotValue.
            setWbCalculationBrutoUurloon(excelWorkBook, testBrutoUurloon);
            evaluateAllFormulaCells(excelWorkBook);
            totalPivotGoalSeek = hourlyRateInput - (readWbCalculationE8() + readWbCalculationF24() + readWbCalculationF30() + readWbCalculationF35());
        }
        duration.stop();

        System.out.println("calculateWithOriginalAgorithm hourlyRate input: " + hourlyRateInput + " Bruto uurloon: " + amountFormat.format(testBrutoUurloon) + ", duration: " + duration);

        return testBrutoUurloon;
    }

    private static double calculateWithExcelLikeAlgorithm(double hourlyRateInput, double testBrutoUurloon) {

        StopWatch duration = new StopWatch();
        duration.start();
        double goalSeekToAlmostZero = calculateTotalPivotGoalSeek(testBrutoUurloon, hourlyRateInput);
        while (goalSeekToAlmostZero > 0.001 || goalSeekToAlmostZero <  - 0.001) {
            testBrutoUurloon -= goalSeekToAlmostZero;

            setWbCalculationBrutoUurloon(excelWorkBook, testBrutoUurloon);
            evaluateAllFormulaCells(excelWorkBook);

            goalSeekToAlmostZero = calculateTotalPivotGoalSeek(testBrutoUurloon, hourlyRateInput);
        }
        duration.stop();

        System.out.println("calculateWithExcelLikeAlgorithm hourlyRate input: " + hourlyRateInput + " Bruto uurloon: " + amountFormat.format(testBrutoUurloon) + ", duration: " + duration);

        return testBrutoUurloon;
    }


    private static void setWbCalculationBrutoUurloon(XSSFWorkbook excelWorkBook, double input) {
        XSSFSheet calculationSheet = excelWorkBook.getSheetAt(2);
        XSSFRow brutoUurloonRow = calculationSheet.getRow(36);
        XSSFCell brutoUurloonCell = brutoUurloonRow.getCell(5);
        brutoUurloonCell.setCellValue(input);
    }

    private static double calculateTotalPivotGoalSeek(double brutoUurloonInput, double hourlyRateInput) {
        return brutoUurloonInput -
                (hourlyRateInput
                        - readWbCalculationE8()
                        - readWbCalculationF24()
                        - readWbCalculationF30()
                        - readWbCalculationF35()
                );
    }

    private static void setWbInputHourlyRate(double hourlyRate) {

        XSSFSheet inputSheet = excelWorkBook.getSheetAt(0);

        XSSFRow inputRow = inputSheet.getRow(45);
        XSSFCell inputCell = inputRow.getCell(2);
        inputCell.setCellValue(hourlyRate);

    }


    private static double readWbCalculationE8() {

        XSSFSheet calculationSheet = excelWorkBook.getSheetAt(2);
        XSSFRow E8Row = calculationSheet.getRow(7);
        XSSFCell E8Cell = E8Row.getCell(4);

        double E8 = Double.parseDouble(E8Cell.getRawValue());
        //System.out.println("E8 : "+E8);
        return (double) Math.round(E8 * 10000) / 10000;
        //return E8;
    }

    private static double readWbCalculationF24() {

        XSSFSheet calculationSheet = excelWorkBook.getSheetAt(2);
        XSSFRow F24Row = calculationSheet.getRow(23);
        XSSFCell F24Cell = F24Row.getCell(5);

        double F24 = Double.parseDouble(F24Cell.getRawValue());
        //System.out.println("F24 : "+F24);
        return (double) Math.round(F24 * 10000) / 10000;

        //return F24;
    }

    private static double readWbCalculationF30() {

        XSSFSheet calculationSheet = excelWorkBook.getSheetAt(2);
        XSSFRow F30Row = calculationSheet.getRow(29);
        XSSFCell F30Cell = F30Row.getCell(5);

        double F30 = Double.parseDouble(F30Cell.getRawValue());
        //System.out.println("F30 : "+F30);
        return (double) Math.round(F30 * 10000) / 10000;

        //return F30;
    }

    private static double readWbCalculationF35() {

        XSSFSheet calculationSheet = excelWorkBook.getSheetAt(2);
        XSSFRow F35Row = calculationSheet.getRow(34);
        XSSFCell F35Cell = F35Row.getCell(5);

        double F35 = Double.parseDouble(F35Cell.getRawValue());
        //System.out.println("F35 : "+F35);
        return (double) Math.round(F35 * 10000) / 10000;

        //return F35;
    }

}
