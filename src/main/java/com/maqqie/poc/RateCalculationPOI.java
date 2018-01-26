package com.maqqie.poc;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class RateCalculationPOI {

	public static void main(String[] args) throws FileNotFoundException, IOException, InterruptedException {
		ClassLoader classloader = Thread.currentThread().getContextClassLoader();

		
		File excelFIle = new File(classloader.getResource("Excel1.xlsm").getFile());

		XSSFWorkbook excelWorkBook = new XSSFWorkbook(new FileInputStream(excelFIle));
		FormulaEvaluator formulaEvaluator = excelWorkBook.getCreationHelper().createFormulaEvaluator();

//Caclulating Burtourloon for 100 input values F37
		for(double input=1.00;input<100;input++){
			double totalPivotGoalSeek = 0.00;

			updateInputCell(excelWorkBook, input);


			double f37 = input;//To bring values closer 
			setF37(excelWorkBook, f37);
			formulaEvaluator.evaluateAll();
			totalPivotGoalSeek = input -(readE8(excelWorkBook)+readF24(excelWorkBook)+readF30(excelWorkBook)+readF35(excelWorkBook));

			while(!((f37-totalPivotGoalSeek < 0.001 && f37-totalPivotGoalSeek > -0.001))){//limit 
				f37 = totalPivotGoalSeek;//reducing the difference between F37 and the pivotValue. 
				setF37(excelWorkBook, f37);
				XSSFFormulaEvaluator.evaluateAllFormulaCells(excelWorkBook);	
				totalPivotGoalSeek = input -(readE8(excelWorkBook)+readF24(excelWorkBook)+readF30(excelWorkBook)+readF35(excelWorkBook));
				
			}

			System.out.println("Goal seek acheived at :"+f37 +" Input : "+input);
		}
		
	}

	public static void updateInputCell(XSSFWorkbook excelWorkBook,double inputValue){

		XSSFSheet inputSheet = excelWorkBook.getSheetAt(0);

		XSSFRow inputRow = inputSheet.getRow(45);
		XSSFCell inputCell = inputRow.getCell(2);
		inputCell.setCellValue(inputValue);

	}

	public static void setF37(XSSFWorkbook excelWorkBook,double input){
		XSSFSheet calculationSheet = excelWorkBook.getSheetAt(2);
		XSSFRow burtoUrloonRow = calculationSheet.getRow(36);
		XSSFCell burtoUrloonCell = burtoUrloonRow.getCell(5);
		burtoUrloonCell.setCellValue(input);
	}

	public static double readF37(XSSFWorkbook excelWorkBook){
		XSSFSheet calculationSheet = excelWorkBook.getSheetAt(2);
		XSSFRow burtoUrloonRow = calculationSheet.getRow(36);
		XSSFCell burtoUrloonCell = burtoUrloonRow.getCell(5);
		return Double.parseDouble(burtoUrloonCell.getRawValue());
	}
	public static double readE8(XSSFWorkbook excelWorkBook){

		XSSFSheet calculationSheet = excelWorkBook.getSheetAt(2);
		XSSFRow E8Row = calculationSheet.getRow(7);
		XSSFCell E8Cell = E8Row.getCell(4);

		double E8 = Double.parseDouble(E8Cell.getRawValue());
		//System.out.println("E8 : "+E8);
		return (double) Math.round(E8 * 10000) / 10000;
		//return E8;
	}

	public static double readF24(XSSFWorkbook excelWorkBook){

		XSSFSheet calculationSheet = excelWorkBook.getSheetAt(2);
		XSSFRow F24Row = calculationSheet.getRow(23);
		XSSFCell F24Cell = F24Row.getCell(5);

		double F24 = Double.parseDouble(F24Cell.getRawValue());
		//System.out.println("F24 : "+F24);
		return (double) Math.round(F24 * 10000) / 10000;

		//return F24;
	}

	public static double readF30(XSSFWorkbook excelWorkBook){

		XSSFSheet calculationSheet = excelWorkBook.getSheetAt(2);
		XSSFRow F30Row = calculationSheet.getRow(29);
		XSSFCell F30Cell = F30Row.getCell(5);

		double F30 = Double.parseDouble(F30Cell.getRawValue());
		//System.out.println("F30 : "+F30);
		return (double) Math.round(F30 * 10000) / 10000;

		//return F30;
	}

	public static double readF35(XSSFWorkbook excelWorkBook){

		XSSFSheet calculationSheet = excelWorkBook.getSheetAt(2);
		XSSFRow F35Row = calculationSheet.getRow(34);
		XSSFCell F35Cell = F35Row.getCell(5);

		double F35 = Double.parseDouble(F35Cell.getRawValue());
		//System.out.println("F35 : "+F35);
		return (double) Math.round(F35 * 10000) / 10000;

		//return F35;
	}

}
