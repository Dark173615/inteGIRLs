/**intGIRLS 2019 Grading
 * Made by Lucinda Zhou
 * Using:
 * Apache (Excel files): https://poi.apache.org/
 * HMMT grading: https://www.hmmt.co/static/scoring-algorithm.pdf
 **/

package integirls;

import java.io.*;
import java.util.*;
import org.apache.poi.openxml4j.exceptions.*;
import org.apache.poi.ss.usermodel.*;


public class Grading {

	public static void main(String[] args) throws IOException, InvalidFormatException {

		//SET THESE:
		String FILENAME = "data.xlsx"; //FILE NAME
		String SHEETNAME = "Algebra"; //SHEET NAME

		//title
		System.out.println("\n\nInteGIRLS Grading: "+FILENAME+" - "+SHEETNAME+"\n");

		//read from excel file
		System.out.println("Reading data from file...");
		Workbook workbook = WorkbookFactory.create(new File(FILENAME));
		Sheet sheet = workbook.getSheet("Algebra");
		DataFormatter dataFormatter = new DataFormatter();

		Iterator<Row> rowIterator = sheet.rowIterator(); //iterator

		int NUM_PEOPLE = sheet.getPhysicalNumberOfRows()-1;

		//get title row, get num questions
		Row headerRow = rowIterator.next();
		int NUM_PROBLEMS =(int)headerRow.getCell(headerRow.getLastCellNum()-1).getNumericCellValue();

		//make arrays
		int[] probTotals = new int[NUM_PROBLEMS];
		double[] probWeighted = new double[NUM_PROBLEMS];
		Person[] allPeople = new Person[NUM_PEOPLE];

		//get info from sheet
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();

			//get name
			Iterator<Cell> cellIterator = row.cellIterator(); //iterator
			Cell cell = cellIterator.next();
			allPeople[row.getRowNum()-1] = new Person(dataFormatter.formatCellValue(cell), NUM_PROBLEMS);
			while (cellIterator.hasNext()) {
				cell = cellIterator.next(); 
				int cellValue = Integer.parseInt(dataFormatter.formatCellValue(cell));
				probTotals[cell.getColumnIndex()-1] += cellValue;
				allPeople[row.getRowNum()-1].setProb(cell.getColumnIndex(),cellValue);
				//System.out.print(cellValue + "\t");
			}
			//System.out.println();
		}

		
		//weighed question algorithm
		System.out.println("Weighing problems...");
		for(int i = 0; i < NUM_PROBLEMS; i++){
			probWeighted[i] = Math.pow(java.lang.Math.E,((i+1)/20.0))+Math.max(8.0-(int)Math.log(probTotals[i]),2.0);
		}

		//calc scores for people
		System.out.println("Calculating scores...");
		for(int i = 0; i < allPeople.length; i++){
			double total = 0;
			for(int j = 0; j < NUM_PROBLEMS; j++){
				total += allPeople[i].getProb(j+1)*probWeighted[j]; 
			}
			allPeople[i].setScore(total);
		}

		//sort person array by score
		System.out.println("Ranking participants...");
		Person temp;
		for (int i = 1; i < allPeople.length; i++) {
			for(int j = i ; j > 0 ; j--){
				if(allPeople[j].getScore() > allPeople[j-1].getScore()){
					temp = allPeople[j];
					allPeople[j] = allPeople[j-1];
					allPeople[j-1] = temp;
				}
			}
		}



		//write to another sheet
		System.out.println("Writing to sheet...");
        Random generator = new Random(); 
		int id = generator.nextInt(1000)+1;
		sheet = workbook.createSheet(SHEETNAME+"_Results"+id);
		Row row = sheet.createRow(0);

		//set titles
		Cell cell = row.createCell(0);
		cell.setCellValue("Name");
		cell.setCellStyle(headerRow.getCell(0).getCellStyle());

		cell = row.createCell(1);
		cell.setCellValue("Score");
		cell.setCellStyle(headerRow.getCell(0).getCellStyle());

		//set values (names/scores)
		for(int i = 0; i < allPeople.length; i++){
			row = sheet.createRow(i+1);
			cell = row.createCell(0);
			cell.setCellValue(allPeople[i].getName());

			cell = row.createCell(1);
			cell.setCellValue(allPeople[i].getScore());

		}

		//write to new file
		FileOutputStream fileOut = new FileOutputStream(FILENAME.substring(0,FILENAME.indexOf("."))+"_RESULTS.xlsx");
		workbook.write(fileOut);
		workbook.close();
		
		System.out.println("\nFinished, stored in "+FILENAME.substring(0,FILENAME.indexOf("."))+"_RESULTS.xlsx"+" in sheet "+SHEETNAME+"_Results"+id);
	}
}
