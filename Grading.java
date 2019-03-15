/**intGIRLS 2019 Grading
 * Made by Lucinda Zhou
 * Using:
 * Apache (Excel files): https://poi.apache.org/
 * HMMT grading: https://www.hmmt.co/static/scoring-algorithm.pdf
 **/

package integirls;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.Font;

import java.io.*;
import java.util.*;

import javax.swing.BoxLayout;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextField;

import org.apache.poi.openxml4j.exceptions.*;
import org.apache.poi.ss.usermodel.*;


public class Grading {

	public static void main(String[] args) throws IOException, InvalidFormatException {

	

		//set up frame
		JFrame frame = new JFrame("FrameDemo");
		frame.setTitle("inteGIRLS Grading");
		frame.setSize(1000,400);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

		Dimension dim = Toolkit.getDefaultToolkit().getScreenSize();
		frame.setLocation(dim.width/2-frame.getSize().width/2, dim.height/2-frame.getSize().height/2);

		//title (panel/label)
		JPanel titlep = new JPanel();
		titlep.setLayout(new BoxLayout(titlep, BoxLayout.Y_AXIS));

		//title
		JLabel title = new JLabel("InteGIRLs Grading");
		title.setFont(new Font("TimesRoman", Font.BOLD, 25));
		titlep.add(title);

		//label above textbox
		JLabel textfieldlabel = new JLabel("Enter your filepath to the excel file here:");
		textfieldlabel.setFont(new Font("TimesRoman", Font.PLAIN, 20));
		titlep.add(textfieldlabel);

		frame.getContentPane().add(titlep, BorderLayout.NORTH);
		
		//prev input label
		JLabel oldInput = new JLabel();
		oldInput.setFont(new Font("TimesRoman", Font.PLAIN, 15));
		titlep.add(oldInput);


		
		//user input boxes
		JPanel fieldp = new JPanel();

		JTextField textfield = new JTextField("",70);
		textfield.setFont(new Font("TimesRoman", Font.PLAIN, 17));

		fieldp.add(textfield);

		frame.getContentPane().add(fieldp, BorderLayout.CENTER);

		//button
		JButton button = new JButton("Submit");
		fieldp.add(button);

		button.addActionListener(new ActionListener(){
			public void actionPerformed(ActionEvent e){
				oldInput.setText(textfield.getText().trim());
			}
		});

		frame.setVisible(true);
		
		//attempt to read from excel file, repeat
		boolean flag = true;
		while(flag){
			try{
				Workbook workbook = WorkbookFactory.create(new File(oldInput.getText()));
				flag = false;
			}
			catch(Exception e){

			}
		}
		
		//set valid filename
		String FILENAME = oldInput.getText().trim();
		Workbook workbook = WorkbookFactory.create(new File(FILENAME));
		
		//update titles
		textfieldlabel.setText("File found. Enter your sheet name here:");

		//get sheet
		flag = true;
		while(flag){
			if(workbook.getSheetIndex(oldInput.getText()) != -1)
				flag = false;
		}
		
		//set valid sheetname
		String SHEETNAME = oldInput.getText().trim();
		Sheet sheet = workbook.getSheet(SHEETNAME);
		
		//update titles
		textfieldlabel.setText("Sheet found.");
		oldInput.setText("");

		//title
		System.out.println("\n\nInteGIRLS Grading: "+FILENAME+" - "+SHEETNAME+"\n");

		//read from excel file
		System.out.println("Reading data from file...");
		JLabel label2 = new JLabel("Reading data from file...");
		label2.setFont(new Font("TimesRoman", Font.PLAIN, 20));
		label2.setLocation(200,200);
		frame.getContentPane().add(label2, BorderLayout.SOUTH);

		frame.repaint();

		//Sheet sheet = workbook.getSheet("Algebra");
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
		label2.setText("<html>Reading data from file..."
				+ "<br>Weighing problems...</html>");
		for(int i = 0; i < NUM_PROBLEMS; i++){
			probWeighted[i] = Math.pow(java.lang.Math.E,((i+1)/20.0))+Math.max(8.0-(int)Math.log(probTotals[i]),2.0);
		}

		//calc scores for people
		System.out.println("Calculating scores...");
		label2.setText("<html>Reading data from file..."
				+ "<br>Weighing problems..."
				+ "<br>Calculating scores...</html>");
		for(int i = 0; i < allPeople.length; i++){
			double total = 0;
			for(int j = 0; j < NUM_PROBLEMS; j++){
				total += allPeople[i].getProb(j+1)*probWeighted[j]; 
			}
			allPeople[i].setScore(total);
		}

		//sort person array by score
		System.out.println("Ranking participants...");
		label2.setText("<html>Reading data from file..."
				+ "<br>Weighing problems..."
				+ "<br>Calculating scores..."
				+ "<br>Ranking participants...</html>");
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
		label2.setText("<html>Reading data from file..."
				+ "<br>Weighing problems..."
				+ "<br>Calculating scores..."
				+ "<br>Ranking participants..."
				+ "<br>Writing to sheet...</html>");
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
		label2.setText("<html>Reading data from file..."
				+ "<br>Weighing problems..."
				+ "<br>Calculating scores..."
				+ "<br>Ranking participants..."
				+ "<br>Writing to sheet..."
				+ "<br><br>Finished, stored in "+FILENAME.substring(0,FILENAME.indexOf("."))+"_RESULTS.xlsx"+" in sheet "+SHEETNAME+"_Results"+id+"</html>");
	}
}
