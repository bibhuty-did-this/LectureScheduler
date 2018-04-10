package controller;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;


import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class Initializer{
	
	private static String DIRECTORY="Database/Department";
	private static String PATH_nameOfTheDepartment="Database/Department/nameOfTheDepartment.xls";
	private static String PATH_nameOfTheCoursesAndYears="Database/Department/nameOfTheCoursesAndYears.xls";
	private static String PATH_facultyDetails="Database/Department/facultyDetails.xls";
	private static String PATH_classroomDetails="Database/Department/classroomDetails.xls";
	private static String PATH_labroomDetails="Database/Department/labroomDetails.xls";
	private static String ODD="Database/Odd";
	private static String EVEN="Database/Even";
	private static String details_PATH="/details.xls";
	private static String labDetails_PATH="/labDetails.xls";
	private static String subjectDetatils_PATH="/subjectDetails.xls";

	public Initializer() throws IOException{
		createDirectory(DIRECTORY);
		createExcelFile_nameOfTheDepartment();
		createExcelFile_nameOfTheCoursesAndYears();
		createExcelFile_facultyDetails();
		createExcelFile_classroomDetails();
		createExcelFile_labroomDetails();
		createDirectory(ODD);
		createDirectory(EVEN);
		createDatabaseForOddAndEvenSemesters();
	}

	public void createExcelFile_nameOfTheDepartment(){

		if(fileExists(PATH_nameOfTheDepartment))return;

		//Creation of the excel file for the first time
		Workbook workbook=new HSSFWorkbook();

		//Creation of sheet "nameOfTheDepartment" for the first time
		Sheet nameOfTheDepartment=workbook.createSheet("nameOfTheDepartment");
		Row r01=nameOfTheDepartment.createRow(1);
		Cell r01_c01=r01.createCell(0);
		r01_c01.setCellValue("Name of the department");
		Cell r01_c02=r01.createCell(1);
		r01_c02.setCellValue("Computer Science & Engineering");
		CellStyle style=workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.BLACK.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);

		Font font=workbook.createFont();
		font.setColor(IndexedColors.WHITE.getIndex());
		font.setFontHeight((short)200);

		style.setFont(font);

		r01_c01.setCellStyle(style);

		nameOfTheDepartment.setColumnWidth(0,7000);
		nameOfTheDepartment.setColumnWidth(1,20000);

		String instructionValue="Please edit the name of your department, if necessary";
		Cell instruction=instructionCell(workbook,nameOfTheDepartment,instructionValue);

		try{
			FileOutputStream output=new FileOutputStream(PATH_nameOfTheDepartment);
			workbook.write(output);
			output.close();
			openExcelFile(PATH_nameOfTheDepartment);
		}catch(Exception ex){
			ex.printStackTrace();
		}

	}
	public void createExcelFile_nameOfTheCoursesAndYears(){

		if(fileExists(PATH_nameOfTheCoursesAndYears))return;

		//Creation of the excel file for the first time
		Workbook workbook=new HSSFWorkbook();

		//Creation of sheet "nameOfTheDepartment" for the first time
		Sheet nameOfTheCourses=workbook.createSheet("nameOfTheCoursesAndYears");
		Row r01=nameOfTheCourses.createRow(1);
		Cell r01_c01=r01.createCell(0);
		r01_c01.setCellValue("Name of the courses");
		Cell r01_c02=r01.createCell(1);
		r01_c02.setCellValue("Description");
		CellStyle style=workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.BLACK.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);

		Font font=workbook.createFont();
		font.setColor(IndexedColors.WHITE.getIndex());
		font.setFontHeight((short)200);

		style.setFont(font);

		r01_c01.setCellStyle(style);
		r01_c02.setCellStyle(style);
		nameOfTheCourses.setColumnWidth(0,7000);
		nameOfTheCourses.setColumnWidth(1,20000);


		String instructionValue="Please enter the data according to the examples and\n keep the naming convention " +
				"according to your comfortness.";
		Cell instruction=instructionCell(workbook,nameOfTheCourses,instructionValue);

		Row r02=nameOfTheCourses.createRow(2);
		Cell r02_c01=r02.createCell(0);
		Cell r02_c02=r02.createCell(1);
		r02_c01.setCellValue("ug02");
		r02_c02.setCellValue("Undergraduate 2nd year students");

		try{
			FileOutputStream output=new FileOutputStream(PATH_nameOfTheCoursesAndYears);
			workbook.write(output);
			output.close();
			openExcelFile(PATH_nameOfTheCoursesAndYears);
		}catch(Exception ex){
			ex.printStackTrace();
		}
	}
	public void createExcelFile_facultyDetails(){

		if(fileExists(PATH_facultyDetails))return;

		//Creation of the excel file for the first time
		Workbook workbook=new HSSFWorkbook();

		//Creation of sheet "nameOfTheDepartment" for the first time
		Sheet facultyDetails=workbook.createSheet("facultyDetails");
		Row r01=facultyDetails.createRow(1);
		Cell r01_c01=r01.createCell(0);
		r01_c01.setCellValue("Enrolment No");
		Cell r01_c02=r01.createCell(1);
		r01_c02.setCellValue("Name of the faculty");
		Cell r01_c03=r01.createCell(2);
		r01_c03.setCellValue("Initials of the faculty");
		CellStyle style=workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.BLACK.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);

		Font font=workbook.createFont();
		font.setColor(IndexedColors.WHITE.getIndex());
		font.setFontHeight((short)200);

		style.setFont(font);

		r01_c01.setCellStyle(style);
		r01_c02.setCellStyle(style);
		r01_c03.setCellStyle(style);
		facultyDetails.setColumnWidth(0,7000);
		facultyDetails.setColumnWidth(1,20000);
		facultyDetails.setColumnWidth(2,20000);

		String instructionValue="Please enter the data according to the examples.";
		Cell instruction=instructionCell(workbook,facultyDetails,instructionValue);

		Row r02=facultyDetails.createRow(2);
		Cell r02_c01=r02.createCell(0);
		Cell r02_c02=r02.createCell(1);
		Cell r02_c03=r02.createCell(2);
		r02_c01.setCellValue("cse01");
		r02_c02.setCellValue("Dhrubajyoti Bhowmik");
		r02_c03.setCellValue("DJB");

		try{
			FileOutputStream output=new FileOutputStream(PATH_facultyDetails);
			workbook.write(output);
			output.close();
			openExcelFile(PATH_facultyDetails);
		}catch(Exception ex){
			ex.printStackTrace();
		}
	}
	public void createExcelFile_classroomDetails(){

		if(fileExists(PATH_classroomDetails))return;
		//Creation of the excel file for the first time
		Workbook workbook=new HSSFWorkbook();

		//Creation of sheet "nameOfTheDepartment" for the first time
		Sheet facultyDetails=workbook.createSheet("classroomDetails");
		Row r01=facultyDetails.createRow(1);
		Cell r01_c01=r01.createCell(0);
		r01_c01.setCellValue("Classroom No");
		Cell r01_c02=r01.createCell(1);
		r01_c02.setCellValue("Capacity");
		CellStyle style=workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.BLACK.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);

		Font font=workbook.createFont();
		font.setColor(IndexedColors.WHITE.getIndex());
		font.setFontHeight((short)200);

		style.setFont(font);

		r01_c01.setCellStyle(style);
		r01_c02.setCellStyle(style);
		facultyDetails.setColumnWidth(0,7000);
		facultyDetails.setColumnWidth(1,20000);

		String instructionValue="Please enter the data according to the examples.";
		Cell instruction=instructionCell(workbook,facultyDetails,instructionValue);

		Row r02=facultyDetails.createRow(2);
		Cell r02_c01=r02.createCell(0);
		Cell r02_c02=r02.createCell(1);
		r02_c01.setCellValue("G-08");
		r02_c02.setCellValue("80");

		try{
			FileOutputStream output=new FileOutputStream(PATH_classroomDetails);
			workbook.write(output);
			output.close();
			openExcelFile(PATH_classroomDetails);
		}catch(Exception ex){
			ex.printStackTrace();
		}
	}
	public void createExcelFile_labroomDetails(){

		if(fileExists(PATH_labroomDetails))return;
		//Creation of the excel file for the first time
		Workbook workbook=new HSSFWorkbook();

		//Creation of sheet "nameOfTheDepartment" for the first time
		Sheet facultyDetails=workbook.createSheet("classroomDetails");
		Row r01=facultyDetails.createRow(1);
		Cell r01_c01=r01.createCell(0);
		r01_c01.setCellValue("Lab No");
		Cell r01_c02=r01.createCell(1);
		r01_c02.setCellValue("Name of the lab");
		CellStyle style=workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.BLACK.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);

		Font font=workbook.createFont();
		font.setColor(IndexedColors.WHITE.getIndex());
		font.setFontHeight((short)200);

		style.setFont(font);

		r01_c01.setCellStyle(style);
		r01_c02.setCellStyle(style);
		facultyDetails.setColumnWidth(0,7000);
		facultyDetails.setColumnWidth(1,20000);

		String instructionValue="Please enter the data according to the examples.";
		Cell instruction=instructionCell(workbook,facultyDetails,instructionValue);

		Row r02=facultyDetails.createRow(2);
		Cell r02_c01=r02.createCell(0);
		Cell r02_c02=r02.createCell(1);
		r02_c01.setCellValue("L-08");
		r02_c02.setCellValue("Microprocessor Lab");

		try{
			FileOutputStream output=new FileOutputStream(PATH_labroomDetails);
			workbook.write(output);
			output.close();
			openExcelFile(PATH_labroomDetails);
		}catch(Exception ex){
			ex.printStackTrace();
		}
	}
	public void createExcelFile_details(String path){

		if(fileExists(path))return;
		//Creation of the excel file for the first time
		Workbook workbook=new HSSFWorkbook();

		//Creation of sheet "nameOfTheDepartment" for the first time
		Sheet details=workbook.createSheet("details");
		Row r01=details.createRow(1);
		Cell r01_c01=r01.createCell(0);
		r01_c01.setCellValue("No of sections");
		Cell r01_c02=r01.createCell(1);
		r01_c02.setCellValue(2);

		CellStyle style=workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.BLACK.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);

		Font font=workbook.createFont();
		font.setColor(IndexedColors.WHITE.getIndex());
		font.setFontHeight((short)200);

		style.setFont(font);

		r01_c01.setCellStyle(style);
		details.setColumnWidth(0,7000);

		String instructionValue="Please enter the section details\n(by default no of sections is taken as 2)";
		Cell instruction=instructionCell(workbook,details,instructionValue);


		try{
			FileOutputStream output=new FileOutputStream(path);
			workbook.write(output);
			output.close();
			openExcelFile(path);
		}catch(Exception ex){
			ex.printStackTrace();
		}
	}

	public void createExcelFile_labDetails(String path){
		if(fileExists(path))return;
		//Creation of the excel file for the first time
		Workbook workbook=new HSSFWorkbook();

		//Creation of sheet "nameOfTheDepartment" for the first time
		Sheet details=workbook.createSheet("labDetails");
		Row r01=details.createRow(1);
		Cell r01_c01=r01.createCell(0);
		r01_c01.setCellValue("LabCode");

		Cell r01_c02=r01.createCell(1);
		r01_c02.setCellValue("Name of the lab");

		Cell r01_c03=r01.createCell(2);
		r01_c03.setCellValue("Room no. of the lab");

		Cell r01_c04=r01.createCell(3);
		r01_c04.setCellValue("No of groups");

		Cell r01_c05=r01.createCell(4);
		r01_c05.setCellValue("Id of the teacher assigned");

		CellStyle style=workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.BLACK.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);

		Font font=workbook.createFont();
		font.setColor(IndexedColors.WHITE.getIndex());
		font.setFontHeight((short)200);

		style.setFont(font);

		r01_c01.setCellStyle(style);
		r01_c02.setCellStyle(style);
		r01_c03.setCellStyle(style);
		r01_c04.setCellStyle(style);
		r01_c05.setCellStyle(style);

		details.setColumnWidth(0,7000);
		details.setColumnWidth(1,7000);
		details.setColumnWidth(2,7000);
		details.setColumnWidth(3,7000);
		details.setColumnWidth(4,7000);
		details.setColumnWidth(5,7000);

		String instructionValue="Please enter the lab details";
		Cell instruction=instructionCell(workbook,details,instructionValue);


		try{
			FileOutputStream output=new FileOutputStream(path);
			workbook.write(output);
			output.close();
			openExcelFile(path);
		}catch(Exception ex){
			ex.printStackTrace();
		}
	}

	public void createExcelFile_subjectDetails(String path){
		if(fileExists(path))return;
		//Creation of the excel file for the first time
		Workbook workbook=new HSSFWorkbook();

		//Creation of sheet "nameOfTheDepartment" for the first time
		Sheet details=workbook.createSheet("subjectDetails");
		Row r01=details.createRow(1);
		Cell r01_c01=r01.createCell(0);
		r01_c01.setCellValue("Subject Code");

		Cell r01_c02=r01.createCell(1);
		r01_c02.setCellValue("Name of the subject");

		Cell r01_c03=r01.createCell(2);
		r01_c03.setCellValue("Combined(Y/N)");

		Cell r01_c04=r01.createCell(3);
		r01_c04.setCellValue("No of students");

		Cell r01_c05=r01.createCell(4);
		r01_c05.setCellValue("Id of the teacher assigned");

		CellStyle style=workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.BLACK.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);

		Font font=workbook.createFont();
		font.setColor(IndexedColors.WHITE.getIndex());
		font.setFontHeight((short)200);

		style.setFont(font);

		r01_c01.setCellStyle(style);
		r01_c02.setCellStyle(style);
		r01_c03.setCellStyle(style);
		r01_c04.setCellStyle(style);
		r01_c05.setCellStyle(style);

		details.setColumnWidth(0,7000);
		details.setColumnWidth(1,7000);
		details.setColumnWidth(2,7000);
		details.setColumnWidth(3,7000);
		details.setColumnWidth(4,7000);
		details.setColumnWidth(5,7000);

		String instructionValue="Please enter the lab details";
		Cell instruction=instructionCell(workbook,details,instructionValue);


		try{
			FileOutputStream output=new FileOutputStream(path);
			workbook.write(output);
			output.close();
			openExcelFile(path);
		}catch(Exception ex){
			ex.printStackTrace();
		}
	}
	public void createDatabaseForOddAndEvenSemesters() throws IOException{
		Workbook workbook=new HSSFWorkbook(new FileInputStream(PATH_nameOfTheCoursesAndYears));
		Sheet sheet=workbook.getSheetAt(0);
		int i=0;
		boolean take=i>1;
		for(Row row:sheet){
			for(Cell cell:row){
				if(take){
					String ODD_PATH=ODD+"/"+cell.getStringCellValue()+"/main";
					createDirectory(ODD_PATH);
					createExcelFile_details(ODD_PATH.concat(details_PATH));
					createExcelFile_labDetails(ODD_PATH.concat(labDetails_PATH));
					createExcelFile_subjectDetails(ODD_PATH.concat(subjectDetatils_PATH));
					String EVEN_PATH=EVEN+"/"+cell.getStringCellValue()+"/main";
					createDirectory(EVEN_PATH);
					createExcelFile_details(EVEN_PATH.concat(details_PATH));
					createExcelFile_labDetails(EVEN_PATH.concat(labDetails_PATH));
					createExcelFile_subjectDetails(EVEN_PATH.concat(subjectDetatils_PATH));
					createDirectory(ODD+"/"+cell.getStringCellValue()+"/generated");
					createDirectory(EVEN+"/"+cell.getStringCellValue()+"/generated");
					break;
				}
			}
			take=++i>1;
		}
	}

	public boolean fileExists(String path){
		File file=new File(path);
		if(file.exists()){
			System.out.println(path+ " already exists");
			return true;
		}
		return false;
	}

	void createDirectory(String path){

		File directory=new File(path);
		if(!directory.exists() && !directory.isDirectory()){
			boolean successful=directory.mkdirs();
			if(successful)
				System.out.println(directory.getPath() +" is successfully created");
			else
				System.out.println("Failed to create a directory");
		}else
			System.out.println(path+" already exists");
	}

	void openExcelFile(String path){
		try {
			Desktop.getDesktop().open(new File(path));
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	Cell instructionCell(Workbook workbook,Sheet sheet,String instructionValue){
		Cell instruction=sheet.createRow(0).createCell(0);
		instruction.setCellValue(instructionValue);
		instruction.getRow().setHeightInPoints(30);
		sheet.addMergedRegion(new CellRangeAddress(0,0,0,5));

		CellStyle style=workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);

		Font font=workbook.createFont();
		font.setColor(IndexedColors.BLACK.getIndex());
		font.setFontHeight((short)200);
		font.setBold(true);

		style.setFont(font);
		instruction.setCellStyle(style);

		return instruction;
	}
}
