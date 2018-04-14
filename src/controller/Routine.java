package controller;

import dataStructure.Faculty;
import functionalities.Operator;
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
import java.util.ArrayList;
import java.util.HashMap;

public class Routine implements Operator{

	private static final int NO_OF_WORKING_DAYS=5;
	private static final int NO_OF_WORKING_HOURS=8;

	private static final String PATH_nameOfTheCoursesAndYears="Database/Department/nameOfTheCoursesAndYears.xls";
	private static final String PATH_facultyDetails="Database/Department/facultyDetails.xls";
	private static final String PATH_classroomDetails="Database/Department/classroomDetails.xls";
	private static final String PATH_labroomDetails="Database/Department/labroomDetails.xls";
	private static final String ODD="Database/Odd";
	private static final String EVEN="Database/Even";
	private static final String details_PATH="/details.xls";
	private static final String labDetails_PATH="/labDetails.xls";
	private static final String subjectDetatils_PATH="/subjectDetails.xls";

	private HashMap<String,String> theory_faculty;
	private HashMap<String,Integer> theory_credits;
	private HashMap<String,Integer> lab_hours;
	private HashMap<String,String> lab_faculty;
	private HashMap<String,String> lab_room;
	private HashMap<String,String> labCode_room;
	private HashMap<String,Integer> theory_capacity;
	private HashMap<String,String> faculty_initials;
	private HashMap<String,String> faculty_name;
	private HashMap<String,Integer> classroom_capacity;

	private boolean[][][] labRooms;
	private boolean[][][] classRooms;
	private boolean[][][] faculty;
	private boolean[][][][] slot;


	private int noOfLabRooms;
	private int noOfClassRooms;
	private int noOfFaculties;
	private int noOfCourses;
	private int maximumNoOfSections;
	private int lunchHour;
	private int minLectureTime;
	private int maxLectureTime;

	private HashMap<String,Integer> facultyId;
	private HashMap<String,Integer> labId;
	private HashMap<String,Integer> classroomId;
	private HashMap<String,Integer> theoryCodeId;
	private HashMap<String,Integer> labCodeId;

	private String section[];

	private ArrayList<String> combinedSubjects;

	public Routine() throws IOException{

		initialize();
		loadFacultyDetails();
		loadClassroomDetails();
		loadLabroomDetails();
		generateExcelFileForTheorySubjects(EVEN);
		generateExcelFileForLabSubjects(EVEN);
		//generateExcelFileForReservedClassroom();
		//generateExcelFileForReservedSlots(EVEN);
		loadSemesterDetails(EVEN);
		generateLectureScheduler();
	}

	private void initialize() throws IOException{
		theory_faculty=new HashMap<>();
		theory_credits=new HashMap<>();
		lab_hours=new HashMap<>();
		lab_faculty=new HashMap<>();
		lab_room=new HashMap<>();
		labCode_room=new HashMap<>();
		theory_capacity=new HashMap<>();
		faculty_initials=new HashMap<>();
		faculty_name=new HashMap<>();
		classroom_capacity=new HashMap<>();

		noOfLabRooms=calculateNoOfRows(PATH_labroomDetails);
		noOfClassRooms=calculateNoOfRows(PATH_classroomDetails);
		noOfFaculties=calculateNoOfRows(PATH_facultyDetails);
		noOfCourses=calculateNoOfRows(PATH_nameOfTheCoursesAndYears);
		maximumNoOfSections=5;
		lunchHour=4;

		labRooms=new boolean[noOfLabRooms][NO_OF_WORKING_DAYS][NO_OF_WORKING_HOURS];
		classRooms=new boolean[noOfClassRooms][NO_OF_WORKING_DAYS][NO_OF_WORKING_HOURS];
		faculty=new boolean[noOfFaculties][NO_OF_WORKING_DAYS][NO_OF_WORKING_HOURS];
		slot=new boolean[noOfCourses][maximumNoOfSections][NO_OF_WORKING_DAYS][NO_OF_WORKING_HOURS];

		facultyId=new HashMap<>();
		labId=new HashMap<>();
		classroomId=new HashMap<>();
		theoryCodeId=new HashMap<>();
		labCodeId=new HashMap<>();

		lunchHour=4;
		minLectureTime=1;
		maxLectureTime=2;

		section=new String[26];
		for(char c='A';c<='Z';++c)
			section[c-'A']=Character.toString(c);

		combinedSubjects=new ArrayList<>();

	}


	private void loadFacultyDetails() throws IOException{
		Workbook workbook=new HSSFWorkbook(new FileInputStream(PATH_facultyDetails));
		Sheet sheet=workbook.getSheetAt(0);
		int i=0;
		int id=0;
		for(Row row:sheet){
			if(i>1){
				String enrolmentNo=row.getCell(0).getStringCellValue();
				String nameOfTheFaculty=row.getCell(1).getStringCellValue();
				String initialsOfTheFaculty=row.getCell(2).getStringCellValue();
				//System.out.println(enrolmentNo+" "+nameOfTheFaculty+" "+initialsOfTheFaculty);
				facultyId.put(enrolmentNo,id++);
				faculty_initials.put(enrolmentNo,initialsOfTheFaculty);
				faculty_name.put(enrolmentNo,nameOfTheFaculty);
			}
			++i;
		}
	}

	private void loadClassroomDetails() throws IOException{
		Workbook workbook=new HSSFWorkbook(new FileInputStream(PATH_classroomDetails));
		Sheet sheet=workbook.getSheetAt(0);
		int i=0;
		int id=0;
		for(Row row:sheet){
			if(i++>1){
				String classroomName;
				if(row.getCell(0).getCellType()==Cell.CELL_TYPE_NUMERIC)
					classroomName=Integer.toString((int)row.getCell(0).getNumericCellValue());
				else
					classroomName=row.getCell(0).getStringCellValue();

				int classroomCapacity=(int)row.getCell(1).getNumericCellValue();
				//System.out.println(classroomName+" "+classroomCapacity);
				classroomId.put(classroomName,id++);
				classroom_capacity.put(classroomName,classroomCapacity);
			}
		}
	}

	private void loadLabroomDetails() throws IOException{
		Workbook workbook=new HSSFWorkbook(new FileInputStream(PATH_labroomDetails));
		Sheet sheet=workbook.getSheetAt(0);
		int i=0;
		int id=0;
		for(Row row:sheet){
			if(i++>1){
				String labNo;
				if(row.getCell(0).getCellType()==Cell.CELL_TYPE_STRING)
					labNo=row.getCell(0).getStringCellValue();
				else
					labNo=Integer.toString((int)row.getCell(0).getNumericCellValue());
				String labName=row.getCell(1).getStringCellValue();
				//System.out.println(labNo+" "+labName);
				labId.put(labNo,id++);
				lab_room.put(labNo,labName);
			}
		}
	}

	private void generateExcelFileForLabSubjects(String semester) throws IOException{
		ArrayList<String> nameOfTheCourseAndYears=new ArrayList<>();
		Workbook workbook=new HSSFWorkbook(new FileInputStream(PATH_nameOfTheCoursesAndYears));
		Sheet sheet=workbook.getSheetAt(0);
		int i=0;
		for(Row row:sheet)
			if(i++>1)
				nameOfTheCourseAndYears.add(row.getCell(0).getStringCellValue());
		//System.out.println(nameOfTheCourseAndYears);

		int id=0;
		for(String course:nameOfTheCourseAndYears){
			String path=semester+"/"+course+"/main"+labDetails_PATH;
			workbook=new HSSFWorkbook(new FileInputStream(path));
			sheet=workbook.getSheetAt(0);
			i=0;
			HashMap<String,String> code_name=new HashMap<>();
			HashMap<String,String> code_room=new HashMap<>();
			HashMap<String,Integer> code_groups=new HashMap<>();
			int noOfSections=getNoOfSections(semester+"/"+course+"/main"+details_PATH);
			int noOfRows=0;
			for(Row row:sheet){
				if(i++>1){
					String code=row.getCell(0).getStringCellValue();
					String name=row.getCell(1).getStringCellValue();
					String room;
					if(row.getCell(2).getCellType()==Cell.CELL_TYPE_NUMERIC)
						room=Integer.toString((int)row.getCell(2).getNumericCellValue());
					else
						room=row.getCell(2).getStringCellValue();
					int groups=(int)row.getCell(3).getNumericCellValue();
					code_name.put(code,name);
					code_room.put(code,room);
					code_groups.put(code,groups);
					noOfRows+=noOfSections*groups;
				}
			}
			noOfRows+=2;
			if(!code_name.isEmpty()){
				Row[] rows=new Row[noOfRows];
				Cell[][] cells=new Cell[noOfRows][3];
				workbook=new HSSFWorkbook();
				sheet=workbook.createSheet("subjectAssignment");
				for(int idx=0;idx<noOfRows;++idx){
					rows[idx]=sheet.createRow(idx);
					for(int j=0;j<3;++j){
						cells[idx][j]=rows[idx].createCell(j);
					}
				}
				String instruction="Please enter the faculty details and the \n" +
						"no of students for the subject";
				cells[0][0]=instructionCell(workbook,sheet,instruction);

				cells[1][0].setCellValue("Lab Code");
				cells[1][1].setCellValue("Faculty assigned");

				CellStyle style=workbook.createCellStyle();
				style.setFillForegroundColor(IndexedColors.BLACK.getIndex());
				style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				style.setAlignment(HorizontalAlignment.CENTER);
				style.setVerticalAlignment(VerticalAlignment.CENTER);

				Font font=workbook.createFont();
				font.setColor(IndexedColors.WHITE.getIndex());
				font.setFontHeight((short)200);

				style.setFont(font);

				cells[1][0].setCellStyle(style);
				cells[1][1].setCellStyle(style);

				sheet.setColumnWidth(0,9000);
				sheet.setColumnWidth(1,9000);

				int idx=2;
				int hours=3;
				for(String code:code_name.keySet()){
					for(int sections=0;sections<noOfSections;++sections){
						for(int group=1;group<=code_groups.get(code);++group){

							cells[idx++][0].setCellValue(code.trim().concat(section[sections]).concat(Integer.toString(group)));
							lab_hours.put(code.trim().concat(section[sections]).concat(Integer.toString(group)),hours);
							labCode_room.put(code.trim().concat(section[sections]).concat(Integer.toString(group)),code_room.get(code));
							labCodeId.put(code.trim().concat(section[sections]).concat(Integer.toString(group)),id++);
						}
					}

				}
				try{
					String writePath=semester+"/"+course+"/generated"+labDetails_PATH;
					if(fileExists(writePath))return;
					FileOutputStream output=new FileOutputStream(writePath);
					workbook.write(output);
					output.close();
					try {
						Desktop.getDesktop().open(new File(writePath));
					} catch (IOException e) {
						e.printStackTrace();
					}
				}catch(Exception ex){
					ex.printStackTrace();
				}
			}
		}

	}

	private void generateExcelFileForTheorySubjects(String semester) throws IOException{
		ArrayList<String> nameOfTheCourseAndYears=new ArrayList<>();
		Workbook workbook=new HSSFWorkbook(new FileInputStream(PATH_nameOfTheCoursesAndYears));
		Sheet sheet=workbook.getSheetAt(0);
		int i=0;
		for(Row row:sheet)
			if(i++>1)
				nameOfTheCourseAndYears.add(row.getCell(0).getStringCellValue());
		//System.out.println(nameOfTheCourseAndYears);

		for(String course:nameOfTheCourseAndYears){
			String path=semester+"/"+course+"/main"+subjectDetatils_PATH;
			workbook=new HSSFWorkbook(new FileInputStream(path));
			sheet=workbook.getSheetAt(0);
			i=0;
			HashMap<String,String> code_name=new HashMap<>();
			HashMap<String,String> code_isCombined=new HashMap<>();
			HashMap<String,Integer> code_credits=new HashMap<>();
			int noOfSections=getNoOfSections(semester+"/"+course+"/main"+details_PATH);
			int noOfRows=0;
			int id=0;
			for(Row row:sheet){
				if(i++>1){
					String code=row.getCell(0).getStringCellValue();
					String name=row.getCell(1).getStringCellValue();
					String isCombined=row.getCell(2).getStringCellValue();
					int credits=(int)row.getCell(3).getNumericCellValue();
					code_name.put(code,name);
					code_isCombined.put(code,isCombined);
					code_credits.put(code,credits);
					if(isCombined.equals("Y"))++noOfRows;
					else noOfRows+=noOfSections;
				}
			}
			noOfRows+=2;
			if(!code_name.isEmpty()){
				Row[] rows=new Row[noOfRows];
				Cell[][] cells=new Cell[noOfRows][3];
				workbook=new HSSFWorkbook();
				sheet=workbook.createSheet("subjectAssignment");
				for(int idx=0;idx<noOfRows;++idx){
					rows[idx]=sheet.createRow(idx);
					for(int j=0;j<3;++j){
						cells[idx][j]=rows[idx].createCell(j);
					}
				}
				String instruction="Please enter the faculty details and the \n" +
						"no of students for the subject";
				cells[0][0]=instructionCell(workbook,sheet,instruction);

				cells[1][0].setCellValue("Subject code");
				cells[1][1].setCellValue("Faculty assigned");
				cells[1][2].setCellValue("No of students studying the subject");

				CellStyle style=workbook.createCellStyle();
				style.setFillForegroundColor(IndexedColors.BLACK.getIndex());
				style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
				style.setAlignment(HorizontalAlignment.CENTER);
				style.setVerticalAlignment(VerticalAlignment.CENTER);

				Font font=workbook.createFont();
				font.setColor(IndexedColors.WHITE.getIndex());
				font.setFontHeight((short)200);

				style.setFont(font);

				cells[1][0].setCellStyle(style);
				cells[1][1].setCellStyle(style);
				cells[1][2].setCellStyle(style);

				sheet.setColumnWidth(0,9000);
				sheet.setColumnWidth(1,9000);
				sheet.setColumnWidth(2,9000);

				int idx=2;
				for(String code:code_name.keySet()){
					int credits=code_credits.get(code);
					if(code_isCombined.get(code).equals("Y")){
						cells[idx++][0].setCellValue(code);
						theory_credits.put(code,credits);
						theoryCodeId.put(code,id++);
						combinedSubjects.add(code);
					}else{
						for(int j=0;j<noOfSections;++j){
							cells[idx++][0].setCellValue(code.trim().concat(section[j]));
							theory_credits.put(code.trim().concat(section[j]),credits);
							theoryCodeId.put(code.trim().concat(section[j]),id++);
						}
					}
				}
				//System.out.println(theory_credits.keySet()+"\n"+theory_credits.values());
				try{
					String writePath=semester+"/"+course+"/generated"+subjectDetatils_PATH;
					if(fileExists(writePath))return;
					FileOutputStream output=new FileOutputStream(writePath);
					workbook.write(output);
					output.close();
					try {
						Desktop.getDesktop().open(new File(writePath));
					} catch (IOException e) {
						e.printStackTrace();
					}
				}catch(Exception ex){
					ex.printStackTrace();
				}
			}
		}
	}

	private int getNoOfSections(String path) throws IOException{
		Workbook workbook=new HSSFWorkbook(new FileInputStream(path));
		Sheet sheet=workbook.getSheetAt(0);
		int noOfSections=(int)sheet.getRow(1).getCell(1).getNumericCellValue();
		//System.out.println("No of sections "+noOfSections);
		return noOfSections;
	}

	private void loadSemesterDetails(String semester) throws IOException{
		ArrayList<String> nameOfTheCourseAndYears=new ArrayList<>();
		Workbook workbook=new HSSFWorkbook(new FileInputStream(PATH_nameOfTheCoursesAndYears));
		Sheet sheet=workbook.getSheetAt(0);
		int i=0;
		for(Row row:sheet)
			if(i++>1)
				nameOfTheCourseAndYears.add(row.getCell(0).getStringCellValue());

		//System.out.println(nameOfTheCourseAndYears);
		loadTheoryDetails(semester,nameOfTheCourseAndYears);
		loadLabDetails(semester,nameOfTheCourseAndYears);

	}

	private void loadTheoryDetails(String semester,ArrayList<String> nameOfTheCourseAndYears) throws IOException{
		int id=0;
		for(String course:nameOfTheCourseAndYears){
			String path=semester+"/"+course+"/generated"+subjectDetatils_PATH;
			Workbook workbook=new HSSFWorkbook(new FileInputStream(path));
			Sheet sheet=workbook.getSheetAt(0);
			int i=0;
			for(Row row:sheet){
				if(i++>1){
					String code=row.getCell(0).getStringCellValue();
					String faculty=row.getCell(1).getStringCellValue();
					int capacity=(int)row.getCell(2).getNumericCellValue();
					theory_faculty.put(code.trim(),faculty.trim());
					theory_capacity.put(code.trim(),capacity);
				}
			}
			//System.out.println(theory_faculty.keySet()+"\n"+theory_faculty.values()+"\n"+theory_capacity.values());
		}
	}

	private void loadLabDetails(String semester,ArrayList<String> nameOfTheCourseAndYears) throws IOException{
		for(String course:nameOfTheCourseAndYears){
			String path=semester+"/"+course+"/generated"+labDetails_PATH;
			Workbook workbook=new HSSFWorkbook(new FileInputStream(path));
			Sheet sheet=workbook.getSheetAt(0);
			int i=0;
			for(Row row:sheet){
				if(i++>1){
					String code=row.getCell(0).getStringCellValue().trim();
					String faculty=row.getCell(1).getStringCellValue().trim();
					lab_faculty.put(code,faculty);
				}
			}
			//System.out.println(lab_faculty.keySet()+"\n"+lab_faculty.values());
		}
	}

	private void generateLectureScheduler(){
		/*
		revealHashMap(facultyId);
		revealHashMap(labId);
		revealHashMap(classroomId);
		revealHashMap(theoryCodeId);
		revealHashMap(labCodeId);
		revealHashMap(theory_credits);
		revealHashMap(theory_faculty);
		*/
		int noOfFaculties=facultyId.size();
		Faculty[] faculties=new Faculty[noOfFaculties];

		for(int i=0;i<noOfFaculties;++i)faculties[i]=new Faculty(i);

		for(String code:theory_faculty.keySet()){
			int id=facultyId.get(theory_faculty.get(code));
			faculties[id].addSubjects(code,theory_credits.get(code));
		}
		for(String code:lab_faculty.keySet()){
			int id=facultyId.get(lab_faculty.get(code));
			faculties[id].addLabs(code);
		}

	}



	public int calculateNoOfRows(String path) throws IOException{
		Workbook workbook=new HSSFWorkbook(new FileInputStream(path));
		Sheet sheet=workbook.getSheetAt(0);
		int rows=0;
		for(Row row:sheet)++rows;
		//System.out.println(path+"has "+rows+" no of rows");
		return rows;
	}

	public Cell instructionCell(Workbook workbook,Sheet sheet,String instructionValue){
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

	public boolean fileExists(String path){
		File file=new File(path);
		if(file.exists()){
			//System.out.println(path+ " already exists");
			return true;
		}
		return false;
	}

	public void revealHashMap(HashMap map){
		System.out.println("So the hashmap "+map.toString());
		for(Object key:map.keySet()){
			System.out.println(map.get(key)+" "+key);
		}
	}

}
