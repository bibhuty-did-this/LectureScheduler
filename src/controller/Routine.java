package controller;

import functionalities.Operator;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
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
	private HashMap<String,String> lab_faculty;
	private HashMap<String,String> lab_room;
	private HashMap<String,String> theory_capacity;
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

	private HashMap<String,Integer> facultyId;
	private HashMap<String,Integer> labId;
	private HashMap<String,Integer> classroomId;

	public Routine() throws IOException{

		initialize();
		loadFacultyDetails();
		loadClassroomDetails();
		loadLabroomDetails();
		generateExcelFileForTheorySubjects();
		generateExcelFileForLabSubjects();
		//loadSemesterDetails(ODD);
	}

	private void initialize() throws IOException{
		theory_faculty=new HashMap<>();
		lab_faculty=new HashMap<>();
		lab_room=new HashMap<>();
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

		lunchHour=4;

	}


	private void loadFacultyDetails() throws IOException{
		Workbook workbook=new HSSFWorkbook(new FileInputStream(PATH_facultyDetails));
		Sheet sheet=workbook.getSheetAt(0);
		int i=0;
		for(Row row:sheet){
			if(i>1){
				String enrolmentNo=row.getCell(0).getStringCellValue();
				String nameOfTheFaculty=row.getCell(1).getStringCellValue();
				String initialsOfTheFaculty=row.getCell(2).getStringCellValue();
				System.out.println(enrolmentNo+" "+nameOfTheFaculty+" "+initialsOfTheFaculty);
				facultyId.put(enrolmentNo,i-1);
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
		for(Row row:sheet){
			if(i++>1){
				String classroomName=row.getCell(0).getStringCellValue();
				int classroomCapacity=Integer.parseInt(row.getCell(1).getStringCellValue());
				//System.out.println(classroomName+" "+classroomCapacity);
				classroomId.put(classroomName,i-1);
				classroom_capacity.put(classroomName,classroomCapacity);
			}
		}
	}

	private void loadLabroomDetails() throws IOException{
		Workbook workbook=new HSSFWorkbook(new FileInputStream(PATH_labroomDetails));
		Sheet sheet=workbook.getSheetAt(0);
		int i=0;
		for(Row row:sheet){
			if(i++>1){
				String labNo=row.getCell(0).getStringCellValue();
				String labName=row.getCell(1).getStringCellValue();
				System.out.println(labNo+" "+labName);
				labId.put(labNo,i-1);
				lab_room.put(labNo,labName);
			}
		}
	}

	private void generateExcelFileForLabSubjects(){

	}

	private void generateExcelFileForTheorySubjects(){

	}

	private void loadSemesterDetails(String semester) throws IOException{
		ArrayList<String> nameOfTheCourseAndYears=new ArrayList<>();
		Workbook workbook=new HSSFWorkbook(new FileInputStream(PATH_nameOfTheCoursesAndYears));
		Sheet sheet=workbook.getSheetAt(0);
		int i=0;
		for(Row row:sheet)
			if(i++>1)
				nameOfTheCourseAndYears.add(row.getCell(0).getStringCellValue());

		System.out.println(nameOfTheCourseAndYears);
		loadTheoryDetails(semester,nameOfTheCourseAndYears);
		loadLabDetails(semester,nameOfTheCourseAndYears);

	}

	private void loadTheoryDetails(String semester,ArrayList<String> nameOfTheCourseAndYears) throws IOException{
		for(String course:nameOfTheCourseAndYears){
			String path=semester+"/"+course+"/generated"+subjectDetatils_PATH;
			Workbook workbook=new HSSFWorkbook(new FileInputStream(path));
			Sheet sheet=workbook.getSheetAt(0);
			int i=0;
			for(Row row:sheet){
				if(i++>1){

				}
			}
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

				}
			}
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

}
