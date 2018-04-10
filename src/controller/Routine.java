package controller;

import functionalities.Operator;

import java.util.HashMap;

public class Routine implements Operator{

	private HashMap<String,String> theory_faculty;
	private HashMap<String,String> lab_faculty;
	private HashMap<String,String> lab_room;
	private HashMap<String,String> theory_capacity;

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

	public Routine(){

	}

	private void initialize(){
		theory_faculty=new HashMap<>();
		lab_faculty=new HashMap<>();
		lab_room=new HashMap<>();
		theory_capacity=new HashMap<>();
	}


}
