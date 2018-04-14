package dataStructure;

import java.util.ArrayList;
import java.util.Collections;

public class Faculty{
	private int id;
	private ArrayList<String> subjectCodes;
	private ArrayList<String> labCodes;
	private int theoryHours;

	public Faculty(int id){
		this.id=id;
		subjectCodes=new ArrayList<>();
		labCodes=new ArrayList<>();
		theoryHours=0;
	}

	public void addSubjects(String subjectCode, int credits){
		theoryHours+=credits;
		subjectCodes.add(subjectCode);
	}
	
	public void addLabs(String labCode){
		labCodes.add(labCode);
	}

	public ArrayList<String> getSubjectCodes(){
		return subjectCodes;
	}

	public ArrayList<String> getLabCodes(){
		return labCodes;
	}

	public int getTheoryHours(){
		return theoryHours;
	}

	public void decreaseHours(int dec){
		theoryHours-=dec;
	}

	public void shuffleSubjectList(){
		Collections.shuffle(subjectCodes);
	}
}