package functionalities;

import java.io.IOException;

public interface Operator{
	/**
	 *
	 */
	default void callForInitialization(){};
	default void initializeTheDatabase(){};
	default void showLabDatabase(){};
	default void showSubjectDatabase(){};
	default void showGeneratedSubjectFiles(){};
	default void showGeneratedLabFiles(){};
	default void showClassRoomDetails(){};
	default void showLabRoomsDetails(){};
	default void printFacultyDetails(){};
	default void goForOddSemester(){};
	default void goForEvenSemester(){};
	default void twoHoursContinuouslyAvailable(){};
	default void assignLabSubjects(){};
	default void assignTheorySubjects(){};
	default void printTheRoutine(){};
}
