package fi.tuni.prog3.sisu;

import java.util.TreeSet;

/**
 * Vertaa Sisusta luettuja kurssimoduuleja keskenään, auttaa järjestelyssä
 * @author Buraq
 */

public class StudyModule extends DegreeModule implements Comparable<StudyModule> {
    private String name;
    private String groupId;
    private int minCredits;
    private String code;
    
    private TreeSet<CourseUnit> courseList = new TreeSet<>(); 
    
    public StudyModule(String name, String code, String groupId, int minCredits) {
        super(name, null, groupId, minCredits);
       
        this.name=name;
        this.groupId=groupId;
        this.minCredits=minCredits;
        this.code=code;
    } 
    @Override
    public int compareTo(StudyModule other) {
        // Define the comparison logic based on the programme's attributes
        int result = this.name.compareTo(other.name);
        if (result == 0) {
            result = this.groupId.compareTo(other.groupId);
        }
        return result;
    }
    public int getCoursesPoints(){
        int points = 0;
        for(CourseUnit i :this.courseList){
            if(i.isChecked())
            points += i.getMinCredits();
        }
        return points;
    }
    
    public boolean isRequirementFulfilled() {
        if(getMinCredits() > 0)
        if(getCoursesPoints() >= getMinCredits()){
            return true;
        }
        if(getMinCredits() == 0){
            if(getCoursesPoints() > getMinCredits()){
            return true;
        }
        }
        return false;
    }
      
    public void addCourse(CourseUnit course){
    this.courseList.add(course);
    }
    
  
    
    public TreeSet<CourseUnit> getCourses(){
        return this.courseList;
    }
    
    public int geStudyModulesize(){
        
        return this.courseList.size();
    }

    public String getCode() {
        return code;
    }

    public void setCode(String type) {
        this.code = code;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getGroupId() {
        return groupId;
    }

    public void setGroupId(String groupId) {
        this.groupId = groupId;
    }

    public int getMinCredits() {
        return minCredits;
    }

    public void setMinCredits(int minCredits) {
        this.minCredits = minCredits;
    }
  
}
