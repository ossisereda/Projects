package fi.tuni.prog3.sisu;

/**
 * Vertaa Sisusta luettuja kursseja keskenään, auttaa järjestelyssä
 * @author Buraq
 */

public class CourseUnit extends DegreeModule implements Comparable<CourseUnit>{
    private String name;
    private int minCredits;
    private String code;
    private String groupId;
    private boolean isChecked = false;
    
    
    public CourseUnit(String name, String code, String groupId, int minCredits) {
        super(name, code, groupId, minCredits);
        
    this.name=name;
    this.minCredits=minCredits;
    this.groupId=groupId;
    this.code=code;
    }
    
    @Override
    public int compareTo(CourseUnit other) {
        int result = this.name.compareTo(other.name);
        if (result == 0) {
            result = this.groupId.compareTo(other.groupId);
        }
        return result;
    }
    
    public Boolean isChecked(){
        return isChecked;
        
    }
    public void changeChecked(Boolean checkedOrNot){
        isChecked = checkedOrNot;
    }
    
    
    @Override
    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    @Override
    public int getMinCredits() {
        return minCredits;
    }

    public void setMinCredits(int minCredits) {
        this.minCredits = minCredits;
    }

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    @Override
    public String getGroupId() {
        return groupId;
    }

    public void setGroupId(String groupId) {
        this.groupId = groupId;
    }
}
