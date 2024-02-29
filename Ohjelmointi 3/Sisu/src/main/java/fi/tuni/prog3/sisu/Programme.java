package fi.tuni.prog3.sisu;

import java.util.TreeSet;

/**
 * Vertaa Sisusta luettuja tutkinto-ohjelmia keskenään, auttaa järjestelyssä
 * @author Buraq
 */

public class Programme extends DegreeModule implements Comparable<Programme>{
    private String name;
    private String id;
    private String groupId;
    private int minCredits;  
    private TreeSet<StudyModule> studyModuleList = new TreeSet<>();

    public Programme(String name, String id, String groupId, int minCredits) {
        super(name, id, groupId, minCredits);
        
        this.name = name;
        this.id = id;
        this.groupId = groupId;
        this.minCredits = minCredits;
     

   
        this.name = name;
        this.id = id;
        this.groupId = groupId;
        this.minCredits = minCredits;
    }
    
    @Override
    public int compareTo(Programme other) {
        int result = this.name.compareTo(other.name);
        if (result == 0) {
            result = this.id.compareTo(other.id);
        }
        return result;
    }
    
    public void add(StudyModule studymodule){
    this.studyModuleList.add(studymodule);
    }
    
    public TreeSet<StudyModule> getStudyModule(){
        return this.studyModuleList;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
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
