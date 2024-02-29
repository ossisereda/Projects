package fi.tuni.prog3.sisu;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.node.ArrayNode;
import java.net.URL;
import java.io.IOException;
import java.util.List;
import java.util.TreeSet;

/**
 * Toteuttaa rajapinnan, jonka avulla haetaan tutkinto-ohjelmien
 * tiedot Sisun tietokannasta.
 * @author Buraq
 */

public class SisuApi implements iAPI {
private Programme programme;
private StudyModule studyModule;
   /**
     * Returns a JsonObject that is extracted from the Sisu API.
     * @param urlString URL for retrieving information from the Sisu API.
     * @return JsonObject.
     * 
     */

//Hakee tietoa annetusta verkko-osoitteesta ja palauttaa JsonNode arvon. 
    public JsonNode getJsonObjectFromApi(String urlString) {
           try {
                URL url = new URL(urlString);
                com.fasterxml.jackson.databind.ObjectMapper objectMapper 
                        = new com.fasterxml.jackson.databind.ObjectMapper();
                
                JsonNode jsonNode = objectMapper.readTree(url);
      
             return jsonNode;

        } catch (IOException e) {
            e.printStackTrace();
        }
    return null;
    
    }
    
public TreeSet<Programme> getDegreeProgrammes() throws IOException{

        TreeSet<Programme> listOfProgrammes = new TreeSet<>();  
        JsonNode jsonNode = getJsonObjectFromApi("https://sis-tuni.funidata.fi/kori/api/module-search?curriculumPeriodId=uta-lvv-2021&universityId=tuni-university-root-id&moduleType=DegreeProgramme&limit=1000");
        JsonNode searchResults = jsonNode.get("searchResults");

        int a = 0; 
        for( JsonNode i : searchResults){
        a++;          
        System.out.println(a);
        Programme programme = new Programme(i.get("name").asText(),
                                  i.get("id").asText(),
                                  i.get("groupId").asText(),
                                  i.get("credits")
                                          .get("min").asInt());
        listOfProgrammes.add(programme);
                  }
        
        return listOfProgrammes;
    }
    public Programme getModule(Programme programme, int i) throws JsonProcessingException{
       
        this.programme = programme;
        String urlString = String.format("https://sis-tuni.funidata.fi/kori/api/modules/by-group-id?groupId=%s&universityId=tuni-university-root-id", programme.getGroupId());
        
        JsonNode jsonModule = getJsonObjectFromApi(urlString);
        if(jsonModule.get(0).has("rule")){
            if(i==1){
           this.studyModule = new StudyModule(jsonModule.get(0).get("name").get("fi").asText(),
                                               jsonModule.get(0).get("code").asText(),
                                               jsonModule.get(0).get("groupId").asText(),
                                               jsonModule.get(0).get("targetCredits").get("min").asInt());
       }
            JsonNode rule = jsonModule.get(0).get("rule");
            whichRuleRecursive(rule);
        }

        StudyModule currentModule = this.studyModule;
        
        if(currentModule.geStudyModulesize() > 0){
           this.programme.add(currentModule);
        }

        
    return this.programme;
    }

    private void whichRuleRecursive(JsonNode rule) throws JsonProcessingException {
        if(rule.get("type").asText().equals("CreditsRule")){
            creditsRuleRecursive(rule);
        }else if(rule.get("type").asText().equals("CompositeRule")){
            compositeRuleRecursive(rule);
        }else if(rule.get("type").asText().equals("ModuleRule")){
            moduleRuleRecursive(rule);
        }else if(rule.get("type").asText().equals("CourseUnitRule")){
            courseUnitRuleRecursive(rule);
        }else if(rule.get("type").asText().equals("StudyModule")
                || rule.get("type").asText().equals("GroupingModule")){
            whichRuleRecursive(rule.get("rule"));
        }
    }
    
    private void creditsRuleRecursive(JsonNode rule) throws JsonProcessingException {
        JsonNode nextRule = rule.get("rule");
        whichRuleRecursive(nextRule);
    }
    
    private void compositeRuleRecursive(JsonNode rule) throws JsonProcessingException {
  
    com.fasterxml.jackson.databind.ObjectMapper objectMapper = new com.fasterxml.jackson.databind.ObjectMapper();
  
    List<JsonNode> rulesList = rule.findValues("rules");
    ArrayNode rulesArray = objectMapper.valueToTree(rulesList);
    for (JsonNode nextRule : rulesArray) {
        
        for(JsonNode i:nextRule){
            whichRuleRecursive(i);
        }
       
    }
}  

    private void moduleRuleRecursive(JsonNode rule) throws JsonProcessingException {        
        if(this.studyModule != null){
            if(this.studyModule.geStudyModulesize()>0){
                StudyModule studMod = this.studyModule;
                this.programme.add(studMod);
            }
        }
        String urlString = String.format("https://sis-tuni.funidata.fi/kori/api/modules/by-group-id?groupId=%s&universityId=tuni-university-root-id", rule.get("moduleGroupId").asText());
        JsonNode jsonModule = getJsonObjectFromApi(urlString).get(0);
        String name = "";
        String code= "";
        String groupId = jsonModule.get("groupId").asText();
        int minCredits = 0;
        JsonNode jsonName = jsonModule.get("name");
        if(jsonName.has("fi")){
            name = jsonName.get("fi").asText();
        }else{
            name = jsonName.get("en").asText();
        }
        
       
        if(jsonModule.get("code").isNull()){    
        }else {
            code = jsonModule.get("code").asText();
        }
        
        if(jsonModule.has("targetCredits")){
            JsonNode jsonMinCredits = jsonModule.get("targetCredits");
            if(jsonMinCredits.get("min").isNull()){
            }else{
                minCredits = jsonMinCredits.get("min").asInt();
            }
        }
        this.studyModule = new StudyModule(name, code, groupId, minCredits);
        whichRuleRecursive(jsonModule);

    }

    private void courseUnitRuleRecursive(JsonNode rule) {
        String urlString = String.format("https://sis-tuni.funidata.fi/kori/api/course-units/by-group-id?groupId=%s&universityId=tuni-university-root-id", rule.get("courseUnitGroupId").asText());
        JsonNode jsonCourse = getJsonObjectFromApi(urlString).get(0);
        JsonNode jsonName = jsonCourse.get("name");
        String name = "";
        String code = "";
        String groupId = jsonCourse.get("groupId").asText();
        int minCredits = jsonCourse.get("credits")
                                    .get("min").asInt();
        
        if(jsonName.has("fi")){
            name = jsonName.get("fi").asText();
        }else{
            name = jsonName.get("en").asText();
        }
        
        
        if(jsonCourse.get("code").isNull()){

            
        }else{
            code = jsonCourse.get("code").asText();
        }

        CourseUnit course = new CourseUnit(name, code, groupId, minCredits);
        this.studyModule.addCourse(course);
    }   
}