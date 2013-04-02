/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package process;

import com.extentech.ExtenXLS.CellHandle;
import com.extentech.ExtenXLS.WorkBookHandle;
import com.extentech.ExtenXLS.WorkSheetHandle;
import com.extentech.formats.XLS.CellNotFoundException;
import com.extentech.formats.XLS.WorkBook;
import com.extentech.formats.XLS.WorkBookFactory;
import com.extentech.formats.XLS.WorkSheetNotFoundException;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import ui.MainFrame1;

/**
 *
 * @author Nikhil
 */
public class SpreadSheetCreator {
    
    List<StudentResult> results;
   
    public SpreadSheetCreator(List<StudentResult> results){
        this.results=results;
    }
    
    public void createXLS(String fileName){
        FileOutputStream fo = null;
        int totalColumns;
        try {
            WorkBookHandle h=new WorkBookHandle();
            WorkSheetHandle sheet1 = h.getWorkSheet("Sheet1");
            
            int currentRow=1;
            char currentColumn;
            int underSessions;
            for(StudentResult i:results){
                if(results.indexOf(i)==0){
                    sheet1.add("Register No.", "A1");
                    sheet1.add("Name of Student", "B1");
                    currentColumn='C';
                    for(ExamResult j:i.subjectResults){
                        sheet1.add(j.subCode, String.valueOf(currentColumn)+String.valueOf(currentRow));
                        currentColumn++;
                    }
                    sheet1.add("SGPA", String.valueOf(currentColumn)+String.valueOf(currentRow));
                    currentColumn++;
                    sheet1.add("No. of Under-sessions", String.valueOf(currentColumn)+String.valueOf(currentRow));
                    totalColumns=currentColumn;
                }
                
                currentRow++;
                sheet1.add(i.regNo,"A"+String.valueOf(currentRow));
                sheet1.add(i.studentName,"B"+String.valueOf(currentRow));
                currentColumn='C';
                underSessions=0;
                for(ExamResult j:i.subjectResults){
                    sheet1.add(j.grade, String.valueOf(currentColumn)+String.valueOf(currentRow));
                    if(MainFrame1.UNDER_SESION_GRADES.contains(j.grade)){
                        underSessions++;
                    } 
                    currentColumn++;
                }
                sheet1.add(i.averageGradePoint, String.valueOf(currentColumn)+String.valueOf(currentRow));
                currentColumn++;
                sheet1.add(underSessions, String.valueOf(currentColumn)+String.valueOf(currentRow)); 
            }
            fo = new FileOutputStream("Result.xls");
            h.writeBytes(fo);
            fo.close();
        } catch (IOException ex) {
            Logger.getLogger(SpreadSheetCreator.class.getName()).log(Level.SEVERE, null, ex);
        }  catch (WorkSheetNotFoundException ex) {
            Logger.getLogger(SpreadSheetCreator.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                fo.close();
            } catch (IOException ex) {
                Logger.getLogger(SpreadSheetCreator.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }
    
     //For unit testing only
    public static void main(String args[]) throws WorkSheetNotFoundException, CellNotFoundException, FileNotFoundException, IOException{
        /*
        //Test 1 - Creating empty XLS
        FileOutputStream fo=new FileOutputStream("Result.xls");
        WorkBookHandle h=new WorkBookHandle();
        WorkSheetHandle y = h.getWorkSheet("Sheet1");
        y.add("aa", "A1");
        h.writeBytes(fo);
        fo.close();
        */  
        
    }
    
}
