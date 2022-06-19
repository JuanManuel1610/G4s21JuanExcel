/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Main.java to edit this template
 */
package g4s21juanexcel;

import java.io.FileOutputStream;
import java.io.OutputStream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
/**
 *
 * @author user
 */
public class G4s21JuanExcel {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        // TODO code application logic here
        Workbook wb = new HSSFWorkbook();
        try ( OutputStream fileOut = new FileOutputStream("archivo.xls")) {
            Sheet sheet1 = wb.createSheet("Primer Hoja");
            Sheet sheet2 = wb.createSheet("Segunda Hoja");
            Sheet sheet = wb.createSheet("Tercer Hoja");
            Row row = sheet.createRow(0);                      
            row.createCell(0).setCellValue("Numero");  
            row.createCell(1).setCellValue("Nombre");    
            row.createCell(2).setCellValue("Edad");  
            row.createCell(3).setCellValue("Correo");  

            row = sheet.createRow(1); 
            row.createCell(0).setCellValue(1); 
            row.createCell(1).setCellValue("Juan Manuel");   
            row.createCell(2).setCellValue(20); 
            row.createCell(3).setCellValue("juan.alonzo@gmail.com"); 
            
            row = sheet.createRow(2); 
            row.createCell(0).setCellValue(2);  
            row.createCell(1).setCellValue("Dulce Citlalli");  
            row.createCell(2).setCellValue(18);
            row.createCell(3).setCellValue("dulceciaca@gmail.com");
            
                         
            
            wb.write(fileOut);
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }
    }
    

