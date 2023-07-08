package org.example;

import java.io.File;
import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


// Press Shift twice to open the Search Everywhere dialog and type `show whitespaces`,
// then press Enter. You can now see whitespace characters in your code.
public class Main {

    public static final String path_excel = "C:\\Users\\tusar\\PROJECT\\MTP_POC_JAVA_PG_EXCEL\\UserDetails.xlsx";
    public static void main(String[] args) {
        // Press Alt+Enter with your caret at the highlighted text to see how
        // IntelliJ IDEA suggests fixing it.
        try{
            FileInputStream file = new FileInputStream(new File(path_excel));
            Workbook wb = new XSSFWorkbook(file);
            Sheet sheet0 = wb.getSheet("Sheet1");
            Integer rowCount = sheet0.getPhysicalNumberOfRows();

            System.out.println(rowCount);

            try{
                Connection connection =null;
                Class.forName("org.postgresql.Driver");
                connection = DriverManager.getConnection("jdbc:postgresql://localhost:5432/MTPUsers","postgres","tusar123");

                if(connection!=null){
                    System.out.println("Database connected successfully");
                    PreparedStatement ps = connection.prepareStatement("select * from gco_users");
                    ResultSet rs = ps.executeQuery();
                    while(rs.next()){
                        Integer wwid = rs.getInt("wwid");
                        System.out.println("wwid is : "+wwid);
                    }

                    ps = connection.prepareStatement("INSERT INTO public.gco_users(wwid, email, manager_wwid, manager_email) VALUES (?, ?, ?,?);");
                    //Insert the excel data to DB
                    for(int i=1;i<rowCount;i++){
                        final Row row = sheet0.getRow(i);
                        String wwid = String.valueOf(row.getCell(0));
                        String email = String.valueOf(row.getCell(1));
                        String manager_wwid = String.valueOf(row.getCell(2));
                        String manager_email = String.valueOf(row.getCell(3));
                        System.out.println("insert into user_gco"+row.getCell(0)+row.getCell(1)+row.getCell(2)+row.getCell(3));
                        ps.setString(1,wwid);
                        ps.setString(2,email);
                        ps.setString(3,manager_wwid);
                        ps.setString(4,manager_email);
                        ps.executeUpdate();
                    }

                }
            }
            catch (Exception e){
                System.out.println(e);
            }


        }
        catch(Exception e){
        e.printStackTrace();
        }
        System.out.printf(path_excel);


    }
}