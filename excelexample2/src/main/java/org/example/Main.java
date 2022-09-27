package org.example;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.*;


import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Date;
import java.time.Instant;
import java.time.LocalDate;
import java.time.LocalDateTime;


public class Main {
    public static void main(String[] args) {
    /*
        //Create a workbook
        XSSFWorkbook workbook = new XSSFWorkbook();

        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("Employee data");

        //long[] serialNr = new long[10];
        // for (long i = 0; i < serialNr.length; i++ ) {
        //}

        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[] {"ID", "NAME", "LASTNAME"});
        data.put("2", new Object[] {1, "Amit", "Shukla"});
        data.put("3", new Object[] {2, "Lokesh", "Gupta"});
        data.put("4", new Object[] {3, "John", "Adwards"});
        data.put("5", new Object[] {4, "Brian", "Schultz"});

        //Iterate over data and write to sheet
        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset)
        {
            Row row = sheet.createRow(rownum++);
            Object [] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr)
            {
                Cell cell = row.createCell(cellnum++);
                if(obj instanceof String)
                    cell.setCellValue((String)obj);
                else if(obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }
        try
        {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File("howtodoinjava_demo.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("howtodoinjava_demo.xlsx written successfully on disk.");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }

     */
        long startTime = System.nanoTime();

        //Alternativ 2
        int[] serial = new int[6];
        for (int i = 0; i< serial.length; i++) {
            serial[i] = i+1;
        }
        String[] name = new String[10];
        name[0] = "Kristian";
        name[1] = "Marcus";
        name[2] = "Richard";
        name[3] = "Stefan";
        name[4] = "Rasmus";
        name[5] = "Jennifer";

        String[] result = new String[10];
        result[0] = "Junior";
        result[1] = "Junior";
        result[2] = "Senior";
        result[3] = "Senior";

        String[] styling = new String[10];


        styling[0] = "Sample";
        styling[1] = "Sample 1";
        styling[2] = "Sample";
        styling[3] = "Sample";
        styling[4] = "Helllo";

        //Create a workbook
        XSSFWorkbook workbook1 = new XSSFWorkbook();
        //Create a spreedsheet
        XSSFSheet sheet2 = workbook1.createSheet("hello");
        //Crate a Row Object
        XSSFRow row; //Horizontel


        //Create cells & set values - Labels
        row = sheet2.createRow(0);
        XSSFCell cell0 = (XSSFCell)row.createCell(0);
        XSSFCell cell1 = (XSSFCell)row.createCell(1);
        XSSFCell cell2 = (XSSFCell)row.createCell(2);
        XSSFCell cell3 = row.createCell(3);
        XSSFCellStyle style5 = workbook1.createCellStyle();
        row.setHeight((short)500); //Setting row-height



        cell0.setCellValue("Serial no." );
        cell1.setCellValue("Name of the people." );
        cell2.setCellValue("Level" );
        cell3.setCellValue("Sample output");

        //Create cells & rows for data
        for(int i = 0; i <serial.length; i++) {
            row = sheet2.createRow(i+1);;

            LocalDateTime now = LocalDateTime.now();


            for (int j = 0; j < 4; j++) {
                Cell cell = row.createCell(j);


                if (cell.getColumnIndex() == 0) {
                    cell.setCellValue(serial[i]);
                }
                else if (cell.getColumnIndex() == 1) {
                    cell.setCellValue(name[i]);
                }
                else if (cell.getColumnIndex() == 2) {
                    cell.setCellValue(result[i]);
                }
                else if (cell.getColumnIndex() == 3) {
                    cell.setCellValue(styling[i]);
                }
            }
        }

        //Writing the created excel file
        try {
            File f = new File("Results.xlsx");
            if (f.exists()) {
                FileOutputStream newOut = new FileOutputStream((LocalDate.now() + " - " + "results.xlsx"));
                workbook1.write(newOut);
                newOut.close();
                System.out.println("File is allready created before, file added your data.");
                long endTime = System.nanoTime();
                long timeElapsed = endTime - startTime;
                System.out.println("Execution time in milliseconds: " + timeElapsed / 1000000);

            } else if (!f.exists()) {
                File fn = new File("Results.xlsx");
                FileOutputStream out = new FileOutputStream((fn));
                workbook1.write(out);
                out.close();
                System.out.println("Excel file is created sucessfully!");
            }
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
