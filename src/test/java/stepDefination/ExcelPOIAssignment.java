package stepDefination;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelPOIAssignment {

    private static final String Excel_Name = "./src/test/java/utilities/Excelbook1.xlsx";

    public static void main(String[] args) {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Employee Data");
        Object[][] EmpData = {
                {"EmployeeNum", "EmployeeName", "Salary"},
                {"101", "Emp101", 2100},
                {"102", "Emp102", 2200},
                {"103", "Emp103", 2300},
                {"104", "Emp104", 2400},
                {"105", "Emp105", 2500},

        };

        XSSFSheet sheet1 = workbook.createSheet("Department Data");

        Object[][] DeptData = {
                {"DeptNum", "DeptName","DeptLocation"},
                {"10", "Dept10", "India"},
                {"20", "Dept20", "UK"},
                {"30", "Dept30", "USA"},
                {"40", "Dept40", "Japan"},
                {"50", "Dept50", "Russia"},

        };


        System.out.println("Creating and Printing excel");

        int rowNum = 0;

        for (Object[] EmpData1 : EmpData) {
            Row row = sheet.createRow(rowNum++);
            int colNum = 0;
            for (Object field : EmpData1) {
                Cell cell = row.createCell(colNum++);
                if (field instanceof String)
                {
                    cell.setCellValue((String) field);
                    System.out.print(cell.getStringCellValue()+" ");
                } else if (field instanceof Integer)
                {
                    cell.setCellValue((Integer) field);
                    System.out.print(cell.getNumericCellValue()+" ");
                }
            }
            System.out.println();
        }

        int rowNum1 = 0;

        for (Object[] Deptdata1 : DeptData) {
            Row row = sheet1.createRow(rowNum1++);
            int colNum = 0;
            for (Object field : Deptdata1) {
                Cell cell = row.createCell(colNum++);
                if (field instanceof String)
                {
                    cell.setCellValue((String) field);
                    System.out.print(cell.getStringCellValue()+" ");
                } else if (field instanceof Integer)
                {
                    cell.setCellValue((Integer) field);
                    System.out.print(cell.getNumericCellValue()+" ");
                }
            }
            System.out.println();
        }

        try {
            FileOutputStream outputStream = new FileOutputStream(Excel_Name);
            workbook.write(outputStream);
            workbook.close();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        System.out.println("Done");

    }
}
