import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.*;

public class ExcelToList {
    public static void excelToList() throws IOException {
        File file = new File("C:\\Users\\ef-lorenz\\Desktop\\test.xlsx");
        FileInputStream fis = new FileInputStream(file);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheetAt(0);
        Iterator<Row> itr = sheet.iterator();

        List<Employee> employeeList = new ArrayList<>();

        while (itr.hasNext()) {
            Row row = itr.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();

                Employee employee = new Employee();

                switch (cell.getCellType()) {
                    case NUMERIC:
                        System.out.print(cell.getNumericCellValue() + "\t\t");
                        break;
                    case STRING:
                        System.out.print(cell.getStringCellValue() + "\t\t");
                        break;
                    case BOOLEAN:
                        System.out.print(cell.getBooleanCellValue() + "\t\t");
                        break;
                    case BLANK:
                        break;
                    default:
                }

            }

            System.out.println("");
        }
    }

    public static void searchEmployee(String str) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(("C:\\Users\\ef-lorenz\\Desktop\\test.xlsx")));
        XSSFSheet sheet = workbook.getSheetAt(0);
        try {
            for (Row row : sheet) {
                Iterator<Cell> cellIterator = row.cellIterator();
                if (row.getCell(0).toString().equals(str)) {
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        System.out.println(row.getCell(5));
                        switch (cell.getCellType()) {
                            case NUMERIC:
                                System.out.print(cell.getNumericCellValue() + " ");
                                break;
                            case STRING:
                                System.out.print(cell.getStringCellValue() + " ");
                                break;
                        }
                    }
                    if (!cellIterator.hasNext()) {
                        break;
                    }
                }
            }
            System.out.println();
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            workbook.close();
        }
    }

    public static void employeeExcelToObjectJava() throws IOException {
        try {
            FileInputStream file = new FileInputStream(new File("C:\\Users\\ef-lorenz\\Desktop\\test.xlsx"));
            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);
            ArrayList<Employee> employeeList = new ArrayList<>();
            //I've Header and I'm ignoring header for that I've +1 in loop
            for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
                Employee e = new Employee();
                Row ro = sheet.getRow(i);
                for (int j = 0; j <= ro.getLastCellNum(); j++) {
                    Cell ce = ro.getCell(j);
                    if (j == 0) {
                        //If you have Header in text It'll throw exception because it won't get NumericValue
                        e.setEmployeeId(ce.getStringCellValue());
                    }
                    if (j == 1) {
                        e.setFullName(ce.getStringCellValue());
                    }
                    if (j == 2) {
                        e.setNickName(ce.getStringCellValue());
                    }
                    if (j == 3) {
                        e.setNewNickname(ce.getStringCellValue());
                    }
                    if (j == 4) {
                        e.setUuid(ce.getNumericCellValue());
                    }
                    if (j == 5) {
                        e.setDepositAddress(ce.getStringCellValue());
                    }
                }
                if (e.getUuid() == null) {
                    break;
                }
                employeeList.add(e);
            }
            for (int i = 0; i < employeeList.size() - 1; i++) {
                System.out.println(employeeList.get(i));
            }
            file.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void guidesExcelToObjectJava() throws IOException {
        try {
            FileInputStream file = new FileInputStream(new File("C:\\Users\\ef-lorenz\\Desktop\\Guides2.xlsx"));
            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);
            ArrayList<Guides> guidesArrayList = new ArrayList<>();
            //I've Header and I'm ignoring header for that I've +1 in loop
            for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
                Guides guides = new Guides();
                Row ro = sheet.getRow(i);
                for (int j = 0; j <= ro.getLastCellNum(); j++) {
                    Cell ce = ro.getCell(j);
                    if (j == 0) {
                        guides.setLegalTopic(ce.getStringCellValue());
                    }
                    if (j == 1) {
                        guides.setSubTopic(ce.getStringCellValue());
                    }
                    if (j == 2) {
                        guides.setTitle(ce.getStringCellValue());
                    }
                    if (j == 3) {
                        guides.setSummary(ce.getStringCellValue());
                    }
                    if (j == 4) {
                        guides.setContent(ce.getStringCellValue());
                    }
                }
                guidesArrayList.add(guides);
            }
            for (int i = 0; i < guidesArrayList.size() - 1; i++) {
                System.out.println("Legal Topic = " + i + " " + guidesArrayList.get(i).getLegalTopic());
                System.out.println("Sub Topic = " + i + " " + guidesArrayList.get(i).getSubTopic());
                System.out.println("Title = " + i + " " + guidesArrayList.get(i).getTitle());
                System.out.println("Summary = " + i + " " + guidesArrayList.get(i).getSummary());
                System.out.println("Content = " + i + " " + guidesArrayList.get(i).getContent());
            }
            file.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void main(String[] args) throws IOException {
//        excelToList();
//        searchEmployee("ZFYW93");

//        employeeExcelToObjectJava();
        guidesExcelToObjectJava();
    }

}
