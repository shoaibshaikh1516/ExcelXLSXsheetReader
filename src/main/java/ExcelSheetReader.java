import model.User;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Date;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

public class ExcelSheetReader {
    public static void main(String[] args) throws Exception {

        //courtesy: http://file-examples.com/ for providing sample data
        Path rootLocation = Paths.get("src/main/resources/data/Sampledata.xlsx");
//        System.out.println("" + rootLocation.toAbsolutePath());
        List<User> users = getExcelData(rootLocation);  //Main Function
        System.out.println(users);

        users.forEach(System.out::println);

    }

    private static List<User> getExcelData(Path rootLocation) throws Exception {
        InputStream ExcelFileToRead = new FileInputStream(String.valueOf(rootLocation.toAbsolutePath()).replace(" ", ""));
        XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);
        XSSFCell cell;
        Sheet sheet = wb.getSheetAt(0);
        //Skip 1 Row
        int rowStart = Math.max(1, sheet.getFirstRowNum());
        int rowEnd = sheet.getLastRowNum();

        List<User> users = new ArrayList<>();
        for (int rowNum = rowStart; rowNum <= rowEnd; rowNum++) {
            Row r = sheet.getRow(rowNum);
            if (r == null) {
                continue;   //here Row isEmpty
            }
            List<String> cellValues = new ArrayList<>();
            for (int cn = 0; cn < 9; cn++) {
                Cell c = r.getCell(cn, Row.RETURN_BLANK_AS_NULL);
                if (c == null) {
                    // The spreadsheet isEmpty in this cell
                    cellValues.add(""); //
                } else {
                    cell = (XSSFCell) c;
                    cellValueToStringList(cell, cellValues); //method
                }
            }
            User user = fromCellValueToEmpModel(cellValues);
            users.add(user);
        }
        return users;
    }

    private static void cellValueToStringList(XSSFCell cell, List<String> cellValues) {
        String cellVal;
        switch (cell.getCellType()) {
            case XSSFCell.CELL_TYPE_BOOLEAN:
                cellVal = String.valueOf(cell.getBooleanCellValue());
                cellValues.add(cellVal);
                break;
            case XSSFCell.CELL_TYPE_STRING:
                cellVal = cell.getStringCellValue();
                cellValues.add(cellVal);
                break;
            case XSSFCell.CELL_TYPE_NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    cellVal = cell.getDateCellValue().toString();
                    cellValues.add(cellVal);
                } else {
                    cellVal = String.valueOf(cell.getNumericCellValue());
                    cellValues.add(cellVal);
                }
                break;
            case XSSFCell.CELL_TYPE_FORMULA:
                cellVal = cell.getCellFormula();
                cellValues.add(cellVal);
                break;
            case XSSFCell.CELL_TYPE_BLANK:
                cellVal = " ";
                cellValues.add(cellVal);
                break;
            default:
                  //  System.out.print("");
        }
    }

    private static User fromCellValueToEmpModel(List<String> cellValues) throws Exception {
        User user = new User();
        user.setIndexId(convertStringToInteger(cellValues.get(0)));
        user.setFirstName(cellValues.get(1));
        user.setLastName(cellValues.get(2));
        user.setGender(cellValues.get(3));
        user.setAge(convertStringToInteger(cellValues.get(5)));
        if (StringUtils.isNotEmpty(cellValues.get(4))) {
            user.setCountry(cellValues.get(4));
        }
        if (NumberUtils.isDigits(String.valueOf(convertStringToInteger(cellValues.get(5))))) {
            user.setAge(convertStringToInteger(cellValues.get(5)));
        }
        if (convertStringToDate(cellValues.get(6)) != null) {
            user.setDate(convertStringToDate(cellValues.get(6)));
        }
        user.setId(convertStringToInteger(cellValues.get(7)));
        return user;
    }

    private static Integer convertStringToInteger(String colval) {
        Integer numValue = null;
        if (!StringUtils.isEmpty(colval)) {
            numValue = (int) Math.round(Double.valueOf(colval));
        }
        return numValue;
    }

    private static Date convertStringToDate(String cellValue) throws Exception {
        Date result;
        if (StringUtils.isEmpty(cellValue)) {
            return null;
        } else {
             // java.util.Date parse = new SimpleDateFormat("E MMM dd HH:mm:ss Z yyyy").parse(cellValue);
            //use SimpleDateFormat as per your usecase
            result = new Date(new SimpleDateFormat("dd/MM/yyyy").parse(cellValue).getTime());
            return result;
        }
    }

}
