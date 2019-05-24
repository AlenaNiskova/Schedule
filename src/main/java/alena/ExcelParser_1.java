package alena;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExcelParser_1 {

    public static final Pattern gruppa = Pattern.compile("[а-яА-Я]{4,5}+[\\-]+[0-9]{2}+[\\-]+[0-9]{2}+\\s*");

    public static boolean doMatch(String word) {
        Matcher matcher = gruppa.matcher(word);
        return matcher.matches();
    }

    public static String group(String fileName) {

        String name = null;
        int rows=1;
        InputStream inputStream = null;
        XSSFWorkbook workBook = null;
        try {
            inputStream = new FileInputStream(fileName);
            workBook = new XSSFWorkbook(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        Sheet sheet = workBook.getSheetAt(0);
        Iterator<Row> it = sheet.iterator();
        while (it.hasNext() & rows<=2) {
            Row row = it.next();
            Iterator<Cell> cells = row.iterator();
            int cols=1;
            while (cells.hasNext() & (cols<=9)) {
                Cell cell = cells.next();
                int cellType = cell.getCellType();
                switch (cellType) {
                    case Cell.CELL_TYPE_STRING:
                        if (doMatch(cell.getStringCellValue())) {
                            name = cell.getStringCellValue();
                        }
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        //result += "[" + cell.getNumericCellValue() + "]";
                        break;
                    case Cell.CELL_TYPE_FORMULA:
                        //result += "[" + cell.getNumericCellValue() + "]";
                        break;
                    default:
                        //result += "|";
                        break;
                }
                cols=cols+1;
            }
            rows=rows+1;
        }

        return name;
    }

    public static String parse(String fileName) {

        String group = null;
        int rows=1;
        //инициализируем потоки
        String result = "";
        InputStream inputStream = null;
        XSSFWorkbook workBook = null;
        try {
            inputStream = new FileInputStream(fileName);
            workBook = new XSSFWorkbook(inputStream);
        } catch (IOException e) {
            e.printStackTrace();
        }
        //разбираем первый лист входного файла на объектную модель
        Sheet sheet = workBook.getSheetAt(0);
        Iterator<Row> it = sheet.iterator();
        //проходим по всему листу
        while (it.hasNext()) {
            Row row = it.next();
            Iterator<Cell> cells = row.iterator();
            int cols=1;
            while (cells.hasNext() & (cols<=9)) {
                if ((cols%17==6) & (rows==2)) {
                    Cell cell = cells.next();
                    group = cell.getStringCellValue();
                }
                    Cell cell = cells.next();
                    int cellType = cell.getCellType();
                    //перебираем возможные типы ячеек
                    switch (cellType) {
                        case Cell.CELL_TYPE_STRING:
                            result += cell.getStringCellValue() + "=";
                            break;
                        case Cell.CELL_TYPE_NUMERIC:
                            result += "[" + cell.getNumericCellValue() + "]";
                            break;
                        case Cell.CELL_TYPE_FORMULA:
                            result += "[" + cell.getNumericCellValue() + "]";
                            break;
                        default:
                            result += "|";
                            break;
                    }
                cols=cols+1;
            }
            rows=rows+1;
            result += "\n";
        }
        result += group;

        return result;
    }



    public static final Pattern urok = Pattern.compile("");

    public static boolean doMatch2(String word) {
        Matcher matcher = urok.matcher(word);
        return matcher.matches();
    }

    public static String lesson(String fileName) {

        String name = null;



        return name;
    }

}
