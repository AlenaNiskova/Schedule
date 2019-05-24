package alena;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.ParseException;
import java.util.Calendar;
import java.util.GregorianCalendar;
import java.util.Iterator;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExcelWorker_1 {

    static String namegroup;
    static String lesson1, cabin1;
    static String lesson2, cabin2;
    static int groupnumber=0;

    public static final Pattern gruppa = Pattern.compile("[а-яА-Я]{4}+[\\-]+[0-9]{2}+[\\-]+[0-9]{2}+\\s*");

    public static boolean gruppaMatch(String word) {
        Matcher matcher = gruppa.matcher(word);
        return matcher.matches();
    }

    public static void group(String fileName, int i, int k) {

        int j=0;
        int col=0;
        int line=0;
        int gpcol=0;
        lesson1="";
        lesson2="";
        cabin1="";
        cabin2="";

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
        while (it.hasNext()) {
            Row row = it.next();
            Iterator<Cell> cells = row.iterator();
            col=0;
            while (cells.hasNext()) {
                Cell cell = cells.next();
                int cellType = cell.getCellType();
                switch (cellType) {
                    case Cell.CELL_TYPE_STRING:
                        if (gruppaMatch(cell.getStringCellValue())) {
                            if (i==j) {
                                namegroup = cell.getStringCellValue();
                                gpcol=col;
                            }
                            j+=1;
                            groupnumber+=1;
                        }
                        if ((line==(k-1)) & (col==gpcol)) {
                            lesson1 = cell.getStringCellValue();
                        }
                        if ((line==(k-1)) & (col==(gpcol+3))) {
                            cabin1 = cell.getStringCellValue();
                        }
                        if ((line==k) & (col==gpcol)) {
                            lesson2 = cell.getStringCellValue();
                        }
                        if ((line==k) & (col==(gpcol+3))) {
                            cabin2 = cell.getStringCellValue();
                        }
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        if ((line==(k-1)) & (col==(gpcol+3))) {
                            cabin1 = Double.toString(cell.getNumericCellValue());
                        }
                        if ((line==k) & (col==(gpcol+3))) {
                            cabin2 = Double.toString(cell.getNumericCellValue());
                        }
                        break;
                    case Cell.CELL_TYPE_FORMULA:
                        break;
                    default:
                        break;
                }
                col+=1;
            }
            line+=1;
        }
    }

    public static void main(String[] args) throws ParseException {

        String flnm = "C:\\Users\\Alena\\Desktop\\Учёба\\KBiSP-2-kurs-1-sem";
        String extntn = ".xlsx";
        int weeks = 1;
        int maxcol, maxrow;
        Calendar c = new GregorianCalendar();
        c.set(c.get(Calendar.YEAR), Calendar.SEPTEMBER, 1);
        while (c.get(Calendar.DAY_OF_WEEK)>2) {
            c.add(Calendar.DAY_OF_YEAR, -1);
        }
        System.out.print(c.get(Calendar.YEAR)+" ");
        System.out.print(c.get(Calendar.MONTH)+" ");
        System.out.println(c.get(Calendar.DATE)+" ");

        // создание самого excel файла в памяти
        XSSFWorkbook workbook = new XSSFWorkbook();

        group(flnm+extntn, 1, 6);
        XSSFSheet[] sheet = new XSSFSheet[groupnumber];

        for (int i=1; i<=2; i+=1) {
            group(flnm+extntn, i, 6);
            sheet[i] = workbook.createSheet(namegroup);
        }

        // создание листа с названием
        //XSSFSheet sheet = workbook.createSheet(namegroup);

        // записываем созданный в памяти Excel документ в файл
        try (FileOutputStream out = new FileOutputStream(new File(flnm+"-Java"+extntn))) {
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Excel файл успешно создан!");
    }
}
