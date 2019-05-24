package alena;

import org.apache.poi.ss.usermodel.Cell;
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

public class Excel /*extends JFrame*/{

    public static final Pattern gruppa = Pattern.compile("\\s*[а-яА-Я]{4}+[\\-]+[0-9]{2}+[\\-]+[0-9]{2}+.*");

    public static boolean gruppaMatch(String word) {
        Matcher matcher = gruppa.matcher(word);
        return matcher.matches();
    }

    public static final Pattern lssn = Pattern.compile("^\\s*[а-яА-Я ]*");

    public static boolean lssnMatch(String word) {
        Matcher matcher = lssn.matcher(word);
        return matcher.matches();
    }

    public static final Pattern lssnwk = Pattern.compile("^\\s*[0-9, ]+[а-яА-Я ]*\\s*");

    public static boolean lssnwkMatch(String word) {
        Matcher matcher = lssnwk.matcher(word);
        return matcher.matches();
    }

    public static final Pattern lssn_wk = Pattern.compile("^\\s*кр+\\s*[0-9,]+.*");

    public static boolean lssn_wkMatch(String word) {
        Matcher matcher = lssn_wk.matcher(word);
        return matcher.matches();
    }

    public static final Pattern chislo = Pattern.compile("[0-9]{1,3}+\\.+0+");

    public static boolean chisloMatch(String word) {
        Matcher matcher = chislo.matcher(word);
        return matcher.matches();
    }

    static String flnm = "C:\\Users\\Alena\\Desktop\\Учёба\\KBiSP_2_kurs_1_sem";
    static String extntn = ".xlsx";
    static String namegroup;
    static int type;
    /*
    0 тип - лекция,
    1 тип - практика,
    2 тип - лаба.
    */
    static String lesson1, cabin1;
    static String lesson2, cabin2;
    static int groupnumber=0;

    public static void group(String fileName, int i, int k) {

        int j=0;
        int col;
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
                                gpcol=cell.getColumnIndex();
                            }
                            j+=1;
                            groupnumber+=1;
                        }
                        if ((cell.getRowIndex()==(k-1)) & (cell.getColumnIndex()==gpcol)) {
                            lesson1 = cell.getStringCellValue();
                        }
                        if ((cell.getRowIndex()==(k-1)) & (cell.getColumnIndex()==(gpcol+3))) {
                            cabin1 = cell.getStringCellValue();
                        }
                        if ((cell.getRowIndex()==k) & (cell.getColumnIndex()==gpcol)) {
                            lesson2 = cell.getStringCellValue();
                        }
                        if ((cell.getRowIndex()==k) & (cell.getColumnIndex()==(gpcol+3))) {
                            cabin2 = cell.getStringCellValue();
                        }
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        if ((cell.getRowIndex()==(k-1)) & (cell.getColumnIndex()==(gpcol+3))) {
                            cabin1 = Double.toString(cell.getNumericCellValue());
                            if (chisloMatch(cabin1)) {
                                cabin1 = cabin1.substring(0,cabin1.length()-2);
                            }
                        }
                        if ((cell.getRowIndex()==k) & (cell.getColumnIndex()==(gpcol+3))) {
                            cabin2 = Double.toString(cell.getNumericCellValue());
                            if (chisloMatch(cabin2)) {
                                cabin2 = cabin2.substring(0,cabin2.length()-2);
                            }
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

        int y;
        int maxcol, maxrow;
        Calendar c = new GregorianCalendar();

        // создание самого excel файла в памяти
        XSSFWorkbook wb = new XSSFWorkbook();
        group(flnm+extntn, 3, 4);
        int number = groupnumber;
        System.out.println(number);
        XSSFSheet[] sheet = new XSSFSheet[number+1];

        for (y=12; y<13; y+=1) {
            int weeks = 1;
            c.set(c.get(Calendar.YEAR), Calendar.SEPTEMBER, 1);
            while (c.get(Calendar.DAY_OF_WEEK) > 2) {
                c.add(Calendar.DAY_OF_YEAR, -1);
            }
            System.out.print(c.get(Calendar.YEAR) + " ");
            System.out.print(c.get(Calendar.MONTH) + " ");
            System.out.println(c.get(Calendar.DATE) + " ");

            // создание листа с названием
            group(flnm + extntn, y, 6);
            System.out.println(namegroup);
            sheet[y] = wb.createSheet(namegroup);

            XSSFFont font = wb.createFont();
            font.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
            XSSFCellStyle style = wb.createCellStyle();
            style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
            style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
            style.setBorderBottom(XSSFCellStyle.BORDER_MEDIUM);
            style.setBorderTop(XSSFCellStyle.BORDER_MEDIUM);
            style.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM);
            style.setBorderRight(XSSFCellStyle.BORDER_MEDIUM);
            style.setFont(font);

            // счетчик для строк
            int rowNum;
            Row row;

            for (rowNum = 0; rowNum < 44; rowNum += 1) {
                row = sheet[y].createRow(rowNum);
                for (int a = 0; a < 36; a += 1) {
                    row.createCell(a).setCellValue("");
                }
            }

            rowNum = 0;
            row = sheet[y].getRow(rowNum);

            //ДИСЦИПЛИНЫ И АУДИТОРИИ
            row.getCell(0).setCellValue("");
            row.getCell(1).setCellValue("");
            int i = 2;
            maxcol = 1;
            while ((c.get(Calendar.DATE) < 22) || (c.get(Calendar.MONTH) < 11)) {
                sheet[y].setColumnWidth(i, 4000);
                row.getCell(i).setCellValue("дисц");
                sheet[y].setColumnWidth(i + 1, 1500);
                row.getCell(i + 1).setCellValue("ауд");
                row.getCell(i).setCellStyle(style);
                row.getCell(i + 1).setCellStyle(style);
                i += 2;
                maxcol += 2;
                c.add(Calendar.DATE, 7);
                weeks += 1;
            }
            weeks -= 1;
            System.out.println(weeks);
            c.add(Calendar.WEEK_OF_YEAR, -weeks);

            //ДАТЫ
            rowNum += 1;
            for (int k = 1; k <= 7; k += 1) {
                row = sheet[y].getRow(rowNum);
                if (k < 7) {
                    row.getCell(0).setCellValue("");
                    row.getCell(1).setCellValue("");
                } else {
                    row.getCell(0).setCellValue("Неделя");
                    row.getCell(1).setCellValue("");
                }
                if (rowNum != 1) {
                    row.getCell(0).setCellStyle(style);
                    row.getCell(1).setCellStyle(style);
                    CellRangeAddress region = new CellRangeAddress(rowNum, rowNum, 0, 1);
                    sheet[y].addMergedRegion(region);
                }
                i = 2;
                int j = 1;
                while (i < maxcol) {
                    if (k < 7) {
                        if (c.get(Calendar.DATE) < 10) {
                            if (c.get(Calendar.MONTH) < 9) {
                                row.getCell(i).setCellValue("0" + c.get(Calendar.DATE) + "." + "0" + (c.get(Calendar.MONTH) + 1) + "." + c.get(Calendar.YEAR));
                            } else
                                row.getCell(i).setCellValue("0" + c.get(Calendar.DATE) + "." + (c.get(Calendar.MONTH) + 1) + "." + c.get(Calendar.YEAR));
                        } else if (c.get(Calendar.MONTH) < 9) {
                            row.getCell(i).setCellValue(c.get(Calendar.DATE) + "." + "0" + (c.get(Calendar.MONTH) + 1) + "." + c.get(Calendar.YEAR));
                        } else
                            row.getCell(i).setCellValue(c.get(Calendar.DATE) + "." + (c.get(Calendar.MONTH) + 1) + "." + c.get(Calendar.YEAR));
                    } else row.getCell(i).setCellValue(j);
                    row.getCell(i + 1).setCellValue("");
                    row.getCell(i).setCellStyle(style);
                    row.getCell(i + 1).setCellStyle(style);
                    CellRangeAddress region1 = new CellRangeAddress(rowNum, rowNum, i, i + 1);
                    sheet[y].addMergedRegion(region1);
                    c.add(Calendar.DATE, 7);
                    i += 2;
                    j += 1;
                }
                c.add(Calendar.WEEK_OF_YEAR, -weeks);
                c.add(Calendar.DAY_OF_YEAR, 1);

                rowNum += 7;
            }
            rowNum -= 7;
            maxrow = rowNum;
            System.out.println(maxrow);

            XSSFFont font1 = wb.createFont();
            font1.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
            font1.setColor(XSSFFont.COLOR_RED);
            XSSFCellStyle style1 = wb.createCellStyle();
            style1.setAlignment(XSSFCellStyle.ALIGN_CENTER);
            style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
            style1.setBorderBottom(XSSFCellStyle.BORDER_MEDIUM);
            style1.setBorderTop(XSSFCellStyle.BORDER_MEDIUM);
            style1.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM);
            style1.setBorderRight(XSSFCellStyle.BORDER_MEDIUM);
            style1.setFont(font1);


            String[] array;
            array = new String[6];
            array[0] = "Пн";
            array[1] = "Вт";
            array[2] = "Ср";
            array[3] = "Чт";
            array[4] = "Пт";
            array[5] = "Сб";
            sheet[y].setColumnWidth(0, 1000);
            sheet[y].setColumnWidth(1, 1500);

            //ВРЕМЯ
            i = 2;
            while (i < maxrow) {
                rowNum = i;
                c.set(Calendar.HOUR, 9);
                c.set(Calendar.MINUTE, 0);
                row = sheet[y].getRow(rowNum);
                row.getCell(0).setCellValue(array[i / 7]);
                for (rowNum = i; rowNum < (i + 6); rowNum += 1) {
                    String st;
                    if (rowNum != i) row = sheet[y].getRow(rowNum);
                    if (c.get(Calendar.HOUR) < 9) st = Integer.toString(c.get(Calendar.HOUR) + 12);
                    else st = Integer.toString(c.get(Calendar.HOUR));
                    if (c.get(Calendar.MINUTE) == 0) st += ":00";
                    else st += ":" + Integer.toString(c.get(Calendar.MINUTE));
                    row.getCell(1).setCellValue(st);
                    if (rowNum != i) row.getCell(0).setCellValue("");
                    row.getCell(0).setCellStyle(style);
                    row.getCell(1).setCellStyle(style1);
                    if (c.get(Calendar.HOUR) != 10) {
                        c.add(Calendar.HOUR, 1);
                        c.add(Calendar.MINUTE, 40);
                    } else {
                        c.set(Calendar.HOUR, 13);
                        c.set(Calendar.MINUTE, 0);
                    }
                }
                i += 7;
            }

            XSSFFont less = wb.createFont();
            less.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
            XSSFCellStyle stless = wb.createCellStyle();
            stless.setAlignment(XSSFCellStyle.ALIGN_CENTER);
            stless.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
            stless.setBorderBottom(XSSFCellStyle.BORDER_THIN);
            stless.setBorderTop(XSSFCellStyle.BORDER_THIN);
            stless.setBorderLeft(XSSFCellStyle.BORDER_THIN);
            stless.setBorderRight(XSSFCellStyle.BORDER_THIN);
            stless.setFont(less);

            XSSFFont cab = wb.createFont();
            cab.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
            XSSFCellStyle stcab = wb.createCellStyle();
            stcab.setAlignment(XSSFCellStyle.ALIGN_CENTER);
            stcab.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
            stcab.setBorderBottom(XSSFCellStyle.BORDER_THIN);
            stcab.setBorderTop(XSSFCellStyle.BORDER_THIN);
            stcab.setBorderLeft(XSSFCellStyle.BORDER_THIN);
            stcab.setBorderRight(XSSFCellStyle.BORDER_MEDIUM);
            stcab.setFont(cab);

            boolean[] week = new boolean[18];
            String forweek = "";
            String for2week = "";
            String fuckless1 = "";
            String fuckless2 = "";
            String fuck1 = "";
            String fuck2 = "";
            boolean[] week2 = new boolean[18];

            //ЦИКЛ ДЛЯ ЗАПОЛНЕНИЯ ТАБЛИЦЫ
            rowNum = 2;
            for (rowNum = 2; rowNum < maxrow; rowNum += 1) {
                for (i = 1; i <= 17; i += 1) week[i] = false;
                for (i = 1; i <= 17; i += 1) week2[i] = false;
                row = sheet[y].getRow(rowNum);
                group(flnm + extntn, y, (rowNum * 2 - (2 * ((rowNum - 2) / 7))));
                if (lesson1.indexOf("\n")!=(-1)) {
                    fuckless1 = lesson1.substring(lesson1.indexOf("\n")+1, lesson1.length());
                    lesson1 = lesson1.substring(0, lesson1.indexOf("\n"));
                    fuck1 = cabin1.substring(cabin1.indexOf(" ")+1, cabin1.length());
                    cabin1 = cabin1.substring(0, cabin1.indexOf(" "));
                }
                if (lesson2.indexOf("\n")!=(-1)) {
                    fuckless2 = lesson2.substring(lesson2.indexOf("\n")+1, lesson2.length());
                    lesson2 = lesson2.substring(0, lesson2.indexOf("\n"));
                    fuck2 = cabin2.substring(cabin2.indexOf(" ")+1, cabin2.length());
                    cabin2 = cabin2.substring(0, cabin2.indexOf(" "));
                }
                if (lssnMatch(lesson1)) {
                    for (i = 1; i <= 17; i += 2) week[i] = true;
                }
                if (lssnMatch(lesson2)) {
                    for (i = 2; i <= 17; i += 2) week[i] = true;
                }
                if (lssnwkMatch(lesson1)) {
                    forweek = lesson1.substring(0, lesson1.indexOf("н"));
                    forweek = forweek.trim();
                    lesson1 = lesson1.substring((lesson1.indexOf("н") + 2), lesson1.length());
                    while (forweek.indexOf(",") != (-1)) {
                        week[Integer.valueOf(forweek.substring(0, forweek.indexOf(",")))] = true;
                        forweek = forweek.substring(forweek.indexOf(",") + 1, forweek.length());
                        forweek = forweek.trim();
                    }
                    if (forweek.indexOf(",") == (-1)) {
                        week[Integer.valueOf(forweek.substring(0, forweek.length()))] = true;
                    }
                }
                if (lssnwkMatch(lesson2)) {
                    forweek = lesson2.substring(0, lesson2.indexOf("н"));
                    forweek = forweek.trim();
                    lesson2 = lesson2.substring((lesson2.indexOf("н") + 2), lesson2.length());
                    while (forweek.indexOf(",") != (-1)) {
                        week[Integer.valueOf(forweek.substring(0, forweek.indexOf(",")))] = true;
                        forweek = forweek.substring(forweek.indexOf(",") + 1, forweek.length());
                        forweek = forweek.trim();
                    }
                    if (forweek.indexOf(",") == (-1)) {
                        week[Integer.valueOf(forweek.substring(0, forweek.length()))] = true;
                    }
                }
                if (lssnwkMatch(fuckless1)) {
                    forweek = fuckless1.substring(0, fuckless1.indexOf("н"));
                    forweek = forweek.trim();
                    fuckless1 = fuckless1.substring((fuckless1.indexOf("н") + 2), fuckless1.length());
                    while (forweek.indexOf(",") != (-1)) {
                        week2[Integer.valueOf(forweek.substring(0, forweek.indexOf(",")))] = true;
                        forweek = forweek.substring(forweek.indexOf(",") + 1, forweek.length());
                        forweek = forweek.trim();
                    }
                    if (forweek.indexOf(",") == (-1)) {
                        week2[Integer.valueOf(forweek.substring(0, forweek.length()))] = true;
                    }
                }
                if (lssnwkMatch(fuckless2)) {
                    forweek = fuckless2.substring(0, fuckless2.indexOf("н"));
                    forweek = forweek.trim();
                    fuckless2 = fuckless2.substring((fuckless2.indexOf("н") + 2), fuckless2.length());
                    while (forweek.indexOf(",") != (-1)) {
                        week2[Integer.valueOf(forweek.substring(0, forweek.indexOf(",")))] = true;
                        forweek = forweek.substring(forweek.indexOf(",") + 1, forweek.length());
                        forweek = forweek.trim();
                    }
                    if (forweek.indexOf(",") == (-1)) {
                        week2[Integer.valueOf(forweek.substring(0, forweek.length()))] = true;
                    }
                }
                if (lssn_wkMatch(lesson1)) {
                    for (i = 1; i <= 17; i += 1) week[i] = true;
                    forweek = lesson1.substring(lesson1.indexOf("кр") + 2, lesson1.indexOf("н"));
                    forweek = forweek.trim();
                    lesson1 = lesson1.substring((lesson1.indexOf("н") + 2), lesson1.length());
                    while (forweek.indexOf(",") != (-1)) {
                        week[Integer.valueOf(forweek.substring(0, forweek.indexOf(",")))] = false;
                        forweek = forweek.substring(forweek.indexOf(",") + 1, forweek.length());
                        forweek = forweek.trim();
                    }
                    if (forweek.indexOf(",") == (-1)) {
                        week[Integer.valueOf(forweek.substring(0, forweek.length()))] = false;
                    }
                }
                if (lssn_wkMatch(lesson2)) {
                    for (i = 1; i <= 17; i += 1) week[i] = true;
                    forweek = lesson2.substring(lesson2.indexOf("кр") + 2, lesson2.indexOf("н"));
                    forweek = forweek.trim();
                    lesson2 = lesson2.substring((lesson2.indexOf("н") + 2), lesson2.length());
                    while (forweek.indexOf(",") != (-1)) {
                        week[Integer.valueOf(forweek.substring(0, forweek.indexOf(",")))] = false;
                        forweek = forweek.substring(forweek.indexOf(",") + 1, forweek.length());
                        forweek = forweek.trim();
                    }
                    if (forweek.indexOf(",") == (-1)) {
                        week[Integer.valueOf(forweek.substring(0, forweek.length()))] = false;
                    }
                }
                if (lssn_wkMatch(fuckless1)) {
                    for (i = 1; i <= 17; i += 1) week2[i] = true;
                    forweek = fuckless1.substring(fuckless1.indexOf("кр") + 2, fuckless1.indexOf("н"));
                    forweek = forweek.trim();
                    fuckless1 = fuckless1.substring((fuckless1.indexOf("н") + 2), fuckless1.length());
                    while (forweek.indexOf(",") != (-1)) {
                        week2[Integer.valueOf(forweek.substring(0, forweek.indexOf(",")))] = false;
                        forweek = forweek.substring(forweek.indexOf(",") + 1, forweek.length());
                        forweek = forweek.trim();
                    }
                    if (forweek.indexOf(",") == (-1)) {
                        week2[Integer.valueOf(forweek.substring(0, forweek.length()))] = false;
                    }
                }
                if (lssn_wkMatch(fuckless2)) {
                    for (i = 1; i <= 17; i += 1) week2[i] = true;
                    forweek = fuckless2.substring(fuckless2.indexOf("кр") + 2, fuckless2.indexOf("н"));
                    forweek = forweek.trim();
                    fuckless2 = fuckless2.substring((fuckless2.indexOf("н") + 2), fuckless2.length());
                    while (forweek.indexOf(",") != (-1)) {
                        week2[Integer.valueOf(forweek.substring(0, forweek.indexOf(",")))] = false;
                        forweek = forweek.substring(forweek.indexOf(",") + 1, forweek.length());
                        forweek = forweek.trim();
                    }
                    if (forweek.indexOf(",") == (-1)) {
                        week2[Integer.valueOf(forweek.substring(0, forweek.length()))] = false;
                    }
                }
                if ((rowNum - 2) % 7 != 6) {
                    for (i = 2; i < 36; i += 2) {
                        if (week[i / 2] & ((i / 2) % 2 != 0)) {
                            row.getCell(i).setCellValue(lesson1);
                            row.getCell(i).setCellStyle(stless);
                            row.getCell(i + 1).setCellValue(cabin1);
                            row.getCell(i + 1).setCellStyle(stcab);
                        } else {
                            row.getCell(i).setCellStyle(stless);
                            row.getCell(i + 1).setCellStyle(stcab);
                        }
                        if (week[i / 2] & ((i / 2) % 2 == 0)) {
                            row.getCell(i).setCellValue(lesson2);
                            row.getCell(i).setCellStyle(stless);
                            row.getCell(i + 1).setCellValue(cabin2);
                            row.getCell(i + 1).setCellStyle(stcab);
                        } else {
                            row.getCell(i).setCellStyle(stless);
                            row.getCell(i + 1).setCellStyle(stcab);
                        }
                        if (week2[i / 2] & ((i / 2) % 2 != 0)) {
                            row.getCell(i).setCellValue(fuckless1);
                            row.getCell(i).setCellStyle(stless);
                            row.getCell(i + 1).setCellValue(fuck1);
                            row.getCell(i + 1).setCellStyle(stcab);
                        }
                        if (week2[i / 2] & ((i / 2) % 2 == 0)) {
                            row.getCell(i).setCellValue(fuckless2);
                            row.getCell(i).setCellStyle(stless);
                            row.getCell(i + 1).setCellValue(fuck2);
                            row.getCell(i + 1).setCellStyle(stcab);
                        }
                    }
                }
            }

            i = 2;
            while (i < maxrow) {
                CellRangeAddress region = new CellRangeAddress(i, i + 5, 0, 0);
                sheet[y].addMergedRegion(region);
                i += 7;
            }

            CellRangeAddress region = new CellRangeAddress(0, 1, 0, 1);
            sheet[y].addMergedRegion(region);

        }

        // записываем созданный в памяти Excel документ в файл
        try (FileOutputStream out = new FileOutputStream(new File(flnm+"-Java"+extntn))) {
            wb.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Excel файл успешно создан!");
    }

}
