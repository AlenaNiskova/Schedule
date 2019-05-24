package alena;

import alena.DataModel;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.List;

public class ExcelWorker {

    public static void main(String[] args) throws ParseException {

        // создание самого excel файла в памяти
        XSSFWorkbook workbook = new XSSFWorkbook();
        // создание листа с названием "Просто лист"
        XSSFSheet sheet = workbook.createSheet("Просто лист");

        // заполняем список какими-то данными
        List<DataModel> dataList = fillData();

        // счетчик для строк
        int rowNum = 0;

        // создаем подписи к столбцам (это будет первая строчка в листе Excel файла)
        Row row = sheet.createRow(rowNum);
        row.createCell(0).setCellValue("Имя");
        row.createCell(1).setCellValue("Фамилия");
        row.createCell(2).setCellValue("Город");
        row.createCell(3).setCellValue("Зарплата");

        // заполняем лист данными
        for (DataModel dataModel : dataList) {
            createSheetHeader(sheet, ++rowNum, dataModel);
        }

        // записываем созданный в памяти Excel документ в файл
        try (FileOutputStream out = new FileOutputStream(new File("C:\\Users\\Alena\\Desktop\\Учёба\\KBiSP-2-kurs-1-sem-Java.xlsx"))) {
            workbook.write(out);
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Excel файл успешно создан!");
    }

    // заполнение строки (rowNum) определенного листа (sheet)
    // данными  из dataModel созданного в памяти Excel файла
    private static void createSheetHeader(XSSFSheet sheet, int rowNum, DataModel dataModel) {
        Row row = sheet.createRow(rowNum);

        row.createCell(0).setCellValue(dataModel.getName());
        row.createCell(1).setCellValue(dataModel.getSurname());
        row.createCell(2).setCellValue(dataModel.getCity());
        row.createCell(3).setCellValue(dataModel.getSalary());
    }

    // заполняем список рандомными данными
    // в реальных приложениях данные будут из БД или интернета
    private static List<DataModel> fillData() {
        List<DataModel> dataModels = new ArrayList<>();
        dataModels.add(new DataModel("Howard", "Wolowitz", "Massachusetts", 90000.0));
        dataModels.add(new DataModel("Leonard", "Hofstadter", "Massachusetts", 95000.0));
        dataModels.add(new DataModel("Sheldon", "Cooper", "Massachusetts", 120000.0));

        return dataModels;
    }

}
