package ru.generator.excel;

import lombok.SneakyThrows;
import lombok.extern.log4j.Log4j2;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Calendar;
import java.util.Date;
import java.util.Properties;

@Log4j2
public class ExcelGenerator {
    static Double costOfOneDfa;
    static Double numberOfSaleDFA;
    static Double summCostSaleDFA;
    static int duration;
    static Double percentForOneDfa;
    static Double rate;
    static String fileName = "config.properties";

    static {
        Properties prop = new Properties();
        try {
            prop.load(Files.newInputStream(Paths.get(fileName)));
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        costOfOneDfa = Double.valueOf(prop.getProperty("app.costOfOneDFA"));
        numberOfSaleDFA = Double.valueOf(prop.getProperty("app.numberOfSaleDFA"));
        duration = Integer.parseInt(prop.getProperty("app.duration"));
        summCostSaleDFA = costOfOneDfa * numberOfSaleDFA;
        rate = Double.parseDouble(prop.getProperty("app.rate"));
        //проценты за 1 ЦФА
        percentForOneDfa = (rate * costOfOneDfa) / 100;
    }

    @SneakyThrows
    public static void main(String[] args) {
        log.info("Start generate Excel files");
        createAccrualNPDSheet();
        createPaymentNPDSheet();
        createAccrualODSheet();
        createPaymentODSheet();
        log.info("Generate Excel files success");
    }

    @SneakyThrows
    public static void createAccrualNPDSheet() {

        Workbook wb = new XSSFWorkbook();
        Sheet sheet0 = wb.createSheet("График начисления НПД");
        Row row0 = sheet0.createRow(0);
        Row row1 = sheet0.createRow(1);

        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(wb.getCreationHelper().createDataFormat().getFormat("dd.MM.yyyy HH:mm:ss"));

        Cell cell01 = row0.createCell(0);
        cell01.setCellValue("Дата начисления");

        Cell cell02 = row0.createCell(1);
        cell02.setCellValue("Сумма к начислению");

        Cell cell11 = row1.createCell(0);
        cell11.setCellValue(createDate(duration));
        cell11.setCellStyle(cellStyle);

        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(percentForOneDfa * numberOfSaleDFA);

        sheet0.autoSizeColumn(0);
        sheet0.autoSizeColumn(1);

        FileOutputStream fileOutputStream;
        try {
            fileOutputStream = new FileOutputStream("1.График начисления НПД.xlsx");
            wb.write(fileOutputStream);
            fileOutputStream.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static void createPaymentNPDSheet() {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet0 = wb.createSheet("График выплаты НПД");
        Row row0 = sheet0.createRow(0);
        Row row1 = sheet0.createRow(1);

        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(wb.getCreationHelper().createDataFormat().getFormat("dd.MM.yyyy HH:mm:ss"));

        Cell cell01 = row0.createCell(0);
        cell01.setCellValue("Дата закрытия реестра");

        Cell cell02 = row0.createCell(1);
        cell02.setCellValue("Дата начала процентного периода");

        Cell cell03 = row0.createCell(2);
        cell03.setCellValue("Дата завершения процентного периода");

        Cell cell04 = row0.createCell(3);
        cell04.setCellValue("Сумма к выплате");

        Cell cell11 = row1.createCell(0);
        cell11.setCellValue(createDate(duration + 1));
        cell11.setCellStyle(cellStyle);

        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(createDate(duration + 2));
        cell12.setCellStyle(cellStyle);

        Cell cell13 = row1.createCell(2);
        cell13.setCellValue(createDate(duration + 3));
        cell13.setCellStyle(cellStyle);

        Cell cell14 = row1.createCell(3);
        cell14.setCellValue((costOfOneDfa * rate) / 100);

        sheet0.autoSizeColumn(0);
        sheet0.autoSizeColumn(1);
        sheet0.autoSizeColumn(2);
        sheet0.autoSizeColumn(3);

        FileOutputStream fileOutputStream;
        try {
            fileOutputStream = new FileOutputStream("2.График выплаты НПД.xlsx");
            wb.write(fileOutputStream);
            fileOutputStream.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static void createAccrualODSheet() {

        Workbook wb = new XSSFWorkbook();
        Sheet sheet0 = wb.createSheet("График начисления ОД");
        Row row0 = sheet0.createRow(0);
        Row row1 = sheet0.createRow(1);

        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setDataFormat(wb.getCreationHelper().createDataFormat().getFormat("dd.MM.yyyy HH:mm:ss"));

        Cell cell01 = row0.createCell(0);
        cell01.setCellValue("Дата начисления");

        Cell cell02 = row0.createCell(1);
        cell02.setCellValue("Сумма к начислению");

        Cell cell11 = row1.createCell(0);
        cell11.setCellValue(createDate(duration + 4));
        cell11.setCellStyle(cellStyle);

        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(costOfOneDfa * numberOfSaleDFA);

        sheet0.autoSizeColumn(0);
        sheet0.autoSizeColumn(1);

        FileOutputStream fileOutputStream;
        try {
            fileOutputStream = new FileOutputStream("3.График начисления ОД.xlsx");
            wb.write(fileOutputStream);
            fileOutputStream.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static void createPaymentODSheet() {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet0 = wb.createSheet("График выплаты ОД");
        Row row0 = sheet0.createRow(0);
        Row row1 = sheet0.createRow(1);

        CellStyle cellStyle = wb.createCellStyle();
        CreationHelper creationHelper = wb.getCreationHelper();
        cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("dd/mm/yy h:mm;@"));

        Cell cell01 = row0.createCell(0);
        cell01.setCellValue("Дата закрытия реестра");

        Cell cell02 = row0.createCell(1);
        cell02.setCellValue("Дата выплаты");

        Cell cell03 = row0.createCell(2);
        cell03.setCellValue("Сумма к выплате");

        Cell cell11 = row1.createCell(0);
        cell11.setCellValue(createDate(duration + 5));
        cell11.setCellStyle(cellStyle);

        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(createDate(duration + 6));
        cell12.setCellStyle(cellStyle);

        Cell cell13 = row1.createCell(2);
        cell13.setCellValue(costOfOneDfa);

        sheet0.autoSizeColumn(0);
        sheet0.autoSizeColumn(1);
        sheet0.autoSizeColumn(2);

        FileOutputStream fileOutputStream;
        try {
            fileOutputStream = new FileOutputStream("4.График выплаты ОД.xlsx");
            wb.write(fileOutputStream);
            fileOutputStream.close();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static Date createDate(int duration) {
        Date currentDate = new Date();
        // Создаем объект типа Calendar и устанавливаем его на текущую дату
        Calendar calendar = Calendar.getInstance();
        calendar.setTime(currentDate);
        // Добавляем 30 минут
        calendar.add(Calendar.MINUTE, duration);
        // Получаем новую дату
        return calendar.getTime();
    }
}
