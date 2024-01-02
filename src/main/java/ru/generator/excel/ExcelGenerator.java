package ru.generator.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

public class ExcelGenerator {
    static Double costOfOneDfa;
    static Double numberOfSaleDFA;
    static Double summCostSaleDFA;
    static long duration;
    static Double percentForOneDfa;
    static Double rate;

    static {
        Properties prop = new Properties();
        String fileName = "src/test/resources/app.config";
        try (FileInputStream fis = new FileInputStream(fileName)) {
            prop.load(fis);
        } catch (IOException ex) {
            ex.printStackTrace();
        }

        costOfOneDfa = Double.valueOf(prop.getProperty("app.costOfOneDFA"));
        numberOfSaleDFA = Double.valueOf(prop.getProperty("app.numberOfSaleDFA"));
        duration = Long.parseLong(prop.getProperty("app.duration"));
        summCostSaleDFA = costOfOneDfa*numberOfSaleDFA;
        rate = Double.parseDouble(prop.getProperty("app.rate"));
        //проценты за 1 ЦФА
        percentForOneDfa = (rate*costOfOneDfa)/100;

    }

    public static void main(String[] args) {
        createAccrualNPDSheet();
        createPaymentNPDSheet();
        createAccrualODSheet();
        createPaymentODSheet();
        System.out.println(createDate(duration));
    }

    public static void createAccrualNPDSheet () {

        Workbook wb = new XSSFWorkbook();
        Sheet sheet0 = wb.createSheet("График начисления НПД");
        Row row0 = sheet0.createRow(0);
        Row row1 = sheet0.createRow(1);

        Cell cell01 = row0.createCell(0);
        cell01.setCellValue("Дата начисления");

        Cell cell02 = row0.createCell(1);
        cell02.setCellValue("Сумма к начислению");

        Cell cell11 = row1.createCell(0);
        cell11.setCellValue(createDate(duration));

        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(percentForOneDfa*numberOfSaleDFA);

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

    public static void createPaymentNPDSheet(){
        Workbook wb = new XSSFWorkbook();
        Sheet sheet0 = wb.createSheet("График выплаты НПД");
        Row row0 = sheet0.createRow(0);
        Row row1 = sheet0.createRow(1);

        Cell cell01 = row0.createCell(0);
        cell01.setCellValue("Дата закрытия реестра");

        Cell cell02 = row0.createCell(1);
        cell02.setCellValue("Дата начала процентного периода");

        Cell cell03 = row0.createCell(2);
        cell03.setCellValue("Дата завершения процентного периода");

        Cell cell04 = row0.createCell(3);
        cell04.setCellValue("Сумма к выплате");

        Cell cell11 = row1.createCell(0);
        cell11.setCellValue(createDate(duration+1));

        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(createDate(duration+2));

        Cell cell13 = row1.createCell(2);
        cell13.setCellValue(createDate(duration+3));

        Cell cell14 = row1.createCell(3);
        cell14.setCellValue((costOfOneDfa*rate)/100);

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

    public static void createAccrualODSheet () {

        Workbook wb = new XSSFWorkbook();
        Sheet sheet0 = wb.createSheet("График начисления ОД");
        Row row0 = sheet0.createRow(0);
        Row row1 = sheet0.createRow(1);

        Cell cell01 = row0.createCell(0);
        cell01.setCellValue("Дата начисления");

        Cell cell02 = row0.createCell(1);
        cell02.setCellValue("Сумма к начислению");

        Cell cell11 = row1.createCell(0);
        cell11.setCellValue(createDate(duration+4));

        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(costOfOneDfa*numberOfSaleDFA);

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

    public static void createPaymentODSheet(){
        Workbook wb = new XSSFWorkbook();
        Sheet sheet0 = wb.createSheet("График выплаты ОД");
        Row row0 = sheet0.createRow(0);
        Row row1 = sheet0.createRow(1);

        Cell cell01 = row0.createCell(0);
        cell01.setCellValue("Дата закрытия реестра");

        Cell cell02 = row0.createCell(1);
        cell02.setCellValue("Дата выплаты");

        Cell cell03 = row0.createCell(2);
        cell03.setCellValue("Сумма к выплате");

        Cell cell11 = row1.createCell(0);
        cell11.setCellValue(createDate(duration+5));

        Cell cell12 = row1.createCell(1);
        cell12.setCellValue(createDate(duration+6));

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

    public static String createDate(long duration){
        SimpleDateFormat formatter= new SimpleDateFormat("dd.MM.yyyy HH:mm");
        Date date = new Date(System.currentTimeMillis() + TimeUnit.MINUTES.toMillis(duration));
        return formatter.format(date);
    }


}
