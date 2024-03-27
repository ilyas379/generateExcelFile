package ru.generator.excel;

import lombok.SneakyThrows;
import lombok.extern.log4j.Log4j2;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileInputStream;

@Log4j2
public class CellFormat {
    @SneakyThrows
    public static void main(String[] args) {
        //Определить какой формат ячейки у выбранного файла для последующего копирования
        String filePath = "C:\\Users\\admin\\Desktop\\123\\test.xlsx";
        log.info("Получение формата ячейки у файла " + filePath);
        FileInputStream fileIn = new FileInputStream(filePath);
        Workbook workbook = WorkbookFactory.create(fileIn);
        CellStyle cellStyle = workbook.getSheetAt(0).getRow(1).getCell(0).getCellStyle();
        String styleString = cellStyle.getDataFormatString();
        log.info("Формат ячейки следующий:");
        System.out.println(styleString);
    }
}
