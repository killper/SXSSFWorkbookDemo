package com.example.demo.service;

import com.example.demo.util.ExcelUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

/**
 * @author lucasliang
 * @date 09/01/2019 5:42 下午
 * @Description
 */
@Slf4j
@Service
public class ExportExcelService {

  public SXSSFWorkbook exportExcelOne(XSSFWorkbook workbook) {
    //excel属性保存
    POIXMLProperties.CustomProperties customProperties = workbook.getProperties()
        .getCustomProperties();
    customProperties.addProperty("ip", "localhost");
    log.info(customProperties.getProperty("ip").getLpwstr());

    int sheetNumber = workbook.getNumberOfSheets();
    SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(workbook, 10000);
    //excel属性保存
    POIXMLProperties.CustomProperties customProperties1 = sxssfWorkbook.getXSSFWorkbook()
        .getProperties()
        .getCustomProperties();
    customProperties1.addProperty("ip1", "127.0.0.1");
    log.info(customProperties.getProperty("ip1").getLpwstr());
    SXSSFSheet sheet;
    XSSFSheet xssfSheet;
    Row row;
    //循环sheet页，为每个sheet页填充数据
    for (int i = 0; i < sheetNumber; i++) {
      sheet = sxssfWorkbook.getSheetAt(i);
      xssfSheet = sxssfWorkbook.getXSSFWorkbook().getSheetAt(i);
      XSSFRow row1 = xssfSheet.getRow(1);
      //循环创建行
      for (int j = 2; j < 1000000; j++) {
        row = sheet.createRow(j);
        ExcelUtil.copyRow(row1, (SXSSFRow) row);
      }
    }
    return sxssfWorkbook;
  }

  private static int getFirstRowCellNumber(XSSFWorkbook workbook, int i) {
    Sheet sheet = workbook.getSheetAt(i);
    Row firstRow = sheet.getRow(0);
    return firstRow.getPhysicalNumberOfCells();
  }
}
