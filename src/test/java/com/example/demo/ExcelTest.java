package com.example.demo;

import com.example.demo.handler.SheetHandler;
import com.example.demo.service.ExportExcelService;
import com.example.demo.util.ExcelUtil;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Iterator;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

@Slf4j
@SpringBootTest
@RunWith(SpringRunner.class)
public class ExcelTest extends DefaultHandler {

  private static final String PATH = System.getProperty("user.dir").concat(File.separator)
      .concat("excelFile");

  @Autowired
  private ExportExcelService exportExcelService;


  /*
     *@param: []
     *@return void
     *@author lucasliang
     *@date 10/01/2019
     *@Description 1.测试 XSSFWorkbook,单sheet页写入100万数据，时间耗时十几秒
     * 2 .测试在workbook中加入CustomProperties属性，用于vba参数调用
     * SXSSFWorkbook只针对2007以后版本
     * xls文件最大行65536，最大列16384xlsx文件最大行1048576，最大列16384
    */
  @Test
  public void testSxssf() {
    String path = PATH.concat(File.separator).concat("3.xlsx");
    XSSFWorkbook xssfWorkbook;
    Long time = System.currentTimeMillis();
    try (FileInputStream inputStream = new FileInputStream(path)) {
      xssfWorkbook = new XSSFWorkbook(inputStream);
      SXSSFWorkbook sxssfWorkbook = exportExcelService.exportExcelOne(xssfWorkbook);
      //删除flush到本地磁盘的临时文件，避免长时间不删除，磁盘空间占满
   /*   sxssfWorkbook.dispose();*/
      ExcelUtil.output(PATH.concat(File.separator).concat("4.xlsx"), sxssfWorkbook);
    } catch (Exception e) {
      log.error("", e);
    }
    Long time1 = System.currentTimeMillis();
    log.info("消耗时间：{}", (time1 - time) / 1000);
  }

  /*
   *@param: []
   *@return void
   *@author lucasliang
   *@date 10/01/2019
   *@Description 遍历 excel中自定义公式
  */
  @Test
  public void test() throws Exception {
    String path = PATH.concat(File.separator).concat("2.xlsx");
    FileInputStream inputStream = new FileInputStream(path);
    XSSFWorkbook xssfWorkbook = new XSSFWorkbook(inputStream);
    XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(0);
    Iterator<Row> rowIterator = xssfSheet.rowIterator();
    while (rowIterator.hasNext()) {
      Iterator<Cell> cellIterator = rowIterator.next().cellIterator();
      while (cellIterator.hasNext()) {
        Cell cell = cellIterator.next();
        String value = ExcelUtil.getCellValueByType(cell);
        if (value.contains("LOAD")) {
          log.info(String.format(cell.getAddress().toString() + "%s:   ", value));
        }
        cellIterator.remove();
      }
      rowIterator.remove();
    }
  }

  /*
    *@param: []
    *@return void
    *@author lucasliang
    *@date 10/01/2019
    *@Description 测试公式并计算出结果
   */
  @Test
  public void testGs() throws Exception {
    String path = PATH.concat(File.separator).concat("3.xlsx");
    try (FileInputStream inputStream = new FileInputStream(path)) {
      XSSFWorkbook xssfWorkbook = new XSSFWorkbook(inputStream);
      SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(xssfWorkbook, 1000000);
      XSSFFormulaEvaluator.evaluateAllFormulaCells(xssfWorkbook);
      ExcelUtil.output(PATH.concat(File.separator).concat("4.xlsx"), sxssfWorkbook);
    }
  }


  /*
    *@param: []
    *@return void
    *@author lucasliang
    *@date 10/01/2019
    *@Description 测试 CustomProperties 属性
   */
  @Test
  public void testCustomDocProterties() throws Exception {
    String path = PATH.concat(File.separator).concat("3.xlsx");
    FileInputStream inputStream = new FileInputStream(path);
    XSSFWorkbook xssfWorkbook = new XSSFWorkbook(inputStream);
    POIXMLProperties.CustomProperties customProperties = xssfWorkbook.getProperties()
        .getCustomProperties();
    customProperties.addProperty("ip", "localhost");
    log.info(customProperties.getProperty("ip").getLpwstr());
    ExcelUtil.output(PATH.concat(File.separator.concat("cus.xlsx")), xssfWorkbook);
  }


  /*
   *@param: []
   *@return void
   *@author lucasliang
   *@date 10/01/2019
   *@Description 基于poi事件驱动解析excel，此方法抛弃
  */
  @Test
  public void testOPCPackage() {
    String path = PATH.concat(File.separator).concat("3.xlsx");
    XSSFWorkbook xssfWorkbook;
    OPCPackage opcPackage;
    try (FileInputStream inputStream = new FileInputStream(path)) {
      xssfWorkbook = new XSSFWorkbook(inputStream);
      opcPackage = xssfWorkbook.getPackage();
      XSSFReader xssfReader = new XSSFReader(opcPackage);
      xssfReader.getSheetsData();
      xssfReader.getStylesTable();
    } catch (Exception e) {
      log.error("", 3);
    }
  }

  /*
     *@param: []
     *@return void
     *@author lucasliang
     *@date 10/01/2019
     *@Description 测试 OPCPackage ，xml解析excel，但此方式基于事件模型，针对excel读取效果好，故抛弃，不适用此方式
    */
  @Test
  public void testOpcPackage() {
    String path = PATH.concat(File.separator).concat("2.xlsx");
    try (OPCPackage pkg = OPCPackage.open(path, PackageAccess.READ_WRITE)) {
      XSSFReader reader = new XSSFReader(pkg);
      SharedStringsTable sst = reader.getSharedStringsTable();
      StylesTable st = reader.getStylesTable();
      XMLReader parser = XMLReaderFactory.createXMLReader();
      // 处理公共属性：Sheet名，Sheet合并单元格
      parser.setContentHandler(new SheetHandler());
      /**
       * 返回一个迭代器，此迭代器会依次得到所有不同的sheet。
       * 每个sheet的InputStream只有从Iterator获取时才会打开。
       * 解析完每个sheet时关闭InputStream。
       * */
      XSSFReader.SheetIterator sheets = (XSSFReader.SheetIterator) reader.getSheetsData();
      while (sheets.hasNext()) {
        InputStream sheetstream = sheets.next();
        InputSource sheetSource = new InputSource(sheetstream);
        try {
          // 解析sheet: com.sun.org.apache.xerces.internal.jaxp.SAXParserImpl:522
          parser.parse(sheetSource);
        } finally {
          sheetstream.close();
        }
      }
    } catch (Exception e) {
      log.error("", e);
    }
  }

}

