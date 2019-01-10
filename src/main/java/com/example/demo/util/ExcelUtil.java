package com.example.demo.util;

import java.io.FileOutputStream;
import java.util.Iterator;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/*
* lucasliang
* excel utils
* */
public class ExcelUtil {

  private static final Logger logger = LoggerFactory.getLogger(ExcelUtil.class);

  private ExcelUtil() {
  }

  /*
     *@param: [outPath, wb]
     *@return void
     *@author lucasliang
     *@date 11/12/2018
     *@Description output file
    */
  public static void output(String outPath, Workbook wb) {
    try (FileOutputStream fos = new FileOutputStream(outPath)) {
      wb.write(fos);
    } catch (Exception e) {
      logger.error("", e);
    }
  }

  public static void copyRow(XSSFRow xssfRow, SXSSFRow sxssfRow) {
    Iterator<Cell> cellIterator = xssfRow.cellIterator();
    SXSSFCell sxssfCell;
    while (cellIterator.hasNext()) {
      Cell cell = cellIterator.next();
      sxssfCell = sxssfRow.createCell(cell.getColumnIndex(), cell.getCellType());
      setCellValue(cell, sxssfCell);
    }
  }

  public static void setCellValue(Cell cell, SXSSFCell sxssfCell) {
    sxssfCell.setCellStyle(cell.getCellStyle());
    CellType cellType = cell.getCellType();
    if (cellType != null) {
      switch (cellType) {
        case NUMERIC:
          if (HSSFDateUtil.isCellDateFormatted(cell)) {
            sxssfCell.setCellValue(DateUtils.formatDate(cell.getDateCellValue(), "yyyy-MM-dd"));
          } else {
            Long longVal = Math.round(cell.getNumericCellValue());
            Double doubleVal = cell.getNumericCellValue();
            if (Double.parseDouble(longVal + ".0") == doubleVal) {   //判断是否含有小数位.0
              sxssfCell.setCellValue(longVal);
            } else {
              sxssfCell.setCellValue(doubleVal);
            }
          }
          break;
        case FORMULA:
          sxssfCell.setCellFormula(cell.getCellFormula());
          break;
      }
    }
  }

  /*
    *@param: [cellType, cell]
    *@return java.lang.String
    *@author lucasliang
    *@date 14/12/2018
    *@Description get excel cellvalue String by CellType
   */
  public static String getCellValueByType(Cell cell) {
    String result = "";
    CellType cellType = cell.getCellType();
    switch (cellType) {
      case NUMERIC:
        if (HSSFDateUtil.isCellDateFormatted(cell)) {
          result = DateUtils.formatDate(cell.getDateCellValue(), "yyyy-MM-dd");
        } else {
          Long longVal = Math.round(cell.getNumericCellValue());
          Double doubleVal = cell.getNumericCellValue();
          if (Double.parseDouble(longVal + ".0") == doubleVal) {   //判断是否含有小数位.0
            result = String.valueOf(longVal);
          } else {
            result = String.valueOf(doubleVal);
          }
        }
        break;
      case STRING:
        result = StringUtils.trim(cell.getStringCellValue());
        break;
      case BOOLEAN:
        result = StringUtils.trim(String.valueOf(cell.getBooleanCellValue()));
        break;
      case FORMULA:
        result = StringUtils.trim(String.valueOf(cell.getCellFormula()));
        break;
      case _NONE:
        result = null;
        break;
      case ERROR:
        result = "";
        break;
      case BLANK:
        result = "";
        break;
      default:
        break;
    }
    return result;
  }

  /*
    *@param: [cell, workbook]
    *@return org.apache.poi.ss.usermodel.CellValue
    *@author lucasliang
    *@date 10/01/2019
    *@Description 读取excel公式并计算出结果
   */
  public CellValue formulaEvaluation(Cell cell, Workbook workbook) {
    FormulaEvaluator formulaEval = workbook.getCreationHelper().createFormulaEvaluator();
    return formulaEval.evaluate(cell);
  }

 /* public static void copyRow(XSSFWorkbook  wb,XSSFRow fromRow,XSSFRow toRow,boolean copyValueFlag){
    CellCopyPolicy cellCopyPolicy=new CellCopyPolicy();
    XSSFRow xssfRow=null;
    fromRow.copyRowFrom(fromRow, cellCopyPolicy);
    for (Iterator cellIt = fromRow.cellIterator(); cellIt.hasNext();) {
      XSSFCell tmpCell = (XSSFCell) cellIt.next();
      XSSFCell newCell = toRow.createCell(tmpCell.getCellNum());
      copyCell(wb,tmpCell, newCell, copyValueFlag);
    }
  }*/

}