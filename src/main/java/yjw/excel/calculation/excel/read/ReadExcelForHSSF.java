package yjw.excel.calculation.excel.read;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 读取Excel表格，注意，Excel文件结尾后缀为xls
 * 
 * @author YangJianWei
 * @version $Id: ReadExcel.java, v 0.1 2018年7月6日 上午11:36:00 YangJianWei Exp $
 */
public class ReadExcelForHSSF {
    
    /**
     * 读取单元格的值
     * 
     * @param cell
     * @return
     */
    private static String getCellValue(Cell cell) {
        Object result = "";
        if (cell != null) {
          switch (cell.getCellType()) {
          case Cell.CELL_TYPE_STRING:
            result = cell.getStringCellValue();
            break;
          case Cell.CELL_TYPE_NUMERIC:
            result = cell.getNumericCellValue();
            break;
          case Cell.CELL_TYPE_BOOLEAN:
            result = cell.getBooleanCellValue();
            break;
          case Cell.CELL_TYPE_FORMULA:
            result = cell.getCellFormula();
            break;
          case Cell.CELL_TYPE_ERROR:
            result = cell.getErrorCellValue();
            break;
          case Cell.CELL_TYPE_BLANK:
            break;
          default:
            break;
          }
        }
        return result.toString();
    }

    public String readExcel()throws IOException{
        //读取Excel文件，后缀以xls结尾，获取路径
        Workbook wb = new HSSFWorkbook(new FileInputStream("D:\\桌面\\公司测试\\测试读取.xls"));
        //获取sheet数目
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);
            Row row = null;
            int lastRowNum = sheet.getLastRowNum();
            //遍历每行
            for (int j = 0; j < lastRowNum; j++) {
                row = sheet.getRow(j);
                if(row != null) {
                    //遍历每列的值
                    for (int k = 0; k < row.getLastCellNum(); k++) {
                        Cell cell = row.getCell(k);
                        String value = getCellValue(cell) ;
                        if(!value.equals("")){
                          System.out.print(value + " | ");
                        }
                    }
                }
            }
            wb.close();
        }
        return null;
    }
    public static void main(String[] args) throws IOException {
        ReadExcelForHSSF read = new ReadExcelForHSSF();
        String s = read.readExcel();
        System.out.println(s);
    }
}
