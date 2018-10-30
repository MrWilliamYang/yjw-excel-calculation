package yjw.excel.calculation.excel.write;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * 将数据写入到Excel表格中
 * 
 * @author YangJianWei
 * @version $Id: WriteExcelForXSSF.java, v 0.1 2018年7月6日 下午4:53:34 YangJianWei Exp $
 */
public class WriteExcelForXSSF {
    @SuppressWarnings("resource")
    public void write() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("0");
        Row row = sheet.createRow(0);
        CellStyle cellStyle = workbook.createCellStyle();
//        //设置这些样式
//        cellStyle.setFillForegroundColor(HSSFColor.SKY_BLUE.index);
//        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//        cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
//        cellStyle.setBorderLeft(CellStyle.BORDER_THIN);
//        cellStyle.setBorderRight(CellStyle.BORDER_THIN);
//        cellStyle.setBorderTop(CellStyle.BORDER_THIN);
//        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
 
        //新增第  0 列
        row.createCell(0).setCellStyle(cellStyle);
        row.createCell(0).setCellValue("姓名");
 
        //新增第  1 列
        row.createCell(1).setCellStyle(cellStyle);
        row.createCell(1).setCellValue("年龄");
 
        //设置sheet名字
        workbook.setSheetName(0, "信息");
        try {
            File file = new File("D:\\桌面\\公司测试\\测试写入.xlsx");
            FileOutputStream fileoutputStream = new FileOutputStream(file);
            workbook.write(fileoutputStream);
            fileoutputStream.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
 
    public static void main(String[] args) {
        new WriteExcelForXSSF().write();
    }
}
