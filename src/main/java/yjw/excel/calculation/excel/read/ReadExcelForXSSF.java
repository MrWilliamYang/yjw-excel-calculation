package yjw.excel.calculation.excel.read;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * 读取Excel表格，注意，Excel文件结尾后缀为xlsx
 * 
 * @author YangJianWei
 * @version $Id: ReadExcel1.java, v 0.1 2018年7月6日 上午11:37:37 YangJianWei Exp $
 */
public class ReadExcelForXSSF {

    public void read() {
        //获取读取文件的路径
        File file = new File("D:\\桌面\\公司测试\\测试数据--付超修改.xlsx");
        InputStream inputStream = null;
        Workbook workbook = null;
        
        
        try {
            inputStream = new FileInputStream(file);
            workbook = WorkbookFactory.create(inputStream);
            inputStream.close();
            //得到工作表对象
            Sheet sheet = workbook.getSheetAt(1);
            //获取总行数
            int rowLength = sheet.getLastRowNum() + 1;
            //得到工作表的列
            Row row = sheet.getRow(0);
            //获取总列数
            int colLength = row.getLastCellNum();
            //获取指定单元格
            Cell cell = row.getCell(0);
//            //获取单元格格式/样式
//            CellStyle cellStyle = cell.getCellStyle();
//            //统计总行数和列数
//            System.out.println("行数" + rowLength + "列数" + colLength);
//            //计算
//            cell.setCellFormula("sum(A3:B3)");
            for (int i = 0; i < rowLength; i++) {
                row = sheet.getRow(i);
                for (int j = 0; j < colLength; j++) {
                    cell = row.getCell(j);
                    System.out.print(cell + "\t");
                    //Excel数据Cell有不同的类型，当我们试图从一个数字类型的Cell读取出一个字符串时就有可能报异常：
                    //Cannot get a STRING value from a NUMERIC cell
                    //将所有的需要读的Cell表格设置为String格式
                    if(cell != null) {
                        cell.setCellType(Cell.CELL_TYPE_BLANK);
                    }
                    //对表格进行修改
                    if(i > 0 && j == 1) {
                        cell.setCellValue("1000");
                        System.out.print(cell.getStringCellValue() + "\t");
                    }
                }
                System.out.println();
            }
//            //将修改好的数据保存
//            OutputStream out = new FileOutputStream(file);
//            workbook.write(out);
        } catch (Exception e) {
            // TODO: handle exception
            e.printStackTrace();
        }
    }
    public static void main(String[] args) {
        new ReadExcelForXSSF().read();
    }
}
