package yjw.excel.calculation.excel.read;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.LinkedList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Apache POI 解析Excel表格（2003和2007兼容）
 * @author YangJianWei
 * @version $Id: ReadExcelForXSSF1.java, v 0.1 2018年7月10日 下午2:59:46 YangJianWei Exp $
 */
public class ReadExcelForXSSF1 {

    public static void main(String[] args) throws Exception {
        
        //读取2007Excel文件
        String path2007 = System.getProperty("user.dir") + System.getProperty("file.separator") + "测试数据--付超修改.xlsx";// 获取项目文件路径 +2007版文件名
        System.out.println("路径：" + path2007);
        File f2007 = new File(path2007);
        try {
            readExcel(f2007);
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
 
    /**
     * 对外提供读取excel 的方法
     */
    public static List<List<Object>> readExcel(File file) throws IOException {
        
        String fileName = file.getName();
        String extension = fileName.lastIndexOf(".") == -1 ? "" : fileName
                .substring(fileName.lastIndexOf(".") + 1);
        if ("xls".equals(extension)) {
            return read2003Excel(file);
        } else if ("xlsx".equals(extension)) {
            return read2007Excel(file);
        } else {
            throw new IOException("不支持的文件类型");
        }
    }
    
    /**
     * 读取 office 2003 excel
     * 
     * @throws IOException
     * @throws FileNotFoundException
     */
    private static List<List<Object>> read2003Excel(File file) throws IOException {
        
        List<List<Object>> list = new LinkedList<List<Object>>();
        @SuppressWarnings("resource")
        HSSFWorkbook hwb = new HSSFWorkbook(new FileInputStream(file));
        HSSFSheet sheet = hwb.getSheetAt(0);
        Object value = null;
        HSSFRow row = null;
        HSSFCell cell = null;
        System.out.println("读取office 2003 excel内容如下：");
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            List<Object> linked = new LinkedList<Object>();
            for (int j = row.getFirstCellNum(); j <= row.getLastCellNum(); j++) {
                cell = row.getCell(j);
                if (cell == null) {
                    continue;
                }
                DecimalFormat df = new DecimalFormat("0");
                SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                DecimalFormat nf = new DecimalFormat("0.00");
                switch (cell.getCellType()) {
                case XSSFCell.CELL_TYPE_STRING:
                    value = cell.getStringCellValue();
                    System.out.print("\t" + value + "\t");
                    break;
                case XSSFCell.CELL_TYPE_NUMERIC:
                    if ("@".equals(cell.getCellStyle().getDataFormatString())) {
                        value = df.format(cell.getNumericCellValue());
                    } else if ("General".equals(cell.getCellStyle()
                            .getDataFormatString())) {
                        value = nf.format(cell.getNumericCellValue());
                    } else {
                        value = sdf.format(HSSFDateUtil.getJavaDate(cell.getNumericCellValue()));
                    }
                    System.out.print("\t" + value + "\t");
                    break;
                case XSSFCell.CELL_TYPE_BOOLEAN:
                    value = cell.getBooleanCellValue();
                    System.out.print("\t" + value + "\t");
                    break;
                case XSSFCell.CELL_TYPE_BLANK:
                    value = "";
                    System.out.print("\t" + value + "\t");
                    break;
                default:
                    value = cell.toString();
                    System.out.print("\t" + value + "\t");
                }
                if (value == null || "".equals(value)) {
                    continue;
                }
                linked.add(value);
            }
            System.out.println("\t");
            list.add(linked);
        }
        return list;
    }
    
    /**
     * 读取Office 2007 excel
     */
    private static List<List<Object>> read2007Excel(File file)throws IOException {
        
        List<List<Object>> list = new LinkedList<List<Object>>();
        @SuppressWarnings("resource")
        XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(file));
        XSSFSheet sheet = xwb.getSheetAt(1);
        String value = null;
        XSSFRow row = null;
        XSSFCell cell = null;
        for (int i = sheet.getFirstRowNum(); i <= sheet.getPhysicalNumberOfRows(); i++) {
            row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            XSSFCell creatCell = row.createCell(34);
            List<Integer> liDay = new ArrayList<>();
            int sum = 0;
            for (int j = row.getFirstCellNum(); j <= row.getLastCellNum(); j++) {
                cell = row.getCell(j);
                if (cell == null) {
                    continue;
                }
                SimpleDateFormat sdf = new SimpleDateFormat("HH:mm");// 格式化时间字符串
                SimpleDateFormat sd = new SimpleDateFormat("MM-dd");// 格式化日期字符串
                //判断数据类型
                switch (cell.getCellType()) {
                case XSSFCell.CELL_TYPE_STRING:
                    value = cell.getStringCellValue();
                    System.out.print(value + "  ");
                    break;
                case XSSFCell.CELL_TYPE_NUMERIC:
                    Integer num1 = new Integer(18);
                    Integer num2 = new Integer(00);
                    if(i == 0){   
                        creatCell.setCellValue("\t总加班时长");
                        value = sd.format(HSSFDateUtil.getJavaDate(cell.getNumericCellValue()));
                    } else {
                        value = sdf.format(HSSFDateUtil.getJavaDate(cell.getNumericCellValue()));
                        String[] xTime = value.split(":");
                        String str1 = xTime[0];//时
                        String str2 = xTime[1];//分
                        int str11 = Integer.parseInt(str1);
                        int str22 = Integer.parseInt(str2);
                        Integer num3 = str11 - num1;
                        Integer num4 = str22 - num2;
                        int e = num3 * 60 + num4;
                        sum += e;
                        liDay.add(e);
                    }
                    System.out.print("\t\t" + value);
                    break;
                case XSSFCell.CELL_TYPE_BOOLEAN:
                    System.out.print("\t" + value);
                    break;
                case XSSFCell.CELL_TYPE_BLANK:
                    value = "";
                    System.out.print("\t" + value);
                    break;
                case XSSFCell.CELL_TYPE_FORMULA:
                    value = "";
                    System.out.print("\t" + value);
                    break;    
                default:
                    System.out.print("\t" + value + "\t");
                }
                if (value == null || "".equals(value)) {
                    System.out.print("\t" + "空值");
                    continue;
                }
                liDay.add(sum);
                if(i == 0) {
                    creatCell.setCellValue("\t总加班时长");
                } else {
                    creatCell.setCellValue("\t加班" + sum + "分钟");
                }
            }
            System.out.println("");
        }
        return list;
    }
}
