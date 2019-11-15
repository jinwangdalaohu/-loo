package com.jjb.git;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
public class ExcelExamRead {
    /** 读Excel文件内容 */
    public void showExcel(String excelName) {
        File file = new File(excelName);
        FileInputStream in = null;
        try {
// 创建对Excel工作簿文件的引用
            in = new FileInputStream(file);
            HSSFWorkbook hwb = new HSSFWorkbook(in);
            HSSFSheet sheet = hwb.getSheet("myFirstExcel");// 根据指定的名字来引用此Excel中的有效工作表
// 读取Excel 工作表的数据
            System.out.println("下面是Excel文件" + file.getAbsolutePath() + "的内容：");
            HSSFRow row = null;
            HSSFCell cell = null;
            int rowNum = 0;
// 行标
            int colNum = 0;
// 列标
            for (; rowNum < 9; rowNum++) {
// 获取第rowNum行
                row = sheet.getRow((short) rowNum);
                for (colNum = 0; colNum < 5; colNum++) {
                    cell = row.getCell((short) colNum);// 根据当前行的位置来创建一个单元格对象
                    System.out.print(cell.getStringCellValue() + "\t");// 获取当前单元格中的内容
                }
                System.out.println(); // 换行
            }
            in.close();
        } catch (Exception e) {
            System.out
                    .println("读取Excel文件" + file.getAbsolutePath() + "失败：" + e);
        } finally {
            if (in != null) {
                try {
                    in.close();
                } catch (IOException e1) {
                }
            }
        }
    }
    public static void main(String[] args) {
        ExcelExamRead excel = new ExcelExamRead();
        String excelName = "D:/ExcelExamRead.xls";
        excel.showExcel(excelName);
    }
}
