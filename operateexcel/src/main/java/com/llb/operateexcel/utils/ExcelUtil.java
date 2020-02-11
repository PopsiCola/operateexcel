package com.llb.operateexcel.utils;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

/**
 * 操作Excel：导入、导出
 * @Author llb
 * Date on 2020/1/8
 */
public class ExcelUtil {

    /**
     * 导出到excel
     * @param path
     */
    public void export2Excel(String path) {
        //创建工作簿
        HSSFWorkbook hssfWorkbook = new HSSFWorkbook();
        //创建shetname页名
        HSSFSheet sheet = hssfWorkbook.createSheet("员工信息");
        //创建一行，下标0开始
        HSSFRow row = sheet.createRow(0);
        //创建这行的列
        HSSFCell cell = row.createCell(0);
        cell.setCellValue("姓名");
        row.createCell(1).setCellValue("性别");
        row.createCell(2).setCellValue("地址");

        row = sheet.createRow(1);
        row.createCell(0).setCellValue("张三");
        row.createCell(1).setCellValue("男");
        row.createCell(2).setCellValue("山东省济南市");

        //保存位置
        FileOutputStream out = null;
        File file = new File(path);
        try {
            out = new FileOutputStream(file);
            hssfWorkbook.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if(out!=null) {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    /**
     * 从excel导入
     * @param sourcePath 原地址
     * @param destPath 目的地址
     * @param ksName 科室名称
     * @param qnDate 签名日期时间
     */
    public void importFromExcel(String sourcePath, String destPath, String ksName, String qnDate) {
        File file = new File(sourcePath);
        File file1 = new File(destPath);

        if(!file1.exists()) {
            try {
                file1.createNewFile();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        //获取输入流
        FileInputStream in = null;
        FileInputStream in1 = null;
        //获取输出流
        FileOutputStream out = null;

        //创建工作簿
        Workbook wb = null;
        Workbook saveWb = null;
        //创建工作表
        Sheet sheet = null;

        //科室列
        int ksColum = 0;
        //签名日期时间列
        int qmDateColum = 0;

        try {
            in = new FileInputStream(file);
            in1 = new FileInputStream(file1);
            String extString = sourcePath.substring(sourcePath.lastIndexOf("."));
            String destString = destPath.substring(destPath.lastIndexOf("."));
            if (".xls".equals(extString)) {
                wb = new HSSFWorkbook(in);
            } else if (".xlsx".equals(extString)) {
                wb = new XSSFWorkbook(in);
            } else {
                wb = null;
            }

            if(".xls".equals(destString)) {
                saveWb = new HSSFWorkbook(in1);
            } else if(".xlsx".equals(destString)) {
                saveWb = new XSSFWorkbook(in1);
            } else {
                saveWb = null;
            }

            //获取工作表：如果工作表是两张，则选择第二张，如果一张，就默认
            int num = wb.getNumberOfSheets();
            sheet = wb.getSheetAt(num==1?0:1);
            Sheet saveSheet = saveWb.getSheetAt(0);
            //判断保存的工作表有无
            if(saveSheet == null) {
                //保存的工作表
                saveSheet = saveWb.createSheet(sheet.getSheetName());
            }

            //获取最后一行
            int nums = sheet.getLastRowNum();
            for (int i = 1; i <= nums; i++) {
                //获取行
                Row row = sheet.getRow(i);
                if(row == null) {
                    nums = i;
                    System.out.println("执行了");
                    continue;
                }
                Cell cell = row.getCell(0);
                if(cell == null || "".equals(cell)) {
                    nums = i;
                    System.out.println("执行了");
                    break;
                }
            }
            //保存到新建的工作表
            int nums1 = saveSheet.getLastRowNum();
            for (int i = 0; i <= nums1; i++) {
                //获取行
                Row row = saveSheet.getRow(i);
                if(row == null) {
                    nums1 = i;
                    System.out.println("执行了");
                    continue;
                }
                Cell cell = row.getCell(0);
                if(null == cell) {
                    nums1 = i;
                    System.out.println("执行了saveFile");
                    break;
                }
            }
            //找出科室名称，签名日期时间列,将第一个表补全科室名称和签名日期时间
            int colum = sheet.getRow(0).getPhysicalNumberOfCells();
            for (int c = 0; c <=colum; c++) {
                if(sheet.getRow(0).getCell(c) == null || sheet.getRow(0).getCell(0) == null|| "".equals(sheet.getRow(0).getCell(0))) {
                    continue;
                } else {
                    Cell cell = sheet.getRow(0).getCell(c);
                    cell.setCellType(CellType.STRING);
                    String value = cell.getStringCellValue();
                    if("科室名称".equals(value)) {
                        ksColum = c;
                        System.out.println("科室名称：" + c);
                    }
                    if ("签名日期时间".equals(value)) {
                        qmDateColum = c;
                        System.out.println("签名日期时间：" + c);
                    }
                }
            }

            System.out.println("记录数为：" + nums);

            //获取表格内容的行号
            for (int i = 1; i <= nums; i++) {
                //获取行
                Row row = sheet.getRow(i);
                if(row == null) {
                   break;
                }

                //创建下一条数据
                Row row1 = saveSheet.createRow(++nums1);

                //获取总列数
                int colums = sheet.getRow(0).getPhysicalNumberOfCells();

                for (int j = 0; j <= colums; j++) {
//                    System.out.println(j);
                    Cell cell = row.getCell(j);
                    /*if(cell == null || row.getCell(0) == null|| "".equals(row.getCell(0))) {
                        cell = row.createCell(j);
                        cell.setCellType(CellType.STRING);
                    }*/
                    if(cell == null) {
                        cell = row.createCell(j);
                        cell.setCellType(CellType.STRING);
                    }

                    if(j == ksColum) {

                        //保存到新的工作表中
                        Cell cell1 = row1.createCell(j);
                        //获取单元格大小
                        int columnWidth = cell.getSheet().getColumnWidth(j);
                        short height = cell.getRow().getHeight();
                        //设置字体大小
                        Font font = saveWb.createFont();
                        font.setFontName("宋体");
                        font.setFontHeightInPoints((short)11);
                        cell1.getCellStyle().setFont(font);
                        cell1.getCellStyle().setAlignment(HorizontalAlignment.CENTER);
                        cell1.getSheet().setColumnWidth(j, columnWidth);
                        cell1.getRow().setHeight(height);

                        cell1.setCellValue(ksName);

                        System.out.print(ksName+ "=============");
                    } else if(j == qmDateColum) {
                        //保存到新的工作表中
                        Cell cell1 = row1.createCell(j);
                        //获取单元格大小
                        int columnWidth = cell.getSheet().getColumnWidth(j);
                        short height = cell.getRow().getHeight();
                        //设置字体大小
                        Font font = saveWb.createFont();
                        font.setFontName("宋体");
                        font.setFontHeightInPoints((short)11);
                        cell1.getCellStyle().setFont(font);
                        cell1.getCellStyle().setAlignment(HorizontalAlignment.CENTER);
                        cell1.getSheet().setColumnWidth(j, columnWidth);
                        cell1.getRow().setHeight(height);

                        cell1.setCellValue(qnDate);

                        System.out.println(qnDate);
                    } else {
//                        Cell cell = row.getCell(j);
                        //获取单元格大小
                        int columnWidth = cell.getSheet().getColumnWidth(j);
                        short height = cell.getRow().getHeight();
                        //                            CellStyle cellStyle = cell.getCellStyle();
                        cell.setCellType(CellType.STRING);
                        String value = row.getCell(j).getStringCellValue();

                        //保存到新的工作表中
                        Cell cell1 = row1.createCell(j);
                        //设置字体大小
                        Font font = saveWb.createFont();
                        font.setFontName("宋体");
                        font.setFontHeightInPoints((short)11);
                        cell1.getCellStyle().setFont(font);
                        cell1.getCellStyle().setAlignment(HorizontalAlignment.CENTER);
                        cell1.getSheet().setColumnWidth(j, columnWidth);
                        cell1.getRow().setHeight(height);

                        cell1.setCellValue(value);

                        //                        System.out.println(value);
                    }

                }
            }


            //保存位置
            out = new FileOutputStream(file1);
            saveWb.write(out);

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if(saveWb != null) {
                try {
                    saveWb.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if(wb != null) {
                try {
                    wb.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if(out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if(in1 != null) {
                try {
                    in1.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if(in != null) {
                try {
                    in.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

    }

    /**
     * 补全表的格式数据(科室名称，签名日期时间)
     * @param destPath
     */
    public void changeExcel(String destPath, String ksName, String qmDate) {
        File file = new File(destPath);
        if(!file.exists()) {
            try {
                file.createNewFile();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        //获取输入流
        FileInputStream in = null;
        //获取输出流
        FileOutputStream out = null;

        //创建工作簿
        Workbook wb = null;
        //创建工作表
        Sheet sheet = null;
        //科室列
        int ksColum = 0;
        //签名日期时间列
        int qmDateColum = 0;

        try {
            in = new FileInputStream(file);
            String extString = destPath.substring(destPath.lastIndexOf("."));
            if (".xls".equals(extString)) {
                wb = new HSSFWorkbook(in);
            } else if (".xlsx".equals(extString)) {
                wb = new XSSFWorkbook(in);
            } else {
                wb = null;
            }
            //获取工作表：如果工作表是两张，则选择第二张，如果一张，就默认
            int num = wb.getNumberOfSheets();
            sheet = wb.getSheetAt(num==1?0:1);

            //获取最后一行
            int nums = sheet.getLastRowNum();
            for (int i = 1; i < nums; i++) {
                //获取行
                Row row = sheet.getRow(i);
                if(row == null) {
                    nums = i;
                    System.out.println("执行了");
                    continue;
                }
                Cell cell  = row.getCell(0);
                if(cell == null|| "".equals(cell)) {
                    nums = i;
                    System.out.println("执行了");
                    break;
                }
            }
            //找出科室名称，签名日期时间列,将第一个表补全科室名称和签名日期时间
            int colum = sheet.getRow(0).getPhysicalNumberOfCells();
            for (int c = 0; c <=colum; c++) {
                if(sheet.getRow(0).getCell(c) == null || sheet.getRow(0).getCell(0) == null|| "".equals(sheet.getRow(0).getCell(0))) {
                    continue;
                } else {
                    Cell cell = sheet.getRow(0).getCell(c);
                    cell.setCellType(CellType.STRING);
                    String value = cell.getStringCellValue();
                    if("科室名称".equals(value)) {
                        ksColum = c;
                        System.out.println("科室名称：" + c);
                    }
                    if ("签名日期时间".equals(value)) {
                        qmDateColum = c;
                        System.out.println("签名日期时间：" + c);
                    }
                }
            }

            System.out.println("总记录数:" + nums);

            //获取表格内容的行号
            for (int i = 1; i <= nums; i++) {
                //获取行
                Row row = sheet.getRow(i);
                if(row == null) {
                    continue;
                }
                //获取总列数
                int colums = sheet.getRow(0).getPhysicalNumberOfCells();

                for (int j = 0; j <= colums; j++) {

                    if(row.getCell(j) == null || row.getCell(0) == null|| "".equals(row.getCell(0))) {
                        Cell cell = row.createCell(j);
                        cell.setCellType(CellType.STRING);
                        //判断是否是科室或者是签名日期时间列
                        if(j == ksColum) {

                            //保存到新的工作表中
                            Cell cell1 = row.createCell(j);
                            //获取单元格大小
                            int columnWidth = cell.getSheet().getColumnWidth(j);
                            short height = cell.getRow().getHeight();
                            //设置字体大小
                            Font font = wb.createFont();
                            font.setFontName("宋体");
                            font.setFontHeightInPoints((short)11);
                            cell1.getCellStyle().setFont(font);
                            cell1.getCellStyle().setAlignment(HorizontalAlignment.CENTER);
                            cell1.getSheet().setColumnWidth(j, columnWidth);
                            cell1.getRow().setHeight(height);

                            cell1.setCellValue(ksName);

                            System.out.print(ksName + "==============");
                        } else if(j == qmDateColum) {
                            //保存到新的工作表中
                            Cell cell1 = row.createCell(j);
                            //获取单元格大小
                            int columnWidth = cell.getSheet().getColumnWidth(j);
                            short height = cell.getRow().getHeight();
                            //设置字体大小
                            Font font = wb.createFont();
                            font.setFontName("宋体");
                            font.setFontHeightInPoints((short)11);
                            cell1.getCellStyle().setFont(font);
                            cell1.getCellStyle().setAlignment(HorizontalAlignment.CENTER);
                            cell1.getSheet().setColumnWidth(j, columnWidth);
                            cell1.getRow().setHeight(height);

                            cell1.setCellValue(qmDate);

                            System.out.println(qmDate);
                        }
                    }
                }

            }
            //保存位置
            out = new FileOutputStream(file);
            wb.write(out);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if(wb != null) {
                try {
                    wb.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if(out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if(in != null) {
                try {
                    in.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    public static void main(String[] args) {
        ExcelUtil excelUtil = new ExcelUtil();
//        String path = "C:\\Users\\ASUS\\Desktop\\刘乐彬-2.3-职工健康情况排查表.xls";
        String path = "C:\\Users\\ASUS\\Desktop\\新建文件夹 (2)\\皮肤科\\皮肤科移动护理-出入量记录12月.xlsx";
        String dest = "F:\\皮肤科移动护理-出入量记录12月.xlsx";
//        excelUtil.export2Excel(path);
//        excelUtil.importFromExcel(path, dest);
    }

}
