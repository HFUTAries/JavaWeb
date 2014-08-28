/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package exceloperator;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

public class ExcelOperator {

    public String reg = ""; // 文件内容
    public String state = ""; // 机构名称
    public String date = ""; // 日期
    public String sera = ""; // 序列标记

    public InputStream inp;
    public FileOutputStream fileOut;
    public Workbook wb;
    public Sheet sheet;
    public String path = ""; // 文本文件路径
    public String txt = ""; // 文本文件路径
    public String name = ""; // 文本文件路径
    public String tag[] = {"序号", "县级WIS代码", "结婚登记日期", "登记证字号", "身份证件号", "姓名",
        "民族", "户籍地/住址", "县/乡级WIS代码", "文化程度", "出生日期", "是否再婚", "身份证件号",
        "姓名", "民族", "女方基本情况", "县/乡级WIS代码", "文化程度", "出生日期", "是否再婚"};

    public ExcelOperator() {
    }

    public ExcelOperator(String filepath, String txtName, String excelName) {
        path = filepath;
        txt = path + txtName;
        name = excelName;
    }

    public boolean CreateExcel(String path) {
        try {
            wb = new HSSFWorkbook(); // 创建新的Excel工作簿
            sheet = wb.createSheet("民政婚姻登记");
            Row row = sheet.createRow(3);
            Cell cell = row.createCell(0);
            cell.setCellType(Cell.CELL_TYPE_STRING);
            cell.setCellValue(1);
            //OutputStreamReader  in = new OutputStreamReader(new FileOutputStream(path),"UTF-8");
            fileOut = new FileOutputStream(path);
            wb.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            System.out.println("CreateExcel() ERRO\n");
            return false;
        }
        return true;
    }

    /**
     *
     * @param filePath : 写入指定的路径
     * @param args ： 待写入的内容
     * @throws IOException
     * @throws InvalidFormatException ： 抛出格式错误异常
     */
    public void WriteExcel(String filePath, String[] args) throws IOException,
            InvalidFormatException {
        try {
            inp = new FileInputStream(filePath);
            wb = WorkbookFactory.create(inp);
            sheet = wb.getSheetAt(0);
            Row row = sheet.createRow(0);
            Cell cell;
            for (int i = 0; i < 20; i++) {
                cell = row.createCell(i);
                cell.setCellType(Cell.CELL_TYPE_STRING);
                cell.setCellValue(args[i]);
            }
            OutputStream out = new FileOutputStream(filePath);
            //fileOut = new FileOutputStream(filePath);
            wb.write(out);
        } catch (InvalidFormatException e) {
            System.out.println("WriteExcel() ERRO in tag\n");
        }
    }

    public void WriteExcel(String filePath, String cont, int r) {
        String[] couple = cont.split(" ");
        try {
            inp = new FileInputStream(filePath);
            wb = WorkbookFactory.create(inp);
            sheet = wb.getSheetAt(0);
            Row row = sheet.createRow(r);
            Cell cell;
            cell = row.createCell(0);
            cell.setCellType(Cell.CELL_TYPE_STRING);
            cell.setCellValue(couple[0]);
            cell = row.createCell(1);
            cell = row.createCell(2);
            cell.setCellType(Cell.CELL_TYPE_STRING);
            cell.setCellValue(couple[1]);
            cell = row.createCell(3);
            cell.setCellType(Cell.CELL_TYPE_STRING);
            cell.setCellValue(couple[2]);
            cell = row.createCell(4); // 身份证号码
            cell.setCellType(Cell.CELL_TYPE_STRING);
            cell.setCellValue(couple[5]);
            cell = row.createCell(5); // 姓名
            cell.setCellType(Cell.CELL_TYPE_STRING);
            cell.setCellValue(couple[3]);
            cell = row.createCell(6);
            cell = row.createCell(7); // 户籍地址
            cell.setCellType(Cell.CELL_TYPE_STRING);
            cell.setCellValue(couple[6]);
            cell = row.createCell(8);
            cell = row.createCell(9);
            cell = row.createCell(10); // 出生日期
            cell.setCellType(Cell.CELL_TYPE_STRING);
            if (couple[5].length() > 13) {
                cell.setCellValue(couple[5].substring(6, 10) + "-"
                        + couple[5].substring(10, 12) + "-"
                        + couple[5].substring(12, 14));
            } else {
                cell.setCellValue("erro");
            }
            cell = row.createCell(11); // 是否再婚
            cell.setCellType(Cell.CELL_TYPE_STRING);
            cell.setCellValue(couple[7]);

            cell = row.createCell(12); // 身份证号码
            cell.setCellType(Cell.CELL_TYPE_STRING);
            cell.setCellValue(couple[14]);
            cell = row.createCell(13); // 姓名
            cell.setCellType(Cell.CELL_TYPE_STRING);
            cell.setCellValue(couple[12]);
            cell = row.createCell(0xe);
            cell = row.createCell(15); // 户籍地址
            cell.setCellType(Cell.CELL_TYPE_STRING);
            cell.setCellValue(couple[15]);
            cell = row.createCell(16);
            cell = row.createCell(17);
            cell = row.createCell(18); // 出生日期
            cell.setCellType(Cell.CELL_TYPE_STRING);
            if (couple[14].length() > 13) {
                cell.setCellValue(couple[14].substring(6, 10) + "-"
                        + couple[14].substring(10, 12) + "-"
                        + couple[14].substring(12, 14));
            } else {
                cell.setCellValue("erro");
            }
            cell = row.createCell(19); // 是否再婚
            cell.setCellType(Cell.CELL_TYPE_STRING);
            cell.setCellValue(couple[16]);
            fileOut = new FileOutputStream(filePath);
            wb.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            System.out.print("Excel.Write() ERRO, Exist IOException in "
                    + couple[0] + " line\n");
            // e.printStackTrace();
        } catch (InvalidFormatException e) {
            System.out.print("Excel.Write() ERRO, ExistInvalidFormatException in "
                    + couple[0] + " line\n");
            // e.printStackTrace();
        }
    }

    public void readTxt() {
        File file = new File(txt);
        BufferedReader reader;
        String tempString;
        try {
            InputStreamReader read = new InputStreamReader(new FileInputStream(file)); 
            reader = new BufferedReader(read);
            tempString = reader.readLine();
            int line = 1;
            // 去除没有内容的行
            while ("".equals(tempString) || tempString == null) {
                tempString = reader.readLine();
                line++;
            }

            // 遍历文件并写入Excel
            tempString = reader.readLine();
            reg = tempString;
            tempString = reader.readLine();
            date = tempString.split("[ *]")[2].substring(5).replaceAll("-", "");
            state = tempString.split("[ *]")[0];
            sera = reader.readLine();
            name = path + date + "民政新婚登记.xls";
            CreateExcel(name);
            WriteExcel(name, tag);
            while (!"".equals(tempString) || null != tempString) {
                tempString = reader.readLine();
                tempString = tempString + " " + reader.readLine().trim();
                line += 2;

                if (line > 1000) {
                    break;
                }
                if (!tempString.trim().equals("?")) {
                    tempString = tempString.replaceAll("[ *]", " ");
                    WriteExcel(name, tempString, line / 2);
                } else if ((tempString = reader.readLine()) != null) {
                    reg = tempString;
                    tempString = reader.readLine();
                    date = tempString.split("[ *]")[2].substring(5).replaceAll(
                            "-", "");
                    state = tempString.split("[ *]")[0];
                    sera = reader.readLine();
                    name = path + date + "民政新婚登记.xls";
                    line = 2;
                    CreateExcel(name);
                    WriteExcel(name, tag);
                } else {
                    break;
                }
            }
            reader.close();
        } catch (IOException | InvalidFormatException e) {
            //e.printStackTrace();
        }
    }
    /*public static void main(String[] args) {
        String name = "test";
        String path = "C:\\Documents and Settings\\Administrator\\桌面\\";
        String txtName = "13年10.1-14.04.24.txt";
        ExcelOperator test = new ExcelOperator(path, txtName, path + name + ".xls");
        test.readTxt();
    }*/

}
