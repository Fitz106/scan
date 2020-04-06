package com.test;

import com.google.common.base.Strings;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;
import java.util.regex.Pattern;

public class Main {

    public static List<String> regxList = new ArrayList<String>();


    public static void main(String[] args) throws Exception{
        getRegxList();
        String isHandleCurrentDir = scanner("是否处理当前目录下的文件，Y/N");
        if(isHandleCurrentDir.toUpperCase().equals("Y")){
            //处理当前目录
            String currentPath = System.getProperty("user.dir");
            System.out.println(currentPath);//user.dir指定了当前的路径
            Integer colNum = Integer.valueOf(scanner("匹配第几列数据"));

            readDataAndHandle(currentPath,colNum);
        }else{
            String path = scanner("文件路径");
            Integer colNum = Integer.valueOf(scanner("匹配第几列数据"));

            readDataAndHandle(path,colNum);
        }
    }


    /**
     *  根据所给路径处理数据
     * @param path
     * @param colNum
     * @throws Exception
     */
    private static void readDataAndHandle(String path, Integer colNum) throws Exception{
        File file = new File(path);
        if (file.isDirectory()) {
            //如果输入路径是目录，则批量处理目录下的文件
            File[] files = file.listFiles();
            for (File fi : files) {
                // 对文件进行过滤，只读取Excel文件

                String fileName = fi.getName();
                String fileType = fileName.substring(fileName.lastIndexOf(".") + 1, fileName.length());

                if (fileType.equalsIgnoreCase("xls")){
                    System.out.println();

                    System.out.println("开始处理文件：" + fileName);
                    String fileFrontStr = fileName.substring(0,fileName.lastIndexOf("."));
                    String resultFileName = fileFrontStr + "_result.xls";
                    //创建workbook
                    HSSFWorkbook workbookResult = new HSSFWorkbook();
                    //添加Worksheet（不添加sheet时生成的xls文件打开时会报错)
                    workbookResult.createSheet("扫描结果");
                    handleXls(colNum, fi, resultFileName, workbookResult);
                    System.out.println("完成处理，结果存放在：" + resultFileName);

                }
                if (fileType.equalsIgnoreCase("xlsx")){
                    System.out.println();

                    System.out.println("开始处理文件：" + fileName);
                    String fileFrontStr = fileName.substring(0,fileName.lastIndexOf("."));
                    String resultFileName = fileFrontStr + "_result.xls";
                    //创建workbook
                    HSSFWorkbook workbookResult = new HSSFWorkbook();
                    //添加Worksheet（不添加sheet时生成的xls文件打开时会报错)
                    workbookResult.createSheet("扫描结果");
                    handleXlsx(colNum, fi, resultFileName, workbookResult);
                    System.out.println("完成处理，结果存放在：" + resultFileName);

                }
            }
        }else{
            //如果输入路径是文件，则单独处理该文件
            String fileName = file.getName();
            String fileType = fileName.substring(fileName.lastIndexOf(".") + 1, fileName.length());
            String fileFrontStr = fileName.substring(0,fileName.lastIndexOf("."));

            String resultFileName = fileFrontStr + "_result.xls";
            //创建workbook
            HSSFWorkbook workbookResult = new HSSFWorkbook();
            //添加Worksheet（不添加sheet时生成的xls文件打开时会报错)
            workbookResult.createSheet("扫描结果");
            if (fileType.equalsIgnoreCase("xls")){
                System.out.println();

                System.out.println("开始处理文件：" + fileName);
                handleXls(colNum, file, resultFileName, workbookResult);
                System.out.println("完成处理，结果存放在：" + resultFileName);
            }
            if (fileType.equalsIgnoreCase("xlsx")){
                System.out.println();

                System.out.println("开始处理文件：" + fileName);
                handleXlsx(colNum, file, resultFileName, workbookResult);
                System.out.println("完成处理，结果存放在：" + resultFileName);
            }

        }

    }

    /**
     * 处理.xlsx格式excel文件
     * @param colNum
     * @param fi
     * @param resultFileName
     * @param workbookResult
     */
    private static void handleXlsx(Integer colNum, File fi, String resultFileName, HSSFWorkbook workbookResult) {
        XSSFWorkbook xssfWorkbook;
        XSSFSheet sheet = null;
        FileInputStream fs;
        int lastRowNum = 0;
        try{
            fs = new FileInputStream(fi.getAbsolutePath());
            xssfWorkbook = new XSSFWorkbook(fs);
            sheet = xssfWorkbook.getSheetAt(0);
            lastRowNum = sheet.getLastRowNum();
            handleXLSX(xssfWorkbook,sheet,lastRowNum,colNum,workbookResult,resultFileName);
        }catch (Exception e){

        }
    }

    /**
     * 处理.xls格式excel文件
     * @param colNum
     * @param fi
     * @param resultFileName
     * @param workbookResult
     */
    private static void handleXls(Integer colNum, File fi, String resultFileName, HSSFWorkbook workbookResult) {
        HSSFWorkbook hssfWorkbook;
        HSSFSheet sheet = null;
        FileInputStream fs;
        int lastRowNum = 0;
        try{
            fs = new FileInputStream(fi.getAbsolutePath());
            hssfWorkbook = new HSSFWorkbook(fs);
            sheet = hssfWorkbook.getSheetAt(0);
            lastRowNum = sheet.getLastRowNum();
            handleXLS(hssfWorkbook,sheet,lastRowNum,colNum,workbookResult,resultFileName);
        }catch (IOException e){
            e.printStackTrace();
        }
    }

    /**
     * 遍历正则库识别是否命中规则，若命中则返回true
     * @param claimPwd
     * @return
     */
    public static boolean scanPassword(String claimPwd){
//        System.out.println("原文：" + claimPwd);
        boolean empty = regxList.isEmpty();
        if (empty){
            return false;
        }else{
            for(String regx : regxList){
                boolean matches = Pattern.matches(regx, claimPwd);
                if (matches){
                    System.out.println("原文：" + claimPwd +   "   ||   命中正则：" + regx);
                    return matches;
                }
            }
            return false;
        }
    }


    /**
     * 读取本地正则表达式库
     */
    public static void getRegxList(){
        String currentPath = System.getProperty("user.dir");

        File file = new File(currentPath +"\\regx.txt");
        BufferedReader reader = null;
        try {
            reader = new BufferedReader(new FileReader(file));
            String tempStr;
            while ((tempStr = reader.readLine()) != null) {
                regxList.add(tempStr);
            }
            reader.close();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (reader != null) {
                try {
                    reader.close();
                } catch (IOException e1) {
                    e1.printStackTrace();
                }
            }
        }

        System.out.println("----------以下是正则表达式库：");
        for (String regx : regxList){
            System.out.println(regx);
        }
        int num = regxList.size();
        System.out.println("----------共 " + num + "条规则");


    }



    public static void handleXLS(HSSFWorkbook hssfWorkbook, HSSFSheet sheet , int lastRowNum, int colNum,
                                 HSSFWorkbook workbookResult,String reultFileName){
        HSSFSheet sheetResult = workbookResult.getSheet("扫描结果");
        int resultRowIndex = 0;
        for (int i = 0; i <= lastRowNum; i++){
            HSSFRow row = sheet.getRow(i);
            HSSFCell cell = row.getCell(colNum);
            if (cell == null){
                break;
            }else{
                String claimPwd = cell.getStringCellValue();
                if (Strings.isNullOrEmpty(claimPwd)){
                    break;
                }else{
                    boolean flag = scanPassword(claimPwd);
                    if (flag){
                        //命中规则，说明是明文密码
                        HSSFRow rowResult = sheetResult.createRow(resultRowIndex);
                        int lastCellNum = row.getLastCellNum();
                        for (int j = 0; j < lastCellNum; j++){
                            HSSFCell cellResult = rowResult.createCell(j);

                            HSSFCell cellByScan = row.getCell(j);
                            cellByScan.setCellType(Cell.CELL_TYPE_STRING);
                            String scanValue = cellByScan.getStringCellValue();
                            cellResult.setCellValue(scanValue);
                        }

                        resultRowIndex++;
                    }else{
                        //未命中规则，说明不是明文密码
                    }

                }
            }

        }
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(reultFileName);
            workbookResult.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public static void handleXLSX(XSSFWorkbook xssfWorkbook, XSSFSheet sheet , int lastRowNum, int colNum,
                                 HSSFWorkbook workbookResult,String reultFileName){
        HSSFSheet sheetResult = workbookResult.getSheet("扫描结果");
        int resultRowIndex = 0;
        for (int i = 1; i < lastRowNum; i++){
            XSSFRow row = sheet.getRow(i);
            XSSFCell cell = row.getCell(colNum);
            if (cell == null){
                break;
            }else{
                String claimPwd = cell.getStringCellValue();
                if (Strings.isNullOrEmpty(claimPwd)){
                    break;
                }else{
                    boolean flag = scanPassword(claimPwd);
                    if (flag){
                        //命中规则，说明是明文密码
                        HSSFRow rowResult = sheetResult.createRow(resultRowIndex);
                        int lastCellNum = row.getLastCellNum();
                        for (int j = 0; j < lastCellNum; j++){
                            HSSFCell cellResult = rowResult.createCell(j);

                            XSSFCell cellByScan = row.getCell(j);
                            String scanValue = cellByScan.getStringCellValue();
                            cellResult.setCellValue(scanValue);
                        }

                        resultRowIndex++;
                    }else{
                        //未命中规则，说明不是明文密码
                    }

                }
            }

        }
        FileOutputStream out = null;
        try {
            out = new FileOutputStream(reultFileName);
            workbookResult.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 交互工具类，识别控制台输入数据
     * @param tip
     * @return
     * @throws Exception
     */
    public static String scanner(String tip) throws Exception{
        Scanner scanner = new Scanner(System.in);
        StringBuilder help = new StringBuilder();
        help.append("请输入" + tip + ":");
        System.out.println(help.toString());
        if (scanner.hasNext()){
            String ipt = scanner.next();
            if (ipt != null && ipt.length() > 0){
                return ipt;
            }
        }
        throw new Exception("请输入正确的" + tip + "!");
    }

}
