package com.leolau.export;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.Statement;

/**
 * @author g.le
 */
public class ExportExcel {
    public static void main(String[] args) throws Exception {
        Class.forName("com.mysql.jdbc.Driver");
        String url = "jdbc:mysql://127.0.0.1:3306/XXXX?useUnicode=true&amp;characterEncoding=utf-8";
        String user = "root";
        String password = "123456";
        Connection conn = DriverManager.getConnection(url, user, password);
        Statement statement = conn.createStatement();
        String sql = "SQL";
        ResultSet rs = statement.executeQuery(sql);

        //创建工作簿
        HSSFWorkbook workBook = new HSSFWorkbook();
        //创建工作表
        HSSFSheet sheet = workBook.createSheet("XXX");
        int count = 1;
        while (rs.next()) {
            HSSFRow row = sheet.createRow(count);
            HSSFCell cell = row.createCell(0);
            cell.setCellValue(rs.getString("field"));
            HSSFCell cell2 = row.createCell(1);
            cell2.setCellValue(rs.getString("field"));
            JSONArray criminalMsg = JSONArray.parseArray(rs.getString("field"));
            HSSFCell cell4 = row.createCell(2);
            cell4.setCellValue(JSONObject.parseObject(criminalMsg.get(0).toString()).getString("field"));
            HSSFCell cell5 = row.createCell(3);
            cell5.setCellValue(JSONObject.parseObject(criminalMsg.get(0).toString()).getString("field"));
            HSSFCell cell6 = row.createCell(4);
            cell6.setCellValue(JSONObject.parseObject(criminalMsg.get(0).toString()).getString("field"));
            HSSFCell cell7 = row.createCell(5);
            cell7.setCellValue(JSONObject.parseObject(criminalMsg.get(0).toString()).getString("field"));
            HSSFCell cell8 = row.createCell(6);
            cell8.setCellValue(JSONObject.parseObject(criminalMsg.get(0).toString()).getString("field"));
            HSSFCell cell9 = row.createCell(7);
            cell9.setCellValue(JSONObject.parseObject(criminalMsg.get(0).toString()).getString("field"));
            count++;
        }

        File file = new File("d:\\poitest\\XXX.xls");
        FileOutputStream fileOutputStream = new FileOutputStream(file);
        workBook.write(fileOutputStream);
        fileOutputStream.close();
    }
}
