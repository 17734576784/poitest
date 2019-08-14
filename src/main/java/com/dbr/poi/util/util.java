package com.dbr.poi.util;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class util {

	public static void main(String[] args) {
		test();
	}
	public static void test() {
//		String[][] propertyDes = {{"姓名","所属部门","职位","入职日期","转正日期","年假","","带薪病假","","调休剩余可休天数","本月加班天数","本月已休","","","","","",""},
//				{"","","","","","基数","剩余可休天数","基数","剩余可休天数","","","年假","调休","带薪病假","婚假","产假","丧假","事假"},
//				};
	
		String[][]  propertyDes = {{"1","","2","3","","","","4","","","","","","",""},
					{"1","1","","3","","3","","4","","","","4","","",""},
					{"","","","3","3","3","3","4","","4","","4","","4",""},
					{"","","","","","","","4","4","4","4","4","4","4","4"}};
		// 第一步，创建一个webbook，对应一个Excel文件
		XSSFWorkbook workbook = new XSSFWorkbook();
		// 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
		Sheet sheet = workbook.createSheet("测试");
		
		XSSFCellStyle style = workbook.createCellStyle();

		int mergerNum = 0; // 合并数
		// 给单元格设置值
		for (int i = 0; i < propertyDes.length; i++) {
			XSSFRow row = (XSSFRow) sheet.createRow(i);
			row.setHeight((short) 700);
			for (int j = 0; j < propertyDes[i].length; j++) {
				XSSFCell cell = row.createCell(j);
				cell.setCellStyle(style);
				cell.setCellValue(propertyDes[i][j]);
			}
		}
		Map<Integer, List<Integer>> map = new HashMap<Integer, List<Integer>>(); // 合并行时要跳过的行列
		// 合并行
		for (int i = 0; i < propertyDes[propertyDes.length - 1].length; i++) {
			if ("".equals(propertyDes[propertyDes.length - 1][i])) {
				for (int j = propertyDes.length - 2; j >= 0; j--) {
					System.out.println(propertyDes[j][i]);
					if (!"".equals(propertyDes[j][i])) {
						sheet.addMergedRegion(new CellRangeAddress(j, propertyDes.length - 1, i, i)); // 合并单元格
						break;
					} else {
						if (map.containsKey(j)) {
							List<Integer> list = map.get(j);
							list.add(i);
							map.put(j, list);
						} else {
							List<Integer> list = new ArrayList<Integer>();
							list.add(i);
							map.put(j, list);
						}
					}
				}
			}
		}
		// 合并列
		for (int i = 0; i < propertyDes.length - 1; i++) {
			for (int j = 0; j < propertyDes[i].length; j++) {
				List<Integer> list = map.get(i);
				if (list == null || (list != null && !list.contains(j))) {
					if ("".equals(propertyDes[i][j])) {
						mergerNum++;
						if (mergerNum != 0 && j == (propertyDes[i].length - 1)) {
							sheet.addMergedRegion(new CellRangeAddress(i, i, j - mergerNum, j)); // 合并单元格
							mergerNum = 0;
						}
					} else {
						if (mergerNum != 0) {
							sheet.addMergedRegion(new CellRangeAddress(i, i, j - mergerNum - 1, j - 1)); // 合并单元格
							mergerNum = 0;
						}
					}
				}
			}
		}
		
	    try {
            FileOutputStream output = new FileOutputStream("e:\\workbook.xls");
            workbook.write(output);
            output.flush();
        } catch (IOException e) {
            e.printStackTrace();
        }
	}

}