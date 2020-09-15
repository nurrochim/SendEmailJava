package com.nurochim;

import java.util.List;
import java.util.Map;

import com.nurochim.poi.excel.ExcelPOIHelper;

public class ExcelStarter {
	public static void main(String[] args) {
		try {
			ExcelPOIHelper helper = new ExcelPOIHelper();
			Map<Integer, List<String>> mapRows = helper.readExcel("D:\\Kemenristek\\Webinar Brazil\\e-sertifikat\\Penerima e-Sertifikat Webinar v3.xlsx");
			for(Integer key : mapRows.keySet()) {
				System.out.println(mapRows.get(key));
			}
			
			System.out.println(mapRows.size());
			System.out.println("Test");
		}catch (Exception e) {
			e.printStackTrace();
		}
	}
}
