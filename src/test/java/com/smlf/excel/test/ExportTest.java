package com.smlf.excel.test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import com.smlf.excel.ExportUtil;
import com.smlf.excel.test.model.Student;

public class ExportTest {
	
	public static final int[] ageArray = new int[]{11,12,13,14,15,16,17,18,19,20};
	
	@Test
	public void testwith_one_list() throws IOException{
		List<Student> list = new ArrayList<Student>();
		for(int i = 0;i<1000; i++){
			Student student = new Student();
			student.setName("student"+i);
			student.setAge(ageArray[i%10]);
			student.setCadres(i%10==0);
			student.setAddress("address"+i);
			student.setSex(i%10==0?"man":"woman");
			list.add(student);
		}
		
		InputStream in= this.getClass().getResourceAsStream("template.xlsx");
		Workbook workbook = new XSSFWorkbook(in);
		ExportUtil.export(workbook, 0, 1, list);
		FileOutputStream fileOutputStream = new FileOutputStream("D://test.xlsx");
        workbook.write(fileOutputStream);

        // flush
        fileOutputStream.flush();
	}

}
