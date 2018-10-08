package com.smlf.excel.test.model;

import com.smlf.excel.annotation.DataField;
import com.smlf.excel.annotation.Dataset;

@Dataset(sheetName="学生")
public class Student {
	
	@DataField(columnIndex = 0)
	private String name;
	
	@DataField(columnIndex = 1)
	private String sex;
	
	@DataField(columnIndex = 2)
	private int age;
	
	@DataField(columnIndex = 3)
	private String address;
	
	@DataField(columnIndex = 4)
	private boolean isCadres;

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getSex() {
		return sex;
	}

	public void setSex(String sex) {
		this.sex = sex;
	}

	public int getAge() {
		return age;
	}

	public void setAge(int age) {
		this.age = age;
	}

	public String getAddress() {
		return address;
	}

	public void setAddress(String address) {
		this.address = address;
	}

	public boolean isCadres() {
		return isCadres;
	}

	public void setCadres(boolean isCadres) {
		this.isCadres = isCadres;
	}

}
