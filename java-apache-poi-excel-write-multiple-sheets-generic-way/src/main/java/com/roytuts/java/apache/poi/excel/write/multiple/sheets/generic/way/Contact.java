package com.roytuts.java.apache.poi.excel.write.multiple.sheets.generic.way;

public class Contact {

	private String mobile;
	private String phone1;
	private String phone2;

	public Contact(String mobile, String phone1, String phone2) {
		this.mobile = mobile;
		this.phone1 = phone1;
		this.phone2 = phone2;
	}

	public String getMobile() {
		return mobile;
	}

	public void setMobile(String mobile) {
		this.mobile = mobile;
	}

	public String getPhone1() {
		return phone1;
	}

	public void setPhone1(String phone1) {
		this.phone1 = phone1;
	}

	public String getPhone2() {
		return phone2;
	}

	public void setPhone2(String phone2) {
		this.phone2 = phone2;
	}

}
