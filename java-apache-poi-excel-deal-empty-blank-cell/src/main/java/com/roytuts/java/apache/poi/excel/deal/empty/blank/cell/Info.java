package com.roytuts.java.apache.poi.excel.deal.empty.blank.cell;

public class Info {

	private String name;
	private String mobile;
	private String phone;
	private String permAddress;
	private String commAddress;

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getMobile() {
		return mobile;
	}

	public void setMobile(String mobile) {
		this.mobile = mobile;
	}

	public String getPhone() {
		return phone;
	}

	public void setPhone(String phone) {
		this.phone = phone;
	}

	public String getPermAddress() {
		return permAddress;
	}

	public void setPermAddress(String permAddress) {
		this.permAddress = permAddress;
	}

	public String getCommAddress() {
		return commAddress;
	}

	public void setCommAddress(String commAddress) {
		this.commAddress = commAddress;
	}

	@Override
	public String toString() {
		return "Info [name=" + name + ", mobile=" + mobile + ", phone=" + phone + ", permAddress=" + permAddress
				+ ", commAddress=" + commAddress + "]";
	}

}
