package com.roytuts.java.apache.poi.excel.write.multiple.sheets.generic.way;

public class User {

	private String userId;
	private String userPass;

	public User(String userId, String userPass) {
		this.userId = userId;
		this.userPass = userPass;
	}

	public String getUserId() {
		return userId;
	}

	public void setUserId(String userId) {
		this.userId = userId;
	}

	public String getUserPass() {
		return userPass;
	}

	public void setUserPass(String userPass) {
		this.userPass = userPass;
	}

}
