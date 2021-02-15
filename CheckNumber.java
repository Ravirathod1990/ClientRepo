package com.client.program;

public class CheckNumber {

	public static void main(String[] args) {
		String[] num = { "12345678"};
		System.out.println(checkNumber(num));
		
		String[] num2 = { "123456780"};
		System.out.println(checkNumber(num2));
		
		String[] num3 = { "12345678", "123456789", "123456789", "123456785", "12345678" };
		System.out.println(checkNumber(num3));
		
		String[] num4 = { "12345678", "123456789", "123456780", "123456785", "12345678" };
		System.out.println(checkNumber(num4));
	}

	public static boolean checkNumber(String[] num) {
		boolean flag = true;
		for (String str : num) {
			if (str.length() == 9 && str.endsWith("0")) {
				flag = false;
				break;
			}
		}
		return flag;
	}
}
