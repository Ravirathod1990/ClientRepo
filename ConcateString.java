package com.client.program;

public class ConcateString {

	public static void main(String[] args) {

		System.out.println(concateString("410012345"));
		System.out.println(concateString("410012345+410012346"));
		System.out.println(concateString("410012345+410012346+410012347"));
		System.out.println(concateString("123+124"));
	}

	public static String concateString(String str) {
		String newStr = "";
		if (str.contains("+")) {
			newStr = str.substring(0, str.indexOf("+") - 1) + "0+" + str;
		} else {
			newStr = str.substring(0, str.length() - 1) + "0+" + str;
		}
		return newStr;
	}
}
