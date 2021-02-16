package com.client.program;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class CheckNumber {

	private static final Pattern p = Pattern.compile("(?<!\\d)\\d{9}(?!\\d)");

	public static void main(String[] args) {
		String subject = "TAF 100909675 any 100909340 string 100909322 string any";

		System.out.println(checkNumber(subject));
	}

	public static boolean checkNumber(String subject) {
		boolean flag = true;
		List<String> numberList = extractNumbersFromString(subject);
		if (!numberList.isEmpty()) {
			for (String number : numberList) {
				if (number.length() == 9 && number.endsWith("0")) {
					flag = false;
					break;
				}
			}
		}
		return flag;
	}

	public static List<String> extractNumbersFromString(String str) {
		List<String> numbers = new ArrayList<>();
		Matcher m = p.matcher(str);
		while (m.find()) {
			numbers.add(m.group());
		}
		return numbers;
	}
}
