package com.client.program;
import java.util.Arrays;

public class RenameFile {

	public static void main(String[] args) {

		String fileName1 = "123456786+123456781+123456789+123456785_TAF_201897.txt";
		String renamedFileName1 = renameFile(fileName1);
		System.out.println("Renamed File Name ===>>> " + renamedFileName1);
		
		String fileName2 = "00123127+00123124__TAF_201897.csv";
		String renamedFileName2 = renameFile(fileName2);
		System.out.println("Renamed File Name ===>>> " + renamedFileName2);
		
		String fileName3 = "998855663+998855666+998855669+998855668+998855667_TAL_201967.doc";
		String renamedFileName3 = renameFile(fileName3);
		System.out.println("Renamed File Name ===>>> " + renamedFileName3);
	}

	private static String renameFile(String fileName) {
		String[] numbersArr = fileName.substring(0, fileName.indexOf("_")).split("\\+");
		Arrays.sort(numbersArr);

		StringBuilder newStr = new StringBuilder();
		for (int i = 0; i < numbersArr.length; i++) {
			newStr.append(numbersArr[i]);
			if (i != numbersArr.length - 1) {
				newStr.append("+");
			}
		}
		return newStr + fileName.substring(fileName.indexOf("_"), fileName.length());
	}
}
