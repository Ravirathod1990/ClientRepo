package com.client.program;

import java.util.Arrays;

public class RenameFile2 {

	public static void main(String[] args) {

		String fileName1 = "123456786+123456781+123456789+123456785";
		String renamedFileName1 = renameFile(fileName1);
		System.out.println("Renamed File Name ===>>> " + renamedFileName1);

		String fileName2 = "00123127";
		String renamedFileName2 = renameFile(fileName2);
		System.out.println("Renamed File Name ===>>> " + renamedFileName2);

		String fileName3 = "998855663+998855666+998855669+998855668+998855667";
		String renamedFileName3 = renameFile(fileName3);
		System.out.println("Renamed File Name ===>>> " + renamedFileName3);
	}

	private static String renameFile(String fileName) {
		String[] numbersArr = fileName.split("\\+");
		Arrays.sort(numbersArr);

		StringBuilder newStr = new StringBuilder();
		for (int i = 0; i < numbersArr.length; i++) {
			newStr.append(numbersArr[i]);
			if (i != numbersArr.length - 1) {
				newStr.append("+");
			}
		}
		return newStr.toString();
	}
}
