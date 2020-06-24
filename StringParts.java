public class StringParts {

	public static void main(String[] args) {

		firstInputMethod();
		secondInputMethod();
	}

	private static void firstInputMethod() {
		String str = "372601048008004014000000";
		str = str.substring(6, str.length());

		String part1 = str.substring(0, 3);
		System.out.println("part1=" + part1);

		String part2 = str.substring(3, 6);
		System.out.println("part2=" + part2);

		String part3 = str.substring(6, 10);
		System.out.println("part3=" + part3);

		String part4 = str.substring(10, 13);
		System.out.println("part4=" + part4);
	}

	private static void secondInputMethod() {
		String str = "2.16-54-98";
		String[] splitStr = str.split("\\.");

		String part1 = "";
		for (int i = 1; i <= 3 - splitStr[0].length(); i++) {
			part1 = part1 + "0";
		}
		part1 = part1 + splitStr[0];
		System.out.println("part1=" + part1);

		splitStr = splitStr[1].split("\\-");
		String part2 = "";
		for (int i = 1; i <= 3 - splitStr[0].length(); i++) {
			part2 = part2 + "0";
		}
		part2 = part2 + splitStr[0];
		System.out.println("part2=" + part2);

		String part3 = "";
		for (int i = 1; i <= 4 - splitStr[1].length(); i++) {
			part3 = part3 + "0";
		}
		part3 = part3 + splitStr[1];
		System.out.println("part3=" + part3);

		String part4 = "";
		for (int i = 1; i <= 3 - splitStr[2].length(); i++) {
			part4 = part4 + "0";
		}
		part4 = part4 + splitStr[2];
		System.out.println("part4=" + part4);
	}
}
