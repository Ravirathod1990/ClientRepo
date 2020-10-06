
public class EmailSplitter {

	public static void main(String[] args) {
		String emailStr = "abc@test.com,xmz@test.com,pnq@test.com,,'', ,  ,''";

		String[] emailArr = emailStr.split("[,'' \t\n\r]+");
		System.out.println("Total Emails ==>> " + emailArr.length);
		for (String email : emailArr) {
			System.out.println(email);
		}
	}
}
