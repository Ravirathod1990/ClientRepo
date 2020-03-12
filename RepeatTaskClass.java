public class RepeatTaskClass {

	public static int volume = 0;
	public static int counter = 1;

	public static void main(String[] args) throws InterruptedException {

		method1();
		method2();
		method3();
		if (volume == 0 && counter == 3) {
			System.out.println("Stop execution as data not found");
			return;
		}
		method4();
		method5();
	}

	public static void method1() {
		System.out.println("In method 1");
	}

	public static void method2() {
		System.out.println("In method 2");
	}

	public static void method3() throws InterruptedException {
		System.out.println("In method 3");
		while (volume == 0 && counter < 3) {
			counter++;
			Thread.sleep(30000);
			method2();
			method3();
		}
	}

	public static void method4() {
		System.out.println("In method 4");
	}

	public static void method5() {
		System.out.println("In method 5");
	}

}
