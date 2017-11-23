package myProject.excelUtil.pojo;

public enum ExcelEnum {

	TRUE("是"), FALSE("否");

	public  String name;

	 ExcelEnum(String name) {
		this.name = name;
	}

	@Override
	public String toString() {

		return this.name;
	}

	public static void main(String[] args) {
		System.out.println(ExcelEnum.FALSE.toString());
	}
}
