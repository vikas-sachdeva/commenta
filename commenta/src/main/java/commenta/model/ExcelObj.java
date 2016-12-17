package commenta.model;

public class ExcelObj {

	private String filePath;

	private ExcelObj() {

	}

	private static class ClassLazyHolder {
		private static final ExcelObj INSTANCE = new ExcelObj();
	}

	public static ExcelObj getInstance() {
		return ClassLazyHolder.INSTANCE;
	}

	public String getFilePath() {
		return filePath;
	}

	public void setFilePath(String filePath) {
		this.filePath = filePath;
	}

}
