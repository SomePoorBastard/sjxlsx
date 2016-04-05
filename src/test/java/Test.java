import com.incesoft.tools.excel.xlsx.ExcelUtils;

import java.io.IOException;

public class Test {

	public static void main(String[] args) throws IOException {
		ExcelUtils.toCSV("/Users/robert/Downloads/Zynga non-FB costs report march 2016 v7.xlsx", "/tmp/test.csv", 0);
	}

}
