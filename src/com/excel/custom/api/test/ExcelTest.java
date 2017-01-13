package com.excel.custom.api.test;
import java.io.IOException;

import org.junit.Test;

import com.excel.generator.service.ExcelService;


public class ExcelTest {

	@Test
	public void test() throws IOException {
		ExcelService service = new ExcelService();
		service.generateExcel();
	}
}
