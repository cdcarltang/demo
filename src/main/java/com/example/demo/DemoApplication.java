package com.example.demo;

import java.io.FileOutputStream;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.core.RowMapper;

@SpringBootApplication
public class DemoApplication implements CommandLineRunner {

	@Autowired
	JdbcTemplate jdbcTemplate;

	public static void main(String[] args) {
		SpringApplication.run(DemoApplication.class, args);
	}

	@Override
	public void run(String... args) throws Exception {

		// 创建excel文件
		HSSFWorkbook wb = new HSSFWorkbook();

		String[] vins = { "xxx", "yyyy" };

		for (String vin : vins) {
			List<DomainModel> rows = jdbcTemplate.query("sql", args, new RowMapper<DomainModel>() {
				public DomainModel mapRow(ResultSet rs, int rowNum) throws SQLException {
					// 映射SQL result到 java对象.
					return null;
				}

			});

			HSSFSheet sheet = wb.createSheet(vin);

			int rowIndex = 1;
			for (DomainModel data : rows) {
				int columnIndex = 0;
				HSSFRow excelRow = sheet.createRow(rowIndex++);
				HSSFCell cell = excelRow.createCell(columnIndex++);
				cell.setCellValue(data.getVin()); // 映射属性。
			}
		}

		FileOutputStream o = new FileOutputStream("D:/xxxxxx.xlsx");
		wb.write(o);
	}
}
