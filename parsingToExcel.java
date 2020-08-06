package parsing.parsing;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

public class parsingToExcel {

	public static void main(String[] args) throws IOException {
		//creating the object of HSSFWorkbook class
		HSSFWorkbook book = new HSSFWorkbook();
		//creating a sheet
		HSSFSheet sheet = book.createSheet("parsing");
		
		//creating a list and filling it in with data
		List<Article> dataList = fileData();
		//we will need a counter
		int count = 0;
		
		//creating rows
		Row row = sheet.createRow(0);
		row.createCell(0).setCellValue("Link");
		row.createCell(1).setCellValue("Context");
		
		//filling 'artList' in the data
		for (Article artList : dataList) {
			SheetHeader(sheet, ++count, artList);
		}
		
		//writing our data to a file
		try (FileOutputStream out = new FileOutputStream(new File("C:\\Users\\one-w\\Desktop\\Java\\parsing.xls"))){//it's my path to the file
			book.write(out);
		}
		catch (Exception e) {
			e.printStackTrace();
		}
		
		System.out.println("Your excel-book was creating.");
	}
	

//creating a method that will parse the site
	public static List<Article> fileData() throws IOException {
		List<Article> modelsData = new ArrayList<>();
		
		Document doc = Jsoup.connect("https://4pda.ru/").get();//you can take another site
		Elements h2Elements = doc.getElementsByAttributeValue("class", "list-post-title");
		h2Elements.forEach(h2Element -> {//I was being parsed '<h2>' and '<a>' tags
			Element aElement = h2Element.child(0);
			String url = aElement.attr("href");
			String title =  aElement.child(0).text();
			
			modelsData.add(new Article(url, title));
		});
		
		return modelsData;
	}
	
//creating a method where the rows('rowNum') will be filled in on an excel workbook sheet
	public static void SheetHeader(HSSFSheet sheet, int rowNum, Article article) {
		Row row = sheet.createRow(rowNum);
		
		row.createCell(0).setCellValue(article.getUrl());
		row.createCell(1).setCellValue(article.getName());
	}
}

//creating a class that will be a model of the data that we will write to the file
class Article {
	public String getUrl() {
		return url;
	}

	public void setUrl(String url) {
		this.url = url;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	private String url;
	private String name;
	
	Article(String url, String name) {
		this.url = url;
		this.name = name;
	}
}
