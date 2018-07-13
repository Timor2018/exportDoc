package com.test;

import java.awt.Color;
import java.io.FileOutputStream;
import java.io.IOException;

import com.lowagie.text.Cell;
import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.Element;
import com.lowagie.text.Font;
import com.lowagie.text.PageSize;
import com.lowagie.text.Paragraph;
import com.lowagie.text.Table;
import com.lowagie.text.pdf.BaseFont;
import com.lowagie.text.rtf.RtfWriter2;

/**
 * 使用Itext 向word中插入数据表
 * 
 * @author K
 *
 */
public class CreateWordDemo {

	public void createDocContext(String file, String contextString) throws DocumentException, IOException {

		// 设置纸张大小

		Document document = new Document(PageSize.A4);

		// 建立一个书写器，与document对象关联

		RtfWriter2.getInstance(document, new FileOutputStream(file));

		document.open();

		// 设置中文字体

		// BaseFont bfChinese = BaseFont.createFont("STSongStd-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);
		BaseFont bfChinese = BaseFont.createFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);
		// 标题字体风格

		Font titleFont = new Font(bfChinese, 12, Font.BOLD);

		// 正文字体风格

		Font contextFont = new Font(bfChinese, 10, Font.NORMAL);

		Paragraph title = new Paragraph("标题");

		// 设置标题格式对齐方式

		title.setAlignment(Element.ALIGN_CENTER);

		title.setFont(titleFont);

		document.add(title);

		Paragraph context = new Paragraph(contextString);

		context.setAlignment(Element.ALIGN_LEFT);

		context.setFont(contextFont);

		// 段间距

		context.setSpacingBefore(3);

		// 设置第一行空的列数

		context.setFirstLineIndent(20);

		document.add(context);

		// 设置Table表格,创建一个6列的表格
		Table table = new Table(6);
		/*int width[] = { 25, 25, 25, 25, 25, 25 };// 设置每列宽度比例
		table.setWidths(width);*/
		table.setWidth(100);// 占页面宽度比例
		table.setAlignment(Element.ALIGN_CENTER);// 居中
		table.setAlignment(Element.ALIGN_MIDDLE);// 垂直居中
		table.setAutoFillEmptyCells(true);// 自动填满
		table.setBorderWidth(1);// 边框宽度

		// 设置表头
		Font haderFontChinese = new Font(bfChinese, 12, Font.NORMAL, Color.BLACK);
		Cell haderCell1 = new Cell(new Paragraph("编号", haderFontChinese));
		Cell haderCell2 = new Cell(new Paragraph("日期", haderFontChinese));
		Cell haderCell3 = new Cell(new Paragraph("VV1", haderFontChinese));
		Cell haderCell4 = new Cell(new Paragraph("VV2", haderFontChinese));
		Cell haderCell5 = new Cell(new Paragraph("VVA", haderFontChinese));
		Cell haderCell6 = new Cell(new Paragraph("DA", haderFontChinese));

		haderCell1.setHeader(true);
		haderCell2.setHeader(true);
		haderCell3.setHeader(true);
		haderCell4.setHeader(true);
		haderCell5.setHeader(true);
		haderCell6.setHeader(true);

		table.addCell(haderCell1);
		table.addCell(haderCell2);
		table.addCell(haderCell3);
		table.addCell(haderCell4);
		table.addCell(haderCell5);
		table.addCell(haderCell6);
		
		table.endHeaders();

		table.addCell(new Cell("#1"));
		table.addCell(new Cell("#2"));
		table.addCell(new Cell("#3"));
		table.addCell(new Cell("#4"));
		table.addCell(new Cell("#5"));
		table.addCell(new Cell("#6"));
		table.addCell(new Cell("#1"));
		table.addCell(new Cell("#2"));
		table.addCell(new Cell("#3"));
		table.addCell(new Cell("#4"));
		table.addCell(new Cell("#5"));
		table.addCell(new Cell("#6"));

		document.add(table);

		document.close();

	}

	public static void main(String[] args) {

		CreateWordDemo word = new CreateWordDemo();

		String file = "E:\\test\\wordTest.doc";

		try {

			word.createDocContext(file, "测试iText导出Word文档");

		} catch (DocumentException e) {

			e.printStackTrace();

		} catch (IOException e) {

			e.printStackTrace();

		}

	}

}
