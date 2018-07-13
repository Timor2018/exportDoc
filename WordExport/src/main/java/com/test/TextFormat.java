package com.test;

import java.awt.Color;

import com.lowagie.text.Cell;
import com.lowagie.text.Element;
import com.lowagie.text.Font;
import com.lowagie.text.Image;
import com.lowagie.text.Paragraph;
import com.lowagie.text.Table;
import com.lowagie.text.pdf.BaseFont;


/**
 *文、图（多图）、表的创建方法工具类
 * 
 * @author K
 *
 */
public class TextFormat {

	/**
	 * 正文一級標題 格式化
	 * 
	 * @param content
	 *            標題
	 * @return 格式化后的paragraph
	 * @throws Exception
	 *             拋出異常
	 */
	public static Paragraph formatTitle(String content) throws Exception {
		// 设置中文字体
		BaseFont bfChinese = BaseFont.createFont(BaseFont.HELVETICA_BOLD, BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);
		Font titleFont = new Font(bfChinese, 16, Font.BOLD);
		Paragraph paragraph = new Paragraph(content);
		// 一級标题格式左对齐
		paragraph.setAlignment(Element.ALIGN_LEFT);
		paragraph.setFont(titleFont);
		// 离上一段落（标题）空的行数
		paragraph.setSpacingBefore(15);

		return paragraph;
	}

	/**
	 * 正文内容格式化
	 * 
	 * @param content
	 * @return
	 * @throws Exception
	 */
	public static Paragraph formatContent(String content) throws Exception {
		// 设置中文字体
		BaseFont bfChinese = BaseFont.createFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);
		// 正文字体风格
		Font contextFont = new Font(bfChinese, 12, Font.NORMAL);
		Paragraph paragraph = new Paragraph(content);
		// 一級标题格式左对齐
		// 正文格式左对齐
		paragraph.setAlignment(Element.ALIGN_LEFT);
		paragraph.setFont(contextFont);
		// 离上一段落空的行数
		paragraph.setSpacingBefore(10);

		return paragraph;
	}

	/**
	 * 插入图片的格式化
	 * 
	 * @param imgPath
	 * @return
	 * @throws Exception
	 */
	public static Image formatImage(String imgPath) throws Exception {
		Image img = Image.getInstance(imgPath);
		img.setAbsolutePosition(0, 0);
		img.setAlignment(Image.ALIGN_CENTER);// 设置图片显示位置
		// img.scalePercent(30);// 表示显示的大小为原尺寸的50%
		img.scaleAbsolute(250, 250);// 直接设定显示尺寸
		// img.scalePercent(25, 12);// 图像高宽的显示比例
		// img.setRotation(30);//图像旋转一定角度
		img.setSpacingBefore(5);
		return img;
	}

	/**
	 * 创建一个6列的表格 表头为：编号 日期 VV1 VV2 VVA DA
	 * 
	 * @return
	 * @throws Exception
	 */
	public static Table creatTable() throws Exception {
		// 设置中文字体
		BaseFont bfChinese = BaseFont.createFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);
		// 设置Table表格,创建一个6列的表格
		Table table = new Table(6);
		table.setWidth(100);// 占页面宽度比例
		table.setAlignment(Element.ALIGN_CENTER);// 居中
		table.setAlignment(Element.ALIGN_MIDDLE);// 垂直居中
		table.setAutoFillEmptyCells(true);// 自动填满
		table.setBorderWidth(1);// 边框宽度

		// 设置表头
		Font haderFontChinese = new Font(bfChinese, 12, Font.BOLD, Color.BLACK);
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

		return table;

	}

	public static Table rowCellTbale() throws Exception {
		// 设置中文字体
		//BaseFont bfChinese = BaseFont.createFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);
		// 设置Table表格,创建一个2列的表格
		Table table = new Table(2);
		table.setWidth(100);// 占页面宽度比例
		table.setAlignment(Element.ALIGN_CENTER);// 居中
		table.setAlignment(Element.ALIGN_MIDDLE);// 垂直居中
		table.setAutoFillEmptyCells(true);// 自动填满
		table.setBorderWidth(0);// 边框宽度
		
		return table;
	}
}
