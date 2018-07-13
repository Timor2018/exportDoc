package com.test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.util.ArrayList;
import java.util.List;

import com.lowagie.text.BadElementException;
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

public class WordTest {

	public static void main(String[] args) {
		exportDoc();
	}

	/**
	 *
	 * @Description: 将网页内容导出为word @param @param file @param @throws
	 *               DocumentException @param @throws IOException 设定文件 @return void
	 *               返回类型 @throws
	 */
	public static String exportDoc() {

		try {
			// 设置纸张大小
			Document document = new Document(PageSize.A4);

			// 建立一个书写器(Writer)与document对象关联，通过书写器(Writer)可以将文档写入到磁盘中
			// ByteArrayOutputStream baos = new ByteArrayOutputStream();
			// C:\\Users\\orion\\Desktop\\home.jpg
			File file = new File("E:\\test\\WordTest.doc");
			RtfWriter2.getInstance(document, new FileOutputStream(file));
			document.open();

			// 设置中文字体
			BaseFont bfChinese = BaseFont.createFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);
			BaseFont titleChinese = BaseFont.createFont(BaseFont.HELVETICA_BOLD, BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);

			// 标题字体风格
			Font title1Font = new Font(titleChinese, 24, Font.BOLD);
			Paragraph title1 = new Paragraph("广东南天司法鉴定所");
			Font title2Font = new Font(titleChinese, 16, Font.BOLD);
			Paragraph title2 = new Paragraph("色阶法检验结果分析报告");

			// 设置标题格式对齐方式
			title1.setAlignment(Element.ALIGN_CENTER);
			title2.setAlignment(Element.ALIGN_CENTER);
			title1.setFont(title1Font);
			title2.setFont(title2Font);
			document.add(title1);
			title2.setSpacingBefore(20);
			document.add(title2);

			// 正文字体风格
			Font contextFont = new Font(bfChinese, 12, Font.NORMAL);
			// Font contextFont = new Font(bfChinese, 11, Font.NORMAL);
			List<String> list = new ArrayList<>();
			list.add("a");
			
		    
			for (String string : list) {
				// code
				String code = "粤南[2018]色阶第07120001号";
				Paragraph codeStyle = new Paragraph(code);
				// 正文格式右对齐
				codeStyle.setAlignment(Element.ALIGN_RIGHT);
				codeStyle.setFont(contextFont);
				// 离上一段落（标题）空的行数
				codeStyle.setSpacingBefore(20);
				
				// 设置第一行空的列数
				// context.setFirstLineIndent(0);
				document.add(codeStyle);
				
				//添加横线———

			    Paragraph p_line = new Paragraph("———————————————————————————————————————————");
			    p_line.setAlignment(Element.ALIGN_RIGHT);
			    document.add(p_line);
				
				// 一标题
				String titleContent = "一、基本情况";
				document.add(TextFormat.formatTitle(titleContent));

				// 正文
				// agent
				String agent = "经 办 人：张三" ;
				document.add(TextFormat.formatContent(agent));

				// caseID
				String caseID = "案件序号：2018年07120001号";
				document.add(TextFormat.formatContent(caseID));

				// scandate
				String scandate = "扫描日期：2018-07-07";
				document.add(TextFormat.formatContent(scandate));

				// require
				String require = "检验要求：文件形成时间";
				document.add(TextFormat.formatContent(require));

				// creatdate
				String creatdate = "创建时间：2018-07-12 11:46";
				document.add(TextFormat.formatContent(creatdate));

				// 二、扫描材料 一标题
				String title2Content = "二、扫描材料";
				document.add(TextFormat.formatTitle(title2Content));

				// 添加图片 Image.getInstance即可以放路径又可以放二进制字节流
				String imgPath1 = "E:\\test\\img1-1.jpg";
				String imgPath2 = "E:\\test\\img2-1.jpg";
				document.add(TextFormat.formatImage(imgPath1));
				document.add(TextFormat.formatImage(imgPath2));

				// 三、检验数据 -标题
				String title3Content = "三、检验数据";
				document.add(TextFormat.formatTitle(title3Content));
				
				//表格
				Table table = TextFormat.creatTable();
				
				for (int i = 0; i < 4; i++) {
					for (int j = 0; j < 6; j++) {
						table.addCell(new Cell(i+j+""));
					}
				}
				document.add(table);
				
				
				// 色阶图表 -标题
				String title4Content = "四、色阶图表";
				document.add(TextFormat.formatTitle(title4Content));

				// 五、分析结果
				String title5Content = "五、分析结果";
				document.add(TextFormat.formatTitle(title5Content));

				// result
				String result = "  经鉴定，该文件形成与1996年7月12日";
				document.add(TextFormat.formatContent(result));
			}
			document.close();
			// 得到输入流
			// wordFile = new ByteArrayInputStream(baos.toByteArray());
			// baos.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (BadElementException e) {
			e.printStackTrace();
		} catch (MalformedURLException e) {
			e.printStackTrace();
		} catch (DocumentException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return "";

	}

}