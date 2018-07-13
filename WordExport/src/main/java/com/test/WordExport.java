package com.test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.util.ArrayList;
import java.util.List;

import com.lowagie.text.BadElementException;
import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.Element;
import com.lowagie.text.Font;
import com.lowagie.text.Image;
import com.lowagie.text.PageSize;
import com.lowagie.text.Paragraph;
import com.lowagie.text.Phrase;
import com.lowagie.text.Table;
import com.lowagie.text.pdf.BaseFont;
import com.lowagie.text.rtf.RtfWriter2;

public class WordExport {

	public static void main(String[] args) {
		exportDoc();
	}

	/**
	 *
	 * @Description: 将网页内容导出为word 
	 * @param @param file 
	 * @param @throws DocumentException 
	 * @param @throws IOException 设定文件
	 * @return void 返回类型
	 *  @throws
	 */
	public static String exportDoc() {

		try {
			// 设置纸张大小
			Document document = new Document(PageSize.A4);

			// 建立一个书写器(Writer)与document对象关联，通过书写器(Writer)可以将文档写入到磁盘中
			// ByteArrayOutputStream baos = new ByteArrayOutputStream();
			// C:\\Users\\orion\\Desktop\\home.jpg
			File file = new File("E:\\test\\WordExport.doc");
			RtfWriter2.getInstance(document, new FileOutputStream(file));
			document.open();

			// 设置中文字体
			BaseFont bfChinese = BaseFont.createFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);

			// 标题字体风格
			Font title1Font = new Font(bfChinese, 24, Font.BOLD);
			Paragraph title1 = new Paragraph("广东南天司法鉴定所");
			Font title2Font = new Font(bfChinese, 16, Font.BOLD);
			Paragraph title2 = new Paragraph("色阶法检验结果分析报告");

			// 设置标题格式对齐方式
			title1.setAlignment(Element.ALIGN_CENTER);
			title2.setAlignment(Element.ALIGN_CENTER);
			title1.setFont(title1Font);
			title2.setFont(title2Font);
			document.add(title1);
			document.add(title2);
			
			//正文一级标题
			Font title3Font = new Font(bfChinese, 16, Font.BOLD);
			
			
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

				// 一标题
				String titleContent = "一、基本情况";
				Paragraph titleContentStyle = new Paragraph(titleContent);
				// 一級标题格式左对齐
				titleContentStyle.setAlignment(Element.ALIGN_LEFT);
				titleContentStyle.setFont(title3Font);
				// 离上一段落（标题）空的行数
				titleContentStyle.setSpacingBefore(5);
				document.add(titleContentStyle);
				
				//正文
				// agent
				String agent = "经 办 人：张三";
				Paragraph agentStyle = new Paragraph(agent);
				// 正文格式左对齐
				agentStyle.setAlignment(Element.ALIGN_LEFT);
				agentStyle.setFont(contextFont);
				// 离上一段落空的行数
				agentStyle.setSpacingBefore(10);
				// 设置第一行空的列数
				// context.setFirstLineIndent(0);
				document.add(agentStyle);
				
				// caseID
				String caseID = "案件序号：2018年07120001号";
				Paragraph caseIdStyle = new Paragraph(caseID);
				// 正文格式左对齐
				caseIdStyle.setAlignment(Element.ALIGN_LEFT);
				caseIdStyle.setFont(contextFont);
				// 离上一段落空的行数
				caseIdStyle.setSpacingBefore(10);
				// 设置第一行空的列数
				// context.setFirstLineIndent(0);
				document.add(caseIdStyle);
				
				// scandate
				String scandate = "扫描日期：2018-07-07";
				Paragraph scandateStyle = new Paragraph(scandate);
				// 正文格式左对齐
				scandateStyle.setAlignment(Element.ALIGN_LEFT);
				scandateStyle.setFont(contextFont);
				// 离上一段落空的行数
				scandateStyle.setSpacingBefore(10);
				// 设置第一行空的列数
				// context.setFirstLineIndent(0);
				document.add(scandateStyle);
				
				// require
				String require = "检验要求：文件形成时间";
				Paragraph requireStyle = new Paragraph(require);
				// 正文格式左对齐
				requireStyle.setAlignment(Element.ALIGN_LEFT);
				requireStyle.setFont(contextFont);
				// 离上一段落空的行数
				requireStyle.setSpacingBefore(10);
				// 设置第一行空的列数
				// context.setFirstLineIndent(0);
				document.add(requireStyle);
				
				// creatdate
				String creatdate = "经 办 人：张三";
				Paragraph creatdateStyle = new Paragraph(creatdate);
				// 正文格式左对齐
				creatdateStyle.setAlignment(Element.ALIGN_LEFT);
				creatdateStyle.setFont(contextFont);
				// 离上一段落空的行数
				creatdateStyle.setSpacingBefore(10);
				// 设置第一行空的列数
				// context.setFirstLineIndent(0);
				document.add(creatdateStyle);
				
				
				Table table = TextFormat.rowCellTbale();
				table.addCell(new Phrase("经办人：", title3Font));
				table.addCell(new Phrase("李四", contextFont));
				table.addCell(new Phrase("案件序号：", title3Font));
				table.addCell(new Phrase("2018年第07120001号", contextFont));
				table.addCell(new Phrase("扫描日期：", title3Font));
				table.addCell(new Phrase("2018-07-10", contextFont));
				table.addCell(new Phrase("检验要求：", title3Font));
				table.addCell(new Phrase("文件形成时间", contextFont));
				table.addCell(new Phrase("制作日期：", title3Font));
				table.addCell(new Phrase("2018-07-12 15:59:36", contextFont));
				
				document.add(table);
				
				//二、扫描材料 一标题
				String title2Content = "二、扫描材料";
				Paragraph title2ContentStyle = new Paragraph(title2Content);
				// 一級标题格式左对齐
				title2ContentStyle.setAlignment(Element.ALIGN_LEFT);
				title2ContentStyle.setFont(title3Font);
				// 离上一段落（标题）空的行数
				title2ContentStyle.setSpacingBefore(5);
				document.add(title2ContentStyle);
				
				// 添加图片 Image.getInstance即可以放路径又可以放二进制字节流
				String imgPath1 = "E:\\test\\img1-1.jpg";
				String imgPath2 = "E:\\test\\img2-1.jpg";
				Image img1 = Image.getInstance(imgPath1);
				Image img2 = Image.getInstance(imgPath2);
				img1.setAbsolutePosition(0, 0);
				img1.setAlignment(Image.ALIGN_CENTER);// 设置图片显示位置
				img1.scalePercent(30);// 表示显示的大小为原尺寸的50%
				img2.setAbsolutePosition(0, 0);
				img2.setAlignment(Image.ALIGN_CENTER);// 设置图片显示位置
				img2.scalePercent(30);// 表示显示的大小为原尺寸的50%
				// img.scaleAbsolute(60, 60);// 直接设定显示尺寸
				// img.scalePercent(25, 12);//图像高宽的显示比例
				// img.setRotation(30);//图像旋转一定角度
				document.add(img1);
				document.add(img2);
				
				//三、检验数据 -标题
				String title3Content = "二、扫描材料";
				Paragraph title3ContentStyle = new Paragraph(title3Content);
				// 一級标题格式左对齐
				title3ContentStyle.setAlignment(Element.ALIGN_LEFT);
				title3ContentStyle.setFont(title3Font);
				// 离上一段落（标题）空的行数
				title3ContentStyle.setSpacingBefore(5);
				document.add(title3ContentStyle);
				
				// 添加图片 Image.getInstance即可以放路径又可以放二进制字节流
				String imgPath3 = "E:\\test\\img1-1.jpg";
				String imgPath4 = "E:\\test\\img2-1.jpg";
				Image img3 = Image.getInstance(imgPath3);
				Image img4 = Image.getInstance(imgPath4);
				img3.setAbsolutePosition(0, 0);
				img3.setAlignment(Image.ALIGN_CENTER);// 设置图片显示位置
				img3.scalePercent(30);// 表示显示的大小为原尺寸的50%
				img4.setAbsolutePosition(0, 0);
				img4.setAlignment(Image.ALIGN_CENTER);// 设置图片显示位置
				img4.scalePercent(30);// 表示显示的大小为原尺寸的50%
				// img.scaleAbsolute(60, 60);// 直接设定显示尺寸
				// img.scalePercent(25, 12);//图像高宽的显示比例
				// img.setRotation(30);//图像旋转一定角度
				document.add(img3);
				document.add(img4);
				
				//色阶图表 -标题
				String title4Content = "四、色阶图表";
				Paragraph title4ContentStyle = new Paragraph(title4Content);
				// 一級标题格式左对齐
				title4ContentStyle.setAlignment(Element.ALIGN_LEFT);
				title4ContentStyle.setFont(title3Font);
				// 离上一段落（标题）空的行数
				title4ContentStyle.setSpacingBefore(5);
				document.add(title4ContentStyle);
				
				
				//五、分析结果
				String title5Content = "四、色阶图表";
				Paragraph title5ContentStyle = new Paragraph(title5Content);
				// 一級标题格式左对齐
				title5ContentStyle.setAlignment(Element.ALIGN_LEFT);
				title5ContentStyle.setFont(title3Font);
				// 离上一段落（标题）空的行数
				title5ContentStyle.setSpacingBefore(5);
				document.add(title5ContentStyle);
				
				//result
				String result = "  经鉴定，该文件形成与1996年7月12日";
				Paragraph resultStyle = new Paragraph(result);
				// 正文格式左对齐
				resultStyle.setAlignment(Element.ALIGN_LEFT);
				resultStyle.setFont(contextFont);
				// 离上一段落空的行数
				resultStyle.setSpacingBefore(10);
				// 设置第一行空的列数
				// context.setFirstLineIndent(0);
				document.add(resultStyle);
				
				
				
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