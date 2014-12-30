using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Graph;


using CHART = System.Windows.Forms.DataVisualization.Charting.Chart;
using System.Diagnostics;
using System.Xml;
//using System.Windows.Forms.DataVisualization.Charting;
//using DT = System.Data.DataTable;
//using DR = System.Data.DataRow;

namespace word2pdf
{
	class BookmarkReplace
	{
		private Word._Application wApp = null;
		private Word._Document wDoc = null;
		private object nothing = Missing.Value;

		public void openDocument(string doc)
		{
			closeWordProcess();

			Object Template = doc; // Optional Object. The name of the template to be used for the new document. If this argument is omitted, the Normal template is used. 
			Object NewTemplate = false; // Optional Object. True to open the document as a template. The default value is False. 
			Object DocumentType = Word.WdNewDocumentType.wdNewBlankDocument; // Optional Object. Can be one of the following WdNewDocumentType constants: wdNewBlankDocument, wdNewEmailMessage, wdNewFrameset, or wdNewWebPage. The default constant is wdNewBlankDocument. 
			Object Visible = true;
			// Optional Object. True to open the document in a visible window. 
			// If this value is False, Microsoft Word opens the document but sets
			// the Visible property of the document window to False. 
			// The default value is True. 

			try
			{
				wApp = new Word.Application();
				//wApp.Visible = true;
				wDoc = wApp.Documents.Add(ref Template, ref NewTemplate,
					ref DocumentType, ref Visible);
			}
			catch (Exception ex)
			{
				this.close();
				throw ex;
			}
		}

		public void closeWordProcess()
		{
			try
			{
				Process[] ps = Process.GetProcesses();
				foreach (Process item in ps)
				{
					if (item.ProcessName == "WINWORD" || item.ProcessName == "AcroRd32")
					{
						item.Kill();
					}
				}
			}
			catch (Exception)
			{
			}
		}

		public void close()
		{
			// 指定文档的保存操作。可以是下列 WdSaveOptions 值之一：
			// wdDoNotSaveChanges、wdPromptToSaveChanges 或 wdSaveChanges。 
			// Object SaveChanges = Word.WdSaveOptions.wdSaveChanges;
			Object DoNotSaveChanges = Word.WdSaveOptions.wdDoNotSaveChanges;

			// 指定文档的保存格式。可以是下列 WdOriginalFormat 值之一：
			// wdOriginalDocumentFormat、wdPromptUser 或 wdWordDocument。 
			Object OriginalFormat = Word.WdOriginalFormat.wdOriginalDocumentFormat;

			// 如果为 true，则将文档传送给下一个收件人。
			// 如果没有为文档附加传送名单，则忽略此参数。 
			Object RouteDocument = false;


			try
			{
				if (wDoc != null)
					wDoc.Close(ref DoNotSaveChanges, ref OriginalFormat, ref RouteDocument);
				if (wApp != null)
					wApp.Quit(ref DoNotSaveChanges, ref OriginalFormat, ref RouteDocument);
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		public void replaceText(String bookmark, String text)
		{
			try
			{
				object bkObj = bookmark;
				if (wApp.ActiveDocument.Bookmarks.Exists(bookmark) == true)
				{
					wApp.ActiveDocument.Bookmarks.get_Item(ref bkObj).Select();
				}
				else return;

				if (text == "^p")
				{
					wApp.Selection.TypeParagraph();
				}
				//else if (text.IndexOf("^p") > 0)
				//{
				//    string[] v = text.Split(new string[] { "^p" },
				//        StringSplitOptions.None);
				//    for (int i = 0; i < v.Length; i++)
				//    {
				//        wApp.Selection.TypeText(v[i]);
				//        if (i < v.Length - 1)
				//        {
				//            wApp.Selection.TypeParagraph();
				//        }
				//    }
				//}
				else
				{
					wApp.Selection.TypeText(text);
				}
			}
			catch (Exception)
			{
			}

			//wApp.Selection.TypeText(text);
			//wApp.Selection.Delete();
		}

		public void replaceImage(String bookmark, String path)
		{
			try
			{
				object bkObj = bookmark;
				if (wApp.ActiveDocument.Bookmarks.Exists(bookmark) == true)
				{
					wApp.ActiveDocument.Bookmarks.get_Item(ref bkObj).Select();
				}
				else return;
				object anchor = wApp.Selection.Range;
				object LinkToFile = false;
				object SaveWithDocument = true;
				//object left = 0;
				//object top = 0; 
				//object width = 100; 
				//object height =200;

				//wDoc.InlineShapes.AddPicture(@"D:\My Projects\IDB.E-Assessment\Src\Assessment.Manage\Report\Template\MBTI22.jpg", ref LinkToFile, ref SaveWithDocument,
				//    ref left, ref top, ref width, ref height, ref anchor); 				
				wDoc.InlineShapes.AddPicture(Path.GetFullPath(path), ref LinkToFile,
					ref SaveWithDocument, ref anchor);
			}
			catch (Exception)
			{
			}
		}

		public void saveAsPdf(String dst)
		{
			if (!dst.EndsWith(".pdf")) dst += ".pdf";
			Object FileName = dst; // 文档的名称。默认值是当前文件夹名和文件名。如果文档在以前没有保存过，则使用默认名称（例如，Doc1.doc）。如果已经存在具有指定文件名的文档，则会在不先提示用户的情况下改写文档。 
			Object FileFormat = Word.WdSaveFormat.wdFormatPDF;
			// 文档的保存格式。可以是任何 WdSaveFormat 值。要以另一种格式保存文档，
			// 请为 SaveFormat 属性指定适当的值。 
			Object LockComments = false; //  如果为 true，则锁定文档以进行注释。默认值为 false。 
			Object Password = System.Type.Missing; // 用来打开文档的密码字符串。（请参见下面的备注。） 
			Object AddToRecentFiles = false; //  如果为 true，则将该文档添加到“文件”菜单上最近使用的文件列表中。默认值为 true。 
			Object WritePassword = System.Type.Missing; // 用来保存对文件所做更改的密码字符串。（请参见下面的备注。） 
			Object ReadOnlyRecommended = false; //  如果为 true，则让 Microsoft Office Word 在打开文档时建议只读状态。默认值为 false。 
			Object EmbedTrueTypeFonts = false; // 如果为 true，则将 TrueType 字体随文档一起保存。如果省略的话，则 EmbedTrueTypeFonts 参数假定 EmbedTrueTypeFonts 属性的值。 
			Object SaveNativePictureFormat = true; //  如果图形是从另一个平台（例如，Macintosh）导入的，则 true 表示仅保存导入图形的 Windows 版本。 
			Object SaveFormsData = false; //  如果为 true，则将用户在窗体中输入的数据另存为数据记录。 
			Object SaveAsAOCELetter = false; //  如果文档附加了邮件程序，则 true 表示会将文档另存为 AOCE 信函（邮件程序会进行保存）。 
			Object Encoding = System.Type.Missing; // MsoEncoding。要用于另存为编码文本文件的文档的代码页或字符集。默认值是系统代码页。 
			Object InsertLineBreaks = true; //  如果文档另存为文本文件，则 true 表示在每行文本末尾插入分行符。 
			Object AllowSubstitutions = false; // 如果文档另存为文本文件，则 true 允许 Word 将某些符号替换为外观与之类似的文本。例如，将版权符号显示为 (c)。默认值为 false。 
			Object LineEnding = Microsoft.Office.Interop.Word.WdLineEndingType.wdCRLF;// Word 在另存为文本文件的文档中标记分行符和换段符。可以是任何 WdLineEndingType 值。 
			Object AddBiDiMarks = true;//如果为 true，则向输出文件添加控制字符，以便保留原始文档中文本的双向布局。 
			try
			{
				wDoc.SaveAs(ref FileName, ref FileFormat, ref LockComments,
					ref Password, ref AddToRecentFiles, ref WritePassword,
					ref ReadOnlyRecommended,
					ref EmbedTrueTypeFonts, ref SaveNativePictureFormat,
					ref SaveFormsData,
					ref SaveAsAOCELetter, ref Encoding, ref InsertLineBreaks,
					ref AllowSubstitutions, ref LineEnding, ref AddBiDiMarks);

			}
			catch (Exception ex)
			{
				string err = string.Format(" 另存文件出错，错误原因：{0}", ex.Message);
				throw new Exception(err, ex);
			}
		}

		public void saveAsWord(String dst)
		{
			if (!dst.EndsWith(".docx")) dst += ".docx";
			Object FileName = dst; // 文档的名称。默认值是当前文件夹名和文件名。如果文档在以前没有保存过，则使用默认名称（例如，Doc1.doc）。如果已经存在具有指定文件名的文档，则会在不先提示用户的情况下改写文档。 
			Object FileFormat = Word.WdSaveFormat.wdFormatDocumentDefault;
			// 文档的保存格式。可以是任何 WdSaveFormat 值。要以另一种格式保存文档，
			// 请为 SaveFormat 属性指定适当的值。 
			Object LockComments = false; //  如果为 true，则锁定文档以进行注释。默认值为 false。 
			Object Password = System.Type.Missing; // 用来打开文档的密码字符串。（请参见下面的备注。） 
			Object AddToRecentFiles = false; //  如果为 true，则将该文档添加到“文件”菜单上最近使用的文件列表中。默认值为 true。 
			Object WritePassword = System.Type.Missing; // 用来保存对文件所做更改的密码字符串。（请参见下面的备注。） 
			Object ReadOnlyRecommended = false; //  如果为 true，则让 Microsoft Office Word 在打开文档时建议只读状态。默认值为 false。 
			Object EmbedTrueTypeFonts = false; // 如果为 true，则将 TrueType 字体随文档一起保存。如果省略的话，则 EmbedTrueTypeFonts 参数假定 EmbedTrueTypeFonts 属性的值。 
			Object SaveNativePictureFormat = true; //  如果图形是从另一个平台（例如，Macintosh）导入的，则 true 表示仅保存导入图形的 Windows 版本。 
			Object SaveFormsData = false; //  如果为 true，则将用户在窗体中输入的数据另存为数据记录。 
			Object SaveAsAOCELetter = false; //  如果文档附加了邮件程序，则 true 表示会将文档另存为 AOCE 信函（邮件程序会进行保存）。 
			Object Encoding = System.Type.Missing; // MsoEncoding。要用于另存为编码文本文件的文档的代码页或字符集。默认值是系统代码页。 
			Object InsertLineBreaks = true; //  如果文档另存为文本文件，则 true 表示在每行文本末尾插入分行符。 
			Object AllowSubstitutions = false; // 如果文档另存为文本文件，则 true 允许 Word 将某些符号替换为外观与之类似的文本。例如，将版权符号显示为 (c)。默认值为 false。 
			Object LineEnding = Microsoft.Office.Interop.Word.WdLineEndingType.wdCRLF;// Word 在另存为文本文件的文档中标记分行符和换段符。可以是任何 WdLineEndingType 值。 
			Object AddBiDiMarks = true;//如果为 true，则向输出文件添加控制字符，以便保留原始文档中文本的双向布局。 
			try
			{
				wDoc.SaveAs(ref FileName, ref FileFormat, ref LockComments,
					ref Password, ref AddToRecentFiles, ref WritePassword,
					ref ReadOnlyRecommended,
					ref EmbedTrueTypeFonts, ref SaveNativePictureFormat,
					ref SaveFormsData,
					ref SaveAsAOCELetter, ref Encoding, ref InsertLineBreaks,
					ref AllowSubstitutions, ref LineEnding, ref AddBiDiMarks);

			}
			catch (Exception ex)
			{
				string err = string.Format(" 另存文件出错，错误原因：{0}", ex.Message);
				throw new Exception(err, ex);
			}
		}

		public void addChart2(string bookmark, string chartType, string[][] data)
		{
			CHART c = ChartCreator.createChart(chartType, data);
			string img = Path.GetFullPath(DateTime.Now.Ticks + ".bmp");
			c.SaveImage(img,
				System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Bmp);
			replaceImage(bookmark, img);
			try
			{
				File.Delete(img);
			}
			catch (Exception)
			{
			}
		}


		public void addChart(string chartType, string bookmark, string[][] data)
		{
			if (wApp.ActiveDocument.Bookmarks.Exists(bookmark) != true)
			{
				return;
			}

			object oMissing = System.Reflection.Missing.Value;
			object oBookmark = bookmark;
			Word.InlineShape oShape;
			object oClassType = "MSGraph.Chart.8";
			Word.Range wrdRng = wDoc.Bookmarks.get_Item(ref oBookmark).Range;
			oShape = wrdRng.InlineShapes.AddOLEObject(ref oClassType, ref oMissing,
				ref oMissing, ref oMissing, ref oMissing,
				ref oMissing, ref oMissing, ref oMissing);
			//Demonstrate use of late bound oChart and oChartApp objects to   
			//manipulate the chart object with MSGraph.   
			object oChart;
			object oChartApp;
			oChart = oShape.OLEFormat.Object;
			oChartApp = oChart.GetType().InvokeMember("Application",
				BindingFlags.GetProperty, null, oChart, null);
			//Change the chart type
			Chart objChart = (Chart)oShape.OLEFormat.Object;
			switch (chartType)
			{
				case "line":
					objChart.ChartType = XlChartType.xlLine;
					break;
				case "radar":
					objChart.ChartType = XlChartType.xlRadar;
					break;
				case "bar":
					break;
				case "vbar":
					break;
			}
			//objChart.ChartType = XlChartType.xlColumnClustered;
			//objChart.ChartType = XlChartType.xlLine;
			//objChart.ChartArea.Border.LineStyle = XlLineStyle.xlDash;
			//objChart.ChartArea.Fill.BackColor.SchemeColor = 28;
			//Series s = (Series)objChart.SeriesCollection(1);

			objChart.Legend.Position = XlLegendPosition.xlLegendPositionBottom;



			//绑定数据   
			Microsoft.Office.Interop.Graph.DataSheet dataSheet;
			dataSheet = objChart.Application.DataSheet;

			dataSheet.Cells.ClearContents();

			int rows = data.Length;
			for (int i = 0; i < rows; i++)
			{
				int cols = data[i].Length;
				for (int j = 0; j < cols; j++)
				{
					try
					{
						dataSheet.Cells[i + 1, j + 1] = decimal.Parse(data[i][j].Trim());
					}
					catch (Exception)
					{
						dataSheet.Cells[i + 1, j + 1] = data[i][j].Trim();
					}
				}
			}

			objChart.Application.Update();
			objChart.Application.Quit();

			//设置大小
			oShape.Width = wApp.InchesToPoints(6.25f);
			oShape.Height = wApp.InchesToPoints(3.57f);
		}


		public void addTable(object bookmark, System.Data.DataTable dt)
		{
			try
			{
				int cols = dt.Columns.Count;
				int rows = dt.Rows.Count;
				object nothing = Missing.Value;

				Word.Range wrdRng = wDoc.Bookmarks.get_Item(ref bookmark).Range;
				Word.Table table = wDoc.Tables.Add(wrdRng, rows + 1, cols, ref nothing, ref nothing);
				table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
				table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
				table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
				table.Columns[1].Width = 100;
				table.Columns[2].Width = 400;

				// table header
				for (int i = 0; i < cols; i++)
				{
					table.Cell(1, i + 1).Range.Text = dt.Columns[i].ColumnName;
					table.Cell(1, i + 1).Range.Shading.ForegroundPatternColor
						= Word.WdColor.wdColorGray25;
					table.Cell(1, i + 1).Height = 35;
				}

				object idx = 1;
				Word.ListTemplate listTemp = wApp.ListGalleries[Word.WdListGalleryType.wdNumberGallery]
				.ListTemplates.get_Item(ref idx);
				object bContinuousPrev = false;
				object applyTo = Missing.Value;
				object defaultListBehaviour = Missing.Value;

				for (int i = 0; i < rows; i++)
				{
					for (int j = 0; j < cols; j++)
					{
						Word.Cell cell = table.Cell(i + 2, j + 1);

						object obj = dt.Rows[i][j];
						string txt = obj == null ? "" : obj.ToString();
						string[] lines = txt.Split(new string[] { "^n" },
							StringSplitOptions.RemoveEmptyEntries);


						cell.Range.Text = lines[0].Trim();

						if (lines.Length > 1)
						{
							cell.Range.Paragraphs[1].Range.ListFormat.ApplyNumberDefault();
							for (int k = 1; k < lines.Length; k++)
							{
								Word.Paragraph p = cell.Range.Paragraphs.Add();
								p.Range.Text = lines[k].Trim();
								p.Range.ListFormat.ApplyNumberDefault();
								object o = p.Range.ParagraphStyle;
							}

							for (int k = 0; k < cell.Range.Paragraphs.Count; k++)
							{
								Word.Paragraph p = cell.Range.Paragraphs[k + 1];
								p.Range.ListFormat.ApplyListTemplate(listTemp, bContinuousPrev,
									applyTo, defaultListBehaviour);
								p.Range.Select();
								wApp.Selection.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
								wApp.Selection.ParagraphFormat.LineSpacing = 14f;
								wApp.Selection.ParagraphFormat.LineUnitAfter = 0.5f;
							}

						}

					}
				}
			}
			catch (Exception)
			{
			}
		}


		public void addTable2(XmlNode tableNode)
		{
			if (tableNode == null) return;

			string bookmark = tableNode.Attributes["k"].Value;
			if (!wDoc.Bookmarks.Exists(bookmark)) return;

			XmlNodeList rows = tableNode.SelectNodes("tr");
			int c = 0, r = rows.Count;
			object nothing = Missing.Value;

			// get max cols
			foreach (XmlNode row in rows)
			{
				int t = row.SelectNodes("td").Count;
				c = Math.Max(c, t);
			}

			// create table of size c x r
			object bm = bookmark;
			Word.Range wrdRng = wDoc.Bookmarks.get_Item(ref bm).Range;
			Word.Table table = wDoc.Tables.Add(wrdRng, r, c, ref nothing, ref nothing);
			table.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
			table.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
			table.PreferredWidthType = Word.WdPreferredWidthType.wdPreferredWidthPercent;
			table.Columns[1].Width = 100;
			table.Columns[2].Width = 400;

			for (int i = 0; i < r; i++)
			{
				XmlNode row = rows[i];
				XmlNodeList cols = row.SelectNodes("td");

				for (int j = 0; j < c; j++)
				{
					setCellContent(table.Cell(i + 1, j + 1), cols[j]);
				}
			}

			for (int i = 0; i < c; i++)
			{
				Word.Cell cell = table.Cell(1, i + 1);
				cell.Height = 35;
				cell.Range.Shading.ForegroundPatternColor =
					Word.WdColor.wdColorGray25;
			}
		}

		private void setCellContent(Word.Cell cell, XmlNode tdNode)
		{
			XmlNodeList lis = tdNode.SelectNodes("ol/li");
			if (lis.Count == 0)
			{
				cell.Range.Text = tdNode.InnerText.Trim();
				return;
			}
			for (int i = 0; i < lis.Count; i++)
			{
				XmlNode li = lis[i];
				string txt = li.InnerText.Trim();
				if (cell.Range.Paragraphs.Count > i)
				{
					cell.Range.Paragraphs[i + 1].Range.Text = txt;
				}
				else
				{
					Word.Paragraph p = cell.Range.Paragraphs.Add();
					p.Range.Text = txt;
				}
			}

			object idx = 1;
			Word.ListTemplate listTemp = wApp.ListGalleries[Word.WdListGalleryType.wdNumberGallery]
			.ListTemplates.get_Item(ref idx);
			object bContinuousPrev = false;
			object applyTo = Missing.Value;
			object defaultListBehaviour = Missing.Value;

			for (int k = 0; k < cell.Range.Paragraphs.Count; k++)
			{
				Word.Paragraph p = cell.Range.Paragraphs[k + 1];
				p.Range.ListFormat.ApplyListTemplate(listTemp, bContinuousPrev,
					applyTo, defaultListBehaviour);
				p.Range.Select();
				wApp.Selection.ParagraphFormat.LineSpacingRule =
					Word.WdLineSpacing.wdLineSpaceExactly;
				wApp.Selection.ParagraphFormat.LineSpacing = 14f;
				wApp.Selection.ParagraphFormat.LineUnitAfter = 0.5f;
			}

		}

		public void addList(XmlNode olNode)
		{
			if (olNode == null) return;

			string bookmark = olNode.Attributes["k"].Value;
			if (!wDoc.Bookmarks.Exists(bookmark)) return;

			XmlNodeList lis = olNode.SelectNodes("li");
			if (lis.Count == 0) return;

			object bm = bookmark;
			Word.Range wrdRange = wDoc.Bookmarks.get_Item(ref bm).Range;
			wrdRange.Select();

			for (int i = 0; i < lis.Count; i++)
			{
				wApp.Selection.TypeText(lis[i].InnerText.Trim());
				wApp.Selection.TypeParagraph();
			}

			// select these paragraphs
			wrdRange.Select();
			Object unit = Word.WdUnits.wdParagraph;
			Object count = lis.Count;
			Object extend = Word.WdMovementType.wdExtend;
			wApp.Selection.MoveDown(ref unit, ref count, ref extend);

			object idx = 1;
			Word.ListTemplate listTemp = wApp.ListGalleries
				[Word.WdListGalleryType.wdNumberGallery]
				.ListTemplates.get_Item(ref idx);
			object bContinuousPrev = false;

			int c = wApp.Selection.Paragraphs.Count;
			for (int i = 0; i < c; i++)
			{
				wApp.Selection.Paragraphs[i + 1].Range.ListFormat.ApplyListTemplate(
					listTemp, bContinuousPrev, nothing, nothing);
				if (i == 0) bContinuousPrev = true;
			}
		}

		public void addUnorderedList(XmlNode ulNode)
		{
			if (ulNode == null) return;

			string bookmark = ulNode.Attributes["k"].Value;
			if (!wDoc.Bookmarks.Exists(bookmark)) return;

			XmlNodeList lis = ulNode.SelectNodes("li");
			if (lis.Count == 0) return;

			object bm = bookmark;
			Word.Range wrdRange = wDoc.Bookmarks.get_Item(ref bm).Range;
			wrdRange.Select();

			for (int i = 0; i < lis.Count; i++)
			{
				wApp.Selection.TypeText(lis[i].InnerText.Trim());
				wApp.Selection.TypeParagraph();
			}

			// select these paragraphs
			wrdRange.Select();
			Object unit = Word.WdUnits.wdParagraph;
			Object count = lis.Count;
			Object extend = Word.WdMovementType.wdExtend;
			wApp.Selection.MoveDown(ref unit, ref count, ref extend);

			object idx = 1;
			Word.ListTemplate listTemp = wApp.ListGalleries
				[Word.WdListGalleryType.wdBulletGallery]
				.ListTemplates.get_Item(ref idx);
			object bContinuousPrev = false;

			int c = wApp.Selection.Paragraphs.Count;
			for (int i = 0; i < c; i++)
			{
				wApp.Selection.Paragraphs[i + 1].Range.ListFormat.ApplyListTemplate(
					listTemp, bContinuousPrev, nothing, nothing);
				if (i == 0) bContinuousPrev = true;
			}
		}
	}
}
