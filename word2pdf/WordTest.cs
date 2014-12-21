using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using System.Reflection;



using System.ComponentModel;
using System.Data;

using System.IO;

using Microsoft.Office.Interop.Graph;






namespace word2pdf
{
	class WordTest
	{
		public static void foo()
		{
			object oMissing = System.Reflection.Missing.Value;

			object oEndOfDoc = "\\endofdoc";

			Word._Application oWord = new Word.Application();

			Word._Document oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
				ref oMissing, ref oMissing);

			oWord.Visible = true;

			Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
			//Insert a chart.

			Word.InlineShape oShape;

			object oClassType = "MSGraph.Chart.8";

			wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

			oShape = wrdRng.InlineShapes.AddOLEObject(ref oClassType, ref oMissing,
				ref oMissing, ref oMissing, ref oMissing,
				ref oMissing, ref oMissing, ref oMissing);



			//Demonstrate use of late bound oChart and oChartApp objects to
			//manipulate the chart object with MSGraph.

			object oChart;
			object oChartApp;

			oChart = oShape.OLEFormat.Object;
			oChartApp = oChart.GetType().InvokeMember("Application", BindingFlags.GetProperty,
				null, oChart, null);



			//Change the chart type to Line.

			object[] Parameters = new Object[1];
			Parameters[0] = 2; //xlLine = 4
			oChart.GetType().InvokeMember("ChartType", BindingFlags.SetProperty,
				null, oChart, Parameters);


			//Update the chart image and quit MSGraph.

			oChartApp.GetType().InvokeMember("Update", BindingFlags.InvokeMethod, null, oChartApp, null);

			oChartApp.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, oChartApp, null);

			//... If desired, you can proceed from here using the Microsoft Graph 

			//Object model on the oChart and oChartApp objects to make additional

			//changes to the chart.



			//Set the width of the chart.

			oShape.Width = oWord.InchesToPoints(6.25f);

			oShape.Height = oWord.InchesToPoints(3.57f);



			//Add text after the chart.

			wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

			wrdRng.InsertParagraphAfter();

			wrdRng.InsertAfter("THE END.");


		}

		public static void foo2()
		{
			object oMissing = System.Reflection.Missing.Value;
			object oEndOfDoc = "\\endofdoc";
			Word.Application oWord = new Word.Application();
			Word.Document oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
				ref oMissing, ref oMissing);
			oWord.Visible = true;

			string[,] data = { 
							 { "", "1", "2", "3", "4", "5", "6", "7", "8" }, 
							 { "4", "5", "6", "7", "8", "9", "10", "11", "12"} ,
							 {"7", "8", "9", "10", "11", "12","4", "5", "6"}
							 };
			AddSimpleChart(oDoc, oWord, oEndOfDoc, data);
		}



		public static void AddSimpleChart(Word.Document WordDoc, Word.Application WordApp,
			Object oEndOfDoc, string[,] data)
		{
			//插入chart
			object oMissing = System.Reflection.Missing.Value;
			Word.InlineShape oShape;
			object oClassType = "MSGraph.Chart.8";
			Word.Range wrdRng = WordDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
			oShape = wrdRng.InlineShapes.AddOLEObject(ref oClassType, ref oMissing,
				ref oMissing, ref oMissing, ref oMissing,
				ref oMissing, ref oMissing, ref oMissing);
			//Demonstrate use of late bound oChart and oChartApp objects to   
			//manipulate the chart object with MSGraph.   
			object oChart;
			object oChartApp;
			oChart = oShape.OLEFormat.Object;
			oChartApp = oChart.GetType().InvokeMember("Application", BindingFlags.GetProperty,
				null, oChart, null);
			//Change the chart type
			Chart objChart = (Chart)oShape.OLEFormat.Object;
			//objChart.ChartType = XlChartType.xlColumnClustered;
			objChart.ChartType = XlChartType.xlLine;
			objChart.ChartArea.Border.LineStyle = XlLineStyle.xlDash;
			objChart.ChartArea.Fill.BackColor.SchemeColor = 28;
			

			Series s = (Series)objChart.SeriesCollection(1);
			
			

			objChart.ChartArea.Border.ColorIndex = 5;
			objChart.ChartArea.Fill.BackColor.SchemeColor = 1;
			


			

			
			//绑定数据   
			Microsoft.Office.Interop.Graph.DataSheet dataSheet;
			dataSheet = objChart.Application.DataSheet;
			dataSheet.Cells.ClearContents();
			int rownum = data.GetLength(0);
			int columnnum = data.GetLength(1);
			for (int i = 1; i <= rownum; i++)
				for (int j = 1; j <= columnnum; j++)
				{
					dataSheet.Cells[i, j] = data[i - 1, j - 1];
				}

			objChart.Application.Update();
			objChart.Application.Quit();

			//oChartApp.GetType().InvokeMember("Update", 
			//    BindingFlags.InvokeMethod, null, oChartApp, null);
			//oChartApp.GetType().InvokeMember("Quit", 
			//    BindingFlags.InvokeMethod, null, oChartApp, null);
			//设置大小
			oShape.Width = WordApp.InchesToPoints(6.25f);
			oShape.Height = WordApp.InchesToPoints(3.57f);
		}
	}


}
