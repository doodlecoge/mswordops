using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Drawing;
using System.Web.UI.DataVisualization.Charting;
using System.Reflection;
using System.Web.UI.WebControls;

using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using System.Collections;

namespace word2pdf
{
	class WordChart
	{
		public static void addChart()
		{
			object oMissing = System.Reflection.Missing.Value;
			object oEndOfDoc = "\\endofdoc";
			Word.Application oWord = new Word.Application();
			Word.Document oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
				ref oMissing, ref oMissing);
			oWord.Visible = true;
			Word.Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;


			Chart chart = new Chart();
			chart.Width = (Unit)400;
			chart.Height = (Unit)300;

			ChartArea area = new ChartArea("");
			area.AxisX.LineColor = System.Drawing.Color.FromArgb(0, 104, 188);
			area.AxisY.LineColor = System.Drawing.Color.FromArgb(0, 104, 188);
			area.AxisX.LineWidth = 2;
			area.AxisY.LineWidth = 2;
			//area.AxisY.Title = "%";
			area.AxisX.MajorGrid.LineWidth = 0;
			area.AxisX.MajorGrid.LineColor = System.Drawing.Color.FromArgb(214, 214, 214);
			area.AxisY.MajorGrid.LineColor = System.Drawing.Color.FromArgb(214, 214, 214);

			area.AxisY.Maximum = 5;
			area.AxisY.Interval = 1;
			area.AxisX.Interval = 1;
			//area.AxisX.LabelAutoFitStyle = LabelAutoFitStyles..LabelsAngleStep90;
			area.AxisX.LabelStyle.Angle = -70;
			area.AxisX.IsLabelAutoFit = false;

			chart.ChartAreas.Add(area);

			Series series0 = new Series();
			series0.ChartType = SeriesChartType.Column;
			series0.ChartArea = area.Name;
			series0.BorderWidth = 2;                                  //线条宽度  
			series0.ShadowOffset = 1;                                 //阴影宽度  
			series0.IsVisibleInLegend = true;                         //是否显示数据说明  
			series0.IsValueShownAsLabel = false;
			series0.MarkerStyle = MarkerStyle.None;
			series0.MarkerSize = 8;

			series0.Color = System.Drawing.Color.FromArgb(0, 176, 80);
			//series0.Label = "#PERCENT";
			series0.LegendText = "自评得分";

			//Series series1 = new Series();
			//series1.ChartType = SeriesChartType.Column;
			//series1.ChartArea = area.Name;
			//series1.BorderWidth = 2;                                  //线条宽度  
			//series1.ShadowOffset = 1;                                 //阴影宽度  
			//series1.IsVisibleInLegend = true;                         //是否显示数据说明  
			//series1.IsValueShownAsLabel = false;
			//series1.MarkerStyle = MarkerStyle.None;
			//series1.MarkerSize = 8;
			//series1.Color = System.Drawing.Color.FromArgb(255, 217, 102);
			////series0.Label = "#PERCENT";
			//series1.LegendText = "他评得分";


			Dictionary<string, int> data = new Dictionary<string, int>();
			data.Add("a", 1);
			data.Add("b", 2);
			data.Add("c", 3);
			data.Add("d", 4);
			data.Add("e", 3);

			foreach (var e in data) 
			{
				series0.Points.AddXY(e.Key, e.Value);
			}
			
			//foreach (DictionaryEntry de in subDimension)
			//{
			//    series0.Points.AddXY(de.Key, 
			//        decimal.Parse((((int)de.Value + 12) / 4.8).ToString()));
			//    series1.Points.AddXY(de.Key, 
			//        decimal.Parse(((int)htSubDimension2[de.Key.ToString().Trim()] / 4.0).ToString()));
			//    //ht3.Add(de.Key.ToString().Trim(), Math.Abs((int)htSubDimension2[de.Key.ToString().Trim()] - (int)de.Value));

			//}




			chart.Series.Add(series0);
			//chart.Series.Add(series1);
			Legend legend = new Legend();
			legend.Docking = Docking.Bottom;
			chart.Legends.Add(legend);

			//string jpgName3 = HttpContext.Current.Server.MapPath("~/Temp/" + Guid.NewGuid().ToString() + ".jpg");
			string jpg = "E:\\chart.jps";
			chart.SaveImage(jpg);

			wrdRng= oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;

			object anchor = oWord.Selection.Range;
			object LinkToFile = false;
			object SaveWithDocument = true;
			
			oDoc.InlineShapes.AddPicture(jpg, ref LinkToFile, ref SaveWithDocument, ref anchor); 

		}
	}
}
