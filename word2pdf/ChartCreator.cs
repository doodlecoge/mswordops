using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms.DataVisualization.Charting;
using System.Data;
using System.Drawing;
using System.IO;



namespace word2pdf
{
	class ChartCreator
	{
		public static Chart createChart(string chartType, String[][] data)
		{
			DataTable dt = new DataTable();
			foreach (var colName in data[0])
			{
				dt.Columns.Add(colName);
			}

			for (int i = 1; i < data.Length; i++)
			{
				dt.Rows.Add(data[i]);
			}

			return createChart(chartType, dt);
		}

		public static Chart createChart(string chartType, DataTable dt)
		{
			Chart chart = new Chart();
			chart.Width = 600;

			if (chartType == "vbar")
			{
				int aHeight = Math.Min(50, 20 * dt.Rows.Count);
				chart.Height = Math.Max(aHeight * dt.Columns.Count, 300);
			}
			else if (chartType == "radar")
			{
				chart.Height = 450;
			}
			else chart.Height = 300;

			chart.ChartAreas.Add(new ChartArea());
			chart.ChartAreas[0].AxisX.Interval = 1;
			chart.ChartAreas[0].AxisY.Interval = 1;

			chart.ChartAreas[0].Position.Width = 100;
			chart.ChartAreas[0].Position.Height = 100;
			chart.ChartAreas[0].Position.X = 0;
			chart.ChartAreas[0].Position.Y = 0;


			//chart.ChartAreas[0].AxisX.Minimum = 1;
			chart.ChartAreas[0].AxisY.Minimum = 1;
			chart.ChartAreas[0].AxisY.LabelStyle.Enabled = false;

			chart.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
			chart.ChartAreas[0].AxisY.MajorGrid.LineWidth = 1;
			chart.ChartAreas[0].AxisY.MajorGrid.LineColor = Color.LightGray;
			//chart.ChartAreas[0].AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dot;

			int m = max(dt);
			chart.ChartAreas[0].AxisY.Maximum = Math.Max(7, m);
			//chart.ChartAreas[0].AxisY.Enabled = AxisEnabled.False;
			chart.ChartAreas[0].AxisY.MajorTickMark.Enabled = false;


			//chart.ChartAreas[0].AxisY.LabelAutoFitMaxFontSize = 5;
			//chart.ChartAreas[0].AxisX.LabelAutoFitMaxFontSize = 5;
			//chart.ChartAreas[0].AxisY.Crossing = 1;
			//chart.ChartAreas[0].AxisY.MajorGrid.IntervalOffset = 3;

			//chart.ChartAreas[0].AxisX.LabelStyle.Angle = 0;

			//chart.ChartAreas[0].InnerPlotPosition.X = 0;
			//chart.ChartAreas[0].InnerPlotPosition.Y = 50;
			//chart.ChartAreas[0].InnerPlotPosition.Height = 80;
			//chart.ChartAreas[0].InnerPlotPosition.Width = 100;


			if (dt.Rows.Count > 1)
			{

				chart.ChartAreas[0].Position.Y = 10;


				chart.Legends.Add("legend");
				//chart.Legends[0].Position.Width = 100;
				//chart.Legends[0].Position.X = 50;
				//chart.Legends[0].Position.Y = 3;
				//chart.Legends[0].Position.X = 50.5f;
				//chart.Legends[0].LegendStyle = LegendStyle.Row;
				chart.Legends[0].Enabled = true;
				chart.Legends[0].LegendStyle = LegendStyle.Row;
				chart.Legends[0].Position.Width = 100;
				chart.Legends[0].Position.Height = 10;
				chart.Legends[0].Position.X = 0;
				chart.Legends[0].Position.Y = 0;
				//chart.Legends[0].Position.Auto = false;
				//chart.Legends[0].Position.Width = 100;
				//chart.Legends[0].Position.Height = 10;
				//chart.Legends[0].Position.X = 0;
				//chart.Legends[0].Position.Y = 0;
			}

			SeriesChartType t;
			switch (chartType)
			{
				case "vbar":
					t = SeriesChartType.Bar;
					//chart.ChartAreas[0].AxisX.LabelStyle.Angle = -45;

					break;
				case "radar":
					t = SeriesChartType.Radar;
					break;
				case "line":
					t = SeriesChartType.Line;
					chart.ChartAreas[0].AxisX.LabelStyle.Angle = 45;
					//chart.ChartAreas[0].AxisX.LabelStyle.IsStaggered = true;
					break;
				default:
					t = SeriesChartType.Column;
					break;
			}


			int rows = dt.Rows.Count;
			int cols = dt.Columns.Count;

			for (int i = 0; i < rows; i++)
			{
				Series s = chart.Series.Add(dt.Rows[i][0].ToString());
				s.ChartType = t;
				if (chartType == "radar")
				{
					//s.CustomProperties = "RadarDrawingStyle=Line, CircularLabelsStyle=Circular";
					s.CustomProperties = "RadarDrawingStyle=Line";
					s.BorderWidth = 3;
				}
				else if (chartType == "line")
				{
					s.BorderWidth = 3;
				}
				else if (chartType == "bar" || chartType == "vbar")
				{

					var d = dt.Rows.Count * 1.0 / (dt.Rows.Count + 1);
					d = Math.Round(d, 3);

					s.CustomProperties = "PointWidth=" + d;
				}

				//chart.ChartAreas[0].AxisX.LabelStyle.Font = new Font("Microsoft Sans Serif", 8.0f);

				for (int j = 1; j < cols; j++)
				{
					s.Points.AddXY(dt.Columns[j].ColumnName, dt.Rows[i][j]);
				}
			}

			//int cols = dt.Columns.Count;
			//for (int i = 1; i < cols; i++)
			//{
			//    chart.Series.Add("" + i);
			//    chart.Series[i - 1].Points.AddXY(dt.Columns[i].ColumnName);
			//    chart.Series[i - 1].Points.DataBindXY(
			//        dt.DefaultView, dt.Columns[0].ColumnName,
			//        dt.DefaultView, dt.Columns[i].ColumnName
			//        );
			//    chart.Series[i - 1].Name = dt.Columns[i].ColumnName;
			//    chart.Series[i - 1].IsVisibleInLegend = true;

			//    chart.Series[i - 1].ChartType = t;
			//    if (chartType == "radar")
			//    {
			//        chart.Series[i - 1].CustomProperties = 
			//            "RadarDrawingStyle=Line, CircularLabelsStyle=Circular";
			//        chart.Series[i - 1].BorderWidth = 5;
			//    }
			//    chart.ChartAreas[0].AxisX.LabelStyle.Font = new Font("Microsoft Sans Serif", 8.0f);
			//}

			//chart.SaveImage("e:\\chart.jpg", ChartImageFormat.Jpeg);
			return chart;
		}

		private static int max(DataTable dt)
		{
			double m = 0;
			int cols = dt.Columns.Count;
			foreach (DataRow row in dt.Rows)
			{
				for (int i = 0; i < cols; i++)
				{
					string val = row[i] == null ? "n/a" : row[i].ToString();
					try
					{
						double x = Double.Parse(val);
						if (m < x) m = x;
					}
					catch (Exception)
					{
					}
				}
			}
			return (int)Math.Ceiling(m);
		}

		public static void foo()
		{
			DataTable dt = new DataTable();
			dt.Columns.Add("Competence", typeof(String));
			dt.Columns.Add("个人", typeof(int));
			dt.Columns.Add("上级", typeof(int));

			dt.Rows.Add("热爱行业", 1.8, 4.2);
			dt.Rows.Add("细节坚持", 4.2, 3.5);
			dt.Rows.Add("数字敏感", 3.5, 1.4);
			dt.Rows.Add("数据分析及问题解决", 1.4, 3.6);
			dt.Rows.Add("战略分析及推动", 3.6, 5.0);
			dt.Rows.Add("绩效激励管理", 5.0, 1.0);
			dt.Rows.Add("协同合作", 1.0, 4.3);
			dt.Rows.Add("新店经营", 4.3, 2.0);
			dt.Rows.Add("市场商机推动", 2.0, 2.9);
			dt.Rows.Add("团队人才培养", 2.9, 3.8);
			dt.Rows.Add("标准化运营", 3.8, 1.1);

			createChart("vbar", dt);


			//            Chart chart = new Chart();
			//            chart.ChartAreas.Add(new ChartArea());
			//            chart.Width = 800;
			//            chart.Height = 800;

			//            chart.Legends.Add("legend");
			//            chart.Legends[0].Enabled = true;



			//            //格式化标签和间隔
			//            //chart.ChartAreas[0].AxisX.LabelStyle.Format = @"{yyyy-MM-dd HH时}";
			//            //chart.ChartAreas[0].AxisX.IntervalType = DateTimeIntervalType.Hours;
			////			chart.ChartAreas[0].AxisX.Interval = 1;



			//            //绑定数据
			//            chart.Series.Add("a");
			//            chart.Series[0].BorderWidth = 100;
			//            chart.Series[0].ChartType = SeriesChartType.Radar;
			//            chart.Series[0].Points.DataBindXY(dt.DefaultView, "Competence",
			//                dt.DefaultView, "Value");
			//            //chart.Series[0].IsValueShownAsLabel = true;
			//            //chart.Series[0].XValueType = ChartValueType.Date;

			//            chart.Series[0].BorderWidth = 10;
			//            chart.Series[0].Name = "ge ren";
			//            chart.Series[0].IsVisibleInLegend = true;

			//            chart.Series.Add("b");
			//            chart.Series[0].BorderWidth = 100;
			//            chart.Series[0].ChartType = SeriesChartType.Radar;
			//            chart.Series[0].Points.DataBindXY(dt.DefaultView, "Competence",
			//                dt.DefaultView, "Value2");
			//            //chart.Series["b"].IsValueShownAsLabel = true;
			//            chart.Series["b"].Name = "ge ren2";



			//            chart.SaveImage("e:\\chart.jpg",ChartImageFormat.Jpeg);





			//Report2BLL reportBLL = new Report2BLL();
			//Random rd = new Random(175);
			//for (int i = 0; i < dtPhone.Rows.Count; i++)
			//{
			//    Chart2.Series.Add(dtPhone.Rows[i]["PhoneName"].ToString());//
			//    int phoneID = Int32.Parse(dtPhone.Rows[i]["PhoneID"].ToString());
			//    DataTable dt = reportBLL.GetFeatureScore(phoneID);
			//    Chart2.Series[i].Points.DataBind(dt.DefaultView, "FeatureName", "Score", "LegendText=FeatureName,YValues=Score,ToolTip=Score");
			//    Chart2.Series[i].ChartType = SeriesChartType.Radar;
			//    Color color = Color.FromArgb(rd.Next(0, 255), rd.Next(0, 255), rd.Next(0, 255));
			//    Chart2.Series[i].BorderColor = color;
			//    Chart2.Series[i].BorderWidth = 2;
			//    Chart2.Series[i].Color = Color.Transparent;//设置Radar区域颜色透明，以防止前面的遮挡后面的显示
			//    Chart2.Series[i].IsValueShownAsLabel = true;//图表显示具体的数值
			//}
			//Chart2.DataBind();
		}

	}
}
