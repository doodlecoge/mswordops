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
			chart.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
			chart.ChartAreas[0].AxisX.LabelStyle.Font = 
				new Font(new FontFamily("微软雅黑"), 8f);
			chart.ChartAreas[0].AxisX.LabelStyle.TruncatedLabels = false;

			chart.ChartAreas[0].AxisY.Minimum = 1;
			//chart.ChartAreas[0].AxisY.LabelStyle.Enabled = false;
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
		}

	}
}
