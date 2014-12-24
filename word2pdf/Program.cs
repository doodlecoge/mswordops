using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Collections;
using System.IO;
using System.Web;
using System.Web.Security;
using Microsoft.Office.Interop.Word;
using System.Reflection;



namespace word2pdf
{
	class Program
	{
		static void Main(string[] args)
		{
			try
			{
				foo(args);
				Console.Out.Write(0);
			}
			catch (Exception e)
			{
				Console.Out.WriteLine("==================================================");
				Console.Out.WriteLine(e.Message);
				Console.Out.WriteLine(e.StackTrace);
				Console.Out.WriteLine();
			}
		}

		static void foo(string[] args)
		{
			XmlDocument xDoc = new XmlDocument();
			if (args.Length == 1)
			{
				string filename = Path.GetFullPath(args[0]);
				xDoc.Load(filename);
			}
			else
			{
				//xDoc.Load("C:\\Users\\hch\\Desktop\\tmp\\input.xml");

				Stream inp = Console.OpenStandardInput(1);
				StreamReader sr = new StreamReader(inp, Encoding.UTF8);
				xDoc.Load(sr);
			}
			
			XmlElement root = xDoc.DocumentElement;
			XmlNodeList children = root.ChildNodes;

			XmlNodeList tNodes = root.GetElementsByTagName("t");
			XmlNodeList iNodes = root.GetElementsByTagName("i");

			XmlNode nodeSrc = root.GetElementsByTagName("src")[0];
			XmlNode nodeDst = root.GetElementsByTagName("dst")[0];


			string src = nodeSrc.InnerText;
			string dst = nodeDst.InnerText;
			string imgs = "";

			XmlNodeList nodeImgs = root.GetElementsByTagName("imgs");
			if (nodeImgs.Count != 0)
			{
				imgs = nodeImgs[0].InnerText;
			}

			BookmarkReplace br = new BookmarkReplace();
			br.openDocument(src);

			foreach (XmlNode node in children)
			{
				if (node.Name == "t")
				{
					br.replaceText(node.Attributes["k"].Value, node.Attributes["v"].Value);
				}
				else if (node.Name == "i")
				{
					string path = node.Attributes["v"].Value;
					if (!path.Equals(Path.GetFullPath(path)))
					{
						path = imgs + path;
					}
					br.replaceImage(node.Attributes["k"].Value, path);
				}
				else if (node.Name == "c")
				{
					string txt = node.InnerText;
					String[] strRows = txt.Split(new string[] { "^n" },
						StringSplitOptions.None);
					int rows = strRows.Length;
					string[][] data = new string[rows][];


					for (int i = 0; i < rows; i++)
					{
						string strRow = strRows[i];
						String[] strCols = strRow.Split(new string[] { "|" },
							StringSplitOptions.None);
						int cols = strCols.Length;
						data[i] = new string[cols];
						for (var j = 0; j < cols; j++)
						{
							data[i][j] = strCols[j];
						}
					}

					br.addChart2(node.Attributes["k"].Value, node.Attributes["v"].Value, data);
				}
				else if (node.Name == "table")
				{
					br.addTable2(node);
				}
				else if(node.Name == "disc")
				{
					int d0 = Int32.Parse(node.Attributes["d0"].Value);
					int d1 = Int32.Parse(node.Attributes["d1"].Value);
					int d2 = Int32.Parse(node.Attributes["d2"].Value);

					int i0 = Int32.Parse(node.Attributes["i0"].Value);
					int i1 = Int32.Parse(node.Attributes["i1"].Value);
					int i2 = Int32.Parse(node.Attributes["i2"].Value);

					int s0 = Int32.Parse(node.Attributes["s0"].Value);
					int s1 = Int32.Parse(node.Attributes["s1"].Value);
					int s2 = Int32.Parse(node.Attributes["s2"].Value);

					int c0 = Int32.Parse(node.Attributes["c0"].Value);
					int c1 = Int32.Parse(node.Attributes["c1"].Value);
					int c2 = Int32.Parse(node.Attributes["c2"].Value);

					string img = DISC.createDISCImage(
						d0, d1, d2, 
						i0,	i1, i2, 
						s0, s1, s2, 
						c0, c1, c2,
						imgs + "DISC.jpg"
						);

					br.replaceImage(node.Attributes["k"].Value, img);

					try
					{
						File.Delete(img);
					}
					catch (Exception)
					{
					}
				}
				else if (node.Name == "ol")
				{
					br.addList(node);
				}
				else if (node.Name == "ul")
				{
					br.addUnorderedList(node);
				}
			}

			try
			{
				//br.saveAsWord(dst);
				br.saveAsPdf(dst);
			}
			finally
			{
				br.close();
			}
			
		}
	}
}
