using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace word2pdf
{
	class Word2Pdf
	{
		public void convert(string src, string dst)
		{
			src = Path.GetFullPath(src);

			if (!File.Exists(src))
			{
				Console.Out.WriteLine("File does not exist: " + src);
				return;
			}

			dst = Path.GetFullPath(dst);

			object osrc = src;
			object odst = dst;

			object Background = true;
			object PrintToFile = true;
			object varFalseValue = false;
			object varTrueValue = true;
			object FileFormat = Word.WdSaveFormat.wdFormatPDF;
			object Encoding = Office.MsoEncoding.msoEncodingUTF8;
			object varMissing = Type.Missing;

			Word.Application app = null;
			Word.Document doc = null;

			try
			{
				app = new Word.Application();

				doc = app.Documents.Open(ref osrc,
					ref varMissing, ref varTrueValue, ref varMissing, ref varMissing, ref varMissing,
					ref varMissing, ref varMissing, ref varMissing, ref varMissing, ref varMissing,
					ref varMissing, ref varMissing, ref varMissing, ref varMissing, ref varMissing);

				doc.SaveAs(ref odst, ref FileFormat, ref varFalseValue,
					ref varMissing, ref varFalseValue, ref varMissing, ref varFalseValue,
					ref varTrueValue, ref varTrueValue, ref varTrueValue, ref varTrueValue,
					ref Encoding, ref varFalseValue, ref varFalseValue, ref varMissing,
					ref varMissing);
			}
			catch (Exception e)
			{
				Console.Out.WriteLine(e.Message);
			}
			finally
			{
				if (doc != null)
					doc.Close(ref varMissing, ref varMissing, ref varMissing);

				if (app != null)
					app.Quit(ref varMissing, ref varMissing, ref varMissing);
			}
		}
	}
}
