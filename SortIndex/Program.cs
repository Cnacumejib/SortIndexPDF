using System;
using System.IO;
//using iTextSharp.license;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
//using System.Collections;
namespace SortIndex
{
	class MainClass
	{
		public static void Main (string[] args)
		{
            // iText.License.LicenseKey.LoadLicenseFile("path / to / itextkey.xml");
            Console.WriteLine ("Файлы pdf находящиеся в текущей директории" +
				"\nбудут разобраны по страницам и сложены в папку 1000pdf." +
				"\nПотом они будут объеденены по почтовому ИНДЕКСу" +
				"\n\nФайлы находящиеся сейчас в 1000pdf будут удалены\n\nНажмите любую клавишу чтобы продолжить");
			Console.ReadKey ();
			string myDir = Environment.CurrentDirectory;
			string[] myFiles=Directory.GetFiles (myDir,"*.pdf");
			myDir = System.IO.Path.Combine(myDir,"1000pdf");
			if(Directory.Exists(myDir))
				Directory.Delete(myDir,true);				
			Directory.CreateDirectory (myDir);				
			int a = 0;
			string myIndexes="";
			foreach (var myFile in myFiles)
            {
				Console.WriteLine("{0}       ", System.IO.Path.GetFileName(myFile));
				a+=SplitAndSave(myFile,myDir,a, ref myIndexes);
				Console.WriteLine("     {0}       {1}",a,myIndexes);
			}

			string[] Indexes = myIndexes.Split (' ');
			Console.Write ("\n\nПоделили: Всего {0} файлов\n{1}\nСобираем...",a,myIndexes);
			a = 0;
			int b ;
			foreach (var item in Indexes) 
			{
			    if (item == "")
					continue;
				string[] myNewFiles=Directory.GetFiles (myDir,item+"*.pdf");
				Console.WriteLine (item+"   "+myNewFiles.Length);
                if (myNewFiles.Length > 1000)
                {
                    b = 0;
                    string[] newShortList=new string[1000];
                    for (int i = 0; i < myNewFiles.Length; i=i+1000)
                    {
                     //   myNewFiles.CopyTo(newShortList, i);
                        Array.Copy(myNewFiles,i, newShortList, 0,1000);
                        b += MergePDFs(newShortList, System.IO.Path.Combine(myDir, $"Index_{item}_{i}.pdf"));

                    }
                }
                else
                {
                    b = MergePDFs(myNewFiles, System.IO.Path.Combine(myDir, $"Index_{item}.pdf"));

                }

				Console.WriteLine ("   "+item+Environment.NewLine);
				a += b;
			}
			Console.WriteLine ("\r\nСобрали {1} файлов. Отделений связи: {0}  ",Indexes.Length-1,a);
			string[] myIndexFiles=Directory.GetFiles (myDir,"*.pdf");
			int pages = 0;
			int total = 0;
			foreach (var iFile in myIndexFiles) 
			{
				using (PdfReader reader = new PdfReader (iFile)) 
				{	
					pages= reader.NumberOfPages;	
					reader.Close();
				}
				File.Move(iFile, System.IO.Path.Combine(myDir,System.IO.Path.GetFileNameWithoutExtension(iFile)+"_"+pages+".pdf"));
				total += pages;
			}
			Console.WriteLine ("Обработано всего {0}",total);
			//Console.ReadKey();
		}
		public static int MergePDFs(string[] fileNames, string targetPdf)
		{
			int counter = 0;
            if(fileNames.Length==0)
            {
                return 0;

            }

			using (FileStream stream = new FileStream(targetPdf, FileMode.Create))
			{
				Document document = new Document();
				PdfCopy pdf = new PdfCopy(document, stream);
				PdfReader reader = null;
				try
				{
					document.Open();
					foreach (string file in fileNames)
					{
						reader = new PdfReader(file);
						pdf.AddDocument(reader);
						counter++;
						Console.Write ("\r"+counter);
						reader.Close();
						File.Delete(file);
                        if (counter % 1000==0)
                        {
                            document.Close();

                        }
					}
				}
				catch (Exception)
				{
				//	counter = 0;
					if (reader != null)
					{
						reader.Close();
					}
				}
				finally
				{
					if (document != null)
					{
						document.Close();
					}
				}
			}
			return counter;
		}

		static public int SplitAndSave(string inputPath, string outputPath, int start, ref string myIndexes)
		{
			//myIndexes = "";
			string filename = "";
			int counter = 0;
			using (PdfReader reader = new PdfReader(inputPath))
			{
				ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
				for (int pagenumber = 1; pagenumber <=reader.NumberOfPages ; pagenumber++)//
				{
					filename = "Без индекса "+(start+pagenumber).ToString() + ".pdf";
					string[] currentText = PdfTextExtractor.GetTextFromPage(reader, pagenumber, strategy).Split('\n');
					for (int i=0; i< currentText.Length;i++) {
						if (currentText[i] == "Адрес:") {
							filename =currentText[i+1].Substring(0, 6) +"  " +(start+pagenumber).ToString() + ".pdf";
							if(!myIndexes.Contains(currentText[i+1].Substring(0, 6)))
							{
								myIndexes +=" " + currentText [i + 1].Substring(0, 6);
							}
                            break;
						}
					}
					Document document = new Document();
					PdfCopy copy = new PdfCopy(document, new FileStream(System.IO.Path.Combine(outputPath, filename), FileMode.Create));

					document.Open();

					copy.AddPage(copy.GetImportedPage(reader, pagenumber));
					counter++;
					document.Close();
					Console.Write("\r  {0}   ",pagenumber);
				}
				return counter;
			}

		}
	}
}

