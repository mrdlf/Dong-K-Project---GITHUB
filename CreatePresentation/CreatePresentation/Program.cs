using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace CreatePresentation
{
    class Program
    {
        public string dbPath = @"C:\Users\Dong Kieu\Desktop\Dong K Project\1. Source\Database\Files\N5_Volcabulary.csv";
        public string resultPath = @"C:\Users\Dong Kieu\Desktop\Dong K Project\1. Source\Database\SlideMaster-GOI-Trial1.pptx";
        public string slideMasterPath = @"C:\Users\Dong Kieu\Desktop\Dong K Project\1. Source\Database\SlideMaster-GOI.pptx";

        public string[] fData = new string[1000];
        public string[,] fLine = new string[1000, 5];

        public void ReadDataFromFile()
        {
            /*
            string text = System.IO.File.ReadAllText(@"D:\General\GNR\Trial1.csv");
            
            System.Console.WriteLine("Contents of Trial.csv = {0}", text);
            */
            string[] lines = System.IO.File.ReadAllLines(dbPath);

            //System.Console.WriteLine("Contents of Trial.cs = ");
            int i = 0;

            foreach (string line in lines)
            {
                // Read each line
                fData[i] = line;

                // Read data among comma ','
                string[] col = line.Split(new char[] { '/' });
                int colLength = col.Length;

                for (int j = 0; j < colLength; j++)
                {
                    fLine[i, j] = col[j];
                }

                // Up index
                i++;
            }

            fData = fData.Where(s => !string.IsNullOrEmpty(s)).ToArray();
        }

        static void Main(string[] args)
        {
            Program p = new Program();
            p.ReadDataFromFile();

            PowerPoint.Application PowerPoint_App = new PowerPoint.Application();
            PowerPoint.Presentations multi_presentations = PowerPoint_App.Presentations;
            PowerPoint.Presentation presentation = multi_presentations.Open(
                p.slideMasterPath, MsoTriState.msoTrue, MsoTriState.msoFalse, MsoTriState.msoFalse);
            PowerPoint.TextFrame2.TextRange objText;
            
            var slideNum = presentation.Slides.Count;
            
            // i : the index of slide in master presentation
            for (int i = 1; i < slideNum; i++)
            {
                // j: the index of shapes in master presentation
                // j = 1: Word
                // j = 2: Kanji if any
                // j = 3: Meaning in Vietnamese
                // j = 4: Meaning in English
                // j = 5: Image
                for (int j = 0; j < 5; j++)
                {
                    // Get object of the current shape
                    PowerPoint.Shape shape = presentation.Slides[i + 1].Shapes[j + 1];

                    if (j < 4)
                    {
                        // Get object text range of this shape to insert text
                        objText = shape.TextFrame2.TextRange;
                        // Assign text read from data file to shape
                        objText.Text = p.fLine[i - 1, j];
                        // To fit text into shape
                        shape.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeTextToFitShape;
                    } else if (j == 4) // In case of shape to insert image
                    {
                        // Get the image link from data file
                        string imgFile = p.fLine[i - 1, j];

                        // Insert picture into the current shape
                        presentation.Slides[i + 1].Shapes.AddPicture(
                            imgFile, MsoTriState.msoTrue, MsoTriState.msoTrue, 
                            shape.Left, shape.Top, shape.Width, shape.Height);
                    }
                }


                /*
                foreach (var item in presentation.Slides[i+1].Shapes)
                {
                    var shape = (PowerPoint.Shape)item;
                    
                    if (shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                        {
                        
                            shape.TextFrame.TextRange.Delete();
                            objText = shape.TextFrame.TextRange;
                            objText.Text = "I caught ya!!!";
                            objText.Font.Name = "Arial";
                            objText.Font.Size = 18;
                            //shape.TextFrame.TextRange.InsertAfter("Oh yeah I caught you!!!");
                            //textBox.TextFrame.TextRange.InsertAfter("Yes I am fine!");
                        }
                    }
                }
                */
                
            }

            presentation.SaveAs(
                p.resultPath,
                PowerPoint.PpSaveAsFileType.ppSaveAsDefault,
                MsoTriState.msoTriStateMixed);
            
            PowerPoint_App.Quit();
            
        }
    }
}
