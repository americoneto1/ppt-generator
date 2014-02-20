using System;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace LouvorPPT
{
    public class Presentation
    {
        private Configuracao _objConfig = null;

        public Presentation(Configuracao objConfig)
        {
            _objConfig = objConfig;
        }     

        public void Generate(string title, string text)
        {
            Application objApplication = null;
            Microsoft.Office.Interop.PowerPoint.Presentation objPresentation = null;

            try
            {
                objApplication = new Application();
                objApplication.Visible = MsoTriState.msoTrue;
                objApplication.WindowState = PpWindowState.ppWindowMaximized;

                objPresentation = objApplication.Presentations.Open(_objConfig.TemplateFile);

                int count = 1;
                string[] parts = Regex.Split(text, "\r\n\r\n");

                foreach (var item in parts)
                {
                    Slide objSlide = objPresentation.Slides.AddSlide(count, objPresentation.Slides[1].CustomLayout);
                    var _with1 = objSlide.Shapes[1].TextFrame.TextRange;
                    _with1.Text = item;
                    objSlide = null;
                    count++;
                }

                objPresentation.Slides[objPresentation.Slides.Count].Delete();

                title = title.Replace(" ", "_");
                string fileName = string.Format(@"{0}/{1}.pptx", _objConfig.DestinationPath, title);
                objPresentation.SaveAs(fileName);
            }
            catch(Exception ex)
            {
                throw ex;
            }
            finally
            {
                ReleaseObjects(objPresentation, objApplication);
            }
        }

        private void ReleaseObjects(Microsoft.Office.Interop.PowerPoint.Presentation _objPresentation, Application _objApplication)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();

            if (_objPresentation != null)
            {
                _objPresentation.Close();
                //Marshal.FinalReleaseComObject(_objPresentation);
            }

            if (_objApplication != null)
            {
                _objApplication.Quit();
                //Marshal.FinalReleaseComObject(_objApplication);
            }

            System.Diagnostics.Process[] pros = System.Diagnostics.Process.GetProcessesByName("POWERPNT");
            for (int i = 0; i < pros.Length; i++)
            {
                pros[i].Kill();
            }
        }
    }
}
