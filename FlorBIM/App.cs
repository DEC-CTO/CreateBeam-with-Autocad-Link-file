using Autodesk;
using Autodesk.Revit;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Text;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace FlorBIM
{
    public class App : IExternalApplication
    {
        internal static App thisApp = null;
        private Main m_MyForm;

        public static UIControlledApplication UIControlApp = null;
        static string AddInPath = typeof(App).Assembly.Location;
        static string ButtonIconsFolder = Path.GetDirectoryName(AddInPath);
        public static string AppName = Path.GetFileNameWithoutExtension(AddInPath);

        public Result OnStartup(UIControlledApplication app)
        {
            m_MyForm = null;
            thisApp = this;
            UIControlApp = app;

            try
            {
                string tabname = "FlorBIM";
                try { app.CreateRibbonTab(tabname); } catch { }

                RibbonPanel ribbonPanel = null;
                List<RibbonPanel> listRibbonPanel = app.GetRibbonPanels();

                if (listRibbonPanel.Count > 0)
                {
                    foreach (RibbonPanel rp in listRibbonPanel)
                    {
                        if (rp.Name == "FlorBIM")
                        {
                            ribbonPanel = rp;
                            break;
                        }
                    }
                }

                if (ribbonPanel == null)
                    ribbonPanel = app.CreateRibbonPanel(tabname, "Lecture");

                if (ribbonPanel != null)
                {
                    {
                        bool bFound = false;

                        foreach (RibbonItem item in ribbonPanel.GetItems())
                        {
                            if (item.Name == "Lecture_")
                                bFound = true;
                        }
                        if (!bFound)
                        {
                            PushButton pushBotton = ribbonPanel.AddItem(new PushButtonData("_____", "_____", AddInPath, "FlorBIM.Command")) as PushButton;
                            pushBotton.ToolTip = "Create Model";
                            pushBotton.LargeImage = GetBitmapSource(FlorBIM.Properties.Resources.Red_24);
                            //pushBotton.Image = GetBitmapSource(DataFlow.Properties.Resources.MJS);
                        }
                    }
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            if (m_MyForm != null && !m_MyForm.Visible)
            {
                m_MyForm.Close();
            }

            return Result.Succeeded;
        }

        public void ShowForm(UIApplication uiapp)
        {
            if (m_MyForm == null || m_MyForm.IsDisposed)
            {
                RequestHandler handler = new RequestHandler();
                ExternalEvent exEvent = ExternalEvent.Create(handler);
                m_MyForm = new Main(exEvent, handler);
                m_MyForm.Show();
            }
        }

        public void WakeFormUp()
        {
            if (m_MyForm != null)
            {
                m_MyForm.WakeUp();
            }
        }

        private System.Windows.Media.Imaging.BitmapSource GetBitmapSource(System.Drawing.Bitmap _image)
        {
            System.Drawing.Bitmap bitmap = _image;
            BitmapSource bitmapSource = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                                            bitmap.GetHbitmap(),
                                            System.IntPtr.Zero,
                                            System.Windows.Int32Rect.Empty,
                                            System.Windows.Media.Imaging.BitmapSizeOptions.FromEmptyOptions());

            return bitmapSource;
        }
    }
}
