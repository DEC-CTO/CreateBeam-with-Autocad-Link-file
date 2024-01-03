using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design.Serialization;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.Layout;
using System.Windows.Media;
using System.Windows.Media.Imaging;

using ACadSharp.IO;
using ACadSharp;
using ACadSharp.Attributes;
using ACadSharp.Entities;
using ACadSharp.Tables.Collections;

namespace FlorBIM
{
    public class Lib
    {

        public static List<string> GetDWGLink(Document doc)
        {
            List<string> list = new List<string>();
            FilteredElementCollector linkedDWG = new FilteredElementCollector(doc).OfClass(typeof(CADLinkType));

            try
            {
                foreach (CADLinkType item in linkedDWG)
                {
                    if(item.IsExternalFileReference() == true)
                    {
                        ExternalFileReference exref = item.GetExternalFileReference();
                        string path = ModelPathUtils.ConvertModelPathToUserVisiblePath(exref.GetAbsolutePath());
                        list.Add(path);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return list;
        }
    }
}
