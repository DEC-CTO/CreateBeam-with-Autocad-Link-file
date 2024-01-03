﻿using Autodesk.Revit.ApplicationServices;
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
    public class RequestHandler : IExternalEventHandler
    {
        public static UIDocument m_uidoc = null;
        public static UIApplication m_uiapp = null;
        public static Document m_doc = null;

        private Request m_request = new Request();
        public Request Request
        {
            get { return m_request; }
        }
        public String GetName()
        {
            return "testing";
        }

        public void Execute(UIApplication uiapp)
        {
            try
            {
                switch (Request.Take())
                {
                    case RequestId.count:
                        {
                            _count(uiapp);
                            break;
                        }
                    case RequestId.CreateBeam:
                        {
                            _CreateBeam(uiapp);
                            break;
                        }
                }
            }
            finally
            {
                App.thisApp.WakeFormUp();
            }
            return;
        }

        private void _count(UIApplication uiapp)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            m_uidoc = uidoc;
            m_doc = m_uidoc.Document;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.Creation.Application CreationApp = app.Create;

            List<string> filepath = Lib.GetDWGLink(m_doc);

            CadDocument doccad = DwgReader.Read(filepath[0]);
            List<Entity> all = new List<Entity>(doccad.Entities);

            List<string> BeamName = new List<string>();
            foreach (Entity ent in all)
            {
                ACadSharp.ObjectType obj = ent.ObjectType;
                if (obj == ACadSharp.ObjectType.TEXT)
                {
                    TextEntity tn = ent as TextEntity;
                    if (tn != null)
                    {
                        string name = tn.Value;
                        BeamName.Add(name);
                    }
                }
            }

            BeamName = BeamName.Distinct().ToList();
            FamilySymbol fs = new FilteredElementCollector(m_doc).OfCategory(BuiltInCategory.OST_StructuralFraming).OfClass(typeof(FamilySymbol)).FirstElement() as FamilySymbol;
            if (fs != null)
            {
                using (Transaction tran = new Transaction(m_doc, "Insert Beam Type"))
                {
                    tran.Start();
                    foreach (string item in BeamName)
                    {
                        fs.Duplicate(item);
                    }
                    tran.Commit();
                }
            }
        }

        private void _CreateBeam(UIApplication uiapp)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            m_uidoc = uidoc;
            m_doc = m_uidoc.Document;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Autodesk.Revit.Creation.Application CreationApp = app.Create;


            List<string> filepath = Lib.GetDWGLink(m_doc);
            CadDocument doccad = DwgReader.Read(filepath[0]);
            List<Entity> all = new List<Entity>(doccad.Entities);

            List<Autodesk.Revit.DB.Line> longCurves = new List<Autodesk.Revit.DB.Line>();
            Dictionary<XYZ, string> dix = new Dictionary<XYZ, string>();

            foreach (Entity item in all)
            {
                ACadSharp.ObjectType obj = item.ObjectType;
                if(obj == ACadSharp.ObjectType.TEXT)
                {
                    TextEntity tn = item as TextEntity;
                    if(tn != null)
                    {
                        string textname = tn.Value;
                        XYZ p1 = new XYZ(tn.InsertPoint.X, tn.InsertPoint.Y, tn.InsertPoint.Z)/304.8;
                        dix.Add(p1, textname);
                    }
                }

                else if(obj == ACadSharp.ObjectType.LINE)
                {
                    ACadSharp.Entities.Line line = item as ACadSharp.Entities.Line;
                    Autodesk.Revit.DB.Line getline = CovertRevitLine(line);
                    if (getline == null) continue;
                    else if(getline != null)
                    {
                        if(getline.Length * 304.8 > 801)
                        {
                            longCurves.Add(getline);
                        }
                    }
                }
            }

            foreach (Autodesk.Revit.DB.Line item in longCurves)
            {
                XYZ mid = item.Evaluate(0.5, true);
                var sort = from n in dix orderby n.Key.DistanceTo(mid) ascending select n;
                string symbolname = sort.First().Value;

                FamilySymbol fs = new FilteredElementCollector(m_doc).OfCategory(BuiltInCategory.OST_StructuralFraming)
                    .OfClass(typeof(FamilySymbol))
                    .FirstOrDefault(q => q.Name == symbolname) as FamilySymbol;

                if (fs == null) continue;

                using(Transaction trans = new Transaction(m_doc, "createBeam"))
                {
                    trans.Start();
                    if(fs != null)
                    {
                        fs.Activate();
                        FamilyInstance fi = m_doc.Create.NewFamilyInstance(item, fs, m_doc.ActiveView.GenLevel, StructuralType.Beam);
                    }
                    trans.Commit();
                }
            }
        }

        public static Autodesk.Revit.DB.Line CovertRevitLine(ACadSharp.Entities.Line line)
        {
            Autodesk.Revit.DB.Line revitline = null;

            try
            {
                CSMath.XYZ p1 = line.StartPoint;
                CSMath.XYZ p2 = line.EndPoint;

                XYZ pp1 = new XYZ(p1.X, p1.Y, p1.Z) / 304.8;
                XYZ pp2 = new XYZ(p2.X, p2.Y, p2.Z) / 304.8;

                revitline = Autodesk.Revit.DB.Line.CreateBound(pp1, pp2);

            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message);
            }

            return revitline;
        }
    }
}