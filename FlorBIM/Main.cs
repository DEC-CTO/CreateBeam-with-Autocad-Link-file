using Autodesk.Revit.UI;
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

namespace FlorBIM
{
    public partial class Main : Form
    {
        private RequestHandler m_Handler;
        private ExternalEvent m_ExEvent;

        public Main(ExternalEvent exEvent, RequestHandler handler)
        {
            InitializeComponent();
            m_Handler = handler;
            m_ExEvent = exEvent;
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            m_ExEvent.Dispose();
            m_ExEvent = null;
            m_Handler = null;

            base.OnFormClosed(e);
        }

        private void EnableCommands(bool status)
        {
            foreach (Control ctrl in this.Controls)
            {
                ctrl.Enabled = status;
            }
        }

        private void MakeRequest(RequestId request)
        {
            m_Handler.Request.Make(request);
            m_ExEvent.Raise();
            Main main = new Main();
            main.DozeOff();
        }

        private void DozeOff()
        {
            EnableCommands(false);
        }
        public void WakeUp()
        {
            EnableCommands(true);
        }

        public Main()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MakeRequest(RequestId.count);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MakeRequest(RequestId.CreateBeam);
        }
    }
}
