using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using IronPython;
using IronPython.Hosting;

using DataTypes;

namespace IronPythonTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            try
            {
                

                var ipy = Python.CreateRuntime();
                dynamic test = ipy.UseFile("../../../../ModuleConfig.py");
                int i = test.run(1, 2);
                //List<int> lst = test.run(new Test());
            }
            catch (Exception exe)
            {
                MessageBox.Show(exe.Message, "ERROR!");
            }
        }

        private void btnModuleInfo_Click(object sender, EventArgs e)
        {
            try
            {
                InputSpecification isp = new InputSpecification();
                isp.Add(new Number("a"));
                isp.Add(new Number("b"));
                isp.Add(new Number("c"));

                tBox1.Text = isp.ToJason();

                //var ipy = Python.CreateRuntime();
                //dynamic config = ipy.UseFile("../ModuleConfig.py");
                //InputSpecification inputSpec = config.input_specification();
            }
            catch (Exception exe)
            {
                MessageBox.Show(exe.Message, "ERROR!");
            }
        }
        
    }
}
