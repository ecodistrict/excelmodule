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

using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;

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
                var settings = new DataContractJsonSerializerSettings();
                settings.EmitTypeInformation = EmitTypeInformation.Never;
                MemoryStream stream1 = new MemoryStream();
                DataContractJsonSerializer ser = new DataContractJsonSerializer(typeof(InputSpecification), settings);
                InputSpecification inputSpec = new InputSpecification();

                //inputSpec.Add(new Number("a"));
                //List aList = new List("ListLbl");
                //aList.Add(new Number("b"));
                //aList.Add(new Number("c"));
                //List listRoot = new List();
                //listRoot.Add(aList);
                //listRoot.Add(aList);
                //inputSpec.Add(listRoot);
               

                var ipy = Python.CreateRuntime();
                dynamic config = ipy.UseFile("../ModuleConfig.py");
                inputSpec = config.input_specification();

                ser.WriteObject(stream1, inputSpec);
                stream1.Position = 0;
                StreamReader sr = new StreamReader(stream1);
                tBox1.Text = sr.ReadToEnd();
            }
            catch (Exception exe)
            {
                MessageBox.Show(exe.Message, "ERROR!");
            }
        }
        
    }
}
