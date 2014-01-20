using System;
using System.Windows.Forms;
using NPOI.Objects;

namespace NPOI.Example
{
    public partial class FormSample : Form
    {
        public FormSample()
        {
            InitializeComponent();
            TestModel[] locationList;
            using (var factory = ObjectFactory.CreateFactory("input.xls"))
            {
                locationList = factory.SheetToObjects<TestModel>();
                dataGridView1.DataSource = locationList;
            }
            using (var factory = DrawingFactory.CreateFactory(string.Format("{0}.xls", DateTime.Now.ToFileTimeUtc())))
            {
                factory.Draw(0, "Sheet0", locationList);
            }
        }
    }
}