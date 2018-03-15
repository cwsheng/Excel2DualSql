using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WfApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private Bitmap backMap;
        private void button1_Click(object sender, EventArgs e)
        {
            using (PictureBox pic = new PictureBox())
            {
                pic.BackColor = this.BackColor;
                pic.Size = this.Size;
                //if (backMap!=null)
                //    backMap.Dispose();
                backMap = new Bitmap(2000, 2000);
                pic.DrawToBitmap(backMap, new Rectangle(Point.Empty, this.Size));  // 画pictureBox1显示的图，假定它没有边框
                using (Graphics g = Graphics.FromImage(backMap))
                {
                }
            }
        }
    }
}
