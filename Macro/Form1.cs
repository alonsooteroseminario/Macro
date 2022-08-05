using System.Drawing;
using System.Windows.Forms;

namespace Macro
{
    public partial class Form1 : Form
    {
        public Form1(Bitmap ToShow)
        {
            InitializeComponent();

            pictureBox1.BackgroundImageLayout = ImageLayout.Stretch;
            pictureBox1.BackgroundImage = ToShow;
        }
    }
}
