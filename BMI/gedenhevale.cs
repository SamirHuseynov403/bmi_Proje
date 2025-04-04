using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BMI
{
    public partial class gedenhevaleyeni : Form
    {
        Hevale fghev;
        Pul_Kocurmesi fgh;

        public gedenhevaleyeni(Hevale fd)
        {
            InitializeComponent();

            this.fghev = fd;
            //this.fgsef = fs;


        }
        public gedenhevaleyeni(Pul_Kocurmesi fs)
        {
            InitializeComponent();
            this.fgh = fs;
        }

        public gedenhevaleyeni()
        {
            // TODO: Complete member initialization
        }
        public string ghhesNo, ghSAA, ghmebleg, ghalanbank, ghvalyuta, ghMhesab, ghgonderen;

        private void gedenhevale_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}
