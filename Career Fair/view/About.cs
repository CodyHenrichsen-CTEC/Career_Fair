using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Career_Fair.view
{
    public partial class About : Form
    {
        public About()
        {
            InitializeComponent();
            versionLabel.Text = "Current version: " + Application.ProductVersion;
        }

    }
}
