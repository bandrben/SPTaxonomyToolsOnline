using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SPTaxonomyToolsOnline
{
    public partial class AboutForm : Form
    {
        public AboutForm()
        {
            InitializeComponent();

            tbAbout.Text = @" *** Welcome to SPTaxonomyToolsOnline ***

Created by Ben Steinhauser of B&R Business Solutions.

Visit us at http://www.bandrsolutions.com!

Contact bsteinhauser@bandrsolutions.com

";

            tbAbout.AppendText(" ");

        }
    }
}
