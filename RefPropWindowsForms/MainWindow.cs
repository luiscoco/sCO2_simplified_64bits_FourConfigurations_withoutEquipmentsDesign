using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using sc.net;

namespace RefPropWindowsForms
{
    public partial class MainWindow : Form
    {
        public About About_dialog;
      
        public WizardOne Wizard_dialog;
        public Recompression_Brayton_Power_Cycle RCwindow;

        public PCRC PCRC_design_dialog;
        public RCMCI RCMCI_design_dialog;
        public RC_without_ReHeating RC_without_ReHeating;    
        public PCRC_without_ReHeating PCRC_without_ReHeating_Dialog;
        public RCMCI_without_ReHeating RCMCI_without_ReHeating_Dialog;

        public core CoreHX = new core();
        public RefrigerantCategory category;
        public ReferenceState referencestate;

        public String Fluids_Path_LCE;
        
        public MainWindow()
        {
            InitializeComponent();
        }

        //Recompression (RC) Brayton Power cycle Design-Point.
        public void DesignPoint_Click(object sender, EventArgs e)
        {
            //Create a new Form for the RC Design-Point
            RCwindow = new Recompression_Brayton_Power_Cycle(this);
            RCwindow.MdiParent = this;
            RCwindow.Show();

        }

        public void designPointToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PCRC_design_dialog = new PCRC();
            PCRC_design_dialog.MdiParent = this;
            PCRC_design_dialog.Show();
        }

        public void designPointToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            RCMCI_design_dialog = new RCMCI();
            RCMCI_design_dialog.MdiParent = this;
            RCMCI_design_dialog.Show();
        }

        public void designPointToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            RC_without_ReHeating = new RC_without_ReHeating();
            RC_without_ReHeating.MdiParent = this;
            RC_without_ReHeating.Show();
        }

        //PCRC without ReHeating at Design-Point
        public void designPointToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            PCRC_without_ReHeating_Dialog = new PCRC_without_ReHeating();
            PCRC_without_ReHeating_Dialog.MdiParent = this;
            PCRC_without_ReHeating_Dialog.Show();
        }

        //RCMCI without ReHeating at Design-Point
        public void designPointToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            RCMCI_without_ReHeating_Dialog = new RCMCI_without_ReHeating();
            RCMCI_without_ReHeating_Dialog.MdiParent = this;
            RCMCI_without_ReHeating_Dialog.Show();
        }

        public void sensingAnalysisToolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        public void MainWindow_Load(object sender, EventArgs e)
        {
            this.aboutToolStripMenuItem_Click(this, e);

            //Configurations_Summary_dialog = new Configurations_Summary(this);
            //Configurations_Summary_dialog.MdiParent = this;
            //Configurations_Summary_dialog.Show();

            Wizard_dialog = new WizardOne(this);
            Wizard_dialog.MdiParent = this;
            Wizard_dialog.Show();
        }

        //About window
        public void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            About_dialog = new About(this);
            About_dialog.ShowDialog();
        }

        //Wizard
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Wizard_dialog = new WizardOne(this);
            Wizard_dialog.MdiParent = this;
            Wizard_dialog.Show();
        }

        //Wizard Configurations 1-6
        private void configurations16ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Wizard_dialog = new WizardOne(this);
            Wizard_dialog.MdiParent = this;
            Wizard_dialog.Show();
        }
    }
}
