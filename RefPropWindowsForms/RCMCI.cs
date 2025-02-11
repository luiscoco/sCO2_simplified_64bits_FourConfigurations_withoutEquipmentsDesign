using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.IO;
using System.Diagnostics;
using System.Reflection;

using System.Data.Common;
using System.Threading;
using System.Text;

using sc.net;

using Excel = Microsoft.Office.Interop.Excel;

namespace RefPropWindowsForms
{
    public partial class RCMCI : Form
    {
        public double MixtureCriticalPressure = 0.0;
        public double MixtureCriticalTemperature = 0.0;

        public core luis = new core();

        //First calculate the Main Compressor Rotational speed and after send that value to the Turbines
        public Double N_design_Main_Compressor;

        //Input Data:
        public RefrigerantCategory category;
        public ReferenceState referencestate;
        public Int64 Error_code;
        public core.RecompCycle recomp_cycle = new core.RecompCycle();

        public Double wmm;
        
        public RCMCI()
        {
            InitializeComponent();
        }

        public Double specific_work_main_turbine = 0;
        public Double specific_work_reheating_turbine = 0;
        public Double specific_work_compressor1 = 0;
        public Double specific_work_compressor2 = 0;
        public Double specific_work_compressor3 = 0;
        public Double Miscellanous_Auxiliaries = 0;
        public Double Total_Auxiliaries = 0;

        public Double w_dot_net2;
        public Double t_mc1_in2,t_mc2_in2;
        public Double t_t_in2;
        public Double p_rhx_in2;
        public Double t_rht_in2;
        public Double ua_lt2, ua_ht2;
        public Double eta1_mc2;
        public Double eta2_mc2;
        public Double eta_rc2;
        public Double eta_t2;
        public Double eta_trh2;
        public Int64 n_sub_hxrs2;
        public Double p_mc1_in2;
        public Double p_mc1_out2;
        public Double p_mc2_in2;
        public Double p_mc2_out2;
        public Double recomp_frac2;
        public Double tol2;
        public Double eta_thermal2;

        public Double dp2_lt1, dp2_lt2;
        public Double dp2_ht1, dp2_ht2;

        public Double dp11_pc1,dp11_pc2;
        public Double dp12_pc1, dp12_pc2;

        public Double dp2_phx1, dp2_phx2;
        public Double dp2_rhx1, dp2_rhx2;

        public Double temp21;
        public Double temp22;
        public Double temp23;
        public Double temp24;
        public Double temp25;
        public Double temp26;
        public Double temp27;
        public Double temp28;
        public Double temp29;
        public Double temp210;
        public Double temp211;
        public Double temp212;
        public Double temp213;
        public Double temp214;

        public Double pres21;
        public Double pres22;
        public Double pres23;
        public Double pres24;
        public Double pres25;
        public Double pres26;
        public Double pres27;
        public Double pres28;
        public Double pres29;
        public Double pres210;
        public Double pres211;
        public Double pres212;
        public Double pres213;
        public Double pres214;

        public Double enth21;
        public Double enth22;
        public Double enth23;
        public Double enth24;
        public Double enth25;
        public Double enth26;
        public Double enth27;
        public Double enth28;
        public Double enth29;
        public Double enth210;
        public Double enth211;
        public Double enth212;
        public Double enth213;
        public Double enth214;

        public Double entr21;
        public Double entr22;
        public Double entr23;
        public Double entr24;
        public Double entr25;
        public Double entr26;
        public Double entr27;
        public Double entr28;
        public Double entr29;
        public Double entr210;
        public Double entr211;
        public Double entr212;
        public Double entr213;
        public Double entr214;

        public Double massflow2;
        public Double LT_mdoth, LT_mdotc, LT_Tcin, LT_Thin, LT_Pcin, LT_Phin;

        public Double LT_Pcout, LT_Phout, LT_Q, HT_mdoth, HT_mdotc, HT_Tcin, HT_Thin;
        public Double HT_Pcin, HT_Phin, HT_Pcout, HT_Phout, HT_Q, LT_UA, HT_UA;
        public Double LT_Effc, HT_Effc;

        public Double PHX1,RHX1,PC11,PC21;

        private void button1_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        //Mixtures Calculation Button
        private void button2_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == "PureFluid")
            {
                category = RefrigerantCategory.PureFluid;
            }
            if (comboBox1.Text == "PredefinedMixture")
            {
                category = RefrigerantCategory.PredefinedMixture;
            }
            if (comboBox1.Text == "NewMixture")
            {
                category = RefrigerantCategory.NewMixture;
            }
            if (comboBox1.Text == "PseudoPureFluid")
            {
                category = RefrigerantCategory.PseudoPureFluid;
            }

            if (comboBox3.Text == "DEF")
            {
                referencestate = ReferenceState.DEF;
            }
            if (comboBox3.Text == "ASH")
            {
                referencestate = ReferenceState.ASH;
            }
            if (comboBox3.Text == "IIR")
            {
                referencestate = ReferenceState.IIR;
            }
            if (comboBox3.Text == "NBP")
            {
                referencestate = ReferenceState.NBP;
            }

            luis.core1(this.comboBox2.Text, category);
            luis.working_fluid.Category = category;
            luis.working_fluid.reference = referencestate;

            //Store Input Data from Graphical User Interface GUI into variables
            w_dot_net2 = Convert.ToDouble(textBox1.Text);
            t_mc1_in2 = Convert.ToDouble(textBox2.Text);
            t_mc2_in2 = Convert.ToDouble(textBox28.Text);
            t_t_in2 = Convert.ToDouble(textBox4.Text);
            p_mc1_in2 = Convert.ToDouble(textBox3.Text);
            p_mc1_out2 = Convert.ToDouble(textBox8.Text);
            p_mc2_in2 = Convert.ToDouble(textBox23.Text);
            p_mc2_out2 = Convert.ToDouble(textBox22.Text);            
            p_rhx_in2 = Convert.ToDouble(textBox7.Text);
            t_rht_in2 = Convert.ToDouble(textBox6.Text);
            ua_lt2 = Convert.ToDouble(textBox17.Text);
            ua_ht2 = Convert.ToDouble(textBox16.Text);
            recomp_frac2 = Convert.ToDouble(textBox15.Text);
            eta1_mc2 = Convert.ToDouble(textBox14.Text);
            eta2_mc2 = Convert.ToDouble(textBox27.Text);
            eta_rc2 = Convert.ToDouble(textBox13.Text);
            eta_t2 = Convert.ToDouble(textBox19.Text);
            eta_trh2 = Convert.ToDouble(textBox18.Text);
            n_sub_hxrs2 = Convert.ToInt64(textBox20.Text);
            tol2 = Convert.ToDouble(textBox21.Text);
            dp2_lt1 = Convert.ToDouble(textBox5.Text);
            dp2_ht1 = Convert.ToDouble(textBox12.Text);
            dp12_pc1 = Convert.ToDouble(textBox11.Text);
            dp12_pc2 = Convert.ToDouble(textBox24.Text);
            dp2_phx1 = Convert.ToDouble(textBox10.Text);
            dp2_rhx1 = Convert.ToDouble(textBox9.Text);
            dp2_lt2 = Convert.ToDouble(textBox26.Text);
            dp2_ht2 = Convert.ToDouble(textBox25.Text);         
        }

        //RESET Button
        private void button5_Click(object sender, EventArgs e)
        {
            textBox47.Text = "";
            textBox46.Text = "";
            textBox45.Text = "";
            textBox44.Text = "";
            textBox43.Text = "";
            textBox42.Text = "";
            textBox39.Text = "";
            textBox38.Text = "";
            textBox37.Text = "";
            textBox29.Text = "";
            textBox57.Text = "";
            textBox56.Text = "";
            textBox65.Text = "";
            textBox64.Text = "";

            textBox63.Text = "";
            textBox62.Text = "";
            textBox61.Text = "";
            textBox60.Text = "";
            textBox55.Text = "";
            textBox54.Text = "";
            textBox53.Text = "";
            textBox52.Text = "";
            textBox41.Text = "";
            textBox40.Text = "";
            textBox59.Text = "";
            textBox58.Text = "";
            textBox67.Text = "";
            textBox66.Text = "";

            textBox48.Text = "";
            textBox49.Text = "";
            textBox50.Text = "";

            //w_dot_net2 = Convert.ToDouble(textBox1.Text);
            textBox1.Text = "50000";
            //t_mc1_in2 = Convert.ToDouble(textBox2.Text);
            textBox2.Text = "305.15";
            //t_mc2_in2 = Convert.ToDouble(textBox28.Text);
            textBox28.Text = "305.15";
            //t_t_in2 = Convert.ToDouble(textBox4.Text);
            //p_mc1_in2 = Convert.ToDouble(textBox3.Text);
            textBox3.Text = "7400";
            //p_mc1_out2 = Convert.ToDouble(textBox8.Text);
            textBox8.Text = "10300";
            //p_mc2_in2 = Convert.ToDouble(textBox23.Text);
            textBox23.Text = "10300";
            //p_mc2_out2 = Convert.ToDouble(textBox22.Text);
            textBox22.Text = "25000";
            //p_rhx_in2 = Convert.ToDouble(textBox7.Text);
            textBox7.Text = "17300";
            //t_rht_in2 = Convert.ToDouble(textBox6.Text);
            //ua_lt2 = Convert.ToDouble(textBox17.Text);
            textBox17.Text = "5000";
            //ua_ht2 = Convert.ToDouble(textBox16.Text);
            textBox16.Text = "5000";
            //recomp_frac2 = Convert.ToDouble(textBox15.Text);
            textBox15.Text = "0.25";
            //eta1_mc2 = Convert.ToDouble(textBox14.Text);
            textBox14.Text = "0.89";
            //eta2_mc2 = Convert.ToDouble(textBox27.Text);
            textBox27.Text = "0.89";
            //eta_rc2 = Convert.ToDouble(textBox13.Text);
            textBox13.Text = "0.89";
            //eta_t2 = Convert.ToDouble(textBox19.Text);
            textBox19.Text = "0.93";
            //eta_trh2 = Convert.ToDouble(textBox18.Text);
            textBox18.Text = "0.93";
            //n_sub_hxrs2 = Convert.ToInt64(textBox20.Text);
            textBox20.Text = "15";
            //tol2 = Convert.ToDouble(textBox21.Text);
            textBox21.Text = "0.00001";
            //dp2_lt1 = Convert.ToDouble(textBox5.Text);
            textBox5.Text = "0.0";
            //dp2_ht1 = Convert.ToDouble(textBox12.Text);
            textBox12.Text = "0.0";
            //dp12_pc1 = Convert.ToDouble(textBox11.Text);
            textBox11.Text = "0.0";
            //dp12_pc2 = Convert.ToDouble(textBox24.Text);
            textBox24.Text = "0.0";
            //dp2_phx1 = Convert.ToDouble(textBox10.Text);
            textBox10.Text = "0.0";
            //dp2_rhx1 = Convert.ToDouble(textBox9.Text);
            textBox9.Text = "0.0";
            //dp2_lt2 = Convert.ToDouble(textBox26.Text);
            textBox26.Text = "0.0";
            //dp2_ht2 = Convert.ToDouble(textBox25.Text);
            textBox25.Text = "0.0";
        }

        //Mixtures Calculation
        private void button20_Click(object sender, EventArgs e)
        {
            int maxIterations = 5;
            int numIterations = 0;


            //PureFluid
            if (comboBox1.Text == "PureFluid")
            {
                category = RefrigerantCategory.PureFluid;
                luis.core1(this.comboBox1.Text, category);
            }

            //NewMixture
            if (comboBox1.Text == "NewMixture")
            {
                category = RefrigerantCategory.NewMixture;
                luis.core1(this.comboBox2.Text + "=" + textBox35.Text + "," +
                           this.comboBox16.Text + "=" + textBox36.Text + "," +
                           this.comboBox18.Text + "=" + textBox69.Text + "," +
                           this.comboBox17.Text + "=" + textBox70.Text, category);
            }

            if (comboBox1.Text == "PredefinedMixture")
            {
                category = RefrigerantCategory.PredefinedMixture;
            }

            if (comboBox1.Text == "PseudoPureFluid")
            {
                category = RefrigerantCategory.PseudoPureFluid;
            }

            if (comboBox3.Text == "DEF")
            {
                referencestate = ReferenceState.DEF;
            }
            if (comboBox3.Text == "ASH")
            {
                referencestate = ReferenceState.ASH;
            }
            if (comboBox3.Text == "IIR")
            {
                referencestate = ReferenceState.IIR;
            }
            if (comboBox3.Text == "NBP")
            {
                referencestate = ReferenceState.NBP;
            }

            luis.working_fluid.Category = category;
            luis.working_fluid.reference = referencestate;

            luis.wmm = luis.working_fluid.MolecularWeight;

            //Store Input Data from Graphical User Interface GUI into variables
            w_dot_net2 = Convert.ToDouble(textBox1.Text);
            t_mc1_in2 = Convert.ToDouble(textBox2.Text);
            t_mc2_in2 = Convert.ToDouble(textBox28.Text);
            t_t_in2 = Convert.ToDouble(textBox4.Text);
            p_mc1_in2 = Convert.ToDouble(textBox3.Text);
            p_mc1_out2 = Convert.ToDouble(textBox8.Text);
            p_mc2_in2 = Convert.ToDouble(textBox23.Text);
            p_mc2_out2 = Convert.ToDouble(textBox22.Text);
            p_rhx_in2 = Convert.ToDouble(textBox7.Text);
            t_rht_in2 = Convert.ToDouble(textBox6.Text);
            ua_lt2 = Convert.ToDouble(textBox17.Text);
            ua_ht2 = Convert.ToDouble(textBox16.Text);
            recomp_frac2 = Convert.ToDouble(textBox15.Text);
            eta1_mc2 = Convert.ToDouble(textBox14.Text);
            eta2_mc2 = Convert.ToDouble(textBox27.Text);
            eta_rc2 = Convert.ToDouble(textBox13.Text);
            eta_t2 = Convert.ToDouble(textBox19.Text);
            eta_trh2 = Convert.ToDouble(textBox18.Text);
            n_sub_hxrs2 = Convert.ToInt64(textBox20.Text);
            tol2 = Convert.ToDouble(textBox21.Text);
            dp2_lt1 = Convert.ToDouble(textBox5.Text);
            dp2_ht1 = Convert.ToDouble(textBox12.Text);
            dp11_pc1 = Convert.ToDouble(textBox11.Text);
            dp11_pc2 = Convert.ToDouble(textBox24.Text);
            dp2_phx1 = Convert.ToDouble(textBox10.Text);
            dp2_rhx1 = Convert.ToDouble(textBox9.Text);
            dp2_lt2 = Convert.ToDouble(textBox26.Text);
            dp2_ht2 = Convert.ToDouble(textBox25.Text);

            core.RCMCIwithReheating cicloRCMCI_withRH = new core.RCMCIwithReheating();

            luis.RecompCycle_RCMCI_with_Reheating(luis, ref cicloRCMCI_withRH, w_dot_net2,
            t_mc2_in2, t_t_in2, t_rht_in2, p_rhx_in2, p_mc2_in2, p_mc2_out2, p_mc1_in2, t_mc1_in2, p_mc1_out2,
            ua_lt2, ua_ht2, eta2_mc2, eta_rc2, eta1_mc2, eta_t2, eta_trh2, n_sub_hxrs2,
            recomp_frac2, tol2, eta_thermal2, -dp2_lt1, -dp2_lt2, -dp2_ht1, -dp2_ht2,
            -dp11_pc1, -dp12_pc1, -dp2_phx1, -dp2_phx2, -dp2_rhx1, -dp2_rhx2, -dp11_pc2, -dp12_pc2);

            massflow2 = cicloRCMCI_withRH.m_dot_turbine;
            w_dot_net2 = cicloRCMCI_withRH.W_dot_net;
            eta_thermal2 = cicloRCMCI_withRH.eta_thermal;
            recomp_frac2 = cicloRCMCI_withRH.recomp_frac;                       

            temp21 = cicloRCMCI_withRH.temp[0];
            temp22 = cicloRCMCI_withRH.temp[1];
            temp23 = cicloRCMCI_withRH.temp[2];
            temp24 = cicloRCMCI_withRH.temp[3];
            temp25 = cicloRCMCI_withRH.temp[4];
            temp26 = cicloRCMCI_withRH.temp[5];
            temp27 = cicloRCMCI_withRH.temp[6];
            temp28 = cicloRCMCI_withRH.temp[7];
            temp29 = cicloRCMCI_withRH.temp[8];
            temp210 = cicloRCMCI_withRH.temp[9];
            temp211 = cicloRCMCI_withRH.temp[10];
            temp212 = cicloRCMCI_withRH.temp[11];
            temp213 = cicloRCMCI_withRH.temp[12];
            temp214 = cicloRCMCI_withRH.temp[13];

            pres21 = cicloRCMCI_withRH.pres[0];
            pres22 = cicloRCMCI_withRH.pres[1];
            pres23 = cicloRCMCI_withRH.pres[2];
            pres24 = cicloRCMCI_withRH.pres[3];
            pres25 = cicloRCMCI_withRH.pres[4];
            pres26 = cicloRCMCI_withRH.pres[5];
            pres27 = cicloRCMCI_withRH.pres[6];
            pres28 = cicloRCMCI_withRH.pres[7];
            pres29 = cicloRCMCI_withRH.pres[8];
            pres210 = cicloRCMCI_withRH.pres[9];
            pres211 = cicloRCMCI_withRH.pres[10];
            pres212 = cicloRCMCI_withRH.pres[11];
            pres213 = cicloRCMCI_withRH.pres[12];
            pres214 = cicloRCMCI_withRH.pres[13];

            textBox47.Text = Convert.ToString(temp21);
            textBox46.Text = Convert.ToString(temp22);
            textBox45.Text = Convert.ToString(temp23);
            textBox44.Text = Convert.ToString(temp24);
            textBox43.Text = Convert.ToString(temp25);
            textBox42.Text = Convert.ToString(temp26);
            textBox39.Text = Convert.ToString(temp27);
            textBox38.Text = Convert.ToString(temp28);
            textBox37.Text = Convert.ToString(temp29);
            textBox29.Text = Convert.ToString(temp210);
            textBox57.Text = Convert.ToString(temp211);
            textBox56.Text = Convert.ToString(temp212);
            textBox65.Text = Convert.ToString(temp213);
            textBox64.Text = Convert.ToString(temp214);

            textBox63.Text = Convert.ToString(pres21);
            textBox62.Text = Convert.ToString(pres22);
            textBox61.Text = Convert.ToString(pres23);
            textBox60.Text = Convert.ToString(pres24);
            textBox55.Text = Convert.ToString(pres25);
            textBox54.Text = Convert.ToString(pres26);
            textBox53.Text = Convert.ToString(pres27);
            textBox52.Text = Convert.ToString(pres28);
            textBox41.Text = Convert.ToString(pres29);
            textBox40.Text = Convert.ToString(pres210);
            textBox59.Text = Convert.ToString(pres211);
            textBox58.Text = Convert.ToString(pres212);
            textBox67.Text = Convert.ToString(pres213);
            textBox66.Text = Convert.ToString(pres214);

            String point1_state, point2_state, point3_state, point4_state, point5_state, point6_state;
            String point7_state, point8_state, point9_state, point10_state, point11_state, point12_state, point13_state, point14_state;

            luis.working_fluid.FindStateWithTP(temp21, pres21);
            enth21 = luis.working_fluid.Enthalpy;
            entr21 = luis.working_fluid.Entropy;

            luis.working_fluid.FindStateWithTP(temp22, pres22);
            enth22 = luis.working_fluid.Enthalpy;
            entr22 = luis.working_fluid.Entropy;

            luis.working_fluid.FindStateWithTP(temp23, pres23);
            enth23 = luis.working_fluid.Enthalpy;
            entr23 = luis.working_fluid.Entropy;

            luis.working_fluid.FindStateWithTP(temp24, pres24);
            enth24 = luis.working_fluid.Enthalpy;
            entr24 = luis.working_fluid.Entropy;

            luis.working_fluid.FindStateWithTP(temp25, pres25);
            enth25 = luis.working_fluid.Enthalpy;
            entr25 = luis.working_fluid.Entropy;

            luis.working_fluid.FindStateWithTP(temp26, pres26);
            enth26 = luis.working_fluid.Enthalpy;
            entr26 = luis.working_fluid.Entropy;

            luis.working_fluid.FindStateWithTP(temp27, pres27);
            enth27 = luis.working_fluid.Enthalpy;
            entr27 = luis.working_fluid.Entropy;

            luis.working_fluid.FindStateWithTP(temp28, pres28);
            enth28 = luis.working_fluid.Enthalpy;
            entr28 = luis.working_fluid.Entropy;

            luis.working_fluid.FindStateWithTP(temp29, pres29);
            enth29 = luis.working_fluid.Enthalpy;
            entr29 = luis.working_fluid.Entropy;

            luis.working_fluid.FindStateWithTP(temp210, pres210);
            enth210 = luis.working_fluid.Enthalpy;
            entr210 = luis.working_fluid.Entropy;

            luis.working_fluid.FindStateWithTP(temp211, pres211);
            enth211 = luis.working_fluid.Enthalpy;
            entr211 = luis.working_fluid.Entropy;

            luis.working_fluid.FindStateWithTP(temp212, pres212);
            enth212 = luis.working_fluid.Enthalpy;
            entr212 = luis.working_fluid.Entropy;

            luis.working_fluid.FindStateWithTP(temp213, pres213);
            enth213 = luis.working_fluid.Enthalpy;
            entr213 = luis.working_fluid.Entropy;

            luis.working_fluid.FindStateWithTP(temp214, pres214);
            enth214 = luis.working_fluid.Enthalpy;
            entr214 = luis.working_fluid.Entropy;

            point1_state = "Pressure (kPa):" + Convert.ToString(pres21) + Environment.NewLine +
                          "Temperature (K):" + Convert.ToString(temp21) + Environment.NewLine +
                          "Entalphy (kJ/kg):" + Convert.ToString(enth21) + Environment.NewLine +
                          "Entrophy (kJ/kg K):" + Convert.ToString(entr21) + Environment.NewLine;

            point2_state = "Pressure (kPa):" + Convert.ToString(pres22) + Environment.NewLine +
                         "Temperature (K):" + Convert.ToString(temp22) + Environment.NewLine +
                         "Entalphy (kJ/kg):" + Convert.ToString(enth22) + Environment.NewLine +
                         "Entrophy (kJ/kg K):" + Convert.ToString(entr22) + Environment.NewLine;

            point3_state = "Pressure (kPa):" + Convert.ToString(pres23) + Environment.NewLine +
                      "Temperature (K):" + Convert.ToString(temp23) + Environment.NewLine +
                      "Entalphy (kJ/kg):" + Convert.ToString(enth23) + Environment.NewLine +
                      "Entrophy (kJ/kg K):" + Convert.ToString(entr23) + Environment.NewLine;

            point4_state = "Pressure (kPa):" + Convert.ToString(pres24) + Environment.NewLine +
                      "Temperature (K):" + Convert.ToString(temp24) + Environment.NewLine +
                      "Entalphy (kJ/kg):" + Convert.ToString(enth24) + Environment.NewLine +
                      "Entrophy (kJ/kg K):" + Convert.ToString(entr24) + Environment.NewLine;

            point5_state = "Pressure (kPa):" + Convert.ToString(pres25) + Environment.NewLine +
                      "Temperature (K):" + Convert.ToString(temp25) + Environment.NewLine +
                      "Entalphy (kJ/kg):" + Convert.ToString(enth25) + Environment.NewLine +
                      "Entrophy (kJ/kg K):" + Convert.ToString(entr25) + Environment.NewLine;

            point6_state = "Pressure (kPa):" + Convert.ToString(pres26) + Environment.NewLine +
                      "Temperature (K):" + Convert.ToString(temp26) + Environment.NewLine +
                      "Entalphy (kJ/kg):" + Convert.ToString(enth26) + Environment.NewLine +
                      "Entrophy (kJ/kg K):" + Convert.ToString(entr26) + Environment.NewLine;

            point7_state = "Pressure (kPa):" + Convert.ToString(pres27) + Environment.NewLine +
                      "Temperature (K):" + Convert.ToString(temp27) + Environment.NewLine +
                      "Entalphy (kJ/kg):" + Convert.ToString(enth27) + Environment.NewLine +
                      "Entrophy (kJ/kg K):" + Convert.ToString(entr27) + Environment.NewLine;

            point8_state = "Pressure (kPa):" + Convert.ToString(pres28) + Environment.NewLine +
                      "Temperature (K):" + Convert.ToString(temp28) + Environment.NewLine +
                      "Entalphy (kJ/kg):" + Convert.ToString(enth28) + Environment.NewLine +
                      "Entrophy (kJ/kg K):" + Convert.ToString(entr28) + Environment.NewLine;

            point9_state = "Pressure (kPa):" + Convert.ToString(pres29) + Environment.NewLine +
                     "Temperature (K):" + Convert.ToString(temp29) + Environment.NewLine +
                     "Entalphy (kJ/kg):" + Convert.ToString(enth29) + Environment.NewLine +
                     "Entrophy (kJ/kg K):" + Convert.ToString(entr29) + Environment.NewLine;

            point10_state = "Pressure (kPa):" + Convert.ToString(pres210) + Environment.NewLine +
                      "Temperature (K):" + Convert.ToString(temp210) + Environment.NewLine +
                      "Entalphy (kJ/kg):" + Convert.ToString(enth210) + Environment.NewLine +
                      "Entrophy (kJ/kg K):" + Convert.ToString(entr210) + Environment.NewLine;

            point11_state = "Pressure (kPa):" + Convert.ToString(pres211) + Environment.NewLine +
                     "Temperature (K):" + Convert.ToString(temp211) + Environment.NewLine +
                     "Entalphy (kJ/kg):" + Convert.ToString(enth211) + Environment.NewLine +
                     "Entrophy (kJ/kg K):" + Convert.ToString(entr211) + Environment.NewLine;

            point12_state = "Pressure (kPa):" + Convert.ToString(pres212) + Environment.NewLine +
                      "Temperature (K):" + Convert.ToString(temp212) + Environment.NewLine +
                      "Entalphy (kJ/kg):" + Convert.ToString(enth212) + Environment.NewLine +
                      "Entrophy (kJ/kg K):" + Convert.ToString(entr212) + Environment.NewLine;

            point13_state = "Pressure (kPa):" + Convert.ToString(pres213) + Environment.NewLine +
                      "Temperature (K):" + Convert.ToString(temp213) + Environment.NewLine +
                      "Entalphy (kJ/kg):" + Convert.ToString(enth213) + Environment.NewLine +
                      "Entrophy (kJ/kg K):" + Convert.ToString(entr213) + Environment.NewLine;

            point14_state = "Pressure (kPa):" + Convert.ToString(pres214) + Environment.NewLine +
                      "Temperature (K):" + Convert.ToString(temp214) + Environment.NewLine +
                      "Entalphy (kJ/kg):" + Convert.ToString(enth214) + Environment.NewLine +
                      "Entrophy (kJ/kg K):" + Convert.ToString(entr214) + Environment.NewLine;

            textBox48.Text = Convert.ToString(w_dot_net2);
            textBox49.Text = Convert.ToString(massflow2);
            textBox50.Text = Convert.ToString(eta_thermal2 * 100);

            //toolTip1.SetToolTip(label55, point1_state);
            //toolTip2.SetToolTip(label57, point2_state);
            //toolTip3.SetToolTip(label59, point3_state);
            //toolTip4.SetToolTip(label60, point4_state);
            //toolTip5.SetToolTip(label61, point5_state);
            //toolTip6.SetToolTip(label62, point6_state);
            //toolTip7.SetToolTip(label63, point7_state);
            //toolTip8.SetToolTip(label65, point8_state);
            //toolTip9.SetToolTip(label66, point9_state);
            //toolTip10.SetToolTip(label67, point10_state);
            //toolTip11.SetToolTip(label64, point11_state);
            //toolTip12.SetToolTip(label68, point12_state);
                        
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "PureFluid")
            {
                comboBox16.Enabled = false;
                comboBox17.Enabled = false;
                comboBox18.Enabled = false;
                textBox35.Enabled = false;
                textBox36.Enabled = false;
                textBox69.Enabled = false;
                textBox70.Enabled = false;
                button20.Enabled = false;               
            }

            else if (comboBox1.Text == "NewMixture")
            {
                comboBox16.Enabled = true;
                comboBox17.Enabled = true;
                comboBox18.Enabled = true;
                textBox35.Enabled = true;
                textBox36.Enabled = true;
                textBox69.Enabled = true;
                textBox70.Enabled = true;
                button20.Enabled = true;

                Refrigerant working_fluid = new Refrigerant(RefrigerantCategory.NewMixture, this.comboBox2.Text + "=" + textBox35.Text + "," + this.comboBox16.Text + "=" + textBox36.Text, ReferenceState.DEF);

                textBox34.Text = Convert.ToString(working_fluid.CriticalPressure);
                textBox68.Text = Convert.ToString(working_fluid.CriticalTemperature);
                textBox51.Text = Convert.ToString(working_fluid.CriticalDensity);

                MixtureCriticalPressure = working_fluid.CriticalPressure;
                MixtureCriticalTemperature = working_fluid.CriticalTemperature;
            }
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        //Set critical conditions
        private void button31_Click(object sender, EventArgs e)
        {
            double option1 = 0.0;
            double option2 = 0.0;
            double option3 = 0.0;
            double option4 = 0.0;

            option1 = Convert.ToDouble(this.textBox35.Text);
            option2 = Convert.ToDouble(this.textBox36.Text);
            option3 = Convert.ToDouble(this.textBox69.Text);
            option4 = Convert.ToDouble(this.textBox70.Text);

            if ((option1 == 1) || (option2 == 1) || (option3 == 1) || (option4 == 1))
            {
                Refrigerant working_fluid = new Refrigerant(RefrigerantCategory.NewMixture, this.comboBox2.Text + "=" +
                textBox35.Text + "," + this.comboBox16.Text + "=" + textBox36.Text + "," + this.comboBox18.Text + "=" +
                textBox69.Text + "," + this.comboBox17.Text + "=" + textBox70.Text, ReferenceState.DEF);                               

                textBox34.Text = Convert.ToString(working_fluid.CriticalPressure);
                textBox68.Text = Convert.ToString(working_fluid.CriticalTemperature);
                textBox51.Text = Convert.ToString(working_fluid.CriticalDensity);

                textBox3.Text = Convert.ToString(working_fluid.CriticalPressure);
                textBox2.Text = Convert.ToString(working_fluid.CriticalTemperature);
                textBox28.Text = Convert.ToString(working_fluid.CriticalTemperature);

                MixtureCriticalPressure = working_fluid.CriticalPressure;
                MixtureCriticalTemperature = working_fluid.CriticalTemperature;
            }

            else
            {
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Excel.Application();

                xlWorkBook = xlApp.Workbooks.Open("C:\\SCSP-simplified-copia3\\RefPropWindowsForms\\bin\\x64\\Debug\\REFPROP.xls");

                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(9);

                //Fluids selection
                xlWorkSheet.Cells[13, 6] = this.comboBox2.Text;
                xlWorkSheet.Cells[14, 6] = this.comboBox16.Text;
                xlWorkSheet.Cells[15, 6] = this.comboBox18.Text;
                xlWorkSheet.Cells[16, 6] = this.comboBox17.Text;

                // % Compositions
                xlWorkSheet.Cells[13, 7] = this.textBox35.Text;
                xlWorkSheet.Cells[14, 7] = this.textBox36.Text;
                xlWorkSheet.Cells[15, 7] = this.textBox69.Text;
                xlWorkSheet.Cells[16, 7] = this.textBox70.Text;

                //MessageBox.Show(xlWorkSheet.get_Range("D68", "D68").Value2.ToString());
                this.textBox3.Text = xlWorkSheet.get_Range("D69", "D69").Value2.ToString();
                this.textBox2.Text = xlWorkSheet.get_Range("D68", "D68").Value2.ToString();
                this.textBox28.Text = xlWorkSheet.get_Range("D68", "D68").Value2.ToString();

                MixtureCriticalPressure = xlWorkSheet.get_Range("D69", "D69").Value;
                MixtureCriticalTemperature = xlWorkSheet.get_Range("D68", "D68").Value2;

                this.textBox34.Text = xlWorkSheet.get_Range("D69", "D69").Value2.ToString();
                this.textBox68.Text = xlWorkSheet.get_Range("D68", "D68").Value2.ToString();
                this.textBox51.Text = xlWorkSheet.get_Range("D70", "D70").Value2.ToString();

                //xlWorkBook.SaveAs("C:\\SCSP_Gitlab\\RefPropWindowsForms\\Copia de REFPROP.xlS", 
                //Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, 
                //Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, 
                //misValue);

                xlWorkBook.Close(false, misValue, misValue);

                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
            }
        }

        //NLopt Optimization analysis
        private void button2_Click_1(object sender, EventArgs e)
        {
            RCMCI_with_ReHeating_Optimization_Analysis_Results RCMCI_with_ReHeating_Optimization_Analysis_Results_window = new RCMCI_with_ReHeating_Optimization_Analysis_Results(this);
            RCMCI_with_ReHeating_Optimization_Analysis_Results_window.Show();
        }
    }
}
