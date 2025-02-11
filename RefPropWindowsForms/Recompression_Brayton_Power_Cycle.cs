using System;
using System.Collections;
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

using NLoptNet;

using sc.net;

using Excel = Microsoft.Office.Interop.Excel;

using System.Diagnostics;

//using Accord;
//using Accord.Math.Optimization;

namespace RefPropWindowsForms
{
    public partial class Recompression_Brayton_Power_Cycle : Form
    {
        public double MixtureCriticalPressure = 0.0;
        public double MixtureCriticalTemperature = 0.0;

        public MainWindow punteroMainWindow = new MainWindow();

        public core luis = new core();

        //Input Data:
        public RefrigerantCategory category;
        public ReferenceState referencestate;

        public core.RecompCycle recomp_cycle = new core.RecompCycle();

        public Double wmm;

        //Brayton Power Cycle Results
        public Double specific_work_main_turbine = 0;
        public Double specific_work_reheating_turbine = 0;
        public Double specific_work_compressor1 = 0;
        public Double specific_work_compressor2 = 0;

        public Double Miscellanous_Auxiliaries = 0;
        public Double Total_Auxiliaries = 0;

        public Double w_dot_net2;
        public Double t_mc_in2;
        public Double t_t_in2;
        public Double p_rhx_in2;
        public Double t_rht_in2;
        public Double ua_lt2,ua_ht2;
        public Double eta_mc2;
        public Double eta_rc2;
        public Double eta_t2;
        public Double eta_trh2;
        public Int64 n_sub_hxrs2;
        public Double p_mc_in2;
        public Double p_mc_out2;
        public Double recomp_frac2;
        public Double tol2;
        public Double eta_thermal2;

        public Double dp2_lt1, dp2_lt2;
        public Double dp2_ht1, dp2_ht2;
        public Double dp2_pc1, dp2_pc2;
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

        public Double massflow2;
        public Double LT_mdoth, LT_mdotc, LT_Tcin, LT_Thin, LT_Pcin, LT_Phin;
        public Double LT_Pcout, LT_Phout, LT_Q, HT_mdoth, HT_mdotc, HT_Tcin, HT_Thin;

        public Recompression_Brayton_Power_Cycle(MainWindow punteroMainWindow1)
        {
            punteroMainWindow = punteroMainWindow1;
            InitializeComponent();
        }

        public Double HT_Pcin, HT_Phin, HT_Pcout, HT_Phout, HT_Q, LT_UA, HT_UA;
        public Double PHX_Q2, RHX_Q2, PC_Q2;
        public Double LT_Effc, HT_Effc;
        public Double N_design2 = 0;

        private void button1_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }
        
        //RESET Button
        private void button14_Click(object sender, EventArgs e)
        {
           
            //w_dot_net2 = Convert.ToDouble(textBox1.Text);
            textBox1.Text = "50000";
            //t_mc_in2 = Convert.ToDouble(textBox2.Text);
            textBox2.Text = "305.15";
            //t_t_in2 = Convert.ToDouble(textBox4.Text);
            //p_mc_in2 = Convert.ToDouble(textBox3.Text);
            textBox3.Text = "7400";
            //p_mc_out2 = Convert.ToDouble(textBox8.Text);
            textBox8.Text = "25000";
            //p_rhx_in2 = Convert.ToDouble(textBox7.Text);
            textBox7.Text = "17300";
            //t_rht_in2 = Convert.ToDouble(textBox6.Text);
            //ua_lt2 = Convert.ToDouble(textBox17.Text);
            textBox17.Text = "5000";
            //ua_ht2 = Convert.ToDouble(textBox16.Text);
            textBox16.Text = "5000";
            //dp2_lt1 = Convert.ToDouble(textBox5.Text);
            textBox5.Text = "0.0";
            //dp2_lt2 = Convert.ToDouble(textBox26.Text);
            textBox26.Text = "0.0";
            //dp2_ht1 = Convert.ToDouble(textBox12.Text);
            textBox12.Text = "0.0";
            //dp2_ht2 = Convert.ToDouble(textBox25.Text);
            textBox25.Text = "0.0";
            //dp2_pc1 = Convert.ToDouble(textBox11.Text);
            textBox11.Text = "0.0";
            //dp2_phx2 = Convert.ToDouble(textBox10.Text);
            textBox10.Text = "0.0";
            //dp2_rhx2 = Convert.ToDouble(textBox9.Text);
            textBox9.Text = "0.0";
            //recomp_frac2 = Convert.ToDouble(textBox15.Text);
            textBox15.Text = "0.25";
            //eta_mc2 = Convert.ToDouble(textBox14.Text);
            textBox14.Text = "0.89";
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

            textBox22.Text = "";
            textBox23.Text = "";
            textBox27.Text = "";
            textBox24.Text = "";
            textBox29.Text = "";
            textBox28.Text = "";
            textBox41.Text = "";
            textBox40.Text = "";
            textBox39.Text = "";
            textBox38.Text = "";
            textBox37.Text = "";
            textBox36.Text = "";

            textBox47.Text = "";
            textBox46.Text = "";
            textBox45.Text = "";
            textBox44.Text = "";
            textBox43.Text = "";
            textBox42.Text = "";
            textBox35.Text = "";
            textBox34.Text = "";
            textBox33.Text = "";
            textBox32.Text = "";
            textBox31.Text = "";
            textBox30.Text = "";

            textBox48.Text = "";
            textBox49.Text = "";
            textBox50.Text = "";
        }

        //For Mixtures Calculations
        private void button22_Click(object sender, EventArgs e)
        {
            int maxIterations = 100;
            int numIterations = 0;

            //PureFluid
            if (comboBox16.Text == "PureFluid")
            {
                category = RefrigerantCategory.PureFluid;
                luis.core1(this.comboBox1.Text, category);
            }

            //NewMixture
            if (comboBox16.Text == "NewMixture")
            {
                category = RefrigerantCategory.NewMixture;
                luis.core1(this.comboBox2.Text + "=" + textBox60.Text + "," + this.comboBox1.Text + "=" + textBox61.Text + "," + this.comboBox18.Text + "=" + textBox51.Text + "," + this.comboBox17.Text + "=" + textBox80.Text, category);       
            }

            if (comboBox16.Text == "PredefinedMixture")
            {
                category = RefrigerantCategory.PredefinedMixture;
            }

            if (comboBox16.Text == "PseudoPureFluid")
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

            Double w_dot_net2 = Convert.ToDouble(textBox1.Text);
            Double t_mc_in2 = Convert.ToDouble(textBox2.Text);
            Double t_t_in2 = Convert.ToDouble(textBox4.Text);
            Double p_mc_in2 = Convert.ToDouble(textBox3.Text);
            Double p_mc_out2 = Convert.ToDouble(textBox8.Text);
            Double p_rhx_in2 = Convert.ToDouble(textBox7.Text);
            Double t_rht_in2 = Convert.ToDouble(textBox6.Text);
            Double ua_lt2 = Convert.ToDouble(textBox17.Text);
            Double ua_ht2 = Convert.ToDouble(textBox16.Text);

            Double dp2_lt1 = Convert.ToDouble(textBox5.Text);
            Double dp2_lt2 = Convert.ToDouble(textBox26.Text);
            Double dp2_ht1 = Convert.ToDouble(textBox12.Text);
            Double dp2_ht2 = Convert.ToDouble(textBox25.Text);
            Double dp2_pc2 = Convert.ToDouble(textBox11.Text);
            Double dp2_phx1 = Convert.ToDouble(textBox10.Text);
            Double dp2_rhx1 = Convert.ToDouble(textBox9.Text);

            Double recomp_frac = Convert.ToDouble(textBox15.Text);
            Double eta_mc2 = Convert.ToDouble(textBox14.Text);
            Double eta_rc2 = Convert.ToDouble(textBox13.Text);
            Double eta_t2 = Convert.ToDouble(textBox19.Text);
            Double eta_trh2 = Convert.ToDouble(textBox18.Text);
            long n_sub_hxrs2 = Convert.ToInt64(textBox20.Text);
            Double tol2 = Convert.ToDouble(textBox21.Text);

            luis.wmm = luis.working_fluid.MolecularWeight;

            core.RecompCycle cicloRC = new core.RecompCycle();

            luis.RecompCycledesign_withReheating(luis, ref cicloRC, w_dot_net2, t_mc_in2, t_t_in2, p_mc_in2, p_mc_out2, p_rhx_in2, t_rht_in2,
                   -dp2_lt1, -dp2_ht1, -dp2_pc2, -dp2_phx1, -dp2_rhx1, -dp2_lt2, -dp2_ht2, ua_lt2, ua_ht2, recomp_frac,
                   eta_mc2, eta_rc2, eta_t2, eta_trh2, n_sub_hxrs2, tol2);

            massflow2 = cicloRC.m_dot_turbine;
            w_dot_net2 = cicloRC.W_dot_net;
            eta_thermal2 = cicloRC.eta_thermal;
            recomp_frac2 = cicloRC.recomp_frac;           

            temp21 = cicloRC.temp[0];
            temp22 = cicloRC.temp[1];
            temp23 = cicloRC.temp[2];
            temp24 = cicloRC.temp[3];
            temp25 = cicloRC.temp[4];
            temp26 = cicloRC.temp[5];
            temp27 = cicloRC.temp[6];
            temp28 = cicloRC.temp[7];
            temp29 = cicloRC.temp[8];
            temp210 = cicloRC.temp[9];
            temp211 = cicloRC.temp[10];
            temp212 = cicloRC.temp[11];

            pres21 = cicloRC.pres[0];
            pres22 = cicloRC.pres[1];
            pres23 = cicloRC.pres[2];
            pres24 = cicloRC.pres[3];
            pres25 = cicloRC.pres[4];
            pres26 = cicloRC.pres[5];
            pres27 = cicloRC.pres[6];
            pres28 = cicloRC.pres[7];
            pres29 = cicloRC.pres[8];
            pres210 = cicloRC.pres[9];
            pres211 = cicloRC.pres[10];
            pres212 = cicloRC.pres[11];

            textBox50.Text = Convert.ToString(cicloRC.eta_thermal);

            textBox22.Text = Convert.ToString(pres21);
            textBox23.Text = Convert.ToString(pres22);
            textBox27.Text = Convert.ToString(pres23);
            textBox24.Text = Convert.ToString(pres24);
            textBox29.Text = Convert.ToString(pres25);
            textBox28.Text = Convert.ToString(pres26);
            textBox41.Text = Convert.ToString(pres27);
            textBox40.Text = Convert.ToString(pres28);
            textBox39.Text = Convert.ToString(pres29);
            textBox38.Text = Convert.ToString(pres210);
            textBox37.Text = Convert.ToString(pres211);
            textBox36.Text = Convert.ToString(pres212);

            textBox47.Text = Convert.ToString(temp21);
            textBox46.Text = Convert.ToString(temp22);
            textBox45.Text = Convert.ToString(temp23);
            textBox44.Text = Convert.ToString(temp24);
            textBox43.Text = Convert.ToString(temp25);
            textBox42.Text = Convert.ToString(temp26);
            textBox35.Text = Convert.ToString(temp27);
            textBox34.Text = Convert.ToString(temp28);
            textBox33.Text = Convert.ToString(temp29);
            textBox32.Text = Convert.ToString(temp210);
            textBox31.Text = Convert.ToString(temp211);
            textBox30.Text = Convert.ToString(temp212);

            textBox48.Text = Convert.ToString(w_dot_net2);
            textBox49.Text = Convert.ToString(massflow2);
            textBox50.Text = Convert.ToString(eta_thermal2 * 100);

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

            String point1_state, point2_state, point3_state, point4_state, point5_state, point6_state;
            String point7_state, point8_state, point9_state, point10_state, point11_state, point12_state;

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

            toolTip1.SetToolTip(label55, point1_state);
            toolTip2.SetToolTip(label57, point2_state);
            toolTip3.SetToolTip(label59, point3_state);
            toolTip4.SetToolTip(label60, point4_state);
            toolTip5.SetToolTip(label61, point5_state);
            toolTip6.SetToolTip(label62, point6_state);
            toolTip7.SetToolTip(label63, point7_state);
            toolTip8.SetToolTip(label65, point8_state);
            toolTip9.SetToolTip(label66, point9_state);
            toolTip10.SetToolTip(label67, point10_state);
            toolTip11.SetToolTip(label64, point11_state);
            toolTip12.SetToolTip(label68, point12_state);
        }

        private void comboBox16_SelectedValueChanged(object sender, EventArgs e)
        {
            if (comboBox16.Text == "PureFluid")
            {
                comboBox1.Enabled = false;
                comboBox18.Enabled = false;
                comboBox17.Enabled = false;
                textBox60.Enabled = false;
                textBox61.Enabled = false;
                textBox51.Enabled = false;
                textBox80.Enabled = false;
                button22.Enabled = false;
            }

            else if (comboBox16.Text == "NewMixture")
            {
                comboBox1.Enabled = true;
                comboBox18.Enabled = true;
                comboBox17.Enabled = true;
                textBox60.Enabled = true;
                textBox61.Enabled = true;
                textBox51.Enabled = true;
                textBox80.Enabled = true;
                button22.Enabled = true;

                string fluid_name = this.comboBox2.Text + "=" + textBox60.Text + "," + this.comboBox1.Text + "=" + textBox61.Text + "," + this.comboBox18.Text + "=" + textBox51.Text + "," + this.comboBox17.Text + "=" + textBox80.Text;
                Refrigerant working_fluid = new Refrigerant(RefrigerantCategory.NewMixture, fluid_name, ReferenceState.DEF);

                textBox59.Text = Convert.ToString(working_fluid.CriticalPressure);
                textBox58.Text = Convert.ToString(working_fluid.CriticalTemperature - 273.15);
                textBox55.Text = Convert.ToString(working_fluid.CriticalDensity);

                MixtureCriticalPressure = working_fluid.CriticalPressure;
                MixtureCriticalTemperature = working_fluid.CriticalTemperature;
            }
        }

        private void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            Refrigerant working_fluid = new Refrigerant(RefrigerantCategory.NewMixture, this.comboBox2.Text + "=" + textBox60.Text + "," + this.comboBox1.Text + "=" + textBox61.Text, ReferenceState.DEF);

            textBox59.Text = Convert.ToString(working_fluid.CriticalPressure);
            textBox58.Text = Convert.ToString(working_fluid.CriticalTemperature - 273.15);
            textBox55.Text = Convert.ToString(working_fluid.CriticalDensity);
        }


        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            Refrigerant working_fluid = new Refrigerant(RefrigerantCategory.NewMixture, this.comboBox2.Text + "=" + textBox60.Text + "," + this.comboBox1.Text + "=" + textBox61.Text + "," + this.comboBox18.Text + "=" + textBox51.Text + "," + this.comboBox17.Text + "=" + textBox80.Text, ReferenceState.DEF);

            textBox59.Text = Convert.ToString(working_fluid.CriticalPressure);
            textBox58.Text = Convert.ToString(working_fluid.CriticalTemperature - 273.15);
            textBox55.Text = Convert.ToString(working_fluid.CriticalDensity);
        }

        private void button23_Click(object sender, EventArgs e)
        {
            Refrigerant working_fluid = new Refrigerant(RefrigerantCategory.NewMixture, this.comboBox2.Text + "=" + textBox60.Text + "," + this.comboBox1.Text + "=" + textBox61.Text + "," + this.comboBox18.Text + "=" + textBox51.Text + "," + this.comboBox17.Text + "=" + textBox80.Text, ReferenceState.DEF);

            this.textBox2.Text = (working_fluid.CriticalTemperature).ToString();
            this.textBox3.Text = Convert.ToString(working_fluid.CriticalPressure);
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

        //Set Critial conditions
        private void button30_Click(object sender, EventArgs e)
        {
            double option1 = 0.0;
            double option2 = 0.0;
            double option3 = 0.0;
            double option4 = 0.0;

            option1 = Convert.ToDouble(this.textBox60.Text);
            option2 = Convert.ToDouble(this.textBox61.Text);
            option3 = Convert.ToDouble(this.textBox51.Text);
            option4 = Convert.ToDouble(this.textBox80.Text);

            if ((option1 == 1) || (option2 == 1) || (option3 == 1) || (option4 == 1))
            {
                Refrigerant working_fluid = new Refrigerant(RefrigerantCategory.NewMixture, this.comboBox2.Text + "=" + textBox60.Text + "," + this.comboBox1.Text + "=" + textBox61.Text + "," + this.comboBox18.Text + "=" + textBox51.Text + "," + this.comboBox17.Text + "=" + textBox80.Text, ReferenceState.DEF);

                textBox59.Text = Convert.ToString(working_fluid.CriticalPressure);
                textBox58.Text = Convert.ToString(working_fluid.CriticalTemperature);
                textBox55.Text = Convert.ToString(working_fluid.CriticalDensity);

                textBox3.Text = Convert.ToString(working_fluid.CriticalPressure);
                textBox2.Text = Convert.ToString(working_fluid.CriticalTemperature);

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
                xlWorkSheet.Cells[14, 6] = this.comboBox1.Text;
                xlWorkSheet.Cells[15, 6] = this.comboBox18.Text;
                xlWorkSheet.Cells[16, 6] = this.comboBox17.Text;

                // % Compositions
                xlWorkSheet.Cells[13, 7] = this.textBox60.Text;
                xlWorkSheet.Cells[14, 7] = this.textBox61.Text;
                xlWorkSheet.Cells[15, 7] = this.textBox51.Text;
                xlWorkSheet.Cells[16, 7] = this.textBox80.Text;

                //MessageBox.Show(xlWorkSheet.get_Range("D68", "D68").Value2.ToString());
                this.textBox58.Text = xlWorkSheet.get_Range("D68", "D68").Value2.ToString();
                this.textBox59.Text = xlWorkSheet.get_Range("D69", "D69").Value2.ToString();
                this.textBox55.Text = xlWorkSheet.get_Range("D70", "D70").Value2.ToString();

                this.textBox3.Text = xlWorkSheet.get_Range("D69", "D69").Value2.ToString();
                this.textBox2.Text = xlWorkSheet.get_Range("D68", "D68").Value2.ToString();

                MixtureCriticalPressure = xlWorkSheet.get_Range("D69", "D69").Value2;
                MixtureCriticalTemperature = xlWorkSheet.get_Range("D68", "D68").Value2;

                //xlWorkBook.SaveAs("C:\\SCSP_Gitlab\\RefPropWindowsForms\\Copia de REFPROP.xlS", 
                //    Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, 
                //    Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, 
                //    misValue);

                xlWorkBook.Close(false, misValue, misValue);

                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
            }
        }
              
        public double Sum(double[] x)
        {
            return Math.Pow(x[0] * x[0] - x[1], 2.0) + Math.Pow(1.0 + x[0], 2.0);
        }

        private void button35_Click(object sender, EventArgs e)
        {
            RC_with_ReHeating_Optimization_Analysis_Results RC_with_ReHeating_Optimization_Analysis_Results_window = new RC_with_ReHeating_Optimization_Analysis_Results(this);
            RC_with_ReHeating_Optimization_Analysis_Results_window.Show();
        }

    } // End Form class
} // End NameSpace
