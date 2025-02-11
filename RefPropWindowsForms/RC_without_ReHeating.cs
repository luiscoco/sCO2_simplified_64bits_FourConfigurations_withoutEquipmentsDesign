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

using NLoptNet;

using sc.net;

using Excel = Microsoft.Office.Interop.Excel;

//using Accord;
//using Accord.Math.Optimization;

namespace RefPropWindowsForms
{
    public partial class RC_without_ReHeating : Form
    {
        public double MixtureCriticalPressure = 0.0;
        public double MixtureCriticalTemperature = 0.0;

        public core luis = new core();

        //Input Data:
        public RefrigerantCategory category;
        public ReferenceState referencestate;
        public Double W_dot_net, T_mc_in, T_t_in, P_mc_in, P_mc_out;
        public Double UA_LT, UA_HT, recomp_frac, eta_mc, eta_rc, eta_t;
        public Double DP_LT_c;
        public Double DP_HT_c;
        public Double DP_PC;
        public Double DP_PHX;
        public Double DP_LT_h;
        public Double DP_HT_h;
        public Int64 N_sub_hxrs;
        public Double tol;
        public Int64 Error_code;
        public core.RecompCycle recomp_cycle = new core.RecompCycle();

        //Parameters
        public Int64 max_iter = 10;
        public Double temperature_tolerance = 1.0e-6;  // temperature differences below this are considered zero

        //Local Variables
        public Int64 error_code, index;
        public Double w_mc, w_rc, w_t, w_trh, C_dot_min, Q_dot_max;
        public Double T9_lower_bound, T9_upper_bound, T8_lower_bound, T8_upper_bound, last_LT_residual, last_T9_guess;
        public Double last_HT_residual, last_T8_guess, secant_guess;
        public Double m_dot_t, m_dot_mc, m_dot_rc, eta_mc_isen, eta_rc_isen, eta_t_isen;
        public Double min_DT_LT, min_DT_HT, UA_LT_calc, UA_HT_calc, Q_dot_LT, Q_dot_HT, UA_HT_residual, UA_LT_residual;
        public Double[] temp = new Double[10];
        public Double[] pres = new Double[10];
        public Double[] enth = new Double[10];
        public Double[] entr = new Double[10];
        public Double[] dens = new Double[10];

        public Double wmm;

        public RC_without_ReHeating()
        {
            InitializeComponent();
        }

        public Double specific_work_main_turbine = 0;
        public Double specific_work_reheating_turbine = 0;
        public Double specific_work_compressor1 = 0;
        public Double specific_work_compressor2 = 0;
        public Double Miscellanous_Auxiliaries = 0;
        public Double Total_Auxiliaries = 0;

        public Double w_dot_net2;
        public Double t_mc_in2;
        public Double t_t_in2;
        public Double ua_lt2, ua_ht2;
        public Double eta_mc2;
        public Double eta_rc2;
        public Double eta_t2;
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

        public Double massflow2;
        public Double LT_mdoth, LT_mdotc, LT_Tcin, LT_Thin, LT_Pcin, LT_Phin;
        public Double LT_Pcout, LT_Phout, LT_Q, HT_mdoth, HT_mdotc, HT_Tcin, HT_Thin;
        public Double HT_Pcin, HT_Phin, HT_Pcout, HT_Phout, HT_Q, LT_UA, HT_UA;
        public Double LT_Effc, HT_Effc, N_design2;
        public Double PHX_Q2, PC_Q2;

        private void button1_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        //RESET Button
        private void button14_Click(object sender, EventArgs e)
        {
            //t_mc_in2 = Convert.ToDouble(textBox2.Text);
            textBox2.Text = "305.15";
            //t_t_in2 = Convert.ToDouble(textBox4.Text);
            //p_mc_in2 = Convert.ToDouble(textBox3.Text);
            textBox3.Text = "7400";
            //p_mc_out2 = Convert.ToDouble(textBox8.Text);
            textBox8.Text = "25000";
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
            //recomp_frac2 = Convert.ToDouble(textBox15.Text);
            textBox15.Text = "0.25";
            //eta_mc2 = Convert.ToDouble(textBox14.Text);
            textBox14.Text = "0.89";
            //eta_rc2 = Convert.ToDouble(textBox13.Text);
            textBox13.Text = "0.89";
            //eta_t2 = Convert.ToDouble(textBox19.Text);
            textBox19.Text = "0.93";
            //n_sub_hxrs2 = Convert.ToInt64(textBox20.Text);
            textBox20.Text = "15";
            //tol2 = Convert.ToDouble(textBox21.Text);
            textBox21.Text = "0.0001";

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
            textBox49.Text = "";
            textBox50.Text = "";
        }

        //Design Point calculation for Mixtures
        public void button11_Click(object sender, EventArgs e)
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
                luis.core1(this.comboBox5.Text + "=" + textBox31.Text + "," + this.comboBox4.Text + "=" + textBox36.Text + "," + this.comboBox7.Text + "=" + textBox67.Text, category);
                //luis.core1(this.comboBox2.Text + "=" + textBox31.Text + "," + this.comboBox6.Text + "=" + textBox36.Text, category);
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

            w_dot_net2 = Convert.ToDouble(textBox1.Text);
            t_mc_in2 = Convert.ToDouble(textBox2.Text);
            t_t_in2 = Convert.ToDouble(textBox4.Text);
            p_mc_in2 = Convert.ToDouble(textBox3.Text);
            p_mc_out2 = Convert.ToDouble(textBox8.Text);
            ua_lt2 = Convert.ToDouble(textBox17.Text);
            ua_ht2 = Convert.ToDouble(textBox16.Text);

            dp2_lt1 = Convert.ToDouble(textBox5.Text);
            dp2_lt2 = Convert.ToDouble(textBox26.Text);
            dp2_ht1 = Convert.ToDouble(textBox12.Text);
            dp2_ht2 = Convert.ToDouble(textBox25.Text);
            dp2_pc2 = Convert.ToDouble(textBox11.Text);
            dp2_phx1 = Convert.ToDouble(textBox10.Text);

            recomp_frac2 = Convert.ToDouble(textBox15.Text);
            eta_mc2 = Convert.ToDouble(textBox14.Text);
            eta_rc2 = Convert.ToDouble(textBox13.Text);
            eta_t2 = Convert.ToDouble(textBox19.Text);
            n_sub_hxrs2 = Convert.ToInt64(textBox20.Text);
            tol2 = Convert.ToDouble(textBox21.Text);

            luis.wmm = luis.working_fluid.MolecularWeight;                        

            core.RecompCycle_withoutRH cicloRC_withoutRH = new core.RecompCycle_withoutRH();

            increasingCIP:

            luis.RecompCycledesign(luis, ref cicloRC_withoutRH, w_dot_net2, t_mc_in2, t_t_in2, p_mc_in2, p_mc_out2,
            -dp2_lt1, -dp2_ht1, -dp2_pc2, -dp2_phx1, -dp2_lt2, -dp2_ht2, ua_lt2, ua_ht2, recomp_frac2,
            eta_mc2, eta_rc2, eta_t2, n_sub_hxrs2, tol2);      

            if (cicloRC_withoutRH.eta_thermal == 0)
            {
                p_mc_in2 = p_mc_in2 + 10.0;
                numIterations++;

                if (numIterations < maxIterations)
                {
                    goto increasingCIP;
                }
            }

            massflow2 = cicloRC_withoutRH.m_dot_turbine;
            w_dot_net2 = cicloRC_withoutRH.W_dot_net;
            eta_thermal2 = cicloRC_withoutRH.eta_thermal;
            eta_thermal2 = cicloRC_withoutRH.eta_thermal;
            recomp_frac2 = cicloRC_withoutRH.recomp_frac;            

            temp21 = cicloRC_withoutRH.temp[0];
            temp22 = cicloRC_withoutRH.temp[1];
            temp23 = cicloRC_withoutRH.temp[2];
            temp24 = cicloRC_withoutRH.temp[3];
            temp25 = cicloRC_withoutRH.temp[4];
            temp26 = cicloRC_withoutRH.temp[5];
            temp27 = cicloRC_withoutRH.temp[6];
            temp28 = cicloRC_withoutRH.temp[7];
            temp29 = cicloRC_withoutRH.temp[8];
            temp210 = cicloRC_withoutRH.temp[9];

            pres21 = cicloRC_withoutRH.pres[0];
            pres22 = cicloRC_withoutRH.pres[1];
            pres23 = cicloRC_withoutRH.pres[2];
            pres24 = cicloRC_withoutRH.pres[3];
            pres25 = cicloRC_withoutRH.pres[4];
            pres26 = cicloRC_withoutRH.pres[5];
            pres27 = cicloRC_withoutRH.pres[6];
            pres28 = cicloRC_withoutRH.pres[7];
            pres29 = cicloRC_withoutRH.pres[8];
            pres210 = cicloRC_withoutRH.pres[9];

            //Fill results in the Graphical User Interface (GUI)

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

            textBox48.Text = Convert.ToString(w_dot_net2);
            textBox49.Text = Convert.ToString(massflow2);
            textBox50.Text = Convert.ToString(eta_thermal2 * 100);

            String point1_state, point2_state, point3_state, point4_state, point5_state, point6_state;
            String point7_state, point8_state, point9_state, point10_state;

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
        }

        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "PureFluid")
            {
                comboBox4.Enabled = false;
                comboBox7.Enabled = false;
                comboBox5.Enabled = false;
                textBox31.Enabled = false;
                textBox36.Enabled = false;
                textBox67.Enabled = false;
                button11.Enabled = false;
            }

            else if (comboBox1.Text == "NewMixture")
            {
                comboBox4.Enabled = true;
                comboBox7.Enabled = true;
                comboBox5.Enabled = true;
                textBox31.Enabled = true;
                textBox36.Enabled = true;
                textBox67.Enabled = true;
                button11.Enabled = true;

                //Refrigerant working_fluid = new Refrigerant(RefrigerantCategory.NewMixture, this.comboBox2.Text + "=" + textBox31.Text + "," + this.comboBox6.Text + "=" + textBox36.Text, ReferenceState.DEF);
                Refrigerant working_fluid = new Refrigerant(RefrigerantCategory.NewMixture, this.comboBox5.Text + "=" + textBox31.Text + "," + this.comboBox4.Text + "=" + textBox36.Text + "," + this.comboBox7.Text + "=" + textBox67.Text, ReferenceState.DEF);

                textBox37.Text = Convert.ToString(working_fluid.CriticalPressure);              
                textBox51.Text = Convert.ToString(working_fluid.CriticalTemperature);             
                textBox52.Text = Convert.ToString(working_fluid.CriticalDensity);

                MixtureCriticalTemperature = working_fluid.CriticalTemperature;
                MixtureCriticalPressure = working_fluid.CriticalPressure;
            }
        }

        //Set critical conditions button
        private void button13_Click(object sender, EventArgs e)
        {
            double option1 = 0.0;
            double option2 = 0.0;
            double option3 = 0.0;

            option1 = Convert.ToDouble(this.textBox31.Text);
            option2 = Convert.ToDouble(this.textBox36.Text);
            option3 = Convert.ToDouble(this.textBox67.Text);

            if ((option1 == 1) || (option2 == 1) || (option3 == 1))
            {
                Refrigerant working_fluid = new Refrigerant(RefrigerantCategory.NewMixture, this.comboBox5.Text + "=" + textBox31.Text + "," + this.comboBox4.Text + "=" + textBox36.Text + "," + this.comboBox7.Text + "=" + textBox67.Text, ReferenceState.DEF);

                this.textBox2.Text = (working_fluid.CriticalTemperature).ToString();
                this.textBox3.Text = Convert.ToString(working_fluid.CriticalPressure);

                this.textBox37.Text = (working_fluid.CriticalPressure).ToString();
                this.textBox51.Text = (working_fluid.CriticalTemperature).ToString();
                this.textBox52.Text = (working_fluid.CriticalDensity).ToString();

                this.textBox2.Text = (working_fluid.CriticalTemperature).ToString();
                this.textBox3.Text = (working_fluid.CriticalPressure).ToString();

                MixtureCriticalTemperature = working_fluid.CriticalTemperature;
                MixtureCriticalPressure = working_fluid.CriticalPressure;
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
                xlWorkSheet.Cells[13, 6] = this.comboBox5.Text;
                xlWorkSheet.Cells[14, 6] = this.comboBox4.Text;
                xlWorkSheet.Cells[15, 6] = this.comboBox7.Text;

                // % Compositions
                xlWorkSheet.Cells[13, 7] = this.textBox31.Text;
                xlWorkSheet.Cells[14, 7] = this.textBox36.Text;
                xlWorkSheet.Cells[15, 7] = this.textBox67.Text;

                //MessageBox.Show(xlWorkSheet.get_Range("D68", "D68").Value2.ToString());
                this.textBox51.Text = xlWorkSheet.get_Range("D68", "D68").Value2.ToString();
                this.textBox37.Text = xlWorkSheet.get_Range("D69", "D69").Value2.ToString();
                this.textBox52.Text = xlWorkSheet.get_Range("D70", "D70").Value2.ToString();

                this.textBox2.Text = xlWorkSheet.get_Range("D68", "D68").Value2.ToString();
                this.textBox3.Text = xlWorkSheet.get_Range("D69", "D69").Value2.ToString();

                MixtureCriticalTemperature = xlWorkSheet.get_Range("D68", "D68").Value2;
                MixtureCriticalPressure = xlWorkSheet.get_Range("D69", "D69").Value2;            

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

        private void button35_Click(object sender, EventArgs e)
        {
            RC_without_ReHeating_Optimization_Analysis_Results RC_without_ReHeating_Optimization_Analysis_Results_window = new RC_without_ReHeating_Optimization_Analysis_Results(this);
            RC_without_ReHeating_Optimization_Analysis_Results_window.Show();
        }
    }
}
