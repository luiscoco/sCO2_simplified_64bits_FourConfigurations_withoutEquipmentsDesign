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

using NLoptNet;

namespace RefPropWindowsForms
{
    public partial class RCMCI_without_ReHeating : Form
    {
        public double MixtureCriticalPressure = 0.0;
        public double MixtureCriticalTemperature = 0.0;

        public core luis = new core();

        //Input Data:
        public RefrigerantCategory category;
        public ReferenceState referencestate;
        public Int64 Error_code;
        public core.RecompCycle recomp_cycle = new core.RecompCycle();

        public Double wmm;

        public RCMCI_without_ReHeating()
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
        public Double t_mc1_in2, t_mc2_in2;
        public Double t_t_in2;
        public Double ua_lt2, ua_ht2;
        public Double eta1_mc2;
        public Double eta2_mc2;
        public Double eta_rc2;
        public Double eta_t2;
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

        public Double dp11_pc1, dp11_pc2;
        public Double dp12_pc1, dp12_pc2;

        public Double dp2_phx1, dp2_phx2;

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
        public Double PHX, PC11, PC21;

        //OK Button
        private void button1_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        //RESET Button
        private void button14_Click(object sender, EventArgs e)
        {
            textBox57.Text = "";
            textBox46.Text = "";
            textBox45.Text = "";
            textBox44.Text = "";
            textBox43.Text = "";
            textBox42.Text = "";
            textBox35.Text = "";
            textBox18.Text = "";
            textBox7.Text = "";
            textBox6.Text = "";
            textBox56.Text = "";
            textBox47.Text = "";

            textBox59.Text = "";
            textBox53.Text = "";
            textBox52.Text = "";
            textBox9.Text = "";
            textBox37.Text = "";
            textBox36.Text = "";
            textBox41.Text = "";
            textBox40.Text = "";
            textBox39.Text = "";
            textBox38.Text = "";
            textBox58.Text = "";
            textBox54.Text = "";

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
            //dp2_lt2 = Convert.ToDouble(textBox26.Text);
            textBox26.Text = "0.0";
            //dp2_ht2 = Convert.ToDouble(textBox25.Text);
            textBox25.Text = "0.0";
        }
        //Mixture Calculation
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
                luis.core1(this.comboBox2.Text + "=" + textBox33.Text + "," +
                           this.comboBox6.Text + "=" + textBox34.Text + "," +
                           this.comboBox12.Text + "=" + textBox68.Text + "," +
                           this.comboBox7.Text + "=" + textBox69.Text, category);
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

            w_dot_net2 = Convert.ToDouble(textBox1.Text);
            t_mc1_in2 = Convert.ToDouble(textBox2.Text);
            t_mc2_in2 = Convert.ToDouble(textBox28.Text);
            t_t_in2 = Convert.ToDouble(textBox4.Text);
            p_mc1_in2 = Convert.ToDouble(textBox3.Text);
            p_mc1_out2 = Convert.ToDouble(textBox8.Text);
            p_mc2_in2 = Convert.ToDouble(textBox23.Text);
            p_mc2_out2 = Convert.ToDouble(textBox22.Text);
            ua_lt2 = Convert.ToDouble(textBox17.Text);
            ua_ht2 = Convert.ToDouble(textBox16.Text);
            recomp_frac2 = Convert.ToDouble(textBox15.Text);
            eta1_mc2 = Convert.ToDouble(textBox14.Text);
            eta2_mc2 = Convert.ToDouble(textBox27.Text);
            eta_rc2 = Convert.ToDouble(textBox13.Text);
            eta_t2 = Convert.ToDouble(textBox19.Text);
            n_sub_hxrs2 = Convert.ToInt64(textBox20.Text);
            tol2 = Convert.ToDouble(textBox21.Text);
            dp2_lt1 = Convert.ToDouble(textBox5.Text);
            dp2_ht1 = Convert.ToDouble(textBox12.Text);
            dp12_pc1 = Convert.ToDouble(textBox11.Text);
            dp12_pc2 = Convert.ToDouble(textBox24.Text);
            dp2_phx1 = Convert.ToDouble(textBox10.Text);
            dp2_lt2 = Convert.ToDouble(textBox26.Text);
            dp2_ht2 = Convert.ToDouble(textBox25.Text);

            core.RCMCIwithoutReheating cicloRCMCI_withoutRH = new core.RCMCIwithoutReheating();

            luis.RecompCycle_RCMCI_without_Reheating(luis, ref cicloRCMCI_withoutRH, w_dot_net2,
           t_mc2_in2, t_t_in2, p_mc2_in2, p_mc2_out2, p_mc1_in2, t_mc1_in2, p_mc1_out2,
           ua_lt2, ua_ht2, eta2_mc2, eta_rc2, eta1_mc2, eta_t2, n_sub_hxrs2,
           recomp_frac2, tol2, eta_thermal2, -dp2_lt1, -dp2_lt2, -dp2_ht1, -dp2_ht2,
           -dp12_pc1, -dp12_pc2, -dp2_phx1, -dp2_phx2, -dp12_pc2, -dp12_pc2);

            massflow2 = cicloRCMCI_withoutRH.m_dot_turbine;
            w_dot_net2 = cicloRCMCI_withoutRH.W_dot_net;
            eta_thermal2 = cicloRCMCI_withoutRH.eta_thermal;
            eta_thermal2 = cicloRCMCI_withoutRH.eta_thermal;            
            recomp_frac2 = cicloRCMCI_withoutRH.recomp_frac;           

            temp21 = cicloRCMCI_withoutRH.temp[10];
            temp22 = cicloRCMCI_withoutRH.temp[1];
            temp23 = cicloRCMCI_withoutRH.temp[2];
            temp24 = cicloRCMCI_withoutRH.temp[3];
            temp25 = cicloRCMCI_withoutRH.temp[4];
            temp26 = cicloRCMCI_withoutRH.temp[5];
            temp27 = cicloRCMCI_withoutRH.temp[6];
            temp28 = cicloRCMCI_withoutRH.temp[7];
            temp29 = cicloRCMCI_withoutRH.temp[8];
            temp210 = cicloRCMCI_withoutRH.temp[9];
            temp211 = cicloRCMCI_withoutRH.temp[11];
            temp212 = cicloRCMCI_withoutRH.temp[0];

            pres21 = cicloRCMCI_withoutRH.pres[10];
            pres22 = cicloRCMCI_withoutRH.pres[1];
            pres23 = cicloRCMCI_withoutRH.pres[2];
            pres24 = cicloRCMCI_withoutRH.pres[3];
            pres25 = cicloRCMCI_withoutRH.pres[4];
            pres26 = cicloRCMCI_withoutRH.pres[5];
            pres27 = cicloRCMCI_withoutRH.pres[6];
            pres28 = cicloRCMCI_withoutRH.pres[7];
            pres29 = cicloRCMCI_withoutRH.pres[8];
            pres210 = cicloRCMCI_withoutRH.pres[9];
            pres211 = cicloRCMCI_withoutRH.pres[11];
            pres212 = cicloRCMCI_withoutRH.pres[0];

            textBox57.Text = Convert.ToString(temp21); //Point 11
            textBox46.Text = Convert.ToString(temp22); //Point 2
            textBox45.Text = Convert.ToString(temp23); //Point 3
            textBox44.Text = Convert.ToString(temp24); //Point 4
            textBox43.Text = Convert.ToString(temp25); //Point 5
            textBox42.Text = Convert.ToString(temp26); //Point 6
            textBox35.Text = Convert.ToString(temp27); //Point 7
            textBox18.Text = Convert.ToString(temp28); //Point 8
            textBox7.Text = Convert.ToString(temp29); //Point 9
            textBox6.Text = Convert.ToString(temp210); //Point 10
            textBox56.Text = Convert.ToString(temp211); //Point 12
            textBox47.Text = Convert.ToString(temp212); //Point 1

            textBox59.Text = Convert.ToString(pres21);
            textBox53.Text = Convert.ToString(pres22);
            textBox52.Text = Convert.ToString(pres23);
            textBox9.Text = Convert.ToString(pres24);
            textBox37.Text = Convert.ToString(pres25);
            textBox36.Text = Convert.ToString(pres26);
            textBox41.Text = Convert.ToString(pres27);
            textBox40.Text = Convert.ToString(pres28);
            textBox39.Text = Convert.ToString(pres29);
            textBox38.Text = Convert.ToString(pres210);
            textBox58.Text = Convert.ToString(pres211);
            textBox54.Text = Convert.ToString(pres212);

            String point1_state, point2_state, point3_state, point4_state, point5_state, point6_state;
            String point7_state, point8_state, point9_state, point10_state, point11_state, point12_state;

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

            textBox48.Text = Convert.ToString(w_dot_net2);
            textBox49.Text = Convert.ToString(massflow2);
            textBox50.Text = Convert.ToString(eta_thermal2 * 100);        
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "PureFluid")
            {
                comboBox6.Enabled = false;
                comboBox12.Enabled = false;
                comboBox7.Enabled = false;
                textBox33.Enabled = false;
                textBox34.Enabled = false;
                textBox68.Enabled = false;
                textBox69.Enabled = false;
                button20.Enabled = false;
            }

            else if (comboBox1.Text == "NewMixture")
            {
                comboBox6.Enabled = true;
                comboBox12.Enabled = true;
                comboBox7.Enabled = true;
                textBox33.Enabled = true;
                textBox34.Enabled = true;
                textBox68.Enabled = true;
                textBox69.Enabled = true;
                button20.Enabled = true;

                Refrigerant working_fluid = new Refrigerant(RefrigerantCategory.NewMixture, this.comboBox2.Text + "=" + textBox33.Text + "," + this.comboBox6.Text + "=" + textBox34.Text + "," + this.comboBox12.Text + "=" + textBox68.Text + "," + this.comboBox7.Text + "=" + textBox69.Text, ReferenceState.DEF);

                textBox32.Text = Convert.ToString(working_fluid.CriticalPressure);
                textBox51.Text = Convert.ToString(working_fluid.CriticalTemperature);             
                textBox31.Text = Convert.ToString(working_fluid.CriticalDensity);

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

        //Set Critical conditions
        private void button25_Click(object sender, EventArgs e)
        {
            double option1 = 0.0;
            double option2 = 0.0;
            double option3 = 0.0;
            double option4 = 0.0;

            option1 = Convert.ToDouble(this.textBox33.Text);
            option2 = Convert.ToDouble(this.textBox34.Text);
            option3 = Convert.ToDouble(this.textBox68.Text);
            option4 = Convert.ToDouble(this.textBox69.Text);

            if ((option1 == 1) || (option2 == 1) || (option3 == 1) || (option4 == 1))
            {
                Refrigerant working_fluid = new Refrigerant(RefrigerantCategory.NewMixture, 
                           this.comboBox2.Text + "=" + textBox33.Text + "," +
                           this.comboBox6.Text + "=" + textBox34.Text + "," +
                           this.comboBox12.Text + "=" + textBox68.Text + "," +
                           this.comboBox7.Text + "=" + textBox69.Text, ReferenceState.DEF);

                textBox32.Text = Convert.ToString(working_fluid.CriticalPressure);
                textBox51.Text = Convert.ToString(working_fluid.CriticalTemperature);
                textBox31.Text = Convert.ToString(working_fluid.CriticalDensity);

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
                xlWorkSheet.Cells[14, 6] = this.comboBox6.Text;
                xlWorkSheet.Cells[15, 6] = this.comboBox12.Text;
                xlWorkSheet.Cells[16, 6] = this.comboBox7.Text;

                // % Compositions
                xlWorkSheet.Cells[13, 7] = this.textBox33.Text;
                xlWorkSheet.Cells[14, 7] = this.textBox34.Text;
                xlWorkSheet.Cells[15, 7] = this.textBox68.Text;
                xlWorkSheet.Cells[16, 7] = this.textBox69.Text;

                //MessageBox.Show(xlWorkSheet.get_Range("D68", "D68").Value2.ToString());
                this.textBox3.Text = xlWorkSheet.get_Range("D69", "D69").Value2.ToString();
                this.textBox2.Text = xlWorkSheet.get_Range("D68", "D68").Value2.ToString();
                this.textBox28.Text = xlWorkSheet.get_Range("D68", "D68").Value2.ToString();

                MixtureCriticalPressure = xlWorkSheet.get_Range("D69", "D69").Value;
                MixtureCriticalTemperature = xlWorkSheet.get_Range("D68", "D68").Value2;

                this.textBox32.Text = xlWorkSheet.get_Range("D69", "D69").Value2.ToString();
                this.textBox51.Text = xlWorkSheet.get_Range("D68", "D68").Value2.ToString();
                this.textBox31.Text = xlWorkSheet.get_Range("D70", "D70").Value2.ToString();

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

        //NLopt optimization not considering UA (Bobyqa, Cobyla, Subplex, Nelder-Meyer, Newuoa, Praxis)
        //private void Button32_Click(object sender, EventArgs e)
        //{
        //    RCMCI_without_ReHeating_Optimization_Analysis_Results RCMCI_without_ReHeating_Optimization_Analysis_Results_window = new RCMCI_without_ReHeating_Optimization_Analysis_Results(this);
        //    RCMCI_without_ReHeating_Optimization_Analysis_Results_window.Show();
            
        //    //PureFluid
        //    if (comboBox1.Text == "PureFluid")
        //    {
        //        category = RefrigerantCategory.PureFluid;
        //        luis.core1(this.comboBox1.Text, category);
        //    }

        //    //NewMixture
        //    if (comboBox1.Text == "NewMixture")
        //    {
        //        category = RefrigerantCategory.NewMixture;
        //        luis.core1(this.comboBox2.Text + "=" + textBox33.Text + "," + this.comboBox6.Text + "=" + textBox34.Text, category);
        //    }

        //    if (comboBox1.Text == "PredefinedMixture")
        //    {
        //        category = RefrigerantCategory.PredefinedMixture;
        //    }

        //    if (comboBox1.Text == "PseudoPureFluid")
        //    {
        //        category = RefrigerantCategory.PseudoPureFluid;
        //    }

        //    if (comboBox3.Text == "DEF")
        //    {
        //        referencestate = ReferenceState.DEF;
        //    }
        //    if (comboBox3.Text == "ASH")
        //    {
        //        referencestate = ReferenceState.ASH;
        //    }
        //    if (comboBox3.Text == "IIR")
        //    {
        //        referencestate = ReferenceState.IIR;
        //    }
        //    if (comboBox3.Text == "NBP")
        //    {
        //        referencestate = ReferenceState.NBP;
        //    }

        //    luis.working_fluid.Category = category;
        //    luis.working_fluid.reference = referencestate;

        //    luis.wmm = luis.working_fluid.MolecularWeight;

        //    w_dot_net2 = Convert.ToDouble(textBox1.Text);
        //    t_mc1_in2 = Convert.ToDouble(textBox2.Text);
        //    t_mc2_in2 = Convert.ToDouble(textBox28.Text);
        //    t_t_in2 = Convert.ToDouble(textBox4.Text);
        //    p_mc1_in2 = Convert.ToDouble(textBox3.Text);
        //    p_mc1_out2 = Convert.ToDouble(textBox8.Text);
        //    p_mc2_in2 = Convert.ToDouble(textBox23.Text);
        //    p_mc2_out2 = Convert.ToDouble(textBox22.Text);
        //    ua_lt2 = Convert.ToDouble(textBox17.Text);
        //    ua_ht2 = Convert.ToDouble(textBox16.Text);
        //    recomp_frac2 = Convert.ToDouble(textBox15.Text);
        //    eta1_mc2 = Convert.ToDouble(textBox14.Text);
        //    eta2_mc2 = Convert.ToDouble(textBox27.Text);
        //    eta_rc2 = Convert.ToDouble(textBox13.Text);
        //    eta_t2 = Convert.ToDouble(textBox19.Text);
        //    n_sub_hxrs2 = Convert.ToInt64(textBox20.Text);
        //    tol2 = Convert.ToDouble(textBox21.Text);
        //    dp2_lt1 = Convert.ToDouble(textBox5.Text);
        //    dp2_ht1 = Convert.ToDouble(textBox12.Text);
        //    dp12_pc1 = Convert.ToDouble(textBox11.Text);
        //    dp12_pc2 = Convert.ToDouble(textBox24.Text);
        //    dp2_phx1 = Convert.ToDouble(textBox10.Text);
        //    dp2_lt2 = Convert.ToDouble(textBox26.Text);
        //    dp2_ht2 = Convert.ToDouble(textBox25.Text);

        //    core.RCMCIwithoutReheating cicloRCMCI_withoutRH = new core.RCMCIwithoutReheating();

        //    double UA_Total = ua_lt2 + ua_ht2;

        //    double LT_fraction = 0.1;

        //    int counter = 0;

        //    List<Double> recomp_frac2_list = new List<Double>();
        //    List<Double> p_mc1_in2_list = new List<Double>();
        //    List<Double> p_mc1_out2_list = new List<Double>();
        //    List<Double> eta_thermal2_list = new List<Double>();

        //    NLoptAlgorithm algorithm_type = NLoptAlgorithm.LN_BOBYQA;

        //    if (comboBox13.Text == "BOBYQA")
        //        algorithm_type = NLoptAlgorithm.LN_BOBYQA;
        //    else if (comboBox13.Text == "COBYLA")
        //        algorithm_type = NLoptAlgorithm.LN_COBYLA;
        //    else if (comboBox13.Text == "SUBPLEX")
        //        algorithm_type = NLoptAlgorithm.LN_SBPLX;
        //    else if (comboBox13.Text == "NELDER-MEAD")
        //        algorithm_type = NLoptAlgorithm.LN_NELDERMEAD;
        //    else if (comboBox13.Text == "NEWUOA")
        //        algorithm_type = NLoptAlgorithm.LN_NEWUOA;
        //    else if (comboBox13.Text == "PRAXIS")
        //        algorithm_type = NLoptAlgorithm.LN_PRAXIS;

        //    using (var solver = new NLoptSolver(algorithm_type, 3, 0.01, 10000))
        //    {
        //        solver.SetLowerBounds(new[] { 0.1, luis.working_fluid.CriticalPressure, (luis.working_fluid.CriticalPressure + 200) });
        //        solver.SetUpperBounds(new[] { 1.0, 125000, (p_mc2_out2 / 1.5) });

        //        solver.SetInitialStepSize(new[] { 0.05, 100, 100 });

        //        var initialValue = new[] { 0.2, luis.working_fluid.CriticalPressure, (luis.working_fluid.CriticalPressure + 500) };

        //        Func<double[], double> funcion = delegate (double[] variables)
        //        {
        //            luis.RecompCycle_RCMCI_without_Reheating(luis, ref cicloRCMCI_withoutRH, w_dot_net2,
        //            t_mc2_in2, t_t_in2, variables[2], p_mc2_out2, variables[1], t_mc1_in2, variables[2],
        //            ua_lt2, ua_ht2, eta2_mc2, eta_rc2, eta1_mc2, eta_t2, n_sub_hxrs2,
        //            variables[0], tol2, eta_thermal2, -dp2_lt1, -dp2_lt2, -dp2_ht1, -dp2_ht2,
        //            -dp12_pc1, -dp12_pc2, -dp2_phx1, -dp2_phx2, -dp12_pc2, -dp12_pc2);

        //            counter++;

        //            massflow2 = cicloRCMCI_withoutRH.m_dot_turbine;
        //            w_dot_net2 = cicloRCMCI_withoutRH.W_dot_net;

        //            eta_thermal2 = cicloRCMCI_withoutRH.eta_thermal;
        //            recomp_frac2 = variables[0];
        //            p_mc1_in2 = variables[1];
        //            p_mc1_out2 = variables[2];

        //            eta_thermal2_list.Add(eta_thermal2);
        //            recomp_frac2_list.Add(recomp_frac2);
        //            p_mc1_in2_list.Add(p_mc1_in2);
        //            p_mc1_out2_list.Add(p_mc1_out2);

        //            RCMCI_without_ReHeating_Optimization_Analysis_Results_window.listBox1.Items.Add(counter.ToString());
        //            RCMCI_without_ReHeating_Optimization_Analysis_Results_window.listBox2.Items.Add(eta_thermal2.ToString());
        //            RCMCI_without_ReHeating_Optimization_Analysis_Results_window.listBox3.Items.Add(recomp_frac2.ToString());
        //            RCMCI_without_ReHeating_Optimization_Analysis_Results_window.listBox4.Items.Add(p_mc1_in2.ToString());
        //            RCMCI_without_ReHeating_Optimization_Analysis_Results_window.listBox7.Items.Add(p_mc1_out2.ToString());

        //            return eta_thermal2;
        //        };

        //        solver.SetMaxObjective(funcion);

        //        double? finalScore;

        //        var result = solver.Optimize(initialValue, out finalScore);

        //        Double max_eta_thermal = 0.0;

        //        max_eta_thermal = eta_thermal2_list.Max();

        //        var maxIndex = eta_thermal2_list.IndexOf(eta_thermal2_list.Max());

        //        RCMCI_without_ReHeating_Optimization_Analysis_Results_window.textBox91.Text = p_mc1_in2_list[maxIndex].ToString();
        //        RCMCI_without_ReHeating_Optimization_Analysis_Results_window.textBox90.Text = recomp_frac2_list[maxIndex].ToString();
        //        RCMCI_without_ReHeating_Optimization_Analysis_Results_window.textBox2.Text = p_mc1_out2_list[maxIndex].ToString();

        //        RCMCI_without_ReHeating_Optimization_Analysis_Results_window.textBox86.Text = eta_thermal2_list[maxIndex].ToString();
        //    }
        //}

        //NLopt optimization considering UA (Bobyqa, Cobyla, Subplex, Nelder-Meyer, Newuoa, Praxis)
        //private void Button28_Click(object sender, EventArgs e)
        //{
        //    RCMCI_without_ReHeating_Optimization_Analysis_Results RCMCI_without_ReHeating_Optimization_Analysis_Results_window = new RCMCI_without_ReHeating_Optimization_Analysis_Results(this);
        //    RCMCI_without_ReHeating_Optimization_Analysis_Results_window.Show();

        //    //PureFluid
        //    if (comboBox1.Text == "PureFluid")
        //    {
        //        category = RefrigerantCategory.PureFluid;
        //        luis.core1(this.comboBox1.Text, category);
        //    }

        //    //NewMixture
        //    if (comboBox1.Text == "NewMixture")
        //    {
        //        category = RefrigerantCategory.NewMixture;
        //        luis.core1(this.comboBox2.Text + "=" + textBox33.Text + "," + this.comboBox6.Text + "=" + textBox34.Text, category);
        //    }

        //    if (comboBox1.Text == "PredefinedMixture")
        //    {
        //        category = RefrigerantCategory.PredefinedMixture;
        //    }

        //    if (comboBox1.Text == "PseudoPureFluid")
        //    {
        //        category = RefrigerantCategory.PseudoPureFluid;
        //    }

        //    if (comboBox3.Text == "DEF")
        //    {
        //        referencestate = ReferenceState.DEF;
        //    }
        //    if (comboBox3.Text == "ASH")
        //    {
        //        referencestate = ReferenceState.ASH;
        //    }
        //    if (comboBox3.Text == "IIR")
        //    {
        //        referencestate = ReferenceState.IIR;
        //    }
        //    if (comboBox3.Text == "NBP")
        //    {
        //        referencestate = ReferenceState.NBP;
        //    }

        //    luis.working_fluid.Category = category;
        //    luis.working_fluid.reference = referencestate;

        //    luis.wmm = luis.working_fluid.MolecularWeight;

        //    w_dot_net2 = Convert.ToDouble(textBox1.Text);
        //    t_mc1_in2 = Convert.ToDouble(textBox2.Text);
        //    t_mc2_in2 = Convert.ToDouble(textBox28.Text);
        //    t_t_in2 = Convert.ToDouble(textBox4.Text);
        //    p_mc1_in2 = Convert.ToDouble(textBox3.Text);
        //    p_mc1_out2 = Convert.ToDouble(textBox8.Text);
        //    p_mc2_in2 = Convert.ToDouble(textBox23.Text);
        //    p_mc2_out2 = Convert.ToDouble(textBox22.Text);
        //    ua_lt2 = Convert.ToDouble(textBox17.Text);
        //    ua_ht2 = Convert.ToDouble(textBox16.Text);
        //    recomp_frac2 = Convert.ToDouble(textBox15.Text);
        //    eta1_mc2 = Convert.ToDouble(textBox14.Text);
        //    eta2_mc2 = Convert.ToDouble(textBox27.Text);
        //    eta_rc2 = Convert.ToDouble(textBox13.Text);
        //    eta_t2 = Convert.ToDouble(textBox19.Text);
        //    n_sub_hxrs2 = Convert.ToInt64(textBox20.Text);
        //    tol2 = Convert.ToDouble(textBox21.Text);
        //    dp2_lt1 = Convert.ToDouble(textBox5.Text);
        //    dp2_ht1 = Convert.ToDouble(textBox12.Text);
        //    dp12_pc1 = Convert.ToDouble(textBox11.Text);
        //    dp12_pc2 = Convert.ToDouble(textBox24.Text);
        //    dp2_phx1 = Convert.ToDouble(textBox10.Text);
        //    dp2_lt2 = Convert.ToDouble(textBox26.Text);
        //    dp2_ht2 = Convert.ToDouble(textBox25.Text);

        //    core.RCMCIwithoutReheating cicloRCMCI_withoutRH = new core.RCMCIwithoutReheating();

        //    double UA_Total = ua_lt2 + ua_ht2;

        //    double LT_fraction = 0.1;

        //    int counter = 0;

        //    List<Double> recomp_frac2_list = new List<Double>();
        //    List<Double> p_mc1_in2_list = new List<Double>();
        //    List<Double> p_mc1_out2_list = new List<Double>();
        //    List<Double> ua_LT_list = new List<Double>();
        //    List<Double> ua_HT_list = new List<Double>();
        //    List<Double> eta_thermal2_list = new List<Double>();

        //    NLoptAlgorithm algorithm_type = NLoptAlgorithm.LN_BOBYQA;

        //    if (comboBox13.Text == "BOBYQA")
        //        algorithm_type = NLoptAlgorithm.LN_BOBYQA;
        //    else if (comboBox13.Text == "COBYLA")
        //        algorithm_type = NLoptAlgorithm.LN_COBYLA;
        //    else if (comboBox13.Text == "SUBPLEX")
        //        algorithm_type = NLoptAlgorithm.LN_SBPLX;
        //    else if (comboBox13.Text == "NELDER-MEAD")
        //        algorithm_type = NLoptAlgorithm.LN_NELDERMEAD;
        //    else if (comboBox13.Text == "NEWUOA")
        //        algorithm_type = NLoptAlgorithm.LN_NEWUOA;
        //    else if (comboBox13.Text == "PRAXIS")
        //        algorithm_type = NLoptAlgorithm.LN_PRAXIS;

        //    using (var solver = new NLoptSolver(algorithm_type, 4, 0.01, 10000))
        //    {
        //        solver.SetLowerBounds(new[] { 0.1, luis.working_fluid.CriticalPressure, (luis.working_fluid.CriticalPressure + 200), 0.2 });
        //        solver.SetUpperBounds(new[] { 1.0, 125000, (p_mc2_out2 / 1.5), 0.8 });

        //        solver.SetInitialStepSize(new[] { 0.05, 100, 100, 0.05 });

        //        var initialValue = new[] { 0.2, luis.working_fluid.CriticalPressure, (luis.working_fluid.CriticalPressure + 500), 0.5 };

        //        Func<double[], double> funcion = delegate (double[] variables)
        //        {
        //            luis.RecompCycle_RCMCI_without_Reheating_for_Optimization(luis, ref cicloRCMCI_withoutRH, w_dot_net2,
        //            t_mc2_in2, t_t_in2, variables[2], p_mc2_out2, variables[1], t_mc1_in2, variables[2],
        //            variables[3], UA_Total, eta2_mc2, eta_rc2, eta1_mc2, eta_t2, n_sub_hxrs2,
        //            variables[0], tol2, eta_thermal2, -dp2_lt1, -dp2_lt2, -dp2_ht1, -dp2_ht2,
        //            -dp12_pc1, -dp12_pc2, -dp2_phx1, -dp2_phx2, -dp12_pc2, -dp12_pc2);

        //            counter++;

        //            massflow2 = cicloRCMCI_withoutRH.m_dot_turbine;
        //            w_dot_net2 = cicloRCMCI_withoutRH.W_dot_net;

        //            eta_thermal2 = cicloRCMCI_withoutRH.eta_thermal;
        //            recomp_frac2 = variables[0];
        //            p_mc1_in2 = variables[1];
        //            p_mc1_out2 = variables[2];
        //            LT_fraction = variables[3];
        //            ua_lt2 = UA_Total * LT_fraction;
        //            ua_ht2 = UA_Total * (1 - LT_fraction);

        //            eta_thermal2_list.Add(eta_thermal2);
        //            recomp_frac2_list.Add(recomp_frac2);
        //            p_mc1_in2_list.Add(p_mc1_in2);
        //            p_mc1_out2_list.Add(p_mc1_out2);
        //            ua_LT_list.Add(ua_lt2);
        //            ua_HT_list.Add(ua_ht2);

        //            RCMCI_without_ReHeating_Optimization_Analysis_Results_window.listBox1.Items.Add(counter.ToString());
        //            RCMCI_without_ReHeating_Optimization_Analysis_Results_window.listBox2.Items.Add(eta_thermal2.ToString());
        //            RCMCI_without_ReHeating_Optimization_Analysis_Results_window.listBox3.Items.Add(recomp_frac2.ToString());
        //            RCMCI_without_ReHeating_Optimization_Analysis_Results_window.listBox4.Items.Add(p_mc1_in2.ToString());
        //            RCMCI_without_ReHeating_Optimization_Analysis_Results_window.listBox7.Items.Add(p_mc1_out2.ToString());
        //            RCMCI_without_ReHeating_Optimization_Analysis_Results_window.listBox5.Items.Add(ua_lt2.ToString());
        //            RCMCI_without_ReHeating_Optimization_Analysis_Results_window.listBox6.Items.Add(ua_ht2.ToString());

        //            return eta_thermal2;
        //        };

        //        solver.SetMaxObjective(funcion);

        //        double? finalScore;

        //        var result = solver.Optimize(initialValue, out finalScore);

        //        Double max_eta_thermal = 0.0;

        //        max_eta_thermal = eta_thermal2_list.Max();

        //        var maxIndex = eta_thermal2_list.IndexOf(eta_thermal2_list.Max());

        //        RCMCI_without_ReHeating_Optimization_Analysis_Results_window.textBox91.Text = p_mc1_in2_list[maxIndex].ToString();
        //        RCMCI_without_ReHeating_Optimization_Analysis_Results_window.textBox90.Text = recomp_frac2_list[maxIndex].ToString();
        //        RCMCI_without_ReHeating_Optimization_Analysis_Results_window.textBox2.Text = p_mc1_out2_list[maxIndex].ToString();
        //        RCMCI_without_ReHeating_Optimization_Analysis_Results_window.textBox82.Text = ua_LT_list[maxIndex].ToString();
        //        RCMCI_without_ReHeating_Optimization_Analysis_Results_window.textBox83.Text = ua_HT_list[maxIndex].ToString();

        //        RCMCI_without_ReHeating_Optimization_Analysis_Results_window.textBox86.Text = eta_thermal2_list[maxIndex].ToString();
        //    }

        //    //using (var solver = new NLoptSolver(algorithm_type, 3, 0.01, 10000))
        //    //{
        //    //    solver.SetLowerBounds(new[] { 0.1, (luis.working_fluid.CriticalPressure + 200), 0.2 });
        //    //    solver.SetUpperBounds(new[] { 1.0, (p_mc2_out2 / 1.5), 0.8 });

        //    //    solver.SetInitialStepSize(new[] { 0.05, 200, 0.05 });

        //    //    var initialValue = new[] { 0.2, (luis.working_fluid.CriticalPressure + 200), 0.5 };

        //    //    Func<double[], double> funcion = delegate (double[] variables)
        //    //    {
        //    //        luis.RecompCycle_RCMCI_without_Reheating_for_Optimization(luis, ref cicloRCMCI_withoutRH, w_dot_net2,
        //    //        t_mc2_in2, t_t_in2, p_mc2_in2, p_mc2_out2, luis.working_fluid.CriticalPressure, t_mc1_in2, variables[1],
        //    //        variables[2], UA_Total, eta2_mc2, eta_rc2, eta1_mc2, eta_t2, n_sub_hxrs2,
        //    //        variables[0], tol2, eta_thermal2, -dp2_lt1, -dp2_lt2, -dp2_ht1, -dp2_ht2,
        //    //        -dp12_pc1, -dp12_pc2, -dp2_phx1, -dp2_phx2, -dp12_pc2, -dp12_pc2);

        //    //        counter++;

        //    //        massflow2 = cicloRCMCI_withoutRH.m_dot_turbine;
        //    //        w_dot_net2 = cicloRCMCI_withoutRH.W_dot_net;

        //    //        eta_thermal2 = cicloRCMCI_withoutRH.eta_thermal;
        //    //        recomp_frac2 = variables[0];
        //    //        p_mc1_out2 = variables[1];
        //    //        LT_fraction = variables[2];
        //    //        ua_lt2 = UA_Total * LT_fraction;
        //    //        ua_ht2 = UA_Total * (1 - LT_fraction);

        //    //        eta_thermal2_list.Add(eta_thermal2);
        //    //        recomp_frac2_list.Add(recomp_frac2);
        //    //        p_mc1_in2_list.Add(p_mc1_in2);
        //    //        p_mc1_out2_list.Add(p_mc1_out2);
        //    //        ua_LT_list.Add(ua_lt2);
        //    //        ua_HT_list.Add(ua_ht2);

        //    //        RCMCI_without_ReHeating_Optimization_Analysis_Results_window.listBox1.Items.Add(counter.ToString());
        //    //        RCMCI_without_ReHeating_Optimization_Analysis_Results_window.listBox2.Items.Add(eta_thermal2.ToString());
        //    //        RCMCI_without_ReHeating_Optimization_Analysis_Results_window.listBox3.Items.Add(recomp_frac2.ToString());
        //    //        RCMCI_without_ReHeating_Optimization_Analysis_Results_window.listBox4.Items.Add(p_mc1_in2.ToString());
        //    //        RCMCI_without_ReHeating_Optimization_Analysis_Results_window.listBox7.Items.Add(p_mc1_out2.ToString());
        //    //        RCMCI_without_ReHeating_Optimization_Analysis_Results_window.listBox5.Items.Add(ua_lt2.ToString());
        //    //        RCMCI_without_ReHeating_Optimization_Analysis_Results_window.listBox6.Items.Add(ua_ht2.ToString());

        //    //        return eta_thermal2;
        //    //    };

        //    //    solver.SetMaxObjective(funcion);

        //    //    double? finalScore;

        //    //    var result = solver.Optimize(initialValue, out finalScore);

        //    //    Double max_eta_thermal = 0.0;

        //    //    max_eta_thermal = eta_thermal2_list.Max();

        //    //    var maxIndex = eta_thermal2_list.IndexOf(eta_thermal2_list.Max());

        //    //    RCMCI_without_ReHeating_Optimization_Analysis_Results_window.textBox91.Text = p_mc1_in2_list[maxIndex].ToString();
        //    //    RCMCI_without_ReHeating_Optimization_Analysis_Results_window.textBox1.Text = recomp_frac2_list[maxIndex].ToString();
        //    //    RCMCI_without_ReHeating_Optimization_Analysis_Results_window.textBox90.Text = p_mc1_out2_list[maxIndex].ToString();
        //    //    RCMCI_without_ReHeating_Optimization_Analysis_Results_window.textBox82.Text = ua_LT_list[maxIndex].ToString();
        //    //    RCMCI_without_ReHeating_Optimization_Analysis_Results_window.textBox83.Text = ua_HT_list[maxIndex].ToString();

        //    //    RCMCI_without_ReHeating_Optimization_Analysis_Results_window.textBox86.Text = eta_thermal2_list[maxIndex].ToString();
        //    //}
        //}

        //NLoptOptimization analysis
        private void button2_Click(object sender, EventArgs e)
        {
            RCMCI_without_ReHeating_Optimization_Analysis_Results RCMCI_without_ReHeating_Optimization_Analysis_Results_window = new RCMCI_without_ReHeating_Optimization_Analysis_Results(this);
            RCMCI_without_ReHeating_Optimization_Analysis_Results_window.Show();
        }
    }
}
