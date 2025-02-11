﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using sc.net;

using NLoptNet;

using Excel = Microsoft.Office.Interop.Excel;

using System.Reflection;


namespace RefPropWindowsForms
{
    public partial class PCRC_without_ReHeating_Optimization_Analysis_Results : Form
    {
        PCRC_without_ReHeating puntero_aplicacion;

        public PCRC_without_ReHeating_Optimization_Analysis_Results(PCRC_without_ReHeating puntero1)
        {
            puntero_aplicacion = puntero1;
            InitializeComponent();
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

        //Ok button
        private void Button1_Click(object sender, EventArgs e)
        {

        }

        //Close button
        private void Button4_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        //Run Optimization
        private void Button3_Click(object sender, EventArgs e)
        {
            int counter_Excel = 4;

            Excel.Application xlApp1;
            Excel.Workbook xlWorkBook1;
            Excel.Worksheet xlWorkSheet1;
            Excel.Worksheet xlWorkSheet2;

            object misValue1 = System.Reflection.Missing.Value;

            xlApp1 = new Excel.Application();
            xlApp1.DisplayAlerts = false;
            xlWorkBook1 = xlApp1.Workbooks.Add(misValue1);

            xlWorkSheet1 = (Excel.Worksheet)xlWorkBook1.Worksheets.Add();

            //xlWorkSheet1 = (Excel.Worksheet)xlWorkBook1.Worksheets.get_Item(xlWorkBook1.Worksheets.Count);           

            double initial_CIP_value = 0;

            //Optimization UA false
            if (checkBox2.Checked == false)
            {
                //PureFluid
                if (puntero_aplicacion.comboBox1.Text == "PureFluid")
                {
                    puntero_aplicacion.category = RefrigerantCategory.PureFluid;
                    puntero_aplicacion.luis.core1(puntero_aplicacion.comboBox1.Text, puntero_aplicacion.category);
                }

                //NewMixture
                if (puntero_aplicacion.comboBox1.Text == "NewMixture")
                {
                    puntero_aplicacion.category = RefrigerantCategory.NewMixture;
                    puntero_aplicacion.luis.core1(puntero_aplicacion.comboBox2.Text + "=" + puntero_aplicacion.textBox68.Text + "," + puntero_aplicacion.comboBox6.Text + "=" + puntero_aplicacion.textBox69.Text + "," + puntero_aplicacion.comboBox12.Text + "=" + puntero_aplicacion.textBox33.Text + "," + puntero_aplicacion.comboBox7.Text + "=" + puntero_aplicacion.textBox34.Text, puntero_aplicacion.category);
                }

                if (puntero_aplicacion.comboBox1.Text == "PredefinedMixture")
                {
                    puntero_aplicacion.category = RefrigerantCategory.PredefinedMixture;
                }

                if (puntero_aplicacion.comboBox1.Text == "PseudoPureFluid")
                {
                    puntero_aplicacion.category = RefrigerantCategory.PseudoPureFluid;
                }

                if (puntero_aplicacion.comboBox3.Text == "DEF")
                {
                    puntero_aplicacion.referencestate = ReferenceState.DEF;
                }
                if (puntero_aplicacion.comboBox3.Text == "ASH")
                {
                    puntero_aplicacion.referencestate = ReferenceState.ASH;
                }
                if (puntero_aplicacion.comboBox3.Text == "IIR")
                {
                    puntero_aplicacion.referencestate = ReferenceState.IIR;
                }
                if (puntero_aplicacion.comboBox3.Text == "NBP")
                {
                    puntero_aplicacion.referencestate = ReferenceState.NBP;
                }

                puntero_aplicacion.luis.working_fluid.Category = puntero_aplicacion.category;
                puntero_aplicacion.luis.working_fluid.reference = puntero_aplicacion.referencestate;

                puntero_aplicacion.w_dot_net2 = Convert.ToDouble(puntero_aplicacion.textBox1.Text);
                puntero_aplicacion.t_mc_in2 = Convert.ToDouble(puntero_aplicacion.textBox2.Text);
                puntero_aplicacion.t_t_in2 = Convert.ToDouble(puntero_aplicacion.textBox4.Text);
                puntero_aplicacion.ua_lt2 = Convert.ToDouble(puntero_aplicacion.textBox17.Text);
                puntero_aplicacion.ua_ht2 = Convert.ToDouble(puntero_aplicacion.textBox16.Text);
                puntero_aplicacion.eta_mc2 = Convert.ToDouble(puntero_aplicacion.textBox14.Text);
                puntero_aplicacion.eta_rc2 = Convert.ToDouble(puntero_aplicacion.textBox13.Text);
                puntero_aplicacion.eta_pc2 = Convert.ToDouble(puntero_aplicacion.textBox24.Text);
                puntero_aplicacion.eta_t2 = Convert.ToDouble(puntero_aplicacion.textBox19.Text);
                puntero_aplicacion.n_sub_hxrs2 = Convert.ToInt64(puntero_aplicacion.textBox20.Text);
                puntero_aplicacion.p_mc_in2 = Convert.ToDouble(puntero_aplicacion.textBox3.Text);
                puntero_aplicacion.p_mc_out2 = Convert.ToDouble(puntero_aplicacion.textBox28.Text);
                puntero_aplicacion.p_pc_in2 = Convert.ToDouble(puntero_aplicacion.textBox23.Text);
                puntero_aplicacion.t_pc_in2 = Convert.ToDouble(puntero_aplicacion.textBox22.Text);
                puntero_aplicacion.p_pc_out2 = Convert.ToDouble(puntero_aplicacion.textBox8.Text);
                puntero_aplicacion.recomp_frac2 = Convert.ToDouble(puntero_aplicacion.textBox15.Text);
                puntero_aplicacion.tol2 = Convert.ToDouble(puntero_aplicacion.textBox21.Text);

                puntero_aplicacion.dp2_lt1 = Convert.ToDouble(puntero_aplicacion.textBox5.Text);
                puntero_aplicacion.dp2_lt2 = Convert.ToDouble(puntero_aplicacion.textBox26.Text);
                puntero_aplicacion.dp2_ht1 = Convert.ToDouble(puntero_aplicacion.textBox12.Text);
                puntero_aplicacion.dp2_ht2 = Convert.ToDouble(puntero_aplicacion.textBox25.Text);
                puntero_aplicacion.dp2_pc1 = Convert.ToDouble(puntero_aplicacion.textBox11.Text);
                puntero_aplicacion.dp2_phx1 = Convert.ToDouble(puntero_aplicacion.textBox10.Text);
                puntero_aplicacion.dp2_cooler2 = Convert.ToDouble(puntero_aplicacion.textBox27.Text);

                puntero_aplicacion.luis.wmm = puntero_aplicacion.luis.working_fluid.MolecularWeight;

                core.PCRCwithoutReheating cicloPCRC_withoutRH = new core.PCRCwithoutReheating();

                double UA_Total = puntero_aplicacion.ua_lt2 + puntero_aplicacion.ua_ht2;

                double LT_fraction = 0.1;

                int counter = 0;

                List<Double> recomp_frac2_list = new List<Double>();
                List<Double> p_pc_in2_list = new List<Double>();
                List<Double> p_pc_out2_list = new List<Double>();
                List<Double> eta_thermal2_list = new List<Double>();

                NLoptAlgorithm algorithm_type = NLoptAlgorithm.LN_BOBYQA;

                if (comboBox19.Text == "BOBYQA")
                    algorithm_type = NLoptAlgorithm.LN_BOBYQA;
                else if (comboBox19.Text == "COBYLA")
                    algorithm_type = NLoptAlgorithm.LN_COBYLA;
                else if (comboBox19.Text == "SUBPLEX")
                    algorithm_type = NLoptAlgorithm.LN_SBPLX;
                else if (comboBox19.Text == "NELDER-MEAD")
                    algorithm_type = NLoptAlgorithm.LN_NELDERMEAD;
                else if (comboBox19.Text == "NEWUOA")
                    algorithm_type = NLoptAlgorithm.LN_NEWUOA;
                else if (comboBox19.Text == "PRAXIS")
                    algorithm_type = NLoptAlgorithm.LN_PRAXIS;

                if (checkBox6.Checked == true)
                {
                    initial_CIP_value = Convert.ToDouble(textBox1.Text);
                }
                else
                {
                    initial_CIP_value = puntero_aplicacion.MixtureCriticalPressure;
                }

                xlWorkSheet1.Name = puntero_aplicacion.comboBox2.Text + " Mixture";

                xlWorkSheet1.Cells[1, 1] = puntero_aplicacion.comboBox2.Text + ":" + puntero_aplicacion.textBox68.Text + "," + puntero_aplicacion.comboBox6.Text + ":" + puntero_aplicacion.textBox69.Text + "," + puntero_aplicacion.comboBox12.Text + ":" + puntero_aplicacion.textBox33.Text + "," + puntero_aplicacion.comboBox7.Text + ":" + puntero_aplicacion.textBox34.Text;
                xlWorkSheet1.Cells[1, 2] = "Pcrit(kPa)";
                xlWorkSheet1.Cells[1, 3] = "Tcrit(ºC)";

                xlWorkSheet1.Cells[2, 1] = "";
                xlWorkSheet1.Cells[2, 2] = Convert.ToString(puntero_aplicacion.MixtureCriticalPressure);
                xlWorkSheet1.Cells[2, 3] = Convert.ToString(puntero_aplicacion.MixtureCriticalTemperature - 273.15);

                xlWorkSheet1.Cells[3, 1] = "";
                xlWorkSheet1.Cells[3, 2] = "";
                xlWorkSheet1.Cells[4, 3] = "";

                xlWorkSheet1.Cells[4, 1] = "PC_in(kPa)";
                xlWorkSheet1.Cells[4, 2] = "PC_out(kPa)";
                xlWorkSheet1.Cells[4, 3] = "CIT(K)";
                xlWorkSheet1.Cells[4, 4] = "LT UA(kW/K)";
                xlWorkSheet1.Cells[4, 5] = "HT UA(kW/K)";
                xlWorkSheet1.Cells[4, 6] = "Rec.Frac.";
                xlWorkSheet1.Cells[4, 7] = "Eff.(%)";
                xlWorkSheet1.Cells[4, 8] = "LTR Eff.(%)";
                xlWorkSheet1.Cells[4, 9] = "LTR Pinch(ºC)";
                xlWorkSheet1.Cells[4, 10] = "HTR Eff.(%)";
                xlWorkSheet1.Cells[4, 11] = "HTR Pinch(ºC)";
               
                using (var solver = new NLoptSolver(algorithm_type, 3, 0.01, 10000))
                {
                    var initialValue = new[] { 0.0, initial_CIP_value, initial_CIP_value };

                    double ratio_CritialPressure_CIP = Convert.ToDouble(textBox4.Text);

                    //Set Lower Bounds
                    if (checkBox7.Checked == false)
                    {
                       solver.SetLowerBounds(new[] { 0.1, initial_CIP_value, (initial_CIP_value + 500) });
                    }
                    if (checkBox7.Checked == true)
                    {
                       solver.SetLowerBounds(new[] { 0.1, (initial_CIP_value / ratio_CritialPressure_CIP), (initial_CIP_value / ratio_CritialPressure_CIP) });
                    }
                    
                    //Set  Upper Bounds
                    solver.SetUpperBounds(new[] { 1.0, 125000, (puntero_aplicacion.p_mc_out2 / 1.5) });

                    //Set Initial Step Size
                    solver.SetInitialStepSize(new[] { 0.05, 1000, 1000 });

                    //Set Initial Value
                    if (checkBox7.Checked == false)
                    {
                        initialValue = new[] { 0.2, initial_CIP_value, (initial_CIP_value + 500) };
                    }
                    else if (checkBox7.Checked == true)
                    {
                        initialValue = new[] { 0.2, (initial_CIP_value / ratio_CritialPressure_CIP), (initial_CIP_value / (ratio_CritialPressure_CIP - 0.1)) };
                    }

                    Func<double[], double> funcion = delegate (double[] variables)
                    {
                        puntero_aplicacion.luis.RecompCycle_PCRC_without_Reheating(puntero_aplicacion.luis, ref cicloPCRC_withoutRH, puntero_aplicacion.w_dot_net2, puntero_aplicacion.t_mc_in2, puntero_aplicacion.t_t_in2, variables[2], puntero_aplicacion.p_mc_out2, variables[1], puntero_aplicacion.t_pc_in2, variables[2],
                        puntero_aplicacion.ua_lt2, puntero_aplicacion.ua_ht2, puntero_aplicacion.eta_mc2, puntero_aplicacion.eta_rc2, puntero_aplicacion.eta_pc2, puntero_aplicacion.eta_t2, puntero_aplicacion.n_sub_hxrs2, variables[0], puntero_aplicacion.tol2, puntero_aplicacion.eta_thermal2, -puntero_aplicacion.dp2_lt1,
                        -puntero_aplicacion.dp2_lt2, -puntero_aplicacion.dp2_ht1, -puntero_aplicacion.dp2_ht2, -puntero_aplicacion.dp2_pc1, -puntero_aplicacion.dp2_pc2,
                        -puntero_aplicacion.dp2_phx1, -puntero_aplicacion.dp2_phx2, -puntero_aplicacion.dp2_cooler1, -puntero_aplicacion.dp2_cooler2);

                        counter++;

                        puntero_aplicacion.massflow2 = cicloPCRC_withoutRH.m_dot_turbine;
                        puntero_aplicacion.w_dot_net2 = cicloPCRC_withoutRH.W_dot_net;
                        puntero_aplicacion.eta_thermal2 = cicloPCRC_withoutRH.eta_thermal;
                        puntero_aplicacion.recomp_frac2 = variables[0];
                        puntero_aplicacion.p_pc_in2 = variables[1];
                        puntero_aplicacion.p_pc_out2 = variables[2];

                        puntero_aplicacion.temp21 = cicloPCRC_withoutRH.temp[0];
                        puntero_aplicacion.temp22 = cicloPCRC_withoutRH.temp[1];
                        puntero_aplicacion.temp23 = cicloPCRC_withoutRH.temp[2];
                        puntero_aplicacion.temp24 = cicloPCRC_withoutRH.temp[3];
                        puntero_aplicacion.temp25 = cicloPCRC_withoutRH.temp[4];
                        puntero_aplicacion.temp26 = cicloPCRC_withoutRH.temp[5];
                        puntero_aplicacion.temp27 = cicloPCRC_withoutRH.temp[6];
                        puntero_aplicacion.temp28 = cicloPCRC_withoutRH.temp[7];
                        puntero_aplicacion.temp29 = cicloPCRC_withoutRH.temp[8];
                        puntero_aplicacion.temp210 = cicloPCRC_withoutRH.temp[9];
                        puntero_aplicacion.temp213 = cicloPCRC_withoutRH.temp[10];
                        puntero_aplicacion.temp214 = cicloPCRC_withoutRH.temp[11];

                        puntero_aplicacion.pres21 = cicloPCRC_withoutRH.pres[0];
                        puntero_aplicacion.pres22 = cicloPCRC_withoutRH.pres[1];
                        puntero_aplicacion.pres23 = cicloPCRC_withoutRH.pres[2];
                        puntero_aplicacion.pres24 = cicloPCRC_withoutRH.pres[3];
                        puntero_aplicacion.pres25 = cicloPCRC_withoutRH.pres[4];
                        puntero_aplicacion.pres26 = cicloPCRC_withoutRH.pres[5];
                        puntero_aplicacion.pres27 = cicloPCRC_withoutRH.pres[6];
                        puntero_aplicacion.pres28 = cicloPCRC_withoutRH.pres[7];
                        puntero_aplicacion.pres29 = cicloPCRC_withoutRH.pres[8];
                        puntero_aplicacion.pres210 = cicloPCRC_withoutRH.pres[9];
                        puntero_aplicacion.pres213 = cicloPCRC_withoutRH.pres[10];
                        puntero_aplicacion.pres214 = cicloPCRC_withoutRH.pres[11];

                        puntero_aplicacion.PHX1 = cicloPCRC_withoutRH.PHX.Q_dot;

                        puntero_aplicacion.LT_Q = cicloPCRC_withoutRH.LT.Q_dot;
                        puntero_aplicacion.LT_mdotc = cicloPCRC_withoutRH.LT.m_dot_design[0];
                        puntero_aplicacion.LT_mdoth = cicloPCRC_withoutRH.LT.m_dot_design[1];
                        puntero_aplicacion.LT_Tcin = cicloPCRC_withoutRH.LT.T_c_in;
                        puntero_aplicacion.LT_Thin = cicloPCRC_withoutRH.LT.T_h_in;
                        puntero_aplicacion.LT_Pcin = cicloPCRC_withoutRH.LT.P_c_in;
                        puntero_aplicacion.LT_Phin = cicloPCRC_withoutRH.LT.P_h_in;
                        puntero_aplicacion.LT_Pcout = cicloPCRC_withoutRH.LT.P_c_out;
                        puntero_aplicacion.LT_Phout = cicloPCRC_withoutRH.LT.P_h_out;
                        puntero_aplicacion.LT_Effc = cicloPCRC_withoutRH.LT.eff;

                        puntero_aplicacion.HT_Q = cicloPCRC_withoutRH.HT.Q_dot;
                        puntero_aplicacion.HT_mdotc = cicloPCRC_withoutRH.HT.m_dot_design[0];
                        puntero_aplicacion.HT_mdoth = cicloPCRC_withoutRH.HT.m_dot_design[1];
                        puntero_aplicacion.HT_Tcin = cicloPCRC_withoutRH.HT.T_c_in;
                        puntero_aplicacion.HT_Thin = cicloPCRC_withoutRH.HT.T_h_in;
                        puntero_aplicacion.HT_Pcin = cicloPCRC_withoutRH.HT.P_c_in;
                        puntero_aplicacion.HT_Phin = cicloPCRC_withoutRH.HT.P_h_in;
                        puntero_aplicacion.HT_Pcout = cicloPCRC_withoutRH.HT.P_c_out;
                        puntero_aplicacion.HT_Phout = cicloPCRC_withoutRH.HT.P_h_out;
                        puntero_aplicacion.HT_Effc = cicloPCRC_withoutRH.HT.eff;

                        puntero_aplicacion.PC11 = -cicloPCRC_withoutRH.PC.Q_dot;
                        puntero_aplicacion.PC21 = -cicloPCRC_withoutRH.COOLER.Q_dot;
                                              
                        eta_thermal2_list.Add(puntero_aplicacion.eta_thermal2);
                        recomp_frac2_list.Add(puntero_aplicacion.recomp_frac2);
                        p_pc_in2_list.Add(puntero_aplicacion.p_pc_in2);
                        p_pc_out2_list.Add(puntero_aplicacion.p_pc_out2);                      

                        listBox1.Items.Add(counter.ToString());
                        listBox2.Items.Add(puntero_aplicacion.eta_thermal2.ToString());
                        listBox3.Items.Add(puntero_aplicacion.recomp_frac2.ToString());
                        listBox4.Items.Add(puntero_aplicacion.p_pc_in2.ToString());
                        listBox9.Items.Add(puntero_aplicacion.p_pc_out2.ToString());
                        listBox5.Items.Add(puntero_aplicacion.ua_lt2.ToString());
                        listBox6.Items.Add(puntero_aplicacion.ua_ht2.ToString());
                        listBox7.Items.Add(puntero_aplicacion.temp27.ToString());
                        listBox8.Items.Add(puntero_aplicacion.temp28.ToString());

                        double LTR_min_DT_1 = cicloPCRC_withoutRH.temp[7] - cicloPCRC_withoutRH.temp[2];
                        double LTR_min_DT_2 = cicloPCRC_withoutRH.temp[8] - cicloPCRC_withoutRH.temp[1];
                        double LTR_min_DT_paper = Math.Min(LTR_min_DT_1, LTR_min_DT_2);

                        double HTR_min_DT_1 = cicloPCRC_withoutRH.temp[7] - cicloPCRC_withoutRH.temp[3];
                        double HTR_min_DT_2 = cicloPCRC_withoutRH.temp[6] - cicloPCRC_withoutRH.temp[4];
                        double HTR_min_DT_paper = Math.Min(HTR_min_DT_1, HTR_min_DT_2);

                        //PC_in(kPa)
                        xlWorkSheet1.Cells[counter_Excel + 1, 1] = Convert.ToString(puntero_aplicacion.p_pc_in2);
                        //PC_out(kPa)
                        xlWorkSheet1.Cells[counter_Excel + 1, 2] = Convert.ToString(puntero_aplicacion.p_pc_out2);
                        //CIT
                        xlWorkSheet1.Cells[counter_Excel + 1, 3] = Convert.ToString(puntero_aplicacion.t_mc_in2 - 273.15);
                        //LT UA(kW/K)
                        xlWorkSheet1.Cells[counter_Excel + 1, 4] = Convert.ToString(puntero_aplicacion.ua_lt2);
                        //HT UA(kW/K)
                        xlWorkSheet1.Cells[counter_Excel + 1, 5] = Convert.ToString(puntero_aplicacion.ua_ht2);
                        //Rec.Frac.
                        xlWorkSheet1.Cells[counter_Excel + 1, 6] = puntero_aplicacion.recomp_frac2.ToString();
                        //Eff.(%)
                        xlWorkSheet1.Cells[counter_Excel + 1, 7] = (puntero_aplicacion.eta_thermal2 * 100).ToString();
                        //LTR Eff.(%)
                        xlWorkSheet1.Cells[counter_Excel + 1, 8] = cicloPCRC_withoutRH.LT.eff.ToString();
                        //LTR Pinch(ºC)
                        xlWorkSheet1.Cells[counter_Excel + 1, 9] = LTR_min_DT_paper.ToString();
                        //HTR Eff.(%)
                        xlWorkSheet1.Cells[counter_Excel + 1, 10] = cicloPCRC_withoutRH.HT.eff.ToString();
                        //HTR Pinch(ºC)
                        xlWorkSheet1.Cells[counter_Excel + 1, 11] = HTR_min_DT_paper.ToString();
                        
                        counter_Excel++;

                        return puntero_aplicacion.eta_thermal2;
                    };

                    solver.SetMaxObjective(funcion);

                    double? finalScore;

                    var result = solver.Optimize(initialValue, out finalScore);

                    Double max_eta_thermal = 0.0;

                    max_eta_thermal = eta_thermal2_list.Max();

                    var maxIndex = eta_thermal2_list.IndexOf(eta_thermal2_list.Max());

                    textBox91.Text = p_pc_in2_list[maxIndex].ToString();
                    textBox2.Text = p_pc_out2_list[maxIndex].ToString();
                    textBox90.Text = recomp_frac2_list[maxIndex].ToString();
                    textBox86.Text = eta_thermal2_list[maxIndex].ToString();

                    //Copy results as design-point inputs
                    if (checkBox3.Checked == true)
                    {
                        puntero_aplicacion.textBox15.Text = recomp_frac2_list[maxIndex].ToString();
                        puntero_aplicacion.textBox23.Text = p_pc_in2_list[maxIndex].ToString();
                        puntero_aplicacion.textBox8.Text = p_pc_out2_list[maxIndex].ToString();
                        puntero_aplicacion.textBox3.Text = p_pc_out2_list[maxIndex].ToString();
                    }

                    //Closing Excel Book
                    xlWorkBook1.SaveAs(textBox3.Text + "PCRC_without_ReHeating_" + xlWorkSheet1.Name + ".xls", Excel.XlFileFormat.xlWorkbookNormal, misValue1, misValue1, misValue1, misValue1, Excel.XlSaveAsAccessMode.xlExclusive, misValue1, misValue1, misValue1, misValue1, misValue1);

                    xlWorkBook1.Close(true, misValue1, misValue1);
                    xlApp1.Quit();

                    releaseObject(xlWorkSheet1);
                    //releaseObject(xlWorkSheet2);
                    releaseObject(xlWorkBook1);
                    releaseObject(xlApp1);
                }
            }

            //-------------------------------------------------------------------------

            //Optimization UA true
            else if (checkBox2.Checked == true)
            {
                //PureFluid
                if (puntero_aplicacion.comboBox1.Text == "PureFluid")
                {
                    puntero_aplicacion.category = RefrigerantCategory.PureFluid;
                    puntero_aplicacion.luis.core1(puntero_aplicacion.comboBox1.Text, puntero_aplicacion.category);
                }

                //NewMixture
                if (puntero_aplicacion.comboBox1.Text == "NewMixture")
                {
                    puntero_aplicacion.category = RefrigerantCategory.NewMixture;
                    puntero_aplicacion.luis.core1(puntero_aplicacion.comboBox2.Text + "=" + puntero_aplicacion.textBox68.Text + "," + puntero_aplicacion.comboBox6.Text + "=" + puntero_aplicacion.textBox69.Text + "," + puntero_aplicacion.comboBox12.Text + "=" + puntero_aplicacion.textBox33.Text + "," + puntero_aplicacion.comboBox7.Text + "=" + puntero_aplicacion.textBox34.Text, puntero_aplicacion.category);
                }

                if (puntero_aplicacion.comboBox1.Text == "PredefinedMixture")
                {
                    puntero_aplicacion.category = RefrigerantCategory.PredefinedMixture;
                }

                if (puntero_aplicacion.comboBox1.Text == "PseudoPureFluid")
                {
                    puntero_aplicacion.category = RefrigerantCategory.PseudoPureFluid;
                }

                if (puntero_aplicacion.comboBox3.Text == "DEF")
                {
                    puntero_aplicacion.referencestate = ReferenceState.DEF;
                }
                if (puntero_aplicacion.comboBox3.Text == "ASH")
                {
                    puntero_aplicacion.referencestate = ReferenceState.ASH;
                }
                if (puntero_aplicacion.comboBox3.Text == "IIR")
                {
                    puntero_aplicacion.referencestate = ReferenceState.IIR;
                }
                if (puntero_aplicacion.comboBox3.Text == "NBP")
                {
                    puntero_aplicacion.referencestate = ReferenceState.NBP;
                }

                puntero_aplicacion.luis.working_fluid.Category = puntero_aplicacion.category;
                puntero_aplicacion.luis.working_fluid.reference = puntero_aplicacion.referencestate;

                puntero_aplicacion.w_dot_net2 = Convert.ToDouble(puntero_aplicacion.textBox1.Text);
                puntero_aplicacion.t_mc_in2 = Convert.ToDouble(puntero_aplicacion.textBox2.Text);
                puntero_aplicacion.t_t_in2 = Convert.ToDouble(puntero_aplicacion.textBox4.Text);
                puntero_aplicacion.ua_lt2 = Convert.ToDouble(puntero_aplicacion.textBox17.Text);
                puntero_aplicacion.ua_ht2 = Convert.ToDouble(puntero_aplicacion.textBox16.Text);
                puntero_aplicacion.eta_mc2 = Convert.ToDouble(puntero_aplicacion.textBox14.Text);
                puntero_aplicacion.eta_rc2 = Convert.ToDouble(puntero_aplicacion.textBox13.Text);
                puntero_aplicacion.eta_pc2 = Convert.ToDouble(puntero_aplicacion.textBox24.Text);
                puntero_aplicacion.eta_t2 = Convert.ToDouble(puntero_aplicacion.textBox19.Text);
                puntero_aplicacion.n_sub_hxrs2 = Convert.ToInt64(puntero_aplicacion.textBox20.Text);
                puntero_aplicacion.p_mc_in2 = Convert.ToDouble(puntero_aplicacion.textBox3.Text);
                puntero_aplicacion.p_mc_out2 = Convert.ToDouble(puntero_aplicacion.textBox28.Text);
                puntero_aplicacion.p_pc_in2 = Convert.ToDouble(puntero_aplicacion.textBox23.Text);
                puntero_aplicacion.t_pc_in2 = Convert.ToDouble(puntero_aplicacion.textBox22.Text);
                puntero_aplicacion.p_pc_out2 = Convert.ToDouble(puntero_aplicacion.textBox8.Text);
                puntero_aplicacion.recomp_frac2 = Convert.ToDouble(puntero_aplicacion.textBox15.Text);
                puntero_aplicacion.tol2 = Convert.ToDouble(puntero_aplicacion.textBox21.Text);

                puntero_aplicacion.dp2_lt1 = Convert.ToDouble(puntero_aplicacion.textBox5.Text);
                puntero_aplicacion.dp2_lt2 = Convert.ToDouble(puntero_aplicacion.textBox26.Text);
                puntero_aplicacion.dp2_ht1 = Convert.ToDouble(puntero_aplicacion.textBox12.Text);
                puntero_aplicacion.dp2_ht2 = Convert.ToDouble(puntero_aplicacion.textBox25.Text);
                puntero_aplicacion.dp2_pc1 = Convert.ToDouble(puntero_aplicacion.textBox11.Text);
                puntero_aplicacion.dp2_phx1 = Convert.ToDouble(puntero_aplicacion.textBox10.Text);
                puntero_aplicacion.dp2_cooler2 = Convert.ToDouble(puntero_aplicacion.textBox27.Text);

                puntero_aplicacion.luis.wmm = puntero_aplicacion.luis.working_fluid.MolecularWeight;

                core.PCRCwithoutReheating cicloPCRC_withoutRH = new core.PCRCwithoutReheating();

                double UA_Total = puntero_aplicacion.ua_lt2 + puntero_aplicacion.ua_ht2;

                double LT_fraction = 0.1;

                int counter = 0;

                List<Double> recomp_frac2_list = new List<Double>();
                List<Double> p_pc_in2_list = new List<Double>();
                List<Double> p_pc_out2_list = new List<Double>();
                List<Double> eta_thermal2_list = new List<Double>();
                List<Double> ua_LT_list = new List<Double>();
                List<Double> ua_HT_list = new List<Double>();

                NLoptAlgorithm algorithm_type = NLoptAlgorithm.LN_BOBYQA;

                if (comboBox19.Text == "BOBYQA")
                    algorithm_type = NLoptAlgorithm.LN_BOBYQA;
                else if (comboBox19.Text == "COBYLA")
                    algorithm_type = NLoptAlgorithm.LN_COBYLA;
                else if (comboBox19.Text == "SUBPLEX")
                    algorithm_type = NLoptAlgorithm.LN_SBPLX;
                else if (comboBox19.Text == "NELDER-MEAD")
                    algorithm_type = NLoptAlgorithm.LN_NELDERMEAD;
                else if (comboBox19.Text == "NEWUOA")
                    algorithm_type = NLoptAlgorithm.LN_NEWUOA;
                else if (comboBox19.Text == "PRAXIS")
                    algorithm_type = NLoptAlgorithm.LN_PRAXIS;

                if (checkBox6.Checked == true)
                {
                    initial_CIP_value = Convert.ToDouble(textBox1.Text);
                }
                else
                {
                    initial_CIP_value = puntero_aplicacion.MixtureCriticalPressure;
                }

                xlWorkSheet1.Name = puntero_aplicacion.comboBox2.Text + " Mixture";

                xlWorkSheet1.Cells[1, 1] = puntero_aplicacion.comboBox2.Text + ":" + puntero_aplicacion.textBox68.Text + "," + puntero_aplicacion.comboBox6.Text + ":" + puntero_aplicacion.textBox69.Text + "," + puntero_aplicacion.comboBox12.Text + ":" + puntero_aplicacion.textBox33.Text + "," + puntero_aplicacion.comboBox7.Text + ":" + puntero_aplicacion.textBox34.Text;
                xlWorkSheet1.Cells[1, 2] = "Pcrit(kPa)";
                xlWorkSheet1.Cells[1, 3] = "Tcrit(ºC)";

                xlWorkSheet1.Cells[2, 1] = "";
                xlWorkSheet1.Cells[2, 2] = Convert.ToString(puntero_aplicacion.MixtureCriticalPressure);
                xlWorkSheet1.Cells[2, 3] = Convert.ToString(puntero_aplicacion.MixtureCriticalTemperature - 273.15);

                xlWorkSheet1.Cells[3, 1] = "";
                xlWorkSheet1.Cells[3, 2] = "";
                xlWorkSheet1.Cells[4, 3] = "";

                xlWorkSheet1.Cells[4, 1] = "PC_in(kPa)";
                xlWorkSheet1.Cells[4, 2] = "PC_out(kPa)";
                xlWorkSheet1.Cells[4, 3] = "CIT(K)";
                xlWorkSheet1.Cells[4, 4] = "LT UA(kW/K)";
                xlWorkSheet1.Cells[4, 5] = "HT UA(kW/K)";
                xlWorkSheet1.Cells[4, 6] = "Rec.Frac.";
                xlWorkSheet1.Cells[4, 7] = "Eff.(%)";
                xlWorkSheet1.Cells[4, 8] = "LTR Eff.(%)";
                xlWorkSheet1.Cells[4, 9] = "LTR Pinch(ºC)";
                xlWorkSheet1.Cells[4, 10] = "HTR Eff.(%)";
                xlWorkSheet1.Cells[4, 11] = "HTR Pinch(ºC)";               

                using (var solver = new NLoptSolver(algorithm_type, 4, 0.01, 10000))
                {
                    solver.SetLowerBounds(new[] { 0.1, initial_CIP_value, (initial_CIP_value + 500), 0.2 });
                    solver.SetUpperBounds(new[] { 1.0, 125000, (puntero_aplicacion.p_mc_out2 / 1.5), 0.8 });

                    solver.SetInitialStepSize(new[] { 0.05, 100, 100, 0.05 });

                    var initialValue = new[] { 0.2, initial_CIP_value, (initial_CIP_value + 500), 0.5 };

                    Func<double[], double> funcion = delegate (double[] variables)
                    {
                        puntero_aplicacion.luis.RecompCycle_PCRC_without_Reheating_for_Optimization(puntero_aplicacion.luis, ref cicloPCRC_withoutRH, puntero_aplicacion.w_dot_net2, puntero_aplicacion.t_mc_in2, puntero_aplicacion.t_t_in2, variables[2], puntero_aplicacion.p_mc_out2, variables[1], puntero_aplicacion.t_pc_in2, variables[2],
                        variables[3], UA_Total, puntero_aplicacion.eta_mc2, puntero_aplicacion.eta_rc2, puntero_aplicacion.eta_pc2, puntero_aplicacion.eta_t2, puntero_aplicacion.n_sub_hxrs2, variables[0], puntero_aplicacion.tol2, puntero_aplicacion.eta_thermal2, -puntero_aplicacion.dp2_lt1, -puntero_aplicacion.dp2_lt2, -puntero_aplicacion.dp2_ht1, -puntero_aplicacion.dp2_ht2, -puntero_aplicacion.dp2_pc1, -puntero_aplicacion.dp2_pc2,
                        -puntero_aplicacion.dp2_phx1, -puntero_aplicacion.dp2_phx2, -puntero_aplicacion.dp2_cooler1, -puntero_aplicacion.dp2_cooler2);

                        counter++;

                        puntero_aplicacion.massflow2 = cicloPCRC_withoutRH.m_dot_turbine;
                        puntero_aplicacion.w_dot_net2 = cicloPCRC_withoutRH.W_dot_net;
                        puntero_aplicacion.eta_thermal2 = cicloPCRC_withoutRH.eta_thermal;
                        puntero_aplicacion.recomp_frac2 = variables[0];
                        puntero_aplicacion.p_pc_in2 = variables[1];
                        puntero_aplicacion.p_pc_out2 = variables[2];
                        LT_fraction = variables[3];
                        puntero_aplicacion.ua_lt2 = UA_Total * LT_fraction;
                        puntero_aplicacion.ua_ht2 = UA_Total * (1 - LT_fraction);

                        puntero_aplicacion.temp21 = cicloPCRC_withoutRH.temp[0];
                        puntero_aplicacion.temp22 = cicloPCRC_withoutRH.temp[1];
                        puntero_aplicacion.temp23 = cicloPCRC_withoutRH.temp[2];
                        puntero_aplicacion.temp24 = cicloPCRC_withoutRH.temp[3];
                        puntero_aplicacion.temp25 = cicloPCRC_withoutRH.temp[4];
                        puntero_aplicacion.temp26 = cicloPCRC_withoutRH.temp[5];
                        puntero_aplicacion.temp27 = cicloPCRC_withoutRH.temp[6];
                        puntero_aplicacion.temp28 = cicloPCRC_withoutRH.temp[7];
                        puntero_aplicacion.temp29 = cicloPCRC_withoutRH.temp[8];
                        puntero_aplicacion.temp210 = cicloPCRC_withoutRH.temp[9];
                        puntero_aplicacion.temp213 = cicloPCRC_withoutRH.temp[10];
                        puntero_aplicacion.temp214 = cicloPCRC_withoutRH.temp[11];

                        puntero_aplicacion.pres21 = cicloPCRC_withoutRH.pres[0];
                        puntero_aplicacion.pres22 = cicloPCRC_withoutRH.pres[1];
                        puntero_aplicacion.pres23 = cicloPCRC_withoutRH.pres[2];
                        puntero_aplicacion.pres24 = cicloPCRC_withoutRH.pres[3];
                        puntero_aplicacion.pres25 = cicloPCRC_withoutRH.pres[4];
                        puntero_aplicacion.pres26 = cicloPCRC_withoutRH.pres[5];
                        puntero_aplicacion.pres27 = cicloPCRC_withoutRH.pres[6];
                        puntero_aplicacion.pres28 = cicloPCRC_withoutRH.pres[7];
                        puntero_aplicacion.pres29 = cicloPCRC_withoutRH.pres[8];
                        puntero_aplicacion.pres210 = cicloPCRC_withoutRH.pres[9];
                        puntero_aplicacion.pres213 = cicloPCRC_withoutRH.pres[10];
                        puntero_aplicacion.pres214 = cicloPCRC_withoutRH.pres[11];

                        puntero_aplicacion.PHX1 = cicloPCRC_withoutRH.PHX.Q_dot;

                        puntero_aplicacion.LT_Q = cicloPCRC_withoutRH.LT.Q_dot;
                        puntero_aplicacion.LT_mdotc = cicloPCRC_withoutRH.LT.m_dot_design[0];
                        puntero_aplicacion.LT_mdoth = cicloPCRC_withoutRH.LT.m_dot_design[1];
                        puntero_aplicacion.LT_Tcin = cicloPCRC_withoutRH.LT.T_c_in;
                        puntero_aplicacion.LT_Thin = cicloPCRC_withoutRH.LT.T_h_in;
                        puntero_aplicacion.LT_Pcin = cicloPCRC_withoutRH.LT.P_c_in;
                        puntero_aplicacion.LT_Phin = cicloPCRC_withoutRH.LT.P_h_in;
                        puntero_aplicacion.LT_Pcout = cicloPCRC_withoutRH.LT.P_c_out;
                        puntero_aplicacion.LT_Phout = cicloPCRC_withoutRH.LT.P_h_out;
                        puntero_aplicacion.LT_Effc = cicloPCRC_withoutRH.LT.eff;

                        puntero_aplicacion.HT_Q = cicloPCRC_withoutRH.HT.Q_dot;
                        puntero_aplicacion.HT_mdotc = cicloPCRC_withoutRH.HT.m_dot_design[0];
                        puntero_aplicacion.HT_mdoth = cicloPCRC_withoutRH.HT.m_dot_design[1];
                        puntero_aplicacion.HT_Tcin = cicloPCRC_withoutRH.HT.T_c_in;
                        puntero_aplicacion.HT_Thin = cicloPCRC_withoutRH.HT.T_h_in;
                        puntero_aplicacion.HT_Pcin = cicloPCRC_withoutRH.HT.P_c_in;
                        puntero_aplicacion.HT_Phin = cicloPCRC_withoutRH.HT.P_h_in;
                        puntero_aplicacion.HT_Pcout = cicloPCRC_withoutRH.HT.P_c_out;
                        puntero_aplicacion.HT_Phout = cicloPCRC_withoutRH.HT.P_h_out;
                        puntero_aplicacion.HT_Effc = cicloPCRC_withoutRH.HT.eff;

                        puntero_aplicacion.PC11 = -cicloPCRC_withoutRH.PC.Q_dot;
                        puntero_aplicacion.PC21 = -cicloPCRC_withoutRH.COOLER.Q_dot;
                                                
                        eta_thermal2_list.Add(puntero_aplicacion.eta_thermal2);
                        recomp_frac2_list.Add(puntero_aplicacion.recomp_frac2);
                        p_pc_in2_list.Add(puntero_aplicacion.p_pc_in2);
                        p_pc_out2_list.Add(puntero_aplicacion.p_pc_out2);
                        ua_LT_list.Add(puntero_aplicacion.ua_lt2);
                        ua_HT_list.Add(puntero_aplicacion.ua_ht2);

                        listBox1.Items.Add(counter.ToString());
                        listBox2.Items.Add(puntero_aplicacion.eta_thermal2.ToString());
                        listBox3.Items.Add(puntero_aplicacion.recomp_frac2.ToString());
                        listBox4.Items.Add(puntero_aplicacion.p_pc_in2.ToString());
                        listBox9.Items.Add(puntero_aplicacion.p_pc_out2.ToString());
                        listBox5.Items.Add(puntero_aplicacion.ua_lt2.ToString());
                        listBox6.Items.Add(puntero_aplicacion.ua_ht2.ToString());

                        double LTR_min_DT_1 = cicloPCRC_withoutRH.temp[7] - cicloPCRC_withoutRH.temp[2];
                        double LTR_min_DT_2 = cicloPCRC_withoutRH.temp[8] - cicloPCRC_withoutRH.temp[1];
                        double LTR_min_DT_paper = Math.Min(LTR_min_DT_1, LTR_min_DT_2);

                        double HTR_min_DT_1 = cicloPCRC_withoutRH.temp[7] - cicloPCRC_withoutRH.temp[3];
                        double HTR_min_DT_2 = cicloPCRC_withoutRH.temp[6] - cicloPCRC_withoutRH.temp[4];
                        double HTR_min_DT_paper = Math.Min(HTR_min_DT_1, HTR_min_DT_2);

                        //PC_in(kPa)
                        xlWorkSheet1.Cells[counter_Excel + 1, 1] = Convert.ToString(puntero_aplicacion.p_pc_in2);
                        //PC_out(kPa)
                        xlWorkSheet1.Cells[counter_Excel + 1, 2] = Convert.ToString(puntero_aplicacion.p_pc_out2);
                        //CIT
                        xlWorkSheet1.Cells[counter_Excel + 1, 3] = Convert.ToString(puntero_aplicacion.t_mc_in2 - 273.15);
                        //LT UA(kW/K)
                        xlWorkSheet1.Cells[counter_Excel + 1, 4] = Convert.ToString(puntero_aplicacion.ua_lt2);
                        //HT UA(kW/K)
                        xlWorkSheet1.Cells[counter_Excel + 1, 5] = Convert.ToString(puntero_aplicacion.ua_ht2);
                        //Rec.Frac.
                        xlWorkSheet1.Cells[counter_Excel + 1, 6] = puntero_aplicacion.recomp_frac2.ToString();
                        //Eff.(%)
                        xlWorkSheet1.Cells[counter_Excel + 1, 7] = (puntero_aplicacion.eta_thermal2 * 100).ToString();
                        //LTR Eff.(%)
                        xlWorkSheet1.Cells[counter_Excel + 1, 8] = cicloPCRC_withoutRH.LT.eff.ToString();
                        //LTR Pinch(ºC)
                        xlWorkSheet1.Cells[counter_Excel + 1, 9] = LTR_min_DT_paper.ToString();
                        //HTR Eff.(%)
                        xlWorkSheet1.Cells[counter_Excel + 1, 10] = cicloPCRC_withoutRH.HT.eff.ToString();
                        //HTR Pinch(ºC)
                        xlWorkSheet1.Cells[counter_Excel + 1, 11] = HTR_min_DT_paper.ToString();
                       
                        counter_Excel++;

                        return puntero_aplicacion.eta_thermal2;
                    };

                    solver.SetMaxObjective(funcion);

                    double? finalScore;

                    var result = solver.Optimize(initialValue, out finalScore);

                    Double max_eta_thermal = 0.0;

                    max_eta_thermal = eta_thermal2_list.Max();

                    var maxIndex = eta_thermal2_list.IndexOf(eta_thermal2_list.Max());

                    textBox91.Text = p_pc_in2_list[maxIndex].ToString();
                    textBox2.Text = p_pc_out2_list[maxIndex].ToString();
                    textBox90.Text = recomp_frac2_list[maxIndex].ToString();
                    textBox86.Text = eta_thermal2_list[maxIndex].ToString();
                    textBox82.Text = ua_LT_list[maxIndex].ToString();
                    textBox83.Text = ua_HT_list[maxIndex].ToString();                  
                    
                    //Copy results as design-point inputs
                    if (checkBox3.Checked == true)
                    {
                        puntero_aplicacion.textBox15.Text = recomp_frac2_list[maxIndex].ToString();
                        puntero_aplicacion.textBox23.Text = p_pc_in2_list[maxIndex].ToString();
                        puntero_aplicacion.textBox8.Text = p_pc_out2_list[maxIndex].ToString();
                        puntero_aplicacion.textBox3.Text = p_pc_out2_list[maxIndex].ToString();
                        puntero_aplicacion.textBox17.Text = ua_LT_list[maxIndex].ToString();
                        puntero_aplicacion.textBox16.Text = ua_HT_list[maxIndex].ToString();
                    }

                    //Closing Excel Book
                    xlWorkBook1.SaveAs(textBox3.Text + "PCRC_without_ReHeating_" + xlWorkSheet1.Name + ".xls", Excel.XlFileFormat.xlWorkbookNormal, misValue1, misValue1, misValue1, misValue1, Excel.XlSaveAsAccessMode.xlExclusive, misValue1, misValue1, misValue1, misValue1, misValue1);

                    xlWorkBook1.Close(true, misValue1, misValue1);
                    xlApp1.Quit();

                    releaseObject(xlWorkSheet1);
                    //releaseObject(xlWorkSheet2);
                    releaseObject(xlWorkBook1);
                    releaseObject(xlApp1);
                }
            }            
        }

        private void Button6_Click(object sender, EventArgs e)
        {

        }

        private void Button5_Click(object sender, EventArgs e)
        {

        }

        private void Button2_Click(object sender, EventArgs e)
        {

        }

        private void Button9_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
            listBox2.Items.Clear();
            listBox3.Items.Clear();
            listBox4.Items.Clear();
            listBox5.Items.Clear();
            listBox6.Items.Clear();
            listBox7.Items.Clear();
            listBox8.Items.Clear();
            listBox9.Items.Clear();
            listBox10.Items.Clear();
            listBox11.Items.Clear();
            listBox12.Items.Clear();
            listBox13.Items.Clear();
            listBox14.Items.Clear();
            listBox15.Items.Clear();
            listBox16.Items.Clear();
            listBox17.Items.Clear();
            listBox18.Items.Clear();
        }

        //Run CIT Optimization
        private void Button7_Click(object sender, EventArgs e)
        {
            int counter = 0;

            double initial_pc_in_value = 0;
            double initial_pc_out_value = 0;

            int counter_Excel = 4;

            Excel.Application xlApp1;
            Excel.Workbook xlWorkBook1;
            Excel.Worksheet xlWorkSheet1;
            //Excel.Worksheet xlWorkSheet2;

            object misValue1 = System.Reflection.Missing.Value;

            xlApp1 = new Excel.Application();
            xlApp1.DisplayAlerts = false;
            xlWorkBook1 = xlApp1.Workbooks.Add(misValue1);

            xlWorkSheet1 = (Excel.Worksheet)xlWorkBook1.Worksheets.Add();

            //xlWorkSheet1 = (Excel.Worksheet)xlWorkBook1.Worksheets.get_Item(xlWorkBook1.Worksheets.Count);    

            //Loop for UA optimization
            for (double j = Convert.ToDouble(textBox11.Text); j <= Convert.ToDouble(textBox10.Text); j = j + Convert.ToDouble(textBox9.Text))
            {
                puntero_aplicacion.ua_lt2 = j / 2;
                puntero_aplicacion.ua_ht2 = j / 2;


                //Loop for CIT optimization
                //for (double i = 305.15; i <= 335.15; i = i + 5)
                for (double i = Convert.ToDouble(textBox57.Text); i <= Convert.ToDouble(textBox56.Text); i = i + Convert.ToDouble(textBox55.Text))
                {
                    counter = 0;

                    //UA Optimization false
                    if (checkBox2.Checked == false)
                    {
                        //PureFluid
                        if (puntero_aplicacion.comboBox1.Text == "PureFluid")
                        {
                            puntero_aplicacion.category = RefrigerantCategory.PureFluid;
                            puntero_aplicacion.luis.core1(puntero_aplicacion.comboBox1.Text, puntero_aplicacion.category);
                        }

                        //NewMixture
                        if (puntero_aplicacion.comboBox1.Text == "NewMixture")
                        {
                            puntero_aplicacion.category = RefrigerantCategory.NewMixture;
                            puntero_aplicacion.luis.core1(puntero_aplicacion.comboBox2.Text + "=" + puntero_aplicacion.textBox68.Text + "," + puntero_aplicacion.comboBox6.Text + "=" + puntero_aplicacion.textBox69.Text + "," + puntero_aplicacion.comboBox12.Text + "=" + puntero_aplicacion.textBox33.Text + "," + puntero_aplicacion.comboBox7.Text + "=" + puntero_aplicacion.textBox34.Text, puntero_aplicacion.category);
                        }

                        if (puntero_aplicacion.comboBox1.Text == "PredefinedMixture")
                        {
                            puntero_aplicacion.category = RefrigerantCategory.PredefinedMixture;
                        }

                        if (puntero_aplicacion.comboBox1.Text == "PseudoPureFluid")
                        {
                            puntero_aplicacion.category = RefrigerantCategory.PseudoPureFluid;
                        }

                        if (puntero_aplicacion.comboBox3.Text == "DEF")
                        {
                            puntero_aplicacion.referencestate = ReferenceState.DEF;
                        }
                        if (puntero_aplicacion.comboBox3.Text == "ASH")
                        {
                            puntero_aplicacion.referencestate = ReferenceState.ASH;
                        }
                        if (puntero_aplicacion.comboBox3.Text == "IIR")
                        {
                            puntero_aplicacion.referencestate = ReferenceState.IIR;
                        }
                        if (puntero_aplicacion.comboBox3.Text == "NBP")
                        {
                            puntero_aplicacion.referencestate = ReferenceState.NBP;
                        }

                        puntero_aplicacion.luis.working_fluid.Category = puntero_aplicacion.category;
                        puntero_aplicacion.luis.working_fluid.reference = puntero_aplicacion.referencestate;

                        puntero_aplicacion.w_dot_net2 = Convert.ToDouble(puntero_aplicacion.textBox1.Text);
                        puntero_aplicacion.t_mc_in2 = Convert.ToDouble(puntero_aplicacion.textBox2.Text);
                        puntero_aplicacion.t_t_in2 = Convert.ToDouble(puntero_aplicacion.textBox4.Text);
                        //puntero_aplicacion.ua_lt2 = Convert.ToDouble(puntero_aplicacion.textBox17.Text);
                        //puntero_aplicacion.ua_ht2 = Convert.ToDouble(puntero_aplicacion.textBox16.Text);
                        puntero_aplicacion.eta_mc2 = Convert.ToDouble(puntero_aplicacion.textBox14.Text);
                        puntero_aplicacion.eta_rc2 = Convert.ToDouble(puntero_aplicacion.textBox13.Text);
                        puntero_aplicacion.eta_pc2 = Convert.ToDouble(puntero_aplicacion.textBox24.Text);
                        puntero_aplicacion.eta_t2 = Convert.ToDouble(puntero_aplicacion.textBox19.Text);
                        puntero_aplicacion.n_sub_hxrs2 = Convert.ToInt64(puntero_aplicacion.textBox20.Text);
                        puntero_aplicacion.p_mc_in2 = Convert.ToDouble(puntero_aplicacion.textBox3.Text);
                        puntero_aplicacion.p_mc_out2 = Convert.ToDouble(puntero_aplicacion.textBox28.Text);
                        puntero_aplicacion.p_pc_in2 = Convert.ToDouble(puntero_aplicacion.textBox23.Text);
                        puntero_aplicacion.t_pc_in2 = Convert.ToDouble(puntero_aplicacion.textBox22.Text);
                        puntero_aplicacion.p_pc_out2 = Convert.ToDouble(puntero_aplicacion.textBox8.Text);
                        puntero_aplicacion.recomp_frac2 = Convert.ToDouble(puntero_aplicacion.textBox15.Text);
                        puntero_aplicacion.tol2 = Convert.ToDouble(puntero_aplicacion.textBox21.Text);

                        puntero_aplicacion.dp2_lt1 = Convert.ToDouble(puntero_aplicacion.textBox5.Text);
                        puntero_aplicacion.dp2_lt2 = Convert.ToDouble(puntero_aplicacion.textBox26.Text);
                        puntero_aplicacion.dp2_ht1 = Convert.ToDouble(puntero_aplicacion.textBox12.Text);
                        puntero_aplicacion.dp2_ht2 = Convert.ToDouble(puntero_aplicacion.textBox25.Text);
                        puntero_aplicacion.dp2_pc1 = Convert.ToDouble(puntero_aplicacion.textBox11.Text);
                        puntero_aplicacion.dp2_phx1 = Convert.ToDouble(puntero_aplicacion.textBox10.Text);
                        puntero_aplicacion.dp2_cooler2 = Convert.ToDouble(puntero_aplicacion.textBox27.Text);

                        puntero_aplicacion.luis.wmm = puntero_aplicacion.luis.working_fluid.MolecularWeight;

                        core.PCRCwithoutReheating cicloPCRC_withoutRH = new core.PCRCwithoutReheating();

                        double UA_Total = puntero_aplicacion.ua_lt2 + puntero_aplicacion.ua_ht2;

                        double LT_fraction = 0.1;

                        //int counter = 0;

                        List<Double> massflow2_list = new List<Double>();
                        List<Double> recomp_frac2_list = new List<Double>();
                        List<Double> p_pc_in2_list = new List<Double>();
                        List<Double> p_pc_out2_list = new List<Double>();
                        List<Double> eta_thermal2_list = new List<Double>();
                        List<Double> PHX_Q2_list = new List<Double>();

                        List<Double> t1_list = new List<Double>();
                        List<Double> t2_list = new List<Double>();
                        List<Double> t3_list = new List<Double>();
                        List<Double> t4_list = new List<Double>();
                        List<Double> t5_list = new List<Double>();
                        List<Double> t6_list = new List<Double>();
                        List<Double> t7_list = new List<Double>();
                        List<Double> t8_list = new List<Double>();
                        List<Double> t9_list = new List<Double>();
                        List<Double> t10_list = new List<Double>();
                        List<Double> t13_list = new List<Double>();
                        List<Double> t14_list = new List<Double>();

                        List<Double> p1_list = new List<Double>();
                        List<Double> p2_list = new List<Double>();
                        List<Double> p3_list = new List<Double>();
                        List<Double> p4_list = new List<Double>();
                        List<Double> p5_list = new List<Double>();
                        List<Double> p6_list = new List<Double>();
                        List<Double> p7_list = new List<Double>();
                        List<Double> p8_list = new List<Double>();
                        List<Double> p9_list = new List<Double>();
                        List<Double> p10_list = new List<Double>();
                        List<Double> p13_list = new List<Double>();
                        List<Double> p14_list = new List<Double>();

                        List<Double> HT_Eff_list = new List<Double>();
                        List<Double> LT_Eff_list = new List<Double>();

                        NLoptAlgorithm algorithm_type = NLoptAlgorithm.LN_BOBYQA;

                        if (comboBox19.Text == "BOBYQA")
                            algorithm_type = NLoptAlgorithm.LN_BOBYQA;
                        else if (comboBox19.Text == "COBYLA")
                            algorithm_type = NLoptAlgorithm.LN_COBYLA;
                        else if (comboBox19.Text == "SUBPLEX")
                            algorithm_type = NLoptAlgorithm.LN_SBPLX;
                        else if (comboBox19.Text == "NELDER-MEAD")
                            algorithm_type = NLoptAlgorithm.LN_NELDERMEAD;
                        else if (comboBox19.Text == "NEWUOA")
                            algorithm_type = NLoptAlgorithm.LN_NEWUOA;
                        else if (comboBox19.Text == "PRAXIS")
                            algorithm_type = NLoptAlgorithm.LN_PRAXIS;

                        if (i == Convert.ToDouble(textBox57.Text))
                        {
                            if (checkBox6.Checked == true)
                            {
                                initial_pc_in_value = Convert.ToDouble(textBox1.Text);
                                initial_pc_out_value = Convert.ToDouble(textBox1.Text) + 500;
                            }
                            else
                            {
                                initial_pc_in_value = puntero_aplicacion.MixtureCriticalPressure;
                                initial_pc_out_value = puntero_aplicacion.MixtureCriticalPressure + 500;
                            }

                            xlWorkSheet1.Name = puntero_aplicacion.comboBox2.Text + " Mixture";

                            xlWorkSheet1.Cells[1, 1] = puntero_aplicacion.comboBox2.Text + ":" + puntero_aplicacion.textBox68.Text + "," + puntero_aplicacion.comboBox6.Text + ":" + puntero_aplicacion.textBox69.Text + "," + puntero_aplicacion.comboBox12.Text + ":" + puntero_aplicacion.textBox33.Text + "," + puntero_aplicacion.comboBox7.Text + ":" + puntero_aplicacion.textBox34.Text;
                            xlWorkSheet1.Cells[1, 2] = "Pcrit(kPa)";
                            xlWorkSheet1.Cells[1, 3] = "Tcrit(ºC)";

                            xlWorkSheet1.Cells[2, 1] = "";
                            xlWorkSheet1.Cells[2, 2] = Convert.ToString(puntero_aplicacion.MixtureCriticalPressure);
                            xlWorkSheet1.Cells[2, 3] = Convert.ToString(puntero_aplicacion.MixtureCriticalTemperature - 273.15);

                            xlWorkSheet1.Cells[3, 1] = "";
                            xlWorkSheet1.Cells[3, 2] = "";
                            xlWorkSheet1.Cells[4, 3] = "";

                            xlWorkSheet1.Cells[4, 1] = "PC_in(kPa)";
                            xlWorkSheet1.Cells[4, 2] = "PC_out(kPa)";
                            xlWorkSheet1.Cells[4, 3] = "CIT(K)";
                            xlWorkSheet1.Cells[4, 4] = "LT UA(kW/K)";
                            xlWorkSheet1.Cells[4, 5] = "HT UA(kW/K)";
                            xlWorkSheet1.Cells[4, 6] = "Rec.Frac.";
                            xlWorkSheet1.Cells[4, 7] = "Eff.(%)";
                            xlWorkSheet1.Cells[4, 8] = "LTR Eff.(%)";
                            xlWorkSheet1.Cells[4, 9] = "LTR Pinch(ºC)";
                            xlWorkSheet1.Cells[4, 10] = "HTR Eff.(%)";
                            xlWorkSheet1.Cells[4, 11] = "HTR Pinch(ºC)";
                            xlWorkSheet1.Cells[4, 12] = "PTC_Apperture_Area(m2)";
                            xlWorkSheet1.Cells[4, 13] = "PTC_Pressure_Drop(bar)";
                            xlWorkSheet1.Cells[4, 14] = "LF_Apperture_Area(m2)";
                            xlWorkSheet1.Cells[4, 15] = "LF_Pressure_Drop(bar)";
                        }

                        using (var solver = new NLoptSolver(algorithm_type, 3, 0.01, 10000))
                        {
                            solver.SetLowerBounds(new[] { 0.1, initial_pc_in_value, initial_pc_out_value });
                            solver.SetUpperBounds(new[] { 1.0, 125000, (puntero_aplicacion.p_mc_out2 / 1.5) });

                            solver.SetInitialStepSize(new[] { 0.05, 100, 100 });

                            var initialValue = new[] { 0.2, initial_pc_in_value, initial_pc_out_value };

                            Func<double[], double> funcion = delegate (double[] variables)
                            {
                                puntero_aplicacion.luis.RecompCycle_PCRC_without_Reheating(puntero_aplicacion.luis, 
                                ref cicloPCRC_withoutRH, puntero_aplicacion.w_dot_net2, i, puntero_aplicacion.t_t_in2, 
                                variables[2], puntero_aplicacion.p_mc_out2, variables[1], i, variables[2],
                                puntero_aplicacion.ua_lt2, puntero_aplicacion.ua_ht2, puntero_aplicacion.eta_mc2, 
                                puntero_aplicacion.eta_rc2, puntero_aplicacion.eta_pc2, puntero_aplicacion.eta_t2, 
                                puntero_aplicacion.n_sub_hxrs2, variables[0], puntero_aplicacion.tol2, 
                                puntero_aplicacion.eta_thermal2, -puntero_aplicacion.dp2_lt1,
                                -puntero_aplicacion.dp2_lt2, -puntero_aplicacion.dp2_ht1, -puntero_aplicacion.dp2_ht2, 
                                -puntero_aplicacion.dp2_pc1, -puntero_aplicacion.dp2_pc2,
                                -puntero_aplicacion.dp2_phx1, -puntero_aplicacion.dp2_phx2, -puntero_aplicacion.dp2_cooler1, 
                                -puntero_aplicacion.dp2_cooler2);

                                counter++;

                                puntero_aplicacion.massflow2 = cicloPCRC_withoutRH.m_dot_turbine;
                                puntero_aplicacion.w_dot_net2 = cicloPCRC_withoutRH.W_dot_net;
                                puntero_aplicacion.eta_thermal2 = cicloPCRC_withoutRH.eta_thermal;
                                puntero_aplicacion.recomp_frac2 = variables[0];
                                puntero_aplicacion.p_pc_in2 = variables[1];
                                puntero_aplicacion.p_pc_out2 = variables[2];

                                puntero_aplicacion.temp21 = cicloPCRC_withoutRH.temp[0];
                                puntero_aplicacion.temp22 = cicloPCRC_withoutRH.temp[1];
                                puntero_aplicacion.temp23 = cicloPCRC_withoutRH.temp[2];
                                puntero_aplicacion.temp24 = cicloPCRC_withoutRH.temp[3];
                                puntero_aplicacion.temp25 = cicloPCRC_withoutRH.temp[4];
                                puntero_aplicacion.temp26 = cicloPCRC_withoutRH.temp[5];
                                puntero_aplicacion.temp27 = cicloPCRC_withoutRH.temp[6];
                                puntero_aplicacion.temp28 = cicloPCRC_withoutRH.temp[7];
                                puntero_aplicacion.temp29 = cicloPCRC_withoutRH.temp[8];
                                puntero_aplicacion.temp210 = cicloPCRC_withoutRH.temp[9];
                                puntero_aplicacion.temp213 = cicloPCRC_withoutRH.temp[10];
                                puntero_aplicacion.temp214 = cicloPCRC_withoutRH.temp[11];

                                puntero_aplicacion.pres21 = cicloPCRC_withoutRH.pres[0];
                                puntero_aplicacion.pres22 = cicloPCRC_withoutRH.pres[1];
                                puntero_aplicacion.pres23 = cicloPCRC_withoutRH.pres[2];
                                puntero_aplicacion.pres24 = cicloPCRC_withoutRH.pres[3];
                                puntero_aplicacion.pres25 = cicloPCRC_withoutRH.pres[4];
                                puntero_aplicacion.pres26 = cicloPCRC_withoutRH.pres[5];
                                puntero_aplicacion.pres27 = cicloPCRC_withoutRH.pres[6];
                                puntero_aplicacion.pres28 = cicloPCRC_withoutRH.pres[7];
                                puntero_aplicacion.pres29 = cicloPCRC_withoutRH.pres[8];
                                puntero_aplicacion.pres210 = cicloPCRC_withoutRH.pres[9];
                                puntero_aplicacion.pres213 = cicloPCRC_withoutRH.pres[10];
                                puntero_aplicacion.pres214 = cicloPCRC_withoutRH.pres[11];

                                puntero_aplicacion.PHX1 = cicloPCRC_withoutRH.PHX.Q_dot;

                                puntero_aplicacion.LT_Q = cicloPCRC_withoutRH.LT.Q_dot;
                                puntero_aplicacion.LT_mdotc = cicloPCRC_withoutRH.LT.m_dot_design[0];
                                puntero_aplicacion.LT_mdoth = cicloPCRC_withoutRH.LT.m_dot_design[1];
                                puntero_aplicacion.LT_Tcin = cicloPCRC_withoutRH.LT.T_c_in;
                                puntero_aplicacion.LT_Thin = cicloPCRC_withoutRH.LT.T_h_in;
                                puntero_aplicacion.LT_Pcin = cicloPCRC_withoutRH.LT.P_c_in;
                                puntero_aplicacion.LT_Phin = cicloPCRC_withoutRH.LT.P_h_in;
                                puntero_aplicacion.LT_Pcout = cicloPCRC_withoutRH.LT.P_c_out;
                                puntero_aplicacion.LT_Phout = cicloPCRC_withoutRH.LT.P_h_out;
                                puntero_aplicacion.LT_Effc = cicloPCRC_withoutRH.LT.eff;

                                puntero_aplicacion.HT_Q = cicloPCRC_withoutRH.HT.Q_dot;
                                puntero_aplicacion.HT_mdotc = cicloPCRC_withoutRH.HT.m_dot_design[0];
                                puntero_aplicacion.HT_mdoth = cicloPCRC_withoutRH.HT.m_dot_design[1];
                                puntero_aplicacion.HT_Tcin = cicloPCRC_withoutRH.HT.T_c_in;
                                puntero_aplicacion.HT_Thin = cicloPCRC_withoutRH.HT.T_h_in;
                                puntero_aplicacion.HT_Pcin = cicloPCRC_withoutRH.HT.P_c_in;
                                puntero_aplicacion.HT_Phin = cicloPCRC_withoutRH.HT.P_h_in;
                                puntero_aplicacion.HT_Pcout = cicloPCRC_withoutRH.HT.P_c_out;
                                puntero_aplicacion.HT_Phout = cicloPCRC_withoutRH.HT.P_h_out;
                                puntero_aplicacion.HT_Effc = cicloPCRC_withoutRH.HT.eff;

                                puntero_aplicacion.PC11 = -cicloPCRC_withoutRH.PC.Q_dot;
                                puntero_aplicacion.PC21 = -cicloPCRC_withoutRH.COOLER.Q_dot;

                                massflow2_list.Add(puntero_aplicacion.massflow2);
                                eta_thermal2_list.Add(puntero_aplicacion.eta_thermal2);
                                recomp_frac2_list.Add(puntero_aplicacion.recomp_frac2);
                                p_pc_in2_list.Add(puntero_aplicacion.p_pc_in2);
                                p_pc_out2_list.Add(puntero_aplicacion.p_pc_out2);

                                t1_list.Add(puntero_aplicacion.temp21);
                                t2_list.Add(puntero_aplicacion.temp22);
                                t3_list.Add(puntero_aplicacion.temp23);
                                t4_list.Add(puntero_aplicacion.temp24);
                                t5_list.Add(puntero_aplicacion.temp25);
                                t6_list.Add(puntero_aplicacion.temp26);
                                t7_list.Add(puntero_aplicacion.temp27);
                                t8_list.Add(puntero_aplicacion.temp28);
                                t9_list.Add(puntero_aplicacion.temp29);
                                t10_list.Add(puntero_aplicacion.temp210);
                                t13_list.Add(puntero_aplicacion.temp213);
                                t14_list.Add(puntero_aplicacion.temp214);

                                p1_list.Add(puntero_aplicacion.pres21);
                                p2_list.Add(puntero_aplicacion.pres22);
                                p3_list.Add(puntero_aplicacion.pres23);
                                p4_list.Add(puntero_aplicacion.pres24);
                                p5_list.Add(puntero_aplicacion.pres25);
                                p6_list.Add(puntero_aplicacion.pres26);
                                p7_list.Add(puntero_aplicacion.pres27);
                                p8_list.Add(puntero_aplicacion.pres28);
                                p9_list.Add(puntero_aplicacion.pres29);
                                p10_list.Add(puntero_aplicacion.pres210);
                                p13_list.Add(puntero_aplicacion.pres213);
                                p14_list.Add(puntero_aplicacion.pres214);

                                PHX_Q2_list.Add(cicloPCRC_withoutRH.PHX.Q_dot);

                                HT_Eff_list.Add(cicloPCRC_withoutRH.HT.eff);
                                LT_Eff_list.Add(cicloPCRC_withoutRH.LT.eff);

                                listBox1.Items.Add(counter.ToString());
                                listBox2.Items.Add(puntero_aplicacion.eta_thermal2.ToString());
                                listBox3.Items.Add(puntero_aplicacion.recomp_frac2.ToString());
                                listBox4.Items.Add(puntero_aplicacion.p_pc_in2.ToString());
                                listBox9.Items.Add(puntero_aplicacion.p_pc_out2.ToString());

                                return puntero_aplicacion.eta_thermal2;
                            };

                            solver.SetMaxObjective(funcion);

                            double? finalScore;

                            var result = solver.Optimize(initialValue, out finalScore);

                            Double max_eta_thermal = 0.0;

                            max_eta_thermal = eta_thermal2_list.Max();

                            var maxIndex = eta_thermal2_list.IndexOf(eta_thermal2_list.Max());

                            textBox91.Text = p_pc_in2_list[maxIndex].ToString();
                            textBox2.Text = p_pc_out2_list[maxIndex].ToString();
                            textBox90.Text = recomp_frac2_list[maxIndex].ToString();
                            textBox86.Text = eta_thermal2_list[maxIndex].ToString();

                            //Copy results as design-point inputs
                            if (checkBox3.Checked == true)
                            {
                                puntero_aplicacion.textBox15.Text = recomp_frac2_list[maxIndex].ToString();
                                puntero_aplicacion.textBox23.Text = p_pc_in2_list[maxIndex].ToString();
                                puntero_aplicacion.textBox8.Text = p_pc_out2_list[maxIndex].ToString();
                                puntero_aplicacion.textBox3.Text = p_pc_out2_list[maxIndex].ToString();
                            }

                            //The variable 'i' is the loop counter for the CIT
                            listBox18.Items.Add(i.ToString());
                            listBox17.Items.Add(eta_thermal2_list[maxIndex].ToString());
                            listBox16.Items.Add(recomp_frac2_list[maxIndex].ToString());
                            listBox15.Items.Add(p_pc_in2_list[maxIndex].ToString());
                            listBox10.Items.Add(p_pc_out2_list[maxIndex].ToString());
                            listBox11.Items.Add(t8_list[maxIndex].ToString());
                            listBox12.Items.Add(t9_list[maxIndex].ToString());                           

                            //Copy results to EXCEL
                            double LTR_min_DT_1 = t8_list[maxIndex] - t3_list[maxIndex];
                            double LTR_min_DT_2 = t9_list[maxIndex] - t2_list[maxIndex];
                            double LTR_min_DT_paper = Math.Min(LTR_min_DT_1, LTR_min_DT_2);

                            double HTR_min_DT_1 = t8_list[maxIndex] - t4_list[maxIndex];
                            double HTR_min_DT_2 = t7_list[maxIndex] - t5_list[maxIndex];
                            double HTR_min_DT_paper = Math.Min(HTR_min_DT_1, HTR_min_DT_2);

                            //PC_in(kPa)
                            xlWorkSheet1.Cells[counter_Excel + 1, 1] = p_pc_in2_list[maxIndex].ToString();
                            //PC_out(kPa)
                            xlWorkSheet1.Cells[counter_Excel + 1, 2] = p_pc_out2_list[maxIndex].ToString();
                            //CIT
                            xlWorkSheet1.Cells[counter_Excel + 1, 3] = Convert.ToString(i - 273.15);
                            //LT UA(kW/K)
                            xlWorkSheet1.Cells[counter_Excel + 1, 4] = Convert.ToString(puntero_aplicacion.ua_lt2);
                            //HT UA(kW/K)
                            xlWorkSheet1.Cells[counter_Excel + 1, 5] = Convert.ToString(puntero_aplicacion.ua_ht2);
                            //Rec.Frac.
                            xlWorkSheet1.Cells[counter_Excel + 1, 6] = recomp_frac2_list[maxIndex].ToString();
                            //Eff.(%)
                            xlWorkSheet1.Cells[counter_Excel + 1, 7] = (eta_thermal2_list[maxIndex] * 100).ToString();
                            //LTR Eff.(%)
                            xlWorkSheet1.Cells[counter_Excel + 1, 8] = cicloPCRC_withoutRH.LT.eff.ToString();
                            //LTR Pinch(ºC)
                            xlWorkSheet1.Cells[counter_Excel + 1, 9] = LTR_min_DT_paper.ToString();
                            //HTR Eff.(%)
                            xlWorkSheet1.Cells[counter_Excel + 1, 10] = cicloPCRC_withoutRH.HT.eff.ToString();
                            //HTR Pinch(ºC)
                            xlWorkSheet1.Cells[counter_Excel + 1, 11] = HTR_min_DT_paper.ToString();

                            counter_Excel++;

                            initial_pc_in_value = puntero_aplicacion.p_pc_in2;
                            initial_pc_out_value = puntero_aplicacion.p_pc_out2;
                        }
                    }

                    //-------------------------------------------------------------------------

                    //UA Optimization true
                    else if (checkBox2.Checked == true)
                    {
                        //PureFluid
                        if (puntero_aplicacion.comboBox1.Text == "PureFluid")
                        {
                            puntero_aplicacion.category = RefrigerantCategory.PureFluid;
                            puntero_aplicacion.luis.core1(puntero_aplicacion.comboBox1.Text, puntero_aplicacion.category);
                        }

                        //NewMixture
                        if (puntero_aplicacion.comboBox1.Text == "NewMixture")
                        {
                            puntero_aplicacion.category = RefrigerantCategory.NewMixture;
                            puntero_aplicacion.luis.core1(puntero_aplicacion.comboBox2.Text + "=" + puntero_aplicacion.textBox68.Text + "," + puntero_aplicacion.comboBox6.Text + "=" + puntero_aplicacion.textBox69.Text + "," + puntero_aplicacion.comboBox12.Text + "=" + puntero_aplicacion.textBox33.Text + "," + puntero_aplicacion.comboBox7.Text + "=" + puntero_aplicacion.textBox34.Text, puntero_aplicacion.category);
                        }

                        if (puntero_aplicacion.comboBox1.Text == "PredefinedMixture")
                        {
                            puntero_aplicacion.category = RefrigerantCategory.PredefinedMixture;
                        }

                        if (puntero_aplicacion.comboBox1.Text == "PseudoPureFluid")
                        {
                            puntero_aplicacion.category = RefrigerantCategory.PseudoPureFluid;
                        }

                        if (puntero_aplicacion.comboBox3.Text == "DEF")
                        {
                            puntero_aplicacion.referencestate = ReferenceState.DEF;
                        }
                        if (puntero_aplicacion.comboBox3.Text == "ASH")
                        {
                            puntero_aplicacion.referencestate = ReferenceState.ASH;
                        }
                        if (puntero_aplicacion.comboBox3.Text == "IIR")
                        {
                            puntero_aplicacion.referencestate = ReferenceState.IIR;
                        }
                        if (puntero_aplicacion.comboBox3.Text == "NBP")
                        {
                            puntero_aplicacion.referencestate = ReferenceState.NBP;
                        }

                        puntero_aplicacion.luis.working_fluid.Category = puntero_aplicacion.category;
                        puntero_aplicacion.luis.working_fluid.reference = puntero_aplicacion.referencestate;

                        puntero_aplicacion.w_dot_net2 = Convert.ToDouble(puntero_aplicacion.textBox1.Text);
                        puntero_aplicacion.t_mc_in2 = Convert.ToDouble(puntero_aplicacion.textBox2.Text);
                        puntero_aplicacion.t_t_in2 = Convert.ToDouble(puntero_aplicacion.textBox4.Text);
                        //puntero_aplicacion.ua_lt2 = Convert.ToDouble(puntero_aplicacion.textBox17.Text);
                        //puntero_aplicacion.ua_ht2 = Convert.ToDouble(puntero_aplicacion.textBox16.Text);
                        puntero_aplicacion.eta_mc2 = Convert.ToDouble(puntero_aplicacion.textBox14.Text);
                        puntero_aplicacion.eta_rc2 = Convert.ToDouble(puntero_aplicacion.textBox13.Text);
                        puntero_aplicacion.eta_pc2 = Convert.ToDouble(puntero_aplicacion.textBox24.Text);
                        puntero_aplicacion.eta_t2 = Convert.ToDouble(puntero_aplicacion.textBox19.Text);
                        puntero_aplicacion.n_sub_hxrs2 = Convert.ToInt64(puntero_aplicacion.textBox20.Text);
                        puntero_aplicacion.p_mc_in2 = Convert.ToDouble(puntero_aplicacion.textBox3.Text);
                        puntero_aplicacion.p_mc_out2 = Convert.ToDouble(puntero_aplicacion.textBox28.Text);
                        puntero_aplicacion.p_pc_in2 = Convert.ToDouble(puntero_aplicacion.textBox23.Text);
                        puntero_aplicacion.t_pc_in2 = Convert.ToDouble(puntero_aplicacion.textBox22.Text);
                        puntero_aplicacion.p_pc_out2 = Convert.ToDouble(puntero_aplicacion.textBox8.Text);
                        puntero_aplicacion.recomp_frac2 = Convert.ToDouble(puntero_aplicacion.textBox15.Text);
                        puntero_aplicacion.tol2 = Convert.ToDouble(puntero_aplicacion.textBox21.Text);

                        puntero_aplicacion.dp2_lt1 = Convert.ToDouble(puntero_aplicacion.textBox5.Text);
                        puntero_aplicacion.dp2_lt2 = Convert.ToDouble(puntero_aplicacion.textBox26.Text);
                        puntero_aplicacion.dp2_ht1 = Convert.ToDouble(puntero_aplicacion.textBox12.Text);
                        puntero_aplicacion.dp2_ht2 = Convert.ToDouble(puntero_aplicacion.textBox25.Text);
                        puntero_aplicacion.dp2_pc1 = Convert.ToDouble(puntero_aplicacion.textBox11.Text);
                        puntero_aplicacion.dp2_phx1 = Convert.ToDouble(puntero_aplicacion.textBox10.Text);
                        puntero_aplicacion.dp2_cooler2 = Convert.ToDouble(puntero_aplicacion.textBox27.Text);

                        puntero_aplicacion.luis.wmm = puntero_aplicacion.luis.working_fluid.MolecularWeight;

                        core.PCRCwithoutReheating cicloPCRC_withoutRH = new core.PCRCwithoutReheating();

                        double UA_Total = puntero_aplicacion.ua_lt2 + puntero_aplicacion.ua_ht2;

                        double LT_fraction = 0.1;

                        //int counter = 0;

                        List<Double> massflow2_list = new List<Double>();
                        List<Double> recomp_frac2_list = new List<Double>();
                        List<Double> p_pc_in2_list = new List<Double>();
                        List<Double> p_pc_out2_list = new List<Double>();
                        List<Double> eta_thermal2_list = new List<Double>();
                        List<Double> PHX_Q2_list = new List<Double>();
                        List<Double> ua_lt2_list = new List<Double>();
                        List<Double> ua_ht2_list = new List<Double>();

                        List<Double> t1_list = new List<Double>();
                        List<Double> t2_list = new List<Double>();
                        List<Double> t3_list = new List<Double>();
                        List<Double> t4_list = new List<Double>();
                        List<Double> t5_list = new List<Double>();
                        List<Double> t6_list = new List<Double>();
                        List<Double> t7_list = new List<Double>();
                        List<Double> t8_list = new List<Double>();
                        List<Double> t9_list = new List<Double>();
                        List<Double> t10_list = new List<Double>();
                        List<Double> t13_list = new List<Double>();
                        List<Double> t14_list = new List<Double>();

                        List<Double> p1_list = new List<Double>();
                        List<Double> p2_list = new List<Double>();
                        List<Double> p3_list = new List<Double>();
                        List<Double> p4_list = new List<Double>();
                        List<Double> p5_list = new List<Double>();
                        List<Double> p6_list = new List<Double>();
                        List<Double> p7_list = new List<Double>();
                        List<Double> p8_list = new List<Double>();
                        List<Double> p9_list = new List<Double>();
                        List<Double> p10_list = new List<Double>();
                        List<Double> p13_list = new List<Double>();
                        List<Double> p14_list = new List<Double>();

                        List<Double> HT_Eff_list = new List<Double>();
                        List<Double> LT_Eff_list = new List<Double>();

                        NLoptAlgorithm algorithm_type = NLoptAlgorithm.LN_BOBYQA;

                        if (comboBox19.Text == "BOBYQA")
                            algorithm_type = NLoptAlgorithm.LN_BOBYQA;
                        else if (comboBox19.Text == "COBYLA")
                            algorithm_type = NLoptAlgorithm.LN_COBYLA;
                        else if (comboBox19.Text == "SUBPLEX")
                            algorithm_type = NLoptAlgorithm.LN_SBPLX;
                        else if (comboBox19.Text == "NELDER-MEAD")
                            algorithm_type = NLoptAlgorithm.LN_NELDERMEAD;
                        else if (comboBox19.Text == "NEWUOA")
                            algorithm_type = NLoptAlgorithm.LN_NEWUOA;
                        else if (comboBox19.Text == "PRAXIS")
                            algorithm_type = NLoptAlgorithm.LN_PRAXIS;

                        if (i == Convert.ToDouble(textBox57.Text))
                        {
                            if (checkBox6.Checked == true)
                            {
                                initial_pc_in_value = Convert.ToDouble(textBox1.Text);
                                initial_pc_out_value = Convert.ToDouble(textBox1.Text) + 500;
                            }
                            else
                            {
                                initial_pc_in_value = puntero_aplicacion.MixtureCriticalPressure;
                                initial_pc_out_value = puntero_aplicacion.MixtureCriticalPressure + 500;
                            }

                            xlWorkSheet1.Name = puntero_aplicacion.comboBox2.Text + " Mixture";

                            xlWorkSheet1.Cells[1, 1] = puntero_aplicacion.comboBox2.Text + ":" + puntero_aplicacion.textBox68.Text + "," + puntero_aplicacion.comboBox6.Text + ":" + puntero_aplicacion.textBox69.Text + "," + puntero_aplicacion.comboBox12.Text + ":" + puntero_aplicacion.textBox33.Text + "," + puntero_aplicacion.comboBox7.Text + ":" + puntero_aplicacion.textBox34.Text;
                            xlWorkSheet1.Cells[1, 2] = "Pcrit(kPa)";
                            xlWorkSheet1.Cells[1, 3] = "Tcrit(ºC)";

                            xlWorkSheet1.Cells[2, 1] = "";
                            xlWorkSheet1.Cells[2, 2] = Convert.ToString(puntero_aplicacion.MixtureCriticalPressure);
                            xlWorkSheet1.Cells[2, 3] = Convert.ToString(puntero_aplicacion.MixtureCriticalTemperature - 273.15);

                            xlWorkSheet1.Cells[3, 1] = "";
                            xlWorkSheet1.Cells[3, 2] = "";
                            xlWorkSheet1.Cells[4, 3] = "";

                            xlWorkSheet1.Cells[4, 1] = "PC_in(kPa)";
                            xlWorkSheet1.Cells[4, 2] = "PC_out(kPa)";
                            xlWorkSheet1.Cells[4, 3] = "CIT(K)";
                            xlWorkSheet1.Cells[4, 4] = "LT UA(kW/K)";
                            xlWorkSheet1.Cells[4, 5] = "HT UA(kW/K)";
                            xlWorkSheet1.Cells[4, 6] = "Rec.Frac.";
                            xlWorkSheet1.Cells[4, 7] = "Eff.(%)";
                            xlWorkSheet1.Cells[4, 8] = "LTR Eff.(%)";
                            xlWorkSheet1.Cells[4, 9] = "LTR Pinch(ºC)";
                            xlWorkSheet1.Cells[4, 10] = "HTR Eff.(%)";
                            xlWorkSheet1.Cells[4, 11] = "HTR Pinch(ºC)";
                            xlWorkSheet1.Cells[4, 12] = "PTC_Apperture_Area(m2)";
                            xlWorkSheet1.Cells[4, 13] = "PTC_Pressure_Drop(bar)";
                            xlWorkSheet1.Cells[4, 14] = "LF_Apperture_Area(m2)";
                            xlWorkSheet1.Cells[4, 15] = "LF_Pressure_Drop(bar)";
                        }

                        using (var solver = new NLoptSolver(algorithm_type, 4, 0.01, 10000))
                        {
                            solver.SetLowerBounds(new[] { 0.1, initial_pc_in_value, initial_pc_out_value, 0.2 });
                            solver.SetUpperBounds(new[] { 1.0, 125000, (puntero_aplicacion.p_mc_out2 / 1.5), 0.8 });

                            solver.SetInitialStepSize(new[] { 0.05, 100, 100, 0.05 });

                            var initialValue = new[] { 0.2, initial_pc_in_value, initial_pc_out_value, 0.5 };

                            Func<double[], double> funcion = delegate (double[] variables)
                            {
                                puntero_aplicacion.luis.RecompCycle_PCRC_without_Reheating_for_Optimization(puntero_aplicacion.luis, ref cicloPCRC_withoutRH, puntero_aplicacion.w_dot_net2, i, puntero_aplicacion.t_t_in2, variables[2], puntero_aplicacion.p_mc_out2, variables[1], i, variables[2],
                                variables[3], UA_Total, puntero_aplicacion.eta_mc2, puntero_aplicacion.eta_rc2, puntero_aplicacion.eta_pc2, puntero_aplicacion.eta_t2, puntero_aplicacion.n_sub_hxrs2, variables[0], puntero_aplicacion.tol2, puntero_aplicacion.eta_thermal2, -puntero_aplicacion.dp2_lt1, -puntero_aplicacion.dp2_lt2, -puntero_aplicacion.dp2_ht1, -puntero_aplicacion.dp2_ht2, -puntero_aplicacion.dp2_pc1, -puntero_aplicacion.dp2_pc2,
                                -puntero_aplicacion.dp2_phx1, -puntero_aplicacion.dp2_phx2, -puntero_aplicacion.dp2_cooler1, -puntero_aplicacion.dp2_cooler2);

                                counter++;

                                puntero_aplicacion.massflow2 = cicloPCRC_withoutRH.m_dot_turbine;
                                puntero_aplicacion.w_dot_net2 = cicloPCRC_withoutRH.W_dot_net;
                                puntero_aplicacion.eta_thermal2 = cicloPCRC_withoutRH.eta_thermal;
                                puntero_aplicacion.recomp_frac2 = variables[0];
                                puntero_aplicacion.p_pc_in2 = variables[1];
                                puntero_aplicacion.p_pc_out2 = variables[2];
                                LT_fraction = variables[3];
                                puntero_aplicacion.ua_lt2 = UA_Total * LT_fraction;
                                puntero_aplicacion.ua_ht2 = UA_Total * (1 - LT_fraction);

                                puntero_aplicacion.temp21 = cicloPCRC_withoutRH.temp[0];
                                puntero_aplicacion.temp22 = cicloPCRC_withoutRH.temp[1];
                                puntero_aplicacion.temp23 = cicloPCRC_withoutRH.temp[2];
                                puntero_aplicacion.temp24 = cicloPCRC_withoutRH.temp[3];
                                puntero_aplicacion.temp25 = cicloPCRC_withoutRH.temp[4];
                                puntero_aplicacion.temp26 = cicloPCRC_withoutRH.temp[5];
                                puntero_aplicacion.temp27 = cicloPCRC_withoutRH.temp[6];
                                puntero_aplicacion.temp28 = cicloPCRC_withoutRH.temp[7];
                                puntero_aplicacion.temp29 = cicloPCRC_withoutRH.temp[8];
                                puntero_aplicacion.temp210 = cicloPCRC_withoutRH.temp[9];
                                puntero_aplicacion.temp213 = cicloPCRC_withoutRH.temp[10];
                                puntero_aplicacion.temp214 = cicloPCRC_withoutRH.temp[11];

                                puntero_aplicacion.pres21 = cicloPCRC_withoutRH.pres[0];
                                puntero_aplicacion.pres22 = cicloPCRC_withoutRH.pres[1];
                                puntero_aplicacion.pres23 = cicloPCRC_withoutRH.pres[2];
                                puntero_aplicacion.pres24 = cicloPCRC_withoutRH.pres[3];
                                puntero_aplicacion.pres25 = cicloPCRC_withoutRH.pres[4];
                                puntero_aplicacion.pres26 = cicloPCRC_withoutRH.pres[5];
                                puntero_aplicacion.pres27 = cicloPCRC_withoutRH.pres[6];
                                puntero_aplicacion.pres28 = cicloPCRC_withoutRH.pres[7];
                                puntero_aplicacion.pres29 = cicloPCRC_withoutRH.pres[8];
                                puntero_aplicacion.pres210 = cicloPCRC_withoutRH.pres[9];
                                puntero_aplicacion.pres213 = cicloPCRC_withoutRH.pres[10];
                                puntero_aplicacion.pres214 = cicloPCRC_withoutRH.pres[11];

                                puntero_aplicacion.PHX1 = cicloPCRC_withoutRH.PHX.Q_dot;

                                puntero_aplicacion.LT_Q = cicloPCRC_withoutRH.LT.Q_dot;
                                puntero_aplicacion.LT_mdotc = cicloPCRC_withoutRH.LT.m_dot_design[0];
                                puntero_aplicacion.LT_mdoth = cicloPCRC_withoutRH.LT.m_dot_design[1];
                                puntero_aplicacion.LT_Tcin = cicloPCRC_withoutRH.LT.T_c_in;
                                puntero_aplicacion.LT_Thin = cicloPCRC_withoutRH.LT.T_h_in;
                                puntero_aplicacion.LT_Pcin = cicloPCRC_withoutRH.LT.P_c_in;
                                puntero_aplicacion.LT_Phin = cicloPCRC_withoutRH.LT.P_h_in;
                                puntero_aplicacion.LT_Pcout = cicloPCRC_withoutRH.LT.P_c_out;
                                puntero_aplicacion.LT_Phout = cicloPCRC_withoutRH.LT.P_h_out;
                                puntero_aplicacion.LT_Effc = cicloPCRC_withoutRH.LT.eff;

                                puntero_aplicacion.HT_Q = cicloPCRC_withoutRH.HT.Q_dot;
                                puntero_aplicacion.HT_mdotc = cicloPCRC_withoutRH.HT.m_dot_design[0];
                                puntero_aplicacion.HT_mdoth = cicloPCRC_withoutRH.HT.m_dot_design[1];
                                puntero_aplicacion.HT_Tcin = cicloPCRC_withoutRH.HT.T_c_in;
                                puntero_aplicacion.HT_Thin = cicloPCRC_withoutRH.HT.T_h_in;
                                puntero_aplicacion.HT_Pcin = cicloPCRC_withoutRH.HT.P_c_in;
                                puntero_aplicacion.HT_Phin = cicloPCRC_withoutRH.HT.P_h_in;
                                puntero_aplicacion.HT_Pcout = cicloPCRC_withoutRH.HT.P_c_out;
                                puntero_aplicacion.HT_Phout = cicloPCRC_withoutRH.HT.P_h_out;
                                puntero_aplicacion.HT_Effc = cicloPCRC_withoutRH.HT.eff;

                                puntero_aplicacion.PC11 = -cicloPCRC_withoutRH.PC.Q_dot;
                                puntero_aplicacion.PC21 = -cicloPCRC_withoutRH.COOLER.Q_dot;

                                massflow2_list.Add(puntero_aplicacion.massflow2);
                                eta_thermal2_list.Add(puntero_aplicacion.eta_thermal2);
                                recomp_frac2_list.Add(puntero_aplicacion.recomp_frac2);
                                p_pc_in2_list.Add(puntero_aplicacion.p_pc_in2);
                                p_pc_out2_list.Add(puntero_aplicacion.p_pc_out2);
                                ua_lt2_list.Add(puntero_aplicacion.ua_lt2);
                                ua_ht2_list.Add(puntero_aplicacion.ua_ht2);

                                t1_list.Add(puntero_aplicacion.temp21);
                                t2_list.Add(puntero_aplicacion.temp22);
                                t3_list.Add(puntero_aplicacion.temp23);
                                t4_list.Add(puntero_aplicacion.temp24);
                                t5_list.Add(puntero_aplicacion.temp25);
                                t6_list.Add(puntero_aplicacion.temp26);
                                t7_list.Add(puntero_aplicacion.temp27);
                                t8_list.Add(puntero_aplicacion.temp28);
                                t9_list.Add(puntero_aplicacion.temp29);
                                t10_list.Add(puntero_aplicacion.temp210);
                                t13_list.Add(puntero_aplicacion.temp213);
                                t14_list.Add(puntero_aplicacion.temp214);

                                p1_list.Add(puntero_aplicacion.pres21);
                                p2_list.Add(puntero_aplicacion.pres22);
                                p3_list.Add(puntero_aplicacion.pres23);
                                p4_list.Add(puntero_aplicacion.pres24);
                                p5_list.Add(puntero_aplicacion.pres25);
                                p6_list.Add(puntero_aplicacion.pres26);
                                p7_list.Add(puntero_aplicacion.pres27);
                                p8_list.Add(puntero_aplicacion.pres28);
                                p9_list.Add(puntero_aplicacion.pres29);
                                p10_list.Add(puntero_aplicacion.pres210);
                                p13_list.Add(puntero_aplicacion.pres213);
                                p14_list.Add(puntero_aplicacion.pres214);

                                PHX_Q2_list.Add(cicloPCRC_withoutRH.PHX.Q_dot);

                                HT_Eff_list.Add(cicloPCRC_withoutRH.HT.eff);
                                LT_Eff_list.Add(cicloPCRC_withoutRH.LT.eff);

                                listBox1.Items.Add(counter.ToString());
                                listBox2.Items.Add(puntero_aplicacion.eta_thermal2.ToString());
                                listBox3.Items.Add(puntero_aplicacion.recomp_frac2.ToString());
                                listBox4.Items.Add(puntero_aplicacion.p_pc_in2.ToString());
                                listBox9.Items.Add(puntero_aplicacion.p_pc_out2.ToString());
                                listBox5.Items.Add(puntero_aplicacion.ua_lt2.ToString());
                                listBox6.Items.Add(puntero_aplicacion.ua_ht2.ToString());

                                return puntero_aplicacion.eta_thermal2;
                            };

                            solver.SetMaxObjective(funcion);

                            double? finalScore;

                            var result = solver.Optimize(initialValue, out finalScore);

                            Double max_eta_thermal = 0.0;

                            max_eta_thermal = eta_thermal2_list.Max();

                            var maxIndex = eta_thermal2_list.IndexOf(eta_thermal2_list.Max());

                            puntero_aplicacion.ua_lt2 = UA_Total * LT_fraction;
                            puntero_aplicacion.ua_ht2 = UA_Total * (1 - LT_fraction);

                            textBox91.Text = p_pc_in2_list[maxIndex].ToString();
                            textBox2.Text = p_pc_out2_list[maxIndex].ToString();
                            textBox90.Text = recomp_frac2_list[maxIndex].ToString();
                            textBox86.Text = eta_thermal2_list[maxIndex].ToString();
                            textBox82.Text = ua_lt2_list[maxIndex].ToString();
                            textBox83.Text = ua_ht2_list[maxIndex].ToString();

                            //Copy results as design-point inputs
                            if (checkBox3.Checked == true)
                            {
                                puntero_aplicacion.textBox15.Text = recomp_frac2_list[maxIndex].ToString();
                                puntero_aplicacion.textBox23.Text = p_pc_in2_list[maxIndex].ToString();
                                puntero_aplicacion.textBox8.Text = p_pc_out2_list[maxIndex].ToString();
                                puntero_aplicacion.textBox3.Text = p_pc_out2_list[maxIndex].ToString();
                                puntero_aplicacion.textBox17.Text = ua_lt2_list[maxIndex].ToString();
                                puntero_aplicacion.textBox16.Text = ua_ht2_list[maxIndex].ToString();
                            }

                            //The variable 'i' is the loop counter for the CIT
                            listBox18.Items.Add(i.ToString());
                            listBox17.Items.Add(eta_thermal2_list[maxIndex].ToString());
                            listBox16.Items.Add(recomp_frac2_list[maxIndex].ToString());
                            listBox15.Items.Add(p_pc_in2_list[maxIndex].ToString());
                            listBox10.Items.Add(p_pc_out2_list[maxIndex].ToString());
                            listBox14.Items.Add(ua_lt2_list[maxIndex].ToString());
                            listBox13.Items.Add(ua_ht2_list[maxIndex].ToString());
                            listBox11.Items.Add(t8_list[maxIndex].ToString());
                            listBox12.Items.Add(t9_list[maxIndex].ToString());
                                                     
                            //Copy results to EXCEL
                            double LTR_min_DT_1 = t8_list[maxIndex] - t3_list[maxIndex];
                            double LTR_min_DT_2 = t9_list[maxIndex] - t2_list[maxIndex];
                            double LTR_min_DT_paper = Math.Min(LTR_min_DT_1, LTR_min_DT_2);

                            double HTR_min_DT_1 = t8_list[maxIndex] - t4_list[maxIndex];
                            double HTR_min_DT_2 = t7_list[maxIndex] - t5_list[maxIndex];
                            double HTR_min_DT_paper = Math.Min(HTR_min_DT_1, HTR_min_DT_2);

                            //PC_in(kPa)
                            xlWorkSheet1.Cells[counter_Excel + 1, 1] = p_pc_in2_list[maxIndex].ToString();
                            //PC_out(kPa)
                            xlWorkSheet1.Cells[counter_Excel + 1, 2] = p_pc_out2_list[maxIndex].ToString();
                            //CIT
                            xlWorkSheet1.Cells[counter_Excel + 1, 3] = Convert.ToString(i - 273.15);
                            //LT UA(kW/K)
                            xlWorkSheet1.Cells[counter_Excel + 1, 4] = ua_lt2_list[maxIndex].ToString();
                            //HT UA(kW/K)
                            xlWorkSheet1.Cells[counter_Excel + 1, 5] = ua_ht2_list[maxIndex].ToString();
                            //Rec.Frac.
                            xlWorkSheet1.Cells[counter_Excel + 1, 6] = recomp_frac2_list[maxIndex].ToString();
                            //Eff.(%)
                            xlWorkSheet1.Cells[counter_Excel + 1, 7] = (eta_thermal2_list[maxIndex] * 100).ToString();
                            //LTR Eff.(%)
                            xlWorkSheet1.Cells[counter_Excel + 1, 8] = cicloPCRC_withoutRH.LT.eff.ToString();
                            //LTR Pinch(ºC)
                            xlWorkSheet1.Cells[counter_Excel + 1, 9] = LTR_min_DT_paper.ToString();
                            //HTR Eff.(%)
                            xlWorkSheet1.Cells[counter_Excel + 1, 10] = cicloPCRC_withoutRH.HT.eff.ToString();
                            //HTR Pinch(ºC)
                            xlWorkSheet1.Cells[counter_Excel + 1, 11] = HTR_min_DT_paper.ToString();
                
                            counter_Excel++;

                            initial_pc_in_value = puntero_aplicacion.p_pc_in2;
                            initial_pc_out_value = puntero_aplicacion.p_pc_out2;
                        }
                    } //checkBox2.Checked (optimize UA)

                } //loop for CIT optimization analysis

            }//loop for UA optimization analysis

            //Closing Excel Book
            xlWorkBook1.SaveAs(textBox3.Text + "CIT_Optimization_PCRC_without_ReHeating_" + xlWorkSheet1.Name + ".xls", Excel.XlFileFormat.xlWorkbookNormal, misValue1, misValue1, misValue1, misValue1, Excel.XlSaveAsAccessMode.xlExclusive, misValue1, misValue1, misValue1, misValue1, misValue1);

            xlWorkBook1.Close(true, misValue1, misValue1);
            xlApp1.Quit();

            releaseObject(xlWorkSheet1);
            //releaseObject(xlWorkSheet2);
            releaseObject(xlWorkBook1);
            releaseObject(xlApp1);
        }

        private void PCRC_without_ReHeating_Optimization_Analysis_Results_Load(object sender, EventArgs e)
        {
            textBox1.Text = puntero_aplicacion.textBox32.Text;
        }
    }
}
