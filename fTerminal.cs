using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using SajetClass;
using Microsoft.Win32;
using System.Globalization;
using System.Data.OracleClient;

namespace ComomTest
{
    public partial class fTerminal : Form
    {
        public fTerminal()
        {
            InitializeComponent();
        }

        public fTerminal(bool bConfig)
        {
            InitializeComponent();
            btnSave.Enabled = bConfig;
        }
        string g_sIniFactoryID = "0";
        string g_sFCID = "0";
        string g_sTerminalID;
        string sSQL;
        DataSet dsTemp;
        string g_sIniFile = Application.StartupPath + "\\Sajet.ini";
        string g_sIniSection =fMain.g_sFunction;


        private void fMain_Load(object sender, EventArgs e)
        {
            SajetCommon.SetLanguageControl(this);
            this.Text = this.Text + "( " + SajetCommon.g_sFileVersion + " )";
            tabControl1.SelectedIndex = 0;            

            //Read Ini File  
            SajetInifile sajetInifile1 = new SajetInifile();
            g_sIniFactoryID = sajetInifile1.ReadIniFile(g_sIniFile, "System", "Factory", "0");
            g_sTerminalID = sajetInifile1.ReadIniFile(g_sIniFile, g_sIniSection, "Terminal", "0");
            
            //Factory
            sSQL = "SELECT FACTORY_ID,FACTORY_CODE,FACTORY_NAME "
                 + "FROM SAJET.SYS_FACTORY "
                 + "WHERE ENABLED = 'Y' "
                 + "ORDER BY FACTORY_CODE ";
            dsTemp = ClientUtils.ExecuteSQL(sSQL);
                      
            string sFind = "";
            for (int i = 0; i <= dsTemp.Tables[0].Rows.Count - 1; i++)
            {
                combFactory.Items.Add(dsTemp.Tables[0].Rows[i]["FACTORY_CODE"].ToString());
                if (g_sIniFactoryID == dsTemp.Tables[0].Rows[i]["FACTORY_ID"].ToString())
                {
                    sFind = dsTemp.Tables[0].Rows[i]["FACTORY_CODE"].ToString();
                }
            }
            if (sFind != "")
            {
                combFactory.SelectedIndex = combFactory.FindString(sFind);
            }
            else
            {
                combFactory.SelectedIndex = 0;
            }                                 
        }
        
        private void combFactory_SelectedIndexChanged(object sender, EventArgs e)
        {
            LabLine.Text = "";
            LabStage.Text = "";
            LabProcess.Text = "";
            LabTerminal.Text = "";
            g_sFCID = "0";
            LabFactoryName.Text = "";

            sSQL = "SELECT FACTORY_ID,FACTORY_NAME "
                 + "FROM SAJET.SYS_FACTORY "
                 + "WHERE FACTORY_CODE =:FACTORY_CODE ";
            object[][] sqlparams = new object[][] { new object[]{ParameterDirection.Input,OracleType.VarChar,"FACTORY_CODE",combFactory.Text}};
            dsTemp = ClientUtils.ExecuteSQL(sSQL,sqlparams);
            if (dsTemp.Tables[0].Rows.Count > 0)
            {
                g_sFCID = dsTemp.Tables[0].Rows[0]["FACTORY_ID"].ToString();
                LabFactoryName.Text = dsTemp.Tables[0].Rows[0]["FACTORY_NAME"].ToString();
            }

            //Show_Terminal("INPUT");
           Show_Terminal(ClientUtils.fParameter);
        }
        
        public void Show_Terminal(string sProcessType)
        {
            TVTerminal.Nodes.Clear();

            sSQL = "SELECT b.pdline_name, c.stage_code, c.stage_name, d.process_code, d.process_name "
                + "       ,a.terminal_id, a.terminal_name "
                + " FROM sajet.sys_terminal a "
                + "     ,sajet.sys_pdline b "
                + "     ,sajet.sys_stage c "
                + "     ,sajet.sys_process d "
                + "     ,sajet.sys_operate_type e "
                + " WHERE b.factory_id = '" + g_sFCID + "' "
                + " AND a.pdline_id = b.pdline_id "
                + " AND a.stage_id = c.stage_id "
                + " AND a.process_id = d.process_id "
                + " AND d.operate_id = e.operate_id "
                + " AND Upper(e.type_name) = '" + sProcessType.ToUpper() + "' "
                + " AND a.enabled = 'Y' "
                + " AND b.enabled = 'Y' "
                + " AND c.enabled = 'Y' "
                + " AND d.enabled = 'Y' "
                + " ORDER BY b.pdline_name, c.stage_code, d.process_code, a.terminal_name ";
            dsTemp = ClientUtils.ExecuteSQL(sSQL);
            if (dsTemp.Tables[0].Rows.Count == 0)
                return;

            string sPreLine = "";
            string sPreStage = "";
            string sPreProcess = "";

            for (int i = 0; i <= dsTemp.Tables[0].Rows.Count - 1; i++)
            {
                string sLine = dsTemp.Tables[0].Rows[i]["PDLINE_NAME"].ToString();
                string sStage = dsTemp.Tables[0].Rows[i]["STAGE_NAME"].ToString();
                string sProcess = dsTemp.Tables[0].Rows[i]["PROCESS_NAME"].ToString();
                string sTerminal = dsTemp.Tables[0].Rows[i]["TERMINAL_NAME"].ToString();
                
                if (sPreLine != sLine)
                {                    
                    TVTerminal.Nodes.Add(sLine);
                    int iNodeCount = TVTerminal.Nodes.Count - 1;
                    TVTerminal.Nodes[iNodeCount].ImageIndex = 0;
                    
                    TVTerminal.Nodes[iNodeCount].Nodes.Add(sStage);
                    TVTerminal.Nodes[iNodeCount].LastNode.ImageIndex = 1;

                    TVTerminal.Nodes[iNodeCount].LastNode.Nodes.Add(sProcess);
                    TVTerminal.Nodes[iNodeCount].LastNode.LastNode.ImageIndex = 2;

                    TVTerminal.Nodes[iNodeCount].LastNode.LastNode.Nodes.Add(sTerminal);
                    TVTerminal.Nodes[iNodeCount].LastNode.LastNode.LastNode.ImageIndex = 3;
                }
                else if (sPreStage != sStage)
                {
                    int iNodeCount = TVTerminal.Nodes.Count - 1;
                    TVTerminal.Nodes[iNodeCount].Nodes.Add(sStage);
                    TVTerminal.Nodes[iNodeCount].LastNode.ImageIndex = 1;

                    TVTerminal.Nodes[iNodeCount].LastNode.Nodes.Add(sProcess);
                    TVTerminal.Nodes[iNodeCount].LastNode.LastNode.ImageIndex = 2;

                    TVTerminal.Nodes[iNodeCount].LastNode.LastNode.Nodes.Add(sTerminal);
                    TVTerminal.Nodes[iNodeCount].LastNode.LastNode.LastNode.ImageIndex = 3;
                }
                else if (sPreProcess != sProcess)
                {
                    int iNodeCount = TVTerminal.Nodes.Count - 1;
                    TVTerminal.Nodes[iNodeCount].LastNode.Nodes.Add(sProcess);
                    TVTerminal.Nodes[iNodeCount].LastNode.LastNode.ImageIndex = 2;

                    TVTerminal.Nodes[iNodeCount].LastNode.LastNode.Nodes.Add(sTerminal);
                    TVTerminal.Nodes[iNodeCount].LastNode.LastNode.LastNode.ImageIndex = 3;
                }
                else
                {                    
                    int iNodeCount = TVTerminal.Nodes.Count - 1;
                    TVTerminal.Nodes[iNodeCount].LastNode.LastNode.Nodes.Add(sTerminal);
                    TVTerminal.Nodes[iNodeCount].LastNode.LastNode.LastNode.ImageIndex = 3;
                }
                sPreLine = dsTemp.Tables[0].Rows[i]["PDLINE_NAME"].ToString();
                sPreStage = dsTemp.Tables[0].Rows[i]["STAGE_NAME"].ToString();
                sPreProcess = dsTemp.Tables[0].Rows[i]["PROCESS_NAME"].ToString();

                //SajetIni中設定的Terminal
                if (g_sTerminalID == dsTemp.Tables[0].Rows[i]["TERMINAL_ID"].ToString())
                {
                    TVTerminal.SelectedNode = TVTerminal.Nodes[TVTerminal.Nodes.Count - 1].LastNode.LastNode.LastNode;
                    TVTerminal.Focus();                    
                }
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (LabTerminal.Text == "")
            {
                MessageBox.Show("Please Choose Terminal","Error",MessageBoxButtons.OK,MessageBoxIcon.Error);
                return;
            }

            string sTerminalID = Get_TerminalID();
            if (sTerminalID == "0")
            {
                return;
            }
            SajetInifile sajetInifile1 = new SajetInifile();
            sajetInifile1.WriteIniFile(g_sIniFile, g_sIniSection, "Terminal", sTerminalID);
            sajetInifile1.WriteIniFile(g_sIniFile, "System", "Factory", g_sFCID);
            g_sTerminalID = sTerminalID;
            SajetCommon.Show_Message("Assign Terminal OK",-1);
            DialogResult = DialogResult.OK;
        }        
        public string Get_TerminalID()
        {
            sSQL = "SELECT A.TERMINAL_ID "
                 + "FROM SAJET.SYS_TERMINAL A "
                 + "    ,SAJET.SYS_PROCESS B "
                 + "    ,SAJET.SYS_PDLINE C "
                 + "WHERE A.TERMINAL_NAME =:TERMINAL_NAME "
                 + "AND B.PROCESS_NAME =:PROCESS_NAME "
                 + "AND C.PDLINE_NAME =: PDLINE_NAME "
                 + "AND A.PROCESS_ID = B.PROCESS_ID "
                 + "AND A.PDLINE_ID = C.PDLINE_ID ";
            object[][] sqlparams = new object[3][];
            sqlparams[0] = new object[] { ParameterDirection.Input, OracleType.VarChar, "TERMINAL_NAME", LabTerminal.Text };
            sqlparams[1] = new object[] { ParameterDirection.Input, OracleType.VarChar, "PROCESS_NAME", LabProcess.Text };
            sqlparams[2] = new object[] { ParameterDirection.Input, OracleType.VarChar, "PDLINE_NAME", LabLine.Text };
            dsTemp = ClientUtils.ExecuteSQL(sSQL,sqlparams);
            if (dsTemp.Tables[0].Rows.Count == 0)
            {
                MessageBox.Show("Terminal Data Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return "0";
            }
            return dsTemp.Tables[0].Rows[0]["TERMINAL_ID"].ToString();

        }

        private void TVTerminal_AfterSelect(object sender, TreeViewEventArgs e)
        {
            TVTerminal.SelectedNode.SelectedImageIndex = TVTerminal.SelectedNode.ImageIndex;

            LabLine.Text = "";
            LabStage.Text = "";
            LabProcess.Text = "";
            LabTerminal.Text = "";

            if (TVTerminal.SelectedNode.Level != 3)
                return;

            LabLine.Text = TVTerminal.SelectedNode.Parent.Parent.Parent.Text;
            LabStage.Text = TVTerminal.SelectedNode.Parent.Parent.Text;
            LabProcess.Text = TVTerminal.SelectedNode.Parent.Text;
            LabTerminal.Text = TVTerminal.SelectedNode.Text;
        }        
       
        public String GetProcessID(String sProcessName)
        {           
            try
            {
                sSQL = "SELECT PROCESS_ID "
                    + "  FROM SAJET.SYS_PROCESS "
                    + " WHERE PROCESS_NAME =:PROCESS_NAME "
                    + "   AND ROWNUM = 1 ";
                object[][] sqlparams = new object[][] { new object[] { ParameterDirection.Input, OracleType.VarChar, "PROCESS_NAME", sProcessName } };
                dsTemp = ClientUtils.ExecuteSQL(sSQL, sqlparams);
                if (dsTemp.Tables[0].Rows.Count > 0)
                    return dsTemp.Tables[0].Rows[0]["PROCESS_ID"].ToString();
                else
                    return "0";
            }
            catch (Exception ex)
            {                
                MessageBox.Show(ex.Message);
                return "0";
            }
        }
        public String GetProcessName(String sProcessID)
        {            
            DataSet dsTemp1 = new DataSet();
            try
            {
                sSQL = "SELECT PROCESS_NAME "
                    + "  FROM SAJET.SYS_PROCESS "
                    + " WHERE PROCESS_ID =:PROCESS_ID "
                    + "   AND ROWNUM = 1 ";
                object[][] sqlparams = new object[][] { new object[] { ParameterDirection.Input, OracleType.VarChar, "PROCESS_ID", sProcessID } };
                dsTemp1 = ClientUtils.ExecuteSQL(sSQL,sqlparams);
                if (dsTemp1.Tables[0].Rows.Count > 0)
                    return dsTemp1.Tables[0].Rows[0]["PROCESS_NAME"].ToString();
                else
                    return "0";
            }
            catch (Exception ex)
            {                
                MessageBox.Show(ex.Message);
                return "0";
            }
        }                                                      
    }
}

