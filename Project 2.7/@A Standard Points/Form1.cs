using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using SKF.RS.STB.Analyst;
using SKF.RS.STB.DB;
using SKF.RS.STB.Generic;
using System.Collections;
using System.IO;
using System.Reflection.Emit;
using SKF.Standard.Points.Entity;
using SKF.Standard.Points.Update;
using Point = System.Drawing.Point;
using TreeElem = SKF.Standard.Points.Entity.TreeElem;


namespace _A_Standard_Point___SCP
{
    
    public enum Acc
    {
        DetectionType = 20500, //Peak
        SaveData = 20202,      // FFT and Time
        Lines = 6400,
        StartFreq = 0,
        StopFreq = 10000,
        LowCutOff = 10,
        Averages = 1
    }
    
    public enum Env
    {
        DetectionType = 20501, //Peak to Peak
        SaveData = 20202,      // FFT and Time
        StartFreq = 0,
        Averages = 1
    }

    public enum Vel
    {
        DetectionType = 20502, // RMS
        SaveData = 20202,      // FFT and Time
        Lines = 1600,
        StartFreq = 0,
         Averages = 1
    }
    public enum Zoom
    {
        DetectionType = 20502, // RMS
        SaveData = 20202,      // FFT and Time
        Lines = 6400,
        StartFreq = 0,
        Averages = 1
    }       

  

    public partial class Form1 : Form
    {

        static bool _writeLogToFile = false;
        string[] _args;
        private AnalystConnection _analystDB = new AnalystConnection();
        private StandardPointName _standardPoint;
        private List<EstimateName> _newNamesToUpdateList;
        private bool _editProgrammatically = false;

        public Form1()
        {
             
            
            InitializeComponent();

            IniReader INIFile = new IniReader();
            bool debug = bool.Parse(INIFile.ReadString("General", "Debug", "false"));
            GenericTools.Debug = debug;
            _writeLogToFile = debug;


        }

        //AnalystConnection _analystDB = new AnalystConnection();

        static void Log(string Message)
        {            
            Console.WriteLine(Message);
            if (_writeLogToFile) GenericTools.WriteLog(Message);
        }

        private bool AllowDisconnect = true;

        public void Tree_Analyst(AnalystConnection cn, TreeView Tree)
        {
            if (cn.IsConnected)
            {
                Hashtable Nodes_Tree = new Hashtable();

                Tree.Invoke(new MethodInvoker(delegate
                {
                    Log("@A Stand - Conectado");

                    Nodes_Tree.Clear();
                    Nodes_Tree.Add("00321", Tree.Nodes.Add("00322", "Error TAGS"));

                    Tree.Enabled = true;
                    Tree.ImageList = imageList1;
                }));

                DataTable AnTblHierarchy;
                string TBLSETID = "";

                if (txtTBLSETID.Text != "")
                    TBLSETID = " AND TBLSETID=" + txtTBLSETID.Text + " ";

                AnTblHierarchy = cn.DataTable("*", "TREEELEM", "REFERENCEID=0 and HIERARCHYTYPE=1 and CONTAINERTYPE!=4" + TBLSETID + " and Parentid!=2147000000 ORDER BY tblsetid, branchlevel, SLOTNUMBER"); //parentid, containertype, branchlevel, SLOTNUMBER");
                Nodes_Tree = Monta_Arvore(_analystDB, AnTblHierarchy, Tree, Nodes_Tree);

                AnTblHierarchy.Clear();

                string Acc = Registration.RegistrationId(_analystDB, "SKFCM_ASPT_Acc").ToString();
                string Vel = Registration.RegistrationId(_analystDB, "SKFCM_ASPT_AccToVel").ToString();
                string Env = Registration.RegistrationId(_analystDB, "SKFCM_ASPT_AccEnvelope").ToString();
                string DAD = Registration.RegistrationId(_analystDB, "SKFCM_ASDD_MicrologDAD").ToString();
                string CAST = "";

                if (_analystDB.DBType == DBType.MSSQL)
                {
                    CAST = " and CAST(PT1.VALUESTRING AS numeric(17,2)) > " + float.Parse(Acel_GTRPM.Text) / (float)60;
                }
                else
                {
                    CAST = " and TO_NUMBER(PT1.VALUESTRING) > " + float.Parse(Acel_GTRPM.Text) / (float)60;
                }


                StringBuilder SQL_Point = new StringBuilder();
               
                SQL_Point.Append("   select ");
                SQL_Point.Append("   Distinct(TE1.TREEELEMID) ");
                SQL_Point.Append("   	, TE1.*");
                SQL_Point.Append("   	, PT1.VALUESTRING as SPEED_RPM ");
                SQL_Point.Append("   from  ");
                SQL_Point.Append("   	treeelem TE1 ");
                SQL_Point.Append("   	, TABLESET TB ");
                SQL_Point.Append("   	, TREEELEM TE2 ");
                SQL_Point.Append("   	, TREEELEM TE3 ");
                SQL_Point.Append("   	, POINT PT ");
                SQL_Point.Append("   	, POINT PT1 ");
                SQL_Point.Append("   	, POINT PT2 ");
                SQL_Point.Append("   	, REGISTRATION RG ");
                SQL_Point.Append("   	, REGISTRATION RG1 ");
                SQL_Point.Append("   where  ");
                SQL_Point.Append("   	TE1.CONTAINERTYPE=4 ");
                if (txtTBLSETID.Text != "")
                    TBLSETID = " AND TE1.TBLSETID=" + txtTBLSETID.Text + " ";
                SQL_Point.Append(TBLSETID);

                SQL_Point.Append("   	and TE1.parentid!=2147000000 ");
                SQL_Point.Append("   	and TE1.HIERARCHYTYPE=1 ");
                SQL_Point.Append("   	and TE1.ELEMENTENABLE=1 ");
                SQL_Point.Append("   	AND TE1.TBLSETID = TB.TBLSETID ");
                SQL_Point.Append("   	AND TE2.TREEELEMID=TE1.PARENTID ");
                SQL_Point.Append("   	AND TE3.TREEELEMID = TE2.PARENTID ");
                SQL_Point.Append("  	AND PT.ELEMENTID = TE1.TREEELEMID ");
                SQL_Point.Append("   	AND PT.VALUESTRING IN ('" + Acc + "', '" + Vel + "', '" + Env + "') ");
                SQL_Point.Append("   	AND PT1.ELEMENTID = PT.ELEMENTID ");
                SQL_Point.Append("   	AND RG.SIGNATURE = 'SKFCM_ASPF_Speed' ");
                SQL_Point.Append("   	AND PT1.FIELDID	= RG.REGISTRATIONID ");
                SQL_Point.Append("   	AND PT2.VALUESTRING ='" + DAD + "' ");
                SQL_Point.Append("   	AND PT2.ELEMENTID = PT1.ELEMENTID ");
                SQL_Point.Append("   	AND RG1.SIGNATURE = 'SKFCM_ASPF_Dad_Id' ");
                SQL_Point.Append("   	AND PT2.FIELDID	= RG1.REGISTRATIONID ");
                SQL_Point.Append("   order by  ");
                SQL_Point.Append("   	  TE1.TBLSETID");
                SQL_Point.Append("   	  ,TE1.containertype");
                SQL_Point.Append("   	  ,TE1.parentid");
                SQL_Point.Append("   	  ,TE1.branchlevel");
                SQL_Point.Append("   	  ,TE1.SLOTNUMBER");

                Log(SQL_Point.ToString());

                Tree.Invoke(new MethodInvoker(delegate
                {
                    lblPbHierarchy.Text = "Querying Point...";
                }));

                AnTblHierarchy = cn.DataTable(SQL_Point.ToString());
                Nodes_Tree = Monta_Arvore(_analystDB, AnTblHierarchy, Tree, Nodes_Tree);
                button1.Invoke(new MethodInvoker(delegate
                  {
                      button1.Enabled = true;
                      button_Progress_Run.Enabled = true;
                  }
                ));

            }
        }
        public void Tree_Analyst_WorkSpace(AnalystConnection cn, TreeView Tree, string WSFilter)
        {
            if (cn.IsConnected)
            {
                Hashtable Nodes_Tree = new Hashtable();

                Tree.Invoke(new MethodInvoker(delegate
                {
                    Log("@A Stand - Conectado");
                    Nodes_Tree.Clear();
                    Nodes_Tree.Add("00321", Tree.Nodes.Add("00322", "Error TAGS"));

                    Tree.Enabled = true;
                    Tree.ImageList = imageList1;
                }));

                DataTable AnTblHierarchy;
                string TBLSETID = "";

                if (txtTBLSETID.Text != "") 
                    TBLSETID = " AND TBLSETID=" +  txtTBLSETID.Text + " ";

                if (WSFilter == "")
                    AnTblHierarchy = cn.DataTable("*", "TREEELEM", "HIERARCHYTYPE=3 and CONTAINERTYPE!=4 " + TBLSETID + " and Parentid!=2147000000 and ORDER BY tblsetid, branchlevel, SLOTNUMBER"); //parentid, containertype, branchlevel, SLOTNUMBER");
                else
                    AnTblHierarchy = cn.DataTable("*", "TREEELEM", "HIERARCHYTYPE=3 and CONTAINERTYPE!=4 " + TBLSETID + " and Parentid!=2147000000 and hierarchyid in (SELECT TREEELEMID FROM TREEELEM WHERE HIERARCHYTYPE=3 and Parentid!=2147000000 and NAME LIKE '" + WSFilter + "') or NAME LIKE '" + WSFilter + "' ORDER BY tblsetid,  branchlevel, SLOTNUMBER"); //parentid, containertype, branchlevel, SLOTNUMBER");

                Tree.Invoke(new MethodInvoker(delegate
                {
                    lblPbHierarchy.Text = "Querying SET...";
                }));

                Nodes_Tree = Monta_Arvore(_analystDB, AnTblHierarchy, Tree, Nodes_Tree);

                AnTblHierarchy.Clear();

                string Acc = Registration.RegistrationId(_analystDB, "SKFCM_ASPT_Acc").ToString();
                string Vel = Registration.RegistrationId(_analystDB, "SKFCM_ASPT_AccToVel").ToString();
                string Env = Registration.RegistrationId(_analystDB, "SKFCM_ASPT_AccEnvelope").ToString();
                string DAD = Registration.RegistrationId(_analystDB, "SKFCM_ASDD_MicrologDAD").ToString();
                string CAST = "";

                if (_analystDB.DBType == DBType.MSSQL)
                {
                    CAST = " and CAST(PT1.VALUESTRING AS numeric(17,2)) > " + float.Parse(Acel_GTRPM.Text) / (float)60;
                }
                else
                {
                    CAST = " and TO_NUMBER(PT1.VALUESTRING) > " + float.Parse(Acel_GTRPM.Text) / (float)60;
                }

                StringBuilder SQL_Point = new StringBuilder();
                SQL_Point.Append("   select ");
                SQL_Point.Append("   Distinct(TE1.TREEELEMID) ");
                SQL_Point.Append("   	, TE1.*");
                SQL_Point.Append("   	, PT1.VALUESTRING as SPEED_RPM ");
                SQL_Point.Append("   from  ");
                SQL_Point.Append("   	treeelem TE1 ");
                SQL_Point.Append("   	, POINT PT ");
                SQL_Point.Append("   	, POINT PT1 ");
                SQL_Point.Append("   	, POINT PT2 ");
                SQL_Point.Append("   	, REGISTRATION RG ");
                SQL_Point.Append("   	, REGISTRATION RG1 ");
                SQL_Point.Append("   where  ");
                SQL_Point.Append("   	TE1.CONTAINERTYPE=4 ");
                SQL_Point.Append("   	and TE1.parentid!=2147000000 ");
                SQL_Point.Append("   	and TE1.HIERARCHYTYPE=3 ");
                SQL_Point.Append("   	 AND (TE1.PARENTENABLE=1 OR TE1.ELEMENTENABLE=1) ");

                if (txtTBLSETID.Text != "")
                    TBLSETID = " AND TE1.TBLSETID= " + txtTBLSETID.Text + " ";
                SQL_Point.Append(TBLSETID);

                if (WSFilter != "")
                    SQL_Point.Append("  and TE1.HIERARCHYID in (SELECT TREEELEMID FROM TREEELEM WHERE HIERARCHYTYPE=3 and Parentid!=2147000000 and NAME LIKE '" + WSFilter + "') ");

                SQL_Point.Append("  	AND PT.ELEMENTID = TE1.REFERENCEID ");
                SQL_Point.Append("   	AND PT.VALUESTRING IN ('" + Acc + "', '" + Vel + "', '" + Env + "') ");
                SQL_Point.Append("   	AND PT1.ELEMENTID = PT.ELEMENTID ");
                SQL_Point.Append("   	AND RG.SIGNATURE = 'SKFCM_ASPF_Speed' ");
                SQL_Point.Append("   	AND PT1.FIELDID	= RG.REGISTRATIONID ");
                SQL_Point.Append("   	AND PT2.VALUESTRING ='" + DAD + "' ");
                SQL_Point.Append("   	AND PT2.ELEMENTID = PT1.ELEMENTID ");
                SQL_Point.Append("   	AND RG1.SIGNATURE = 'SKFCM_ASPF_Dad_Id' ");
                SQL_Point.Append("   	AND PT2.FIELDID	= RG1.REGISTRATIONID ");
                //SQL_Point.Append(CAST);
                SQL_Point.Append("   order by  ");
                SQL_Point.Append("   	  TE1.TBLSETID");
                SQL_Point.Append("   	  ,TE1.containertype");
                SQL_Point.Append("   	  ,TE1.parentid");
                SQL_Point.Append("   	  ,TE1.branchlevel");
                SQL_Point.Append("   	  ,TE1.SLOTNUMBER");
                Tree.Invoke(new MethodInvoker(delegate
                {
                    lblPbHierarchy.Text = "Querying Point...";
                }));

                Log("SQL WS: " + SQL_Point.ToString());

                AnTblHierarchy = cn.DataTable(SQL_Point.ToString());
                Nodes_Tree = Monta_Arvore(_analystDB, AnTblHierarchy, Tree, Nodes_Tree);

                button1.Invoke(new MethodInvoker(delegate
            {
                button1.Enabled = true;
                button_Progress_Run.Enabled = true;
            }));

            }
        }
        public Hashtable Monta_Arvore(AnalystConnection cn, DataTable AnTblHierarchy, TreeView Tree, Hashtable Nodes_Tree)
        {
            TreeNode NewNode = null;
            TreeNode ParentNode = null;
            Log("Monta Arvore");
            pbHierarchy.Invoke(new MethodInvoker(delegate
            {

                pbHierarchy.Value = 0;
                pbHierarchy.Maximum = AnTblHierarchy.Rows.Count;
            }));

            lblPbHierarchy.Invoke(new MethodInvoker(delegate { lblPbHierarchy.Text = "0/" + AnTblHierarchy.Rows.Count; }));

            SKF.RS.STB.Analyst.Point AnPoint;

            for (Int32 i = 0; i < AnTblHierarchy.Rows.Count; i++)
            {
                Log(" Inicio do processo");
                if (Convert.ToInt64(AnTblHierarchy.Rows[i]["Parentid"]) == 0)
                {
                    Log("@A Stand - Add Hierarchy: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString());
                    Tree.Invoke(new MethodInvoker(delegate
                    {
                        Nodes_Tree.Add(Convert.ToInt64(AnTblHierarchy.Rows[i]["TreeelemId"])
                            , Tree.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString()
                            , AnTblHierarchy.Rows[i]["Name"].ToString()));

                        Log("Arvore Add: " + AnTblHierarchy.Rows[i]["Name"].ToString());
                    }));

                    
                }
                else
                {
                    NewNode = null;
                    ParentNode = (TreeNode)Nodes_Tree[Convert.ToInt64(AnTblHierarchy.Rows[i]["Parentid"])];


                    if ((Convert.ToInt64(AnTblHierarchy.Rows[i]["ContainerType"])) == 4)
                    {
                        Log(" Ponto ");
                        if (Convert.ToUInt32(AnTblHierarchy.Rows[i]["ReferenceID"]) == 0)
                        {
                            AnPoint = new SKF.RS.STB.Analyst.Point(cn, Convert.ToUInt32(AnTblHierarchy.Rows[i]["TreeElemId"]));
                            Log("Treeelem Point");
                        }
                        else
                        {
                            AnPoint = new SKF.RS.STB.Analyst.Point(cn, Convert.ToUInt32(AnTblHierarchy.Rows[i]["ReferenceID"]));
                            Log("Referenced Point");
                        }
                        float RPM = float.Parse(AnTblHierarchy.Rows[i]["SPEED_RPM"].ToString()); //GetPointRPM(cn, Convert.ToUInt32(AnTblHierarchy.Rows[i]["TreeElemId"]));
                        float tmp1 = float.Parse(Acel_GTRPM.Text);

                        //if (RPM > (tmp1 / 60))
                        //{
                            Log("RPM Dentro");
                            Tree.Invoke(new MethodInvoker(delegate
                            {
                                try
                                {

                                    Log("@A Stand - Add Point: " + AnPoint.Name + " - " + AnPoint.TreeElemId + " - RPM do Ponto: " + RPM + " - RPM setado: " + (tmp1 / 60));
                                    NewNode = ParentNode.Nodes.Add(AnPoint.TreeElemId.ToString(), AnPoint.Name.ToString(), 3, 3);
                                    Log(" Tag: " + AnTblHierarchy.Rows[i]["ReferenceId"].ToString());
                                    if (AnTblHierarchy.Rows[i]["ReferenceId"].ToString() != "0")
                                        NewNode.Tag = AnTblHierarchy.Rows[i]["ReferenceId"].ToString();

                                }
                                catch (Exception ex)
                                {
                                    if (ParentNode == null)
                                        ParentNode = (TreeNode)Nodes_Tree["00321"];

                                    NewNode = ParentNode.Nodes.Add(AnPoint.TreeElemId.ToString(), "ParentId: " + AnTblHierarchy.Rows[i]["ParentId"] + " - " + AnPoint.Name.ToString(), 3, 3);
                                    Log(" Tag: " + AnTblHierarchy.Rows[i]["ReferenceId"].ToString());
                                    if (AnTblHierarchy.Rows[i]["ReferenceId"].ToString() != "0")
                                        NewNode.Tag = AnTblHierarchy.Rows[i]["ReferenceId"].ToString();

                                    Log("ELEMENTO DE REFERÊNCIA NÃO ENCONTRADO: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " " + AnPoint.Name.ToString() + " - Exceção: " + ex.ToString());
                                }
                            }));
                    }
                    else if ((Convert.ToInt64(AnTblHierarchy.Rows[i]["ContainerType"])) == 3)
                    {
                        Tree.Invoke(new MethodInvoker(delegate
                        {
                            try
                            {
                                Log("@A Stand - Add Machine: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString());
                                NewNode = ParentNode.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString(), AnTblHierarchy.Rows[i]["Name"].ToString(), 2, 2);
                                Log(" Machine Added: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString());
                            }
                            catch (Exception ex)
                            {
                                if (ParentNode == null)
                                    ParentNode = (TreeNode)Nodes_Tree["00321"];

                                NewNode = ParentNode.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString(), "Parentid: " + AnTblHierarchy.Rows[i]["ParentId"].ToString() + " - " + AnTblHierarchy.Rows[i]["Name"].ToString(), 2, 2);
                                Log("ELEMENTO DE REFERÊNCIA NÃO ENCONTRADO: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " - Exceção: " + ex.ToString());
                            }
                        }));
                    }

                    else
                    {
                        Tree.Invoke(new MethodInvoker(delegate
                        {
                            try
                            {
                                Log("@A Stand - Add SET: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " - " + AnTblHierarchy.Rows[i]["Name"].ToString());

                                NewNode = ParentNode.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString(), AnTblHierarchy.Rows[i]["Name"].ToString(), 1, 1);
                                NewNode.StateImageIndex = 0;

                                Log(" SET Added: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString());

                            }
                            catch (Exception ex)
                            {
                                if (ParentNode == null)
                                    ParentNode = (TreeNode)Nodes_Tree["00321"];

                                NewNode = ParentNode.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString(), "Parentid: " + AnTblHierarchy.Rows[i]["ParentId"].ToString() + " - " + AnTblHierarchy.Rows[i]["Name"].ToString(), 1, 1);
                                NewNode.StateImageIndex = 0;
                                Log("ELEMENTO DE REFERÊNCIA NÃO ENCONTRADO: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " - Exceção: " + ex.ToString());
                            }
                        }));
                    }
                    if (NewNode != null)
                    {
                        Tree.Invoke(new MethodInvoker(delegate
                        {
                            try
                            {
                                Log("@A Stand - Add To Tree: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " - " + AnTblHierarchy.Rows[i]["Name"].ToString());
                                Nodes_Tree.Add(Convert.ToInt64(AnTblHierarchy.Rows[i]["TreeelemId"]), NewNode);
                                Log(" Added: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString());
                            }
                            catch (Exception ex)
                            {
                                Log("ELEMENTO DE REFERÊNCIA NÃO ENCONTRADO: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " - Exceção: " + ex.ToString());
                            }
                        }));
                    }
                }

                pbHierarchy.Invoke(new MethodInvoker(delegate { pbHierarchy.Value = i; }));
                lblPbHierarchy.Invoke(new MethodInvoker(delegate { lblPbHierarchy.Text = i + "/" + AnTblHierarchy.Rows.Count; }));
            }

            AnTblHierarchy.Dispose();

            pbHierarchy.Invoke(new MethodInvoker(delegate { pbHierarchy.Value = pbHierarchy.Maximum; }));
            lblPbHierarchy.Invoke(new MethodInvoker(delegate { lblPbHierarchy.Text = pbHierarchy.Maximum + "/" + AnTblHierarchy.Rows.Count; }));
            Tree.Invoke(new MethodInvoker(delegate { Tree.Enabled = true; }));
            Tree.Invoke(new MethodInvoker(delegate { Tree.SelectedNode = Tree.TopNode; }));

            return Nodes_Tree;
        }
        private float GetPointRPMHz(AnalystConnection cn, uint TreeelemiD)
        {
            float retorno = float.NaN;
            try
            {
                retorno = cn.SQLtoFloat("ValueString", "Point", "ElementId=" + TreeelemiD.ToString() + " AND FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Speed")); // cn.SQLtoString(SQL.ToString());

                Log("RPM Retorno:" + retorno);
                if (!float.IsNaN(retorno))
                {
                    //retorno = 0;
                    #region LinkReferencia
                    Log("RPM Retorno * 60:" + retorno);
                    if (retorno <= 0)
                    {
                        retorno = 0;
                        //retorno = cn.SQLtoFloat("ValueString", "Point", "ElementId=" + TreeelemiD + " AND FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Speed_Reference_Id"));
                        //Log("RPM Retorno Reference:" + retorno);
                        //if ((retorno != 0) && (!float.IsNaN(retorno)))
                        //{
                        //    SKF.RS.STB.Analyst.Point AnPoint = new SKF.RS.STB.Analyst.Point(cn, Convert.ToUInt32(retorno));
                        //    if (AnPoint.LastMeas.OverallReading != null)
                        //    {
                        //        float RPM1 = AnPoint.LastMeas.OverallReading.OverallValue; //cn.SQLtoString(SQL.ToString());
                        //        float Ratio = cn.SQLtoFloat("ValueString", "Point", "ElementId=" + TreeelemiD.ToString() + " and FieldId=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Speed_Ratio")); //cn.SQLtoString(SQL.ToString());

                        //        retorno = RPM1 * Ratio;
                        //        Log("RPM Retorno Ultima Velocidade * Radio: " + Math.Round(retorno, 1));

                        //        retorno = retorno / 60;
                        //    }
                        //    else
                        //    {
                        //        retorno = 0;
                        //    }

                        //}
                        //else
                        //{
                        //    retorno = 0;
                        //    Log("RPM Retorno: " + retorno);
                        //}
                    }
                    #endregion
                }
                else
                {
                    retorno = 0;
                    Log("RPM Retorno: != :" + retorno);
                }
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
                Log("ERRO TreeelemId: " + TreeelemiD + " Retorno: -1  ERRO: " + ex.ToString());

                retorno = 0;
            }
            return retorno;
        }
        private float GetPointRPM(AnalystConnection cn, uint TreeelemiD)
        {
            float retorno = float.NaN;
            try
            {
               retorno = cn.SQLtoFloat("ValueString", "Point", "ElementId=" + TreeelemiD.ToString() + " AND FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Speed")); // cn.SQLtoString(SQL.ToString());


         
                Log("RPM Retorno:" + retorno);
                if (!float.IsNaN(retorno))
                {

                    retorno = retorno * 60;
                    Log("RPM Retorno * 60:" + retorno);
                    if (retorno <= 0)
                    {
                        retorno = cn.SQLtoFloat("ValueString", "Point", "ElementId=" + TreeelemiD + " AND FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Speed_Reference_Id"));

                        Log("RPM Retorno Reference:" + retorno);
                        if ((retorno != 0) && (!float.IsNaN(retorno)))
                        {
                            SKF.RS.STB.Analyst.Point AnPoint = new SKF.RS.STB.Analyst.Point(cn, Convert.ToUInt32(retorno));

                            float RPM1 = AnPoint.LastMeas.OverallReading.OverallValue; //cn.SQLtoString(SQL.ToString());

                            float Ratio = cn.SQLtoFloat("ValueString", "Point", "ElementId=" + TreeelemiD.ToString() + " and FieldId=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Speed_Ratio")); //cn.SQLtoString(SQL.ToString());

                            retorno = RPM1 * Ratio;
                            Log("RPM Retorno Ultima Velocidade * Ratio: " + Math.Round(retorno,1));
                        }
                        else
                        {
                            retorno = -1;
                            Log("RPM Retorno: " + retorno);
                        }
                    }

                }
                else
                {
                    retorno = -1;
                    Log("RPM Retorno: != :" + retorno);
                }
            }
            catch (Exception ex)
            {
                Log(ex.ToString());
                Log("ERRO TreeelemId: " + TreeelemiD + " Retorno: -1  ERRO: " + ex.ToString());

                retorno = -1;
            }
            return retorno;
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            GenericTools.Debug = true;
            string INIFile = GenericTools.GetAuxFileName(".ini");
            IniReader Ini = new IniReader();
            bool IniFound = false;

            _args = Program.args;

            if (_args.Length == 0)
            {
                if (File.Exists(INIFile))
                {
                    Ini.Filename = INIFile;
                    IniFound = true;

                    GenericTools.Debug = Ini.ReadBoolean("General", "Debug", false);
                    if (GenericTools.Debug) MessageBox.Show("Debug mode active!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    GenericTools.Interactive = Ini.ReadBoolean("General", "Interactive", false);
                    if (GenericTools.Interactive) MessageBox.Show("Interactive mode active!", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                _ReadIniDatabases(Ini, IniFound);
            }
            else if (_args.Length == 5)
            {
                Parameter_Connect(_args[0].ToString(), _args[1].ToString(), _args[2].ToString(), _args[3].ToString(), _args[4].ToString());
            }
            else if (_args.Length == 4)
            {
                Parameter_Connect("1", _args[0].ToString(), _args[1].ToString(), _args[2].ToString(), _args[3].ToString());
            }
           
            _RefreshForm();
        }

        private void updateHierarchy(ComboBox cb, bool Option4AllHierarchies = false)
        {

            cb.Items.Clear();

            List<Hierarchy> listHierarchy = new List<Hierarchy>();
            try
            {

                AnalystConnection AnConn = _analystDB;

                if (_analystDB.IsConnected)
                {
                    //AnalystConnection AnConn = new AnalystConnection(DBType.MSSQL, txtAnalystServer.Text, "SKFUser", txtAnalystDBUser.Text, "cm");
                    DataTable tblset = AnConn.DataTable("SELECT * FROM TABLESET");
                    foreach (DataRow dr in tblset.Rows)
                    {
                        Hierarchy itm = new Hierarchy();
                        itm.Name = dr["TBLSETID"].ToString() + " - " + dr["TBLSETNAME"].ToString();
                        itm.TblSetId = uint.Parse(dr["TBLSETID"].ToString());
                        listHierarchy.Add(itm);
                    }

                    cb.DisplayMember = "Text";
                    cb.ValueMember = "Value";

                    if (Option4AllHierarchies)
                        cb.Items.Add(new { Text = "All", Value = 0 });

                    foreach (Hierarchy itm in listHierarchy)
                    {
                        cb.Items.Add(new { Text = itm.Name, Value = itm.TblSetId });
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }
        private void updateWorkspace(int strHierarquia, ComboBox cb)
        {

            cb.Items.Clear();

            List<Hierarchy> listHierarchy = new List<Hierarchy>();
            try
            {
                AnalystConnection AnConn = _analystDB;

                string hierarchy = "";
                if (strHierarquia > 0)
                    hierarchy = "AND TE.TBLSETID=" + strHierarquia;

                DataTable tblset = AnConn.DataTable("select WS.ELEMENTID , TE.NAME from WORKSPACE WS, TREEELEM TE WHERE	TE.TREEELEMID=WS.ELEMENTID " + hierarchy + " ORDER BY NAME ASC");
                foreach (DataRow dr in tblset.Rows)
                {
                    Hierarchy itm = new Hierarchy();
                    itm.Name = dr["NAME"].ToString();
                    itm.TblSetId = uint.Parse(dr["ELEMENTID"].ToString());
                    listHierarchy.Add(itm);
                }

                cb.DisplayMember = "Text";
                cb.ValueMember = "Value";


                foreach (Hierarchy itm in listHierarchy)
                {
                    cb.Items.Add(new { Text = itm.Name, Value = itm.TblSetId });
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }
       
        public class Hierarchy
        {
            public string Name { get; set; }
            public uint TblSetId { get; set; }
        }
        private void Parameter_Connect(string TypeDB, string AnServer, string DsServer, string SAM_DB, string SAM_Path)
        {
            if (TypeDB == "1")
            {
                groupBox_AnDB.Enabled = AllowDisconnect;
                _analystDB.DBType = DBType.Oracle;
                _analystDB.DBName = AnServer;
                _analystDB.InitialCatalog = "";
                _analystDB.Password = "cm";
                _analystDB.User = "SKFUser1";
            }

            else if (TypeDB == "2")
            {

                groupBox_AnDB.Enabled = AllowDisconnect;
                _analystDB.DBType = DBType.MSSQL;
                _analystDB.DBName = "SKF-SQLDB01";
                _analystDB.InitialCatalog = AnServer;
                _analystDB.Password = "cm";
                _analystDB.User = "skfuser" + AnServer;

                button_Connect_AnDB_Click(null, null);
            }
        }
        private void _ReadIniDatabases(IniReader Ini, bool IniFound)
        {
            if (IniFound)
            {

                int TmpDBType = Ini.ReadInteger("Analyst", "DBType", (int)DBType.None);
                if (TmpDBType != (int)DBType.None)
                {
                    groupBox_AnDB.Enabled = AllowDisconnect;
                    _analystDB.DBType = (TmpDBType == 1 ? DBType.Oracle : (TmpDBType == 2 ? DBType.MSSQL : DBType.Access));

                    _analystDB.DBName = Ini.ReadString("Analyst", "DBName", string.Empty);
                    _analystDB.InitialCatalog = Ini.ReadString("Analyst", "InitialCatalog", "");
                    _analystDB.Password = Ini.ReadString("Analyst", "Password", "cm");
                    _analystDB.User = Ini.ReadString("Analyst", "User", "SKFUser1");

                    button_Connect_AnDB_Click(null, null);
                }
            }
        }
        private void button_Connect_AnDB_Click(object sender, EventArgs e)
        {

            if (sender != null)
            {
                if (radioButton_AnDB_Oracle.Checked)
                {
                    _analystDB.DBType = DBType.Oracle;
                }
                else if (radioButton_AnDB_MSSQL.Checked)
                {
                    _analystDB.DBType = DBType.MSSQL;
                }
                else
                {
                    _analystDB.DBType = DBType.None;
                }
                _analystDB.DBName = comboBox_AnDB.Text;
                _analystDB.InitialCatalog = textBox_InitialCatalog_AnDB.Text;
                _analystDB.User = textBox_Username_AnDB.Text;
                _analystDB.Password = textBox_Password_AnDB.Text;
            }
            else
            {
                radioButton_AnDB_Oracle.Checked = (_analystDB.DBType == DBType.Oracle);
                radioButton_AnDB_MSSQL.Checked = (_analystDB.DBType == DBType.MSSQL);
                comboBox_AnDB.Text = _analystDB.DBName;
                textBox_InitialCatalog_AnDB.Text = _analystDB.InitialCatalog;
                textBox_Username_AnDB.Text = _analystDB.User;
                textBox_Password_AnDB.Text = _analystDB.Password;
            }

            _analystDB.Connect();

            _RefreshForm();
        }
        private void _StoreIniFile()
        {
            //_Wait_Show("Writing settings to INI file");

            string INIFile = GenericTools.GetAuxFileName(".ini");
            try
            {
                IniReader Ini = new IniReader(INIFile);

                //Database
                _StoreIniDatabases(Ini);

            }
            catch (Exception ex)
            {
                GenericTools.GetError("Error writing to INI file (" + INIFile + "): " + ex.Message);
            }
        }
        private void _StoreIniDatabases(IniReader Ini)
        {

            // Analyst
            Ini.Write("Analyst", "DBType", (int)_analystDB.DBType);
            Ini.Write("Analyst", "DBName", _analystDB.DBName);
            Ini.Write("Analyst", "InitialCatalog", _analystDB.InitialCatalog);
            Ini.Write("Analyst", "User", _analystDB.User);
            Ini.Write("Analyst", "Password", _analystDB.Password);

        }
        private void button_Disconnect_AnDB_Click(object sender, EventArgs e)
        {
            _analystDB.Close();
            _RefreshForm();
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            _StoreIniFile();
          
        }
        private void _RefreshForm()
        {
            button_Connect_AnDB.Enabled = (comboBox_AnDB.Text != string.Empty) & !_analystDB.IsConnected;
            button_Disconnect_AnDB.Enabled = _analystDB.IsConnected;
            textBox_InitialCatalog_AnDB.Enabled = radioButton_AnDB_MSSQL.Checked & !_analystDB.IsConnected;
            label_InitialCatalog_AnDB.Enabled = radioButton_AnDB_MSSQL.Checked & !_analystDB.IsConnected;
            tableLayoutPanel_Server_AnDB.Enabled = !_analystDB.IsConnected;
            tableLayoutPanel_Connection_AnDB.Enabled = !_analystDB.IsConnected;

            updateHierarchy(cbASS_Hierarchy, true);
            cbASS_Hierarchy.SelectedIndex = 0;

            updateHierarchy(tbl_setCombobox, true);
            tbl_setCombobox.SelectedIndex = 0;
            
            updateHierarchy(cbCSS_Hierarchy, true);
            cbCSS_Hierarchy.SelectedIndex = 0;

            updateHierarchy(cbCOC_Hierarchy, true);
            cbCOC_Hierarchy.SelectedIndex = 0;

            updateHierarchy(cbRRS_Hierarchy, true);
            cbRRS_Hierarchy.SelectedIndex = 0;


        }
        private void label1_Click(object sender, EventArgs e)
        {
            if (Control.ModifierKeys == Keys.Control)
            {
                if (FormBorderStyle == FormBorderStyle.FixedSingle)
                    FormBorderStyle = FormBorderStyle.Sizable;
                else
                    FormBorderStyle = FormBorderStyle.FixedSingle;
            }
        }
        private void CheckAllChildNodes(TreeNode treeNode, bool nodeChecked)
        {
            foreach (TreeNode node in treeNode.Nodes)
            
            {
                node.Checked = nodeChecked;
                if (node.Nodes.Count > 0)
                {
                    this.CheckAllChildNodes(node, nodeChecked);
                }
            }
        }
        private void label4_Click(object sender, EventArgs e)
        {

        }
        private void label12_Click(object sender, EventArgs e)
        {

        }
        private void button_Progress_Run_Click(object sender, EventArgs e)
        {
            tCount = 0;
            progressBar_Points.Value = 0;
            progressBar_Points.Maximum = 0;
            label_Machine_Left.Text = "0/0";

            CallRecursive(treeView_AnStandards);

            progressBar_Points.Maximum = tCount;
            label_Machine_Left.Text = "0/" + progressBar_Points.Maximum.ToString();


            Thread LoadIt = new Thread(() => { Replicate(); });
            LoadIt.Start();

        }

        private void Replicate()
        {

            for (Int32 j = 0; j < treeView_AnStandards.Nodes.Count; j++)
            {
                TreeView_Check(treeView_AnStandards.Nodes[j]);
            }
        }

        private void TreeView_Check(TreeNode treeNode)
        {
            if (treeNode.Checked == true && treeNode.ImageIndex == 3)
            {
                uint PointId;

                if (treeNode.Tag != null)
                    PointId = Convert.ToUInt32(treeNode.Tag);
                else
                    PointId = Convert.ToUInt32(treeNode.Name);

                if (PointId == 626018)
                {
                    string strop = "";
                }
                SKF.RS.STB.Analyst.Point AnPoint = new SKF.RS.STB.Analyst.Point(_analystDB, PointId);

                float RPM = GetPointRPM(_analystDB, PointId);
                float FTF = Get_FTF(_analystDB, PointId);
                float BPFI = Get_Great_BPFI(_analystDB, PointId);

                Log("Point: " + AnPoint.Name + " - TreeelemId: " + AnPoint.TreeElemId);

                #region TRATAMENTO EM ACELERAÇÃO
                if (AnPoint.PointType == PointType.Acc)
                {
                    // ACELERATION

                    
                    uint acel_detection = (uint)Acc.DetectionType;
                    uint acel_savedate = (uint)Acc.SaveData;
                    uint acel_lines = 0;
                    uint acel_averages = (uint)Acc.Averages;
                    float acel_startfreq = (uint)Acc.StartFreq;
                    float acel_stopfreq = 0;
                    float acel_lowcutoff = 0;

                    AccParam(AnPoint, RPM, FTF, BPFI, out acel_detection, out acel_savedate, out acel_lines, out acel_averages, out acel_startfreq, out acel_stopfreq, out acel_lowcutoff);
                    try
                    {
                        Reply_Aceleration(
                             _analystDB
                             , PointId
                             , acel_lines
                             , acel_averages
                             , acel_startfreq
                             , acel_stopfreq
                             , acel_lowcutoff
                             , acel_savedate.ToString()
                             , acel_detection.ToString());
                    }
                    catch (Exception ex)
                    {
                        Log("ERRO REPLY ACC: " + ex.Message.ToString());
                    }
                }


                #endregion

                #region TRATAMENTO ENVELOPE
                if (AnPoint.PointType == PointType.AccEnvelope)
                {

                    // ENVELOPE
                    Log("ENV: Initial Procedure Pt: " + AnPoint.TreeElemId.ToString());

                    uint env_detection = (uint)Env.DetectionType;
                    uint env_savedate = (uint)Env.SaveData;
                    uint env_averages = (uint)Env.Averages;
                    float env_startfreq = 0;
                    float env_stopfreq = 0;  //(float)Math.Ceiling(Convert.ToDecimal(Env_Range(BPFI, RPM)));
                    float env_lowcutoff = 0;
                    uint env_lines = 0;

                    EnvParam(AnPoint, RPM, FTF, BPFI, out env_detection, out env_savedate, out env_lines, out env_averages, out env_startfreq, out env_stopfreq, out env_lowcutoff);

                    try
                    {
                        Reply_Envelope(_analystDB
                            , PointId
                            , env_lines
                            , env_averages
                            , env_startfreq
                            , env_stopfreq
                            , env_lowcutoff
                            , env_savedate.ToString()
                            , env_detection.ToString());
                    }
                    catch (Exception ex)
                    {
                        Log("ERRO REPLY ENV: " + ex.Message.ToString());
                    }
                }
                #endregion

                #region TRATAMENTO VELOCIDADE
                if (AnPoint.PointType == PointType.AccToVel)
                {
                    
                    Log("VEL: Initial Procedure Pt: " + AnPoint.TreeElemId.ToString());

                    uint vel_detection = (uint)Vel.DetectionType;
                    uint vel_savedate = (uint)Vel.SaveData;
                    uint vel_lines = (uint)Vel.Lines;
                    uint vel_averages = (uint)Vel.Averages;
                    float vel_startfreq = (uint)Vel.StartFreq;                   
                    float vel_lowcutoff = 0;
                    float vel_stopfreq = 0;


                    VelParam(AnPoint, RPM, FTF, BPFI, out vel_detection, out vel_savedate, out vel_lines, out vel_averages, out vel_startfreq, out vel_stopfreq, out vel_lowcutoff);
                    
                    try
                    { 
                        if (
                                    treeNode.Text.Length <= 4
                                ||  treeNode.Text.Substring(Math.Max(0, treeNode.Text.Length - 5), 5).ToUpper().Trim() != "SPEED"
                                &&  treeNode.Text.Substring(Math.Max(0, treeNode.Text.Length - 4), 4).ToUpper().Trim() != "ZOOM"
                                &&  treeNode.Text.Substring(Math.Max(0, treeNode.Text.Length - 2), 2).ToUpper().Trim() != "HD"
                                &&  treeNode.Text.Substring(Math.Max(0, treeNode.Text.Length - 2), 2).ToUpper().Trim() != "HR"
                            )
                            Reply_Velocity(_analystDB
                                , PointId
                                , vel_lines
                                , vel_averages
                                , vel_startfreq
                                , vel_stopfreq
                                , vel_lowcutoff
                                , vel_savedate.ToString()
                                , vel_detection.ToString());

                    }
                    catch (Exception ex)
                    {
                        Log("ERRO REPLY VEL: " + ex.Message.ToString());
                    }
                    if (treeNode.Text.Length > 4)
                    {
                        if (
                               treeNode.Text.Substring(Math.Max(0, treeNode.Text.Length - 4), 4).ToUpper().Trim() == "ZOOM"
                            || treeNode.Text.Substring(Math.Max(0, treeNode.Text.Length - 5), 5).ToUpper().Trim() == "SPEED"
                            || treeNode.Text.Substring(Math.Max(0, treeNode.Text.Length - 2), 2).ToUpper().Trim() == "HD"
                            || treeNode.Text.Substring(Math.Max(0, treeNode.Text.Length - 2), 2).ToUpper().Trim() == "HR"
                        )
                        {
                            uint velz_SaveData = (uint)Zoom.SaveData;
                            uint velz_DetectionType = (uint)Zoom.DetectionType;
                            uint velz_lines;
                            uint velz_averages;
                            float velz_stopfreq;
                            float velz_lowcutoff;

                            VelZoomParam(out velz_SaveData, out velz_DetectionType, out velz_stopfreq, out velz_lines, out velz_averages, out velz_lowcutoff);

                            try
                            {
                                Reply_Velocity_ZOOM(_analystDB, PointId, velz_lines, velz_averages, velz_stopfreq, velz_lowcutoff, velz_SaveData.ToString(), velz_DetectionType.ToString());
                            }
                            catch (Exception ex)
                            {
                                Log("ERRO REPLY VEL ZOOM: " + ex.Message.ToString());
                            }
                        }
                    }
                }
                #endregion

                progressBar_Points.Invoke(new MethodInvoker(delegate
                {
                    progressBar_Points.Value = progressBar_Points.Value + 1;
                    label_Machine_Left.Text = progressBar_Points.Value + "/" + progressBar_Points.Maximum.ToString();
                }));

            }

            foreach (TreeNode tn in treeNode.Nodes)
            {
                TreeView_Check(tn);
            }
        }
        private void VelZoomParam(out uint velz_SaveData
            , out uint velz_DetectionType
            , out float velz_stopfreq
            , out uint velz_lines
            , out uint velz_averages
            , out float velz_lowcutoff)
        {
            velz_SaveData = (uint)Zoom.SaveData;
            velz_DetectionType = (uint)Zoom.DetectionType;
            velz_stopfreq = 400;
            velz_lines = 6400;
            velz_averages = 1;
            velz_lowcutoff = 2;
        }

        private void AccParam(SKF.RS.STB.Analyst.Point AnPoint, float RPM, float FTF, float BPFI
            , out uint acel_detection
            , out uint acel_savedate
            , out uint acel_lines
            , out uint acel_averages
            , out float acel_startfreq
            , out float acel_stopfreq
            , out float acel_lowcutoff
            )
        {

                acel_detection = (uint)Acc.DetectionType;
                acel_savedate = (uint)Acc.SaveData;
                acel_lines = 0;
                acel_averages = (uint)Acc.Averages;
                acel_startfreq = (uint)Acc.StartFreq;
                acel_stopfreq = 0;
                acel_lowcutoff = -1; //PC Why it is -1?

                DataTable dt_Freq = getOtherFrequencies(_analystDB, AnPoint.PointId);
                int variavel = 0;
                if (dt_Freq.Rows.Count > 0)
                    variavel = int.Parse(dt_Freq.Rows[0]["RPM"].ToString());

                float RPMHz = GetPointRPMHz(_analystDB, AnPoint.PointId);
                float maiorFrequencia = variavel * RPMHz;

                float maiorFrequencia3Harmonicas = (maiorFrequencia * (float)3.5);
                float maiorFrequenciaRolamento = (BPFI * RPMHz) * (float)5.5;

                float frequenciaFinal;

                if (maiorFrequenciaRolamento > maiorFrequencia3Harmonicas)
                    frequenciaFinal = (float)Math.Ceiling(Convert.ToDecimal(maiorFrequenciaRolamento));
                else
                    frequenciaFinal = (float)Math.Ceiling(Convert.ToDecimal(maiorFrequencia3Harmonicas));

                float RotacaoReferencia = 600;
                float RotacaoReferenciaHz = RotacaoReferencia / 60;

                float RotacaoReferenciaIntermediaria = 200;
                float RotacaoReferenciaIntermediariaHz = RotacaoReferenciaIntermediaria / 60;

                #region ACELERAÇÃO DE BAIXA
                if (AnPoint.Name.Substring(AnPoint.Name.Length - 3, 3) == "LOW" || AnPoint.Name.Substring(AnPoint.Name.Length - 2, 2) == "LF")
                {
                    if (RPMHz == 0 || RPMHz < 0 )
                    {
                        #region Rotação = 0
                        acel_stopfreq = 2000;
                        acel_lines = 3200;
                        #endregion
                    }
                    else if (RPMHz > RotacaoReferenciaHz)
                    {
                        #region Rotação Maior que 600 RPM
                        if (variavel == 0 && BPFI == 0)
                        {
                            //acel_stopfreq = 2000;
                            //acel_lines = NumeroDeLinhasAproximado(3200);

                            acel_stopfreq = ((int)Math.Floor(Math.Ceiling((float)RPMHz * 55))); //PC changed - 50
                            acel_stopfreq = CalculoFrequenciaFinal(acel_stopfreq);
                            acel_lines = NumeroDeLinhasAproximado((int)Math.Floor(Math.Ceiling((((float)2.5 / (float)RPMHz) * 16) * acel_stopfreq)));
                        }
                        else
                        {
                            float maiorFrequenciaHarmonicas = ((variavel * RPMHz) * (float)3.5);
                            maiorFrequenciaRolamento = (BPFI * RPMHz) * (float)5.5;

                            if (maiorFrequenciaRolamento > maiorFrequenciaHarmonicas)
                                frequenciaFinal = (float)Math.Ceiling(Convert.ToDecimal(maiorFrequenciaRolamento));
                            else
                                frequenciaFinal = (float)Math.Ceiling(Convert.ToDecimal(maiorFrequenciaHarmonicas));

                            acel_stopfreq = CalculoFrequenciaFinal(frequenciaFinal);

                           // if (acel_stopfreq > 2000)
                           //     acel_stopfreq = 2000;

                            acel_lines = NumeroDeLinhasAproximado((int)Math.Floor(Math.Ceiling((((float)2.5 / (float)RPMHz) * 16) * acel_stopfreq)));
                        }
                        #endregion
                    }
                    else if (RPMHz >= RotacaoReferenciaIntermediariaHz && RPMHz <= RotacaoReferenciaHz)
                    {
                        #region Rotação Maior que 200 RPM e Menor que 600 RPM
                        if (variavel == 0 && BPFI == 0)
                        {
                            acel_stopfreq = ((int)Math.Floor(Math.Ceiling((float)RPMHz * 105))); //PC changed - 150
                            acel_stopfreq = CalculoFrequenciaFinal(acel_stopfreq);
                            acel_lines = NumeroDeLinhasAproximado((int)Math.Floor(Math.Ceiling((((float)2.5 / (float)RPMHz) * 8) * acel_stopfreq)));
                            // acel_stopfreq = 500;
                            // acel_lines = NumeroDeLinhasAproximado(1600);
                        }
                        else
                        {
                            float maiorFrequenciaHarmonicas = ((variavel * RPMHz) * (float)6.5);
                            maiorFrequenciaRolamento = (BPFI * RPMHz) * (float)10.5;

                            if (maiorFrequenciaRolamento > maiorFrequenciaHarmonicas)
                                frequenciaFinal = CalculoFrequenciaFinal((float)Math.Ceiling(Convert.ToDecimal(maiorFrequenciaRolamento)));
                            else
                                frequenciaFinal = (float)Math.Ceiling(Convert.ToDecimal(maiorFrequenciaHarmonicas));


                            acel_stopfreq = (float)Math.Ceiling(Convert.ToDecimal(frequenciaFinal));

                            if (acel_stopfreq > 2000)
                                acel_stopfreq = 2000;

                            acel_lines = NumeroDeLinhasAproximado((int)Math.Floor(Math.Ceiling((((float)2.5 / (float)RPMHz) * 8) * acel_stopfreq)));

                        }
                        #endregion
                    }
                    else if (RPMHz < RotacaoReferenciaIntermediariaHz)
                    {
                        #region Rotação Menor que 200 RPM
                        if (variavel == 0 && BPFI == 0)
                        {
                            acel_stopfreq = ((int)Math.Floor(Math.Ceiling((float)RPMHz * 105))); //PC changed - 50
                            acel_stopfreq = CalculoFrequenciaFinal(acel_stopfreq);
                            acel_lines = NumeroDeLinhasAproximado((int)Math.Floor(Math.Ceiling((((float)2.5 / (float)RPMHz) * 8) * acel_stopfreq)));
                            // acel_stopfreq = 250;
                            // acel_lines = NumeroDeLinhasAproximado(800);
                        }
                        else
                        {
                            float maiorFrequenciaHarmonicas = ((variavel * RPMHz) * (float)6.5);
                            maiorFrequenciaRolamento = (BPFI * RPMHz) * (float)10.5;

                            if (maiorFrequenciaRolamento > maiorFrequenciaHarmonicas)
                                frequenciaFinal = CalculoFrequenciaFinal((float)Math.Ceiling(Convert.ToDecimal(maiorFrequenciaRolamento)));
                            else
                                frequenciaFinal = (float)Math.Ceiling(Convert.ToDecimal(maiorFrequenciaHarmonicas));


                            acel_stopfreq = (float)Math.Ceiling(Convert.ToDecimal(frequenciaFinal));


                            if (acel_stopfreq > 2000)
                                acel_stopfreq = 2000;

                            // if (acel_stopfreq < 250)
                            //     acel_stopfreq = CalculoFrequenciaFinal(acel_stopfreq);

                            acel_lines = NumeroDeLinhasAproximado((int)Math.Floor(Math.Ceiling((((float)2.5 / (float)RPMHz) * 8) * acel_stopfreq)));
                        }
                        #endregion
                    }
                    acel_lowcutoff = CalculoLowCutoff(RPMHz, acel_stopfreq, acel_lines); 
                    try
                    {
                       
                    }
                    catch (Exception ex)
                    {
                        Log("ERRO REPLY ACC_LOW: " + ex.Message.ToString());
                    }

                }
                #endregion

                #region ACELERAÇÃO TRADICIONAL
                if (AnPoint.Name.Substring(AnPoint.Name.Length - 3, 3) != "LOW" && AnPoint.Name.Substring(AnPoint.Name.Length - 2, 2) != "LF")
                {
                    if (maiorFrequencia == 0)
                    {
                        acel_stopfreq = 5000;
                        acel_lines = 3200;
                    }
                    else if (frequenciaFinal < 10000)
                    {
                        if (frequenciaFinal < 5000)
                        {
                            acel_stopfreq = 5000;
                            acel_lines = 3200;
                        }
                        else
                        {
                            acel_stopfreq = 10000;
                            acel_lines = 6400;
                        }
                    }
                    else
                    {
                        acel_stopfreq = 10000;
                        acel_lines = 6400;
                    }


                    //acel_lowcutoff = CalculoLowCutoff(RPMHz, acel_stopfreq, acel_lines); //(float)Math.Ceiling(Convert.ToDecimal((acel_stopfreq /acel_lines) * 5));  // = (uint)Acc.LowCutOff;
                    acel_lowcutoff = 10;
                #endregion
     
                try
                {
                      
                }
                catch (Exception ex)
                {
                    Log("ERRO REPLY ACC: " + ex.Message.ToString());
                }
            }
        }

        private void EnvParam(SKF.RS.STB.Analyst.Point AnPoint, float RPM, float FTF, float BPFI
            , out uint env_detection
            , out uint env_savedate
            , out uint env_lines
            , out uint env_averages
            , out float env_startfreq
            , out float env_stopfreq
            , out float env_lowcutoff)
        {
            env_detection = (uint)Env.DetectionType;
            env_savedate = (uint)Env.SaveData;
            env_averages = (uint)Env.Averages;
            env_startfreq = 0;
            env_stopfreq = 0;  //(float)Math.Ceiling(Convert.ToDecimal(Env_Range(BPFI, RPM)));
            env_lines = 0;
            env_lowcutoff = 0;

            float RotacaoReferencia = 600;
            float RotacaoReferenciaHz = RotacaoReferencia / 60;

            float RotacaoReferenciaBaixissima = 8;
            float RotacaoReferenciaBaixissimaHz = RotacaoReferenciaBaixissima / 60;


            float RPMHz = GetPointRPMHz(_analystDB, AnPoint.PointId);

            if (RPMHz == 0 || RPMHz < 0)
            {
                env_stopfreq = 2000;
                env_lines = 3200;
            }
            else if (RPMHz > RotacaoReferenciaHz)
            {
                if (BPFI == 0)
                {
                    env_stopfreq = RPMHz * 55; //PC Changed to 55
                    env_stopfreq = CalculoFrequenciaFinal(env_stopfreq);
                    env_lines = NumeroDeLinhasAproximado((int)Math.Floor(Math.Ceiling((((float)2.5 / (float)RPMHz) * 16) * env_stopfreq)));
                }
                else
                {
                    env_stopfreq = (float)Math.Ceiling(Convert.ToDecimal((RPMHz * BPFI) * (float)5.5));
                    env_stopfreq = CalculoFrequenciaFinal(env_stopfreq);
                    env_lines = NumeroDeLinhasAproximado((int)Math.Floor(Math.Ceiling((((float)2.5 / (float)RPMHz) * 16) * env_stopfreq)));
                }
            }
            else if (RPMHz <= RotacaoReferenciaHz)
            {
                if (RPMHz <= RotacaoReferenciaBaixissimaHz)
                {
                    if (BPFI == 0)
                    {
                        env_stopfreq = RPMHz * 105; //PC changed to 105
                        /// ARREDONDAMENTO CONFORME PLANILHA
                        env_stopfreq = CalculoFrequenciaFinal(env_stopfreq);
                        env_lines = NumeroDeLinhasAproximado((int)Math.Floor(Math.Ceiling((((float)2.5 / (float)RPMHz) * 8) * env_stopfreq)));
                    }
                    else
                    {

                        env_stopfreq = (float)Math.Ceiling(Convert.ToDecimal((RPMHz * BPFI) * (float)10.5));
                        env_stopfreq = CalculoFrequenciaFinal(env_stopfreq);
                        env_lines = NumeroDeLinhasAproximado((int)Math.Floor(Math.Ceiling((((float)2.5 / (float)RPMHz) * 8) * env_stopfreq)));
                    }
                }
                else
                {
                    if (BPFI == 0)
                    {
                        env_stopfreq = RPMHz * 105; //PC changed to 105
                        env_stopfreq = CalculoFrequenciaFinal(env_stopfreq);
                        env_lines = NumeroDeLinhasAproximado((int)Math.Floor(Math.Ceiling((((float)2.5 / (float)RPMHz) * 8) * env_stopfreq)));
                    }
                    else
                    {

                        env_stopfreq = (float)Math.Ceiling(Convert.ToDecimal((RPMHz * BPFI) * (float)10.5));
                        env_stopfreq = CalculoFrequenciaFinal(env_stopfreq);
                        env_lines = NumeroDeLinhasAproximado((int)Math.Floor(Math.Ceiling((((float)2.5 / (float)RPMHz) * 8) * env_stopfreq)));
                    }
                }
            }
        }


        private void VelParam(SKF.RS.STB.Analyst.Point AnPoint, float RPM, float FTF, float BPFI
            , out uint vel_detection
            , out uint vel_savedate
            , out uint vel_lines
            , out uint vel_averages
            , out float vel_startfreq
            , out float vel_stopfreq
            , out float vel_lowcutoff)
        {

            vel_detection = (uint)Vel.DetectionType;
            vel_savedate = (uint)Vel.SaveData;
            vel_lines = (uint)Vel.Lines;
            vel_averages = (uint)Vel.Averages;
            vel_startfreq = (uint)Vel.StartFreq;
            vel_lowcutoff = 0;
            vel_stopfreq = 0;


            float RotacaoReferencia = 600;
            float RotacaoReferenciaHz = RotacaoReferencia / 60;

            float RotacaoReferenciaIntermediaria = 200;
            float RotacaoReferenciaIntermediariaHz = RotacaoReferenciaIntermediaria / 60;

            float RPMHz = GetPointRPMHz(_analystDB, AnPoint.PointId);
            float frequencia = 0;

            float frequenciaFinal;

            DataTable dt_Freq = getOtherFrequencies(_analystDB, AnPoint.PointId);
            if (dt_Freq.Rows.Count > 0)
                frequencia = float.Parse(dt_Freq.Rows[0]["RPM"].ToString());

            if (RPMHz == 0 || RPMHz < 0)
            {
                #region Rotação = 0
                vel_stopfreq = 2000;
                vel_lines = 1600;
                #endregion
            }
            else if (RPMHz > RotacaoReferenciaHz)
            {
                #region Rotação Maior que 600 RPM
                if (frequencia == 0 && BPFI == 0)
                //{
                //    vel_stopfreq = 2000;
                //    vel_lines = NumeroDeLinhasAproximado(1600);
                //} PC Changed
                {
                    vel_stopfreq = RPMHz * 55;
                    vel_stopfreq = CalculoFrequenciaFinal(vel_stopfreq);
                    vel_lines = NumeroDeLinhasAproximado((int)Math.Floor(Math.Ceiling((((float)1 / (float)RPMHz) * 16) * vel_stopfreq)));
                }
                
                else
                {
                    float maiorFrequenciaHarmonicas = ((frequencia * RPMHz) * (float)3.5); // PC Mudar aqui para usuario definir frequencias a serem utilizadas
                    float maiorFrequenciaRolamento = (BPFI * RPMHz) * (float)5.5;

                    if (maiorFrequenciaRolamento > maiorFrequenciaHarmonicas)
                        frequenciaFinal = (float)Math.Ceiling(Convert.ToDecimal(maiorFrequenciaRolamento));
                    else
                        frequenciaFinal = (float)Math.Ceiling(Convert.ToDecimal(maiorFrequenciaHarmonicas));


                    vel_stopfreq = CalculoFrequenciaFinal(frequenciaFinal);

                    //if (vel_stopfreq > 2000)
                    //{
                    //    vel_stopfreq = 2000;
                    //    vel_lines = 3200;
                    //}
                    //else
                    //    vel_lines = 1600;

                    vel_lines = NumeroDeLinhasAproximado((int)Math.Floor(Math.Ceiling((((float)1 / (float)RPMHz) * 16) * vel_stopfreq))); 

                    //vel_lines = 1600; // NumeroDeLinhasAproximado((int)Math.Floor(Math.Ceiling((((float)1 / (float)RPMHz) * 16) * vel_stopfreq)));
                }
                #endregion
            }
            else if (RPMHz >= RotacaoReferenciaIntermediariaHz && RPMHz <= RotacaoReferenciaHz)
            {
                #region Rotação Maior que 200 RPM e Menor que 600 RPM
                if (frequencia == 0 && BPFI == 0)
                {
                    vel_stopfreq = RPMHz * 105;
                    vel_stopfreq = CalculoFrequenciaFinal(vel_stopfreq);
                    vel_lines = NumeroDeLinhasAproximado((int)Math.Floor(Math.Ceiling((((float)1 / (float)RPMHz) * 8) * vel_stopfreq)));
                }
                else
                {
                    float maiorFrequenciaHarmonicas = ((frequencia * RPMHz) * (float)6.5);
                    float maiorFrequenciaRolamento = (BPFI * RPMHz) * (float)10.5;

                    if (maiorFrequenciaRolamento > maiorFrequenciaHarmonicas)
                        frequenciaFinal = (float)Math.Ceiling(Convert.ToDecimal(maiorFrequenciaRolamento));
                    else
                        frequenciaFinal = (float)Math.Ceiling(Convert.ToDecimal(maiorFrequenciaHarmonicas));


                    vel_stopfreq = (float)Math.Ceiling(Convert.ToDecimal(frequenciaFinal));

                    if (vel_stopfreq > 2000)
                        vel_stopfreq = 2000;

                 //   if (vel_stopfreq < 500)
                 //       vel_stopfreq = 500;

                    vel_lines = NumeroDeLinhasAproximado((int)Math.Floor(Math.Ceiling((((float)1 / (float)RPMHz) * 8) * vel_stopfreq)));

                }
                #endregion
            }
            else if (RPMHz < RotacaoReferenciaIntermediariaHz)
            {
                #region Rotação Menor que 200 RPM
                if (frequencia == 0 && BPFI == 0)
                {
                    vel_stopfreq = RPMHz * 105;
                    vel_stopfreq = CalculoFrequenciaFinal(vel_stopfreq);
                    vel_lines = NumeroDeLinhasAproximado((int)Math.Floor(Math.Ceiling((((float)1 / (float)RPMHz) * 8) * vel_stopfreq)));
                }

                else
                {
                    float maiorFrequenciaHarmonicas = ((frequencia * RPMHz) * (float)6.5);
                    float maiorFrequenciaRolamento = (BPFI * RPMHz) * (float)10.5;

                    if (maiorFrequenciaRolamento > maiorFrequenciaHarmonicas)
                        frequenciaFinal = (float)Math.Ceiling(Convert.ToDecimal(maiorFrequenciaRolamento));
                    else
                        frequenciaFinal = (float)Math.Ceiling(Convert.ToDecimal(maiorFrequenciaHarmonicas));


                    vel_stopfreq = (float)Math.Ceiling(Convert.ToDecimal(frequenciaFinal));


                    if (vel_stopfreq > 2000)
                        vel_stopfreq = 2000;

                 //   if (vel_stopfreq < 250)
                 //       vel_stopfreq = 250;


                    vel_lines = NumeroDeLinhasAproximado((int)Math.Floor(Math.Ceiling((((float)1 / (float)RPMHz) * 8) * vel_stopfreq)));

                }
                #endregion
            }

            vel_lowcutoff = CalculoLowCutoff(RPMHz, vel_stopfreq, vel_lines); // PC changed - vel_lowcutoff = (float)Math.Ceiling(Convert.ToDecimal((vel_stopfreq / vel_lines) * 5));

        }


        private int FMaxAjust(float FMax, float Number)
        {
            var T1 = Math.Floor(Math.Log(FMax) / Math.Log(10));
            var T2 = Math.Log(Number) / Math.Log(10);
            var T3 = Math.Round(Math.Pow(10, (T1 + T2)), 2);

            //var T4 = Math.Round(Math.Pow(10, (T1 + T2)), 2);


            return (int)Math.Ceiling(T3);
        }
        private int CalculoFrequenciaFinal(float env_stopfreq)
        {

            int env_stopfreq_tmp = 0;

            double test = Math.Log(env_stopfreq);

            /// Se (FreqFinal / Log(10)) - (FreqFinal_Arredondada_Baixo / Log(10)) > 

            if ((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10))) > (Math.Log(0.5) / Math.Log(10))
                && (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10))))) <= (Math.Log(1) / Math.Log(10)))
            {
                env_stopfreq_tmp = FMaxAjust(env_stopfreq, (float)1);
            }
            else if (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10)))) > (Math.Log(1) / Math.Log(10)) && (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10))))) <= (Math.Log(1.2) / Math.Log(10)))
                env_stopfreq_tmp = FMaxAjust(env_stopfreq, (float)1.2);
            else if (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10)))) > (Math.Log(1.2) / Math.Log(10)) && (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10))))) <= (Math.Log(1.5) / Math.Log(10)))
                env_stopfreq_tmp = FMaxAjust(env_stopfreq, (float)1.5);
            else if (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10)))) > (Math.Log(1.5) / Math.Log(10)) && (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10))))) <= (Math.Log(2) / Math.Log(10)))
                env_stopfreq_tmp = FMaxAjust(env_stopfreq, (float)2);
            else if (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10)))) > (Math.Log(2) / Math.Log(10)) && (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10))))) <= (Math.Log(2.5) / Math.Log(10)))
                env_stopfreq_tmp = FMaxAjust(env_stopfreq, (float)2.5);
            else if (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10)))) > (Math.Log(2.5) / Math.Log(10)) && (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10))))) <= (Math.Log(3) / Math.Log(10)))
                env_stopfreq_tmp = FMaxAjust(env_stopfreq, (float)3);
            else if (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10)))) > (Math.Log(3) / Math.Log(10)) && (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10))))) <= (Math.Log(3.5) / Math.Log(10)))
                env_stopfreq_tmp = FMaxAjust(env_stopfreq, (float)3.5);
            else if (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10)))) > (Math.Log(3.5) / Math.Log(10)) && (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10))))) <= (Math.Log(4) / Math.Log(10)))
                env_stopfreq_tmp = FMaxAjust(env_stopfreq, (float)4);
            else if (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10)))) > (Math.Log(4) / Math.Log(10)) && (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10))))) <= (Math.Log(4.5) / Math.Log(10)))
                env_stopfreq_tmp = FMaxAjust(env_stopfreq, (float)4.5);
            else if (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10)))) > (Math.Log(4.5) / Math.Log(10)) && (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10))))) <= (Math.Log(5) / Math.Log(10)))
                env_stopfreq_tmp = FMaxAjust(env_stopfreq, (float)5);
            else if (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10)))) > (Math.Log(5) / Math.Log(10)) && (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10))))) <= (Math.Log(5.5) / Math.Log(10)))
                env_stopfreq_tmp = FMaxAjust(env_stopfreq, (float)5.5);
            else if (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10)))) > (Math.Log(5.5) / Math.Log(10)) && (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10))))) <= (Math.Log(6) / Math.Log(10)))
                env_stopfreq_tmp = FMaxAjust(env_stopfreq, (float)6);
            else if (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10)))) > (Math.Log(6) / Math.Log(10)) && (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10))))) <= (Math.Log(6.5) / Math.Log(10)))
                env_stopfreq_tmp = FMaxAjust(env_stopfreq, (float)6.5);
            else if (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10)))) > (Math.Log(6.5) / Math.Log(10)) && (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10))))) <= (Math.Log(7) / Math.Log(10)))
                env_stopfreq_tmp = FMaxAjust(env_stopfreq, (float)7);
            else if (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10)))) > (Math.Log(7) / Math.Log(10)) && (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10))))) <= (Math.Log(7.5) / Math.Log(10)))
                env_stopfreq_tmp = FMaxAjust(env_stopfreq, (float)7.5);
            else if (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10)))) > (Math.Log(7.5) / Math.Log(10)) && (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10))))) <= (Math.Log(8) / Math.Log(10)))
                env_stopfreq_tmp = FMaxAjust(env_stopfreq, (float)8);
            else if (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10)))) > (Math.Log(8) / Math.Log(10)) && (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10))))) <= (Math.Log(8.5) / Math.Log(10)))
                env_stopfreq_tmp = FMaxAjust(env_stopfreq, (float)8.5);
            else if (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10)))) > (Math.Log(8.5) / Math.Log(10)) && (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10))))) <= (Math.Log(9) / Math.Log(10)))
                env_stopfreq_tmp = FMaxAjust(env_stopfreq, (float)9);
            else if (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10)))) > (Math.Log(9) / Math.Log(10)) && (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10))))) <= (Math.Log(9.5) / Math.Log(10)))
                env_stopfreq_tmp = FMaxAjust(env_stopfreq, (float)9.5);
            else if (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10)))) > (Math.Log(9.5) / Math.Log(10)) && (((Math.Log(env_stopfreq) / Math.Log(10)) - ((int)Math.Floor(Math.Log(env_stopfreq) / Math.Log(10))))) <= (Math.Log(10) / Math.Log(10)))
                env_stopfreq_tmp = FMaxAjust(env_stopfreq, (float)10);
            else
                env_stopfreq_tmp = int.MaxValue;


            return (env_stopfreq_tmp == int.MaxValue ? (int)env_stopfreq : env_stopfreq_tmp);
        }
        private float[] LCO_Array = {
            (float)0.0078
            ,(float)0.0156
            ,(float)0.0313
            ,(float)0.0625
            ,(float)0.125
            ,(float)0.25
            ,(float)0.5
            ,1
            ,2
            ,4
            ,6
            ,8
            ,10
        };


        private float CalculoLowCutoff(float Velocidade, float FMax, uint NumeroLinhas)
        {
            float _return = 0;
            float Calculo1 = (float)(Velocidade * 0.25);
            float Calculo2 = (float)(FMax / NumeroLinhas) * 5;
            float temp;

            if (Velocidade < 2 / 60)
            {
                temp = (float)(FMax / NumeroLinhas * 5);
                _return = ClosestTo(LCO_Array, temp);
            }
            else if (FMax == 0 && (Calculo1 > Calculo2))
            {
                temp = (float)(Velocidade * 0.25 * 1.25);
                _return = ClosestTo(LCO_Array, temp);
            }
            else if (FMax == 0 && (Calculo1 < Calculo2))
            {
                temp = (float)(Velocidade * 5 * 1.25);
                _return = ClosestTo(LCO_Array, temp);
            }
            else if (FMax > 0 && (Calculo1 > Calculo2))
            {
                temp = (float)(Velocidade * 0.25 * 1.25);
                _return = ClosestTo(LCO_Array, temp);
            }
            else if (Calculo1 < Calculo2)
            {
                temp = (float)(FMax / NumeroLinhas * 5 * 1.25);
                _return = ClosestTo(LCO_Array, temp);
            }

            return _return;
        }

        public static float ClosestTo(IEnumerable<float> collection, float target)
        {
            // NB Method will return int.MaxValue for a sequence containing no elements.
            // Apply any defensive coding here as necessary.
            var closest = float.MaxValue;
            var minDifference = float.MaxValue;
            float difference = 0; //PC changed 
            foreach (var element in collection)
            {
                difference = (Math.Abs(element - target)); //PC changed - var difference = Math.Abs((long)element - target)
                if (minDifference > difference)
                {
                    minDifference = difference; // PC changed - minDifference = (int)difference;
                    closest = element;
                }
            }

            return closest;
        }
        public int tCount { get; set; }
        private void CallRecursive(TreeView treeView)
        {
            // Print each node recursively.
            TreeNodeCollection nodes = treeView.Nodes;
            foreach (TreeNode n in nodes)
                TreeView_Check_Count(n);
        }
        private void TreeView_Check_Count(TreeNode treeNode)
        {
            if (treeNode.Checked == true && treeNode.ImageIndex == 3)

                tCount++;

            foreach (TreeNode tn in treeNode.Nodes)
                TreeView_Check_Count(tn);
        }
        private float Env_Range(float BPFI, float RPM)
        {
            float range = 0;

            if (BPFI == 0)
                range = ((float)150 * (RPM / (float)60));
            else
                range = ((BPFI * (float)5.5) * (RPM / (float)60));

            return range;
        }
        private float Env_LowCutOff(float Freq_Max, uint Lines, float RPM)
        {
            float LowCutOff = 0;

            LowCutOff = (Freq_Max / Lines) * 5;
            
            return LowCutOff;
        }
        private uint NumeroDeLinhasAproximado(int NumeroLinhas)
        {
            float value_temp = NumeroLinhas;
            uint lines = 0;
            if (value_temp <= 100) lines = 100;
            if (value_temp > 100 && value_temp <= 200) lines = 200;
            if (value_temp > 200 && value_temp <= 400) lines = 400;
            if (value_temp > 400 && value_temp <= 800) lines = 800;
            if (value_temp > 800 && value_temp <= 1600) lines = 1600;
            if (value_temp > 1600 && value_temp <= 3200) lines = 3200;
            if (value_temp > 3200 && value_temp <= 6400) lines = 6400;
            if (value_temp > 6400 && value_temp <= 12800) lines = 12800;
            if (value_temp > 12800 && value_temp <= 25600) lines = 25600;
            if (value_temp >= 25600) lines = 25600;

            return lines;

        }
        private uint Env_Lines(float FTF, float BPFI, float RPM)
        {

            float value_temp = 0;
            uint lines = 0;

            if (FTF != 0)
            {
                // Retirada combinada com o Mario Eduardo na revisão 9 do Procedimento. 09/06/2014
                //if (FTF < (float)0.2)
                //    value_temp = (float)40 * (BPFI / FTF);
                //else
                value_temp = (float)130 * BPFI;

                if (value_temp <= 100) lines = 100;
                if (value_temp > 100 && value_temp <= 200) lines = 200;
                if (value_temp > 200 && value_temp <= 400) lines = 400;
                if (value_temp > 400 && value_temp <= 800) lines = 800;
                if (value_temp > 800 && value_temp <= 1600) lines = 1600;
                if (value_temp > 1600 && value_temp <= 3200) lines = 3200;
                if (value_temp > 3200 && value_temp <= 6400) lines = 6400;
                if (value_temp > 6400 && value_temp <= 12800) lines = 12800;
                if (value_temp > 12800 && value_temp <= 25600) lines = 25600;
                if (value_temp >= 25600) lines = 25600;
            }
            else
            {
                if (RPM > (float)2000)
                    lines = 3200;
                else
                    lines = 1600;
            }

            return lines;
        }
        private float Vel_LowCutOff(float BPFI, uint Lines, float RPM)
        {
            float LowCutOff = 0;

            if (BPFI == 0)
                LowCutOff = ((((float)55 * RPM) / (float)60) / Lines) * 5;
            else
                LowCutOff = ((float)1000 / Lines) * 5;

            return LowCutOff;
        }
        private int Vel_Lines(int RPM)
        {

            float value_temp = 0;
            int lines = 0;

            if (RPM != 0)
            {
                value_temp = (float)0.2 * RPM;

                if (value_temp >= 0 && value_temp <= 100) lines = 100;
                if (value_temp > 100 && value_temp <= 200) lines = 200;
                if (value_temp > 200 && value_temp <= 400) lines = 400;
                if (value_temp > 400 && value_temp <= 800) lines = 800;
                if (value_temp > 800 && value_temp <= 1600) lines = 1600;
                if (value_temp > 1600 && value_temp <= 3200) lines = 3200;
                if (value_temp > 3200 && value_temp <= 6400) lines = 6400;
                if (value_temp > 6400 && value_temp <= 12800) lines = 12800;
                if (value_temp > 12800 && value_temp <= 25600) lines = 25600;
                if (value_temp >= 25600) lines = 25600;
            }
            else
            {
                lines = 1600;
            }

            return lines;
        }
        
        private DataTable getOtherFrequencies(AnalystConnection cn, uint pt)
        {
            DataTable dt_Bear = new DataTable();
            dt_Bear = cn.DataTable("select RPM from GENERALFREQS WHERE GFID IN (select GEOMETRYID from FREQENTRIES WHERE FSID IN (select FSID from FREQASSIGN WHERE ELEMENTID=" + pt + ") AND GEOMETRYTABLE='SKFCM_ASGI_OtherFreqType') AND TYPE=2 order by RPM DESC");
            return dt_Bear;
        }
        private float Get_GMF(AnalystConnection cn, uint pt)
        {
            
            float GMF = 0;
            float RPM;

            DataTable dt_Bear = new DataTable();

            RPM = GetPointRPM(cn, pt);
            try
            {
                dt_Bear = cn.DataTable("select RPM from GENERALFREQS WHERE GFID IN (select GEOMETRYID from FREQENTRIES WHERE FSID IN (select FSID from FREQASSIGN WHERE ELEMENTID=" + pt + ") AND GEOMETRYTABLE='SKFCM_ASGI_OtherFreqType')");
                //select RPM from FREQENTRIES WHERE FSID IN (select FSID from FREQASSIGN WHERE ELEMENTID=" + pt + ") AND GEOMETRYTABLE='SKFCM_ASGI_BearingFreqType'");
                if (dt_Bear.Rows.Count > 0)
                {
                    float varia = 0;
                    float tGMF = 0;

                    for (int i = 0; i < dt_Bear.Rows.Count; i++)
                    {
                        float Bear_Value = float.Parse(dt_Bear.Rows[i][0].ToString());

                        varia = (Bear_Value * RPM) / 60;
                        if (varia >= tGMF)
                        {
                            tGMF = varia;
                        }
                    }
                    GMF = (float)3.2 * tGMF;
                }
            }
            catch (Exception ex)
            {
                Log ("Error Get_GMF Pt: " + pt.ToString() + " - Message: " + ex.Message.ToString()); 
            }

            return GMF;
        }
        private float Get_FTF(AnalystConnection cn, uint pt)
        {
            float FTF = 0;

            try 
            {
                DataTable dt_Bear = new DataTable();
                dt_Bear = cn.DataTable("select NAME from FREQENTRIES WHERE FSID IN (select FSID from FREQASSIGN WHERE ELEMENTID=" + pt + ") AND GEOMETRYTABLE='SKFCM_ASGI_BearingFreqType'");
                if (dt_Bear.Rows.Count > 0)
                {

                    float varia;
                    string name;

                    for (int i = 0; i < dt_Bear.Rows.Count; i++)
                    {
                        name = dt_Bear.Rows[i][0].ToString();
                        string[] bear = name.Split('(');
                        if (bear.Length > 1)
                        {
                            string[] manu = bear[1].Split(')');

                            varia = cn.SQLtoFloat("select FTF from BEARING where NAME='" + bear[0].Trim() + "' and MANUFACTURE = '" + manu[0].Replace('(', ' ').Trim() + "'");
                            if (varia > FTF)
                            {
                                FTF = varia;
                            }
                        }
                        else
                        {
                            FTF = 0;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log("Error Get_FTF Pt: " + pt.ToString() + " - Message: " + ex.Message.ToString()); 
            }
            return FTF;
        }
        private float Get_Great_BPFI(AnalystConnection cn, uint pt)
        {
            float BPFI = 0;

            try
            {
                DataTable dt_Bear = new DataTable();
                dt_Bear = cn.DataTable("select NAME from FREQENTRIES WHERE FSID IN (select FSID from FREQASSIGN WHERE ELEMENTID=" + pt + ") AND GEOMETRYTABLE='SKFCM_ASGI_BearingFreqType'");
                if (dt_Bear.Rows.Count > 0)
                {

                    float varia;
                    string name;

                    for (int i = 0; i < dt_Bear.Rows.Count; i++)
                    {
                        name = dt_Bear.Rows[i][0].ToString();
                        string[] bear = name.Split('(');

                        if (bear.Length > 1)
                        {
                            string[] manu = bear[1].Split(')');

                            varia = cn.SQLtoFloat("select BPFI from BEARING where NAME='" + bear[0].Trim() + "' and MANUFACTURE = '" + manu[0].Replace('(', ' ').Trim() + "'");
                            if (varia > BPFI)
                            {
                                BPFI = varia;
                            }
                        }
                        else
                            BPFI = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                Log("Error Get_Great_BPFI Pt: " + pt.ToString() + " - Message: " + ex.Message.ToString());
            }
            return BPFI;
        }

        private void Reply_Envelope(AnalystConnection cn, uint pt, uint Lines, uint Averages, float StartFreq, float StopFreq, float LowCuttOffFreq, string SaveData, string Detection)
        {
            try
            {
                if (Lines > 0)
                    cn.SQLUpdate("Point", "ValueString", Lines, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Lines") + " and ELEMENTID=" + pt);

                cn.SQLUpdate("Point", "ValueString", Averages, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Averages") + " and ELEMENTID=" + pt);
                cn.SQLUpdate("Point", "ValueString", StartFreq, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Start_Freq") + " and ELEMENTID=" + pt);
                cn.SQLUpdate("Point", "ValueString", LowCuttOffFreq, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Low_Freq_Cutoff") + " and ELEMENTID=" + pt);
                cn.SQLUpdate("Point", "ValueString", SaveData, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Save_Data") + " and ELEMENTID=" + pt);
                cn.SQLUpdate("Point", "ValueString", Detection, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Detection") + " and ELEMENTID=" + pt);
                cn.SQLUpdate("Point", "ValueString", StopFreq, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_End_Freq") + " and ELEMENTID=" + pt);
            }
            catch (Exception ex)
            {
                Log("Error Reply_Envelope Pt: " + pt.ToString() + " - Message: " + ex.Message.ToString());
            }

        }
        private void Reply_Aceleration(AnalystConnection cn, uint pt, uint Lines, uint Averages, float StartFreq, float StopFreq, float LowCuttOffFreq, string SaveData, string Detection)
        {
            try
            {
                if (Lines > 0)
                    cn.SQLUpdate("Point", "ValueString", Lines, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Lines") + " and ELEMENTID=" + pt);

                cn.SQLUpdate("Point", "ValueString", Averages, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Averages") + " and ELEMENTID=" + pt);
                cn.SQLUpdate("Point", "ValueString", StartFreq, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Start_Freq") + " and ELEMENTID=" + pt);
                cn.SQLUpdate("Point", "ValueString", LowCuttOffFreq, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Low_Freq_Cutoff") + " and ELEMENTID=" + pt);
                cn.SQLUpdate("Point", "ValueString", SaveData, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Save_Data") + " and ELEMENTID=" + pt);
                cn.SQLUpdate("Point", "ValueString", Detection, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Detection") + " and ELEMENTID=" + pt);

                if (StopFreq > 0)
                    cn.SQLUpdate("Point", "ValueString", StopFreq, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_End_Freq") + " and ELEMENTID=" + pt);
            }
            catch (Exception ex)
            {
                Log("Error Reply_Aceleration Pt: " + pt.ToString() + " - Message: " + ex.Message.ToString());
            }
        }
        private void Reply_Velocity(AnalystConnection cn, uint pt, uint Lines, uint Averages, float StartFreq, float StopFreq, float LowCuttOffFreq, string SaveData, string Detection)
        {
            try
            {
                if (Lines > 0)
                    cn.SQLUpdate("Point", "ValueString", Lines, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Lines") + " and ELEMENTID=" + pt);

                cn.SQLUpdate("Point", "ValueString", Averages, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Averages") + " and ELEMENTID=" + pt);
                cn.SQLUpdate("Point", "ValueString", StartFreq, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Start_Freq") + " and ELEMENTID=" + pt);
                cn.SQLUpdate("Point", "ValueString", LowCuttOffFreq, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Low_Freq_Cutoff") + " and ELEMENTID=" + pt);
                cn.SQLUpdate("Point", "ValueString", SaveData, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Save_Data") + " and ELEMENTID=" + pt);
                cn.SQLUpdate("Point", "ValueString", Detection, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Detection") + " and ELEMENTID=" + pt);

                if (StopFreq != 0)
                    cn.SQLUpdate("Point", "ValueString", StopFreq, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_End_Freq") + " and ELEMENTID=" + pt);
            }
            catch (Exception ex)
            {
                Log("Error Reply_Velocity Pt: " + pt.ToString() + " - Message: " + ex.Message.ToString());
            }
        }
        private void Reply_Velocity_ZOOM(AnalystConnection cn, uint pt, uint Lines, uint Averages, float StopFreq, float LowCuttOffFreq, string SaveData, string Detection)
        {
            StringBuilder SQL = new StringBuilder();

            try
            {
                if (Lines > 0)
                    cn.SQLUpdate("Point", "ValueString", Lines, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Lines") + " and ELEMENTID=" + pt);

                cn.SQLUpdate("Point", "ValueString", Averages, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Averages") + " and ELEMENTID=" + pt);
                cn.SQLUpdate("Point", "ValueString", StopFreq, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_End_Freq") + " and ELEMENTID=" + pt);
                cn.SQLUpdate("Point", "ValueString", LowCuttOffFreq, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Low_Freq_Cutoff") + " and ELEMENTID=" + pt);
                cn.SQLUpdate("Point", "ValueString", SaveData, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Save_Data") + " and ELEMENTID=" + pt);
                cn.SQLUpdate("Point", "ValueString", Detection, "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Detection") + " and ELEMENTID=" + pt);
            }
            catch (Exception ex)
            {
                Log("Error Reply_Velocity Pt: " + pt.ToString() + " - Message: " + ex.Message.ToString());
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            button_Progress_Run.Enabled = false;


            treeView_AnStandards.Nodes.Clear();
            if (chkWorkSpace.Checked == false)
            {
                Thread LoadIt = new Thread(() => { Tree_Analyst(_analystDB, treeView_AnStandards); });
                LoadIt.Start();
            }
            else
            {
                if (txtWS.Text != "")
                {
                    Thread LoadIt = new Thread(() => { Tree_Analyst_WorkSpace(_analystDB, treeView_AnStandards, txtWS.Text); });
                    LoadIt.Start();
                }
                else
                {
                    MessageBox.Show("You MUST select a Workspace");
                }
            }
        }

        private void btAES_Load_Click(object sender, EventArgs e)
        {
            StringBuilder SQL = new StringBuilder();

            string TBLSETID = "";
            string Acc = Registration.RegistrationId(_analystDB,"SKFCM_ASPT_Acc").ToString();
            string Vel = Registration.RegistrationId(_analystDB, "SKFCM_ASPT_AccToVel").ToString();
            string Env = Registration.RegistrationId(_analystDB, "SKFCM_ASPT_AccEnvelope").ToString();
            string DAD = Registration.RegistrationId(_analystDB, "SKFCM_ASDD_MicrologDAD").ToString();
            string CAST = "";

            if (_analystDB.DBType == DBType.MSSQL)
            {
                CAST = " and CAST(PT1.VALUESTRING AS numeric(17,2)) > " + float.Parse(Acel_GTRPM.Text) / (float)60;
            }
            else
            {
                CAST = " and TO_NUMBER(PT1.VALUESTRING) > " + float.Parse(Acel_GTRPM.Text) / (float)60;
            }

            SQL.Clear();

            if (chkCSS_Workspace.Checked == true)
            {
                SQL.Append("   select ");
                SQL.Append("   	TB.TBLSETNAME as Hierarchy_Name ");
                SQL.Append("   	, TE3.NAME as Parent_Node");
                SQL.Append("   	, TE2.NAME as Asset_Name ");
                SQL.Append("   	, TE1.REFERENCEID as Point_Id ");
                SQL.Append("   	, TE1.NAME as Point_Name ");
                SQL.Append("   	, PT1.VALUESTRING as Speed_Hz ");
                SQL.Append("   from  ");
                SQL.Append("   	treeelem TE1 ");
                SQL.Append("   	, TABLESET TB ");
                SQL.Append("   	, TREEELEM TE2 ");
                SQL.Append("   	, TREEELEM TE3 ");
                SQL.Append("   	, POINT PT ");
                SQL.Append("   	, POINT PT1 ");
                SQL.Append("   	, POINT PT2 ");
                SQL.Append("   	, REGISTRATION RG ");
                SQL.Append("   	, REGISTRATION RG1 ");
                SQL.Append("   where  ");
                SQL.Append("   	TE1.CONTAINERTYPE=4 ");
                if (txtAES_TBLSETID.Text != "")
                {
                    TBLSETID = " AND TE1.TBLSETID= " + txtAES_TBLSETID.Text + " ";
                    SQL.Append(TBLSETID);
                }
                SQL.Append("   	and TE1.parentid!=2147000000 ");
                SQL.Append("   	and TE1.HIERARCHYTYPE=3 ");
                if (txtAES_WS.Text != "")
                    SQL.Append("  and TE1.HIERARCHYID in (SELECT TREEELEMID FROM TREEELEM WHERE HIERARCHYTYPE=3 and Parentid!=2147000000 and NAME LIKE '" + txtAES_WS.Text + "') ");

                SQL.Append("   	and TE1.ELEMENTENABLE=1 ");
                SQL.Append("   	AND TE1.TBLSETID = TB.TBLSETID ");
                SQL.Append("   	AND TE2.TREEELEMID=TE1.PARENTID ");
                SQL.Append("   	AND TE3.TREEELEMID = TE2.PARENTID ");
                SQL.Append("  	AND PT.ELEMENTID = TE1.REFERENCEID ");
                SQL.Append("   	AND PT.VALUESTRING IN ('" + Acc + "', '" + Vel + "', '" + Env + "') ");
                SQL.Append("   	AND PT1.ELEMENTID = PT.ELEMENTID ");
                SQL.Append("   	AND RG.SIGNATURE = 'SKFCM_ASPF_Speed' ");
                SQL.Append("   	AND PT1.FIELDID	= RG.REGISTRATIONID ");
                SQL.Append("   	AND PT2.VALUESTRING ='" + DAD + "' ");
                SQL.Append("   	AND PT2.ELEMENTID = PT1.ELEMENTID ");
                SQL.Append("   	AND RG1.SIGNATURE = 'SKFCM_ASPF_Dad_Id' ");
                SQL.Append("   	AND PT2.FIELDID	= RG1.REGISTRATIONID ");
               // SQL.Append(CAST);
                SQL.Append("   order by  ");
                SQL.Append("   	tb.TBLSETID ");
                SQL.Append("   	, te2.NAME ");
                SQL.Append("   	, te1.SLOTNUMBER ");
                SQL.Append("   	, te1.NAME ");
            }
            else
            {
                SQL.Append("   select ");
                SQL.Append("   	TB.TBLSETNAME as Hierarchy_Name ");
                SQL.Append("   	, TE3.NAME as Parent_Node");
                SQL.Append("   	, TE2.NAME as Asset_Name ");
                SQL.Append("   	, TE1.TREEELEMID as Point_Id ");
                SQL.Append("   	, TE1.NAME as Point_Name ");
                SQL.Append("   	, PT1.VALUESTRING as Speed_Hz ");
                SQL.Append("   from  ");
                SQL.Append("   	treeelem TE1 ");
                SQL.Append("   	, TABLESET TB ");
                SQL.Append("   	, TREEELEM TE2 ");
                SQL.Append("   	, TREEELEM TE3 ");
                SQL.Append("   	, POINT PT ");
                SQL.Append("   	, POINT PT1 ");
                SQL.Append("   	, POINT PT2 ");
                SQL.Append("   	, REGISTRATION RG ");
                SQL.Append("   	, REGISTRATION RG1 ");
                SQL.Append("   where  ");
                SQL.Append("   	TE1.CONTAINERTYPE=4 ");
                if (txtAES_TBLSETID.Text != "")
                {
                    TBLSETID = " AND TE1.TBLSETID= " + txtAES_TBLSETID.Text + " ";
                    SQL.Append(TBLSETID);
                }
                SQL.Append("   	and TE1.parentid!=2147000000 ");
                SQL.Append("   	and TE1.HIERARCHYTYPE=1 ");
                SQL.Append("   	and TE1.ELEMENTENABLE=1 ");
                SQL.Append("   	AND TE1.TBLSETID = TB.TBLSETID ");
                SQL.Append("   	AND TE2.TREEELEMID=TE1.PARENTID ");
                SQL.Append("   	AND TE3.TREEELEMID = TE2.PARENTID ");
                SQL.Append("  	AND PT.ELEMENTID = TE1.TREEELEMID ");
                SQL.Append("   	AND PT.VALUESTRING IN ('" + Acc + "', '" + Vel + "', '" + Env + "') ");
                SQL.Append("   	AND PT1.ELEMENTID = PT.ELEMENTID ");
                SQL.Append("   	AND RG.SIGNATURE = 'SKFCM_ASPF_Speed' ");
                SQL.Append("   	AND PT1.FIELDID	= RG.REGISTRATIONID ");
                SQL.Append("   	AND PT2.VALUESTRING ='" + DAD + "' ");
                SQL.Append("   	AND PT2.ELEMENTID = PT1.ELEMENTID ");
                SQL.Append("   	AND RG1.SIGNATURE = 'SKFCM_ASPF_Dad_Id' ");
                SQL.Append("   	AND PT2.FIELDID	= RG1.REGISTRATIONID ");
                //SQL.Append(CAST);
                SQL.Append("   order by  ");
                SQL.Append("   	tb.TBLSETID ");
                SQL.Append("   	, te2.NAME ");
                SQL.Append("   	, te1.SLOTNUMBER ");
                SQL.Append("   	, te1.NAME ");
            }
            Log(SQL.ToString());

            lblSSP.Invoke(new MethodInvoker(delegate
              {
                  lblSSP.Text = "Querying...";
              }));
            DataTable SSP_DT = _analystDB.DataTable(SQL.ToString());
            Log("Rows Point: " + SSP_DT.Rows.Count.ToString());

            lblSSP.Invoke(new MethodInvoker(delegate
            {
                lblSSP.Text = "0/" + SSP_DT.Rows.Count.ToString();
                pbSSP.Maximum = SSP_DT.Rows.Count;
                pbSSP.Value = 0;
            }));

            Log("Iniciando Thread");
            Thread LoadIt = new Thread(() => { SSP_Ajust(SSP_DT, dgSSP); });
            LoadIt.Start();
            
        }
        private void SSP_Ajust(DataTable dt, DataGridView dgSSP)
        {
            lblSSP.Invoke(new MethodInvoker(delegate { pbSSP.Value = 0; }));

            dt.Columns.Add(new DataColumn("Spectrum_Speed_Hz", typeof(string))); // PC before:"SpectrumSpeed" 

            int i=0;
            foreach(DataRow dr in dt.Rows)
            {

                //dr["RPM"] = float.Parse(dr["RPM"].ToString()) ; //change the name

                SKF.RS.STB.Analyst.Point Pt = new SKF.RS.STB.Analyst.Point(_analystDB, Convert.ToUInt32(dr["Point_Id"]));
                Measurement Meas = new Measurement(Pt);

                try
                {
                    if (Meas.MeasId > 0)
                        dr["Spectrum_Speed_Hz"] = Convert.ToDecimal(Meas.ReadingFFT.Speed);
                    else
                        dr["Spectrum_Speed_Hz"] = 0;
                }
                catch (Exception ex)
                {
                    dr["Spectrum_Speed_Hz"] = 0;
                    Log("Erro Get FFT: " + ex.Message.ToString());
                }
                i++;
                lblSSP.Invoke(new MethodInvoker(delegate { pbSSP.Value = i; lblSSP.Text = i + "/" + dt.Rows.Count.ToString(); }));
            }

            dgSSP.Invoke(new MethodInvoker(delegate {
                    dgSSP.DataSource = null;
                    dgSSP.Refresh(); 
                    dgSSP.DataSource = dt;
                    ColumnsResize(dgSSP);
                    btSSP_Transfer.Enabled = true;
            }));

        }
        private void ColumnsResize(DataGridView grd)
        {
            for (int w = 0; w < grd.Columns.Count; w++)
            {
                grd.Columns[w].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
            }

            for (int i = 0; i < grd.Columns.Count; i++)
            {
                int colw = grd.Columns[i].Width;
                grd.Columns[i].AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                grd.Columns[i].Width = colw;
            }
        }
        private void btSSP_Transfer_Click(object sender, EventArgs e)
        {
            lblSSP.Text = "0/" + dgSSP.Rows.Count.ToString();
            pbSSP.Maximum = dgSSP.Rows.Count;
            pbSSP.Value = 0;

            DataTable DT = ((DataTable)dgSSP.DataSource).Copy();;

            Thread LoadIt = new Thread(() => { SSP_Apply_Ajust(DT); });
            LoadIt.Start();
        }
        private void SSP_Apply_Ajust(DataTable Dg)
        {
            int i=0;
            foreach (DataRow dr in Dg.Rows)
            {
                if (float.Parse(dr["Spectrum_Speed_Hz"].ToString()) > 0)
                {
                    try
                    {
                        SKF.RS.STB.Analyst.Point Pt = new SKF.RS.STB.Analyst.Point(_analystDB, uint.Parse(dr["Point_Id"].ToString()));
                        float Passo1 = (float)Math.Round(float.Parse(dr["Spectrum_Speed_Hz"].ToString()) * 60, 0);
                        float Passo2 = Passo1 / 60;

                        Pt.Speed = Passo2; // (float)Math.Round(float.Parse(dr["SpectrumSpeed"].ToString()), 3);
                    }
                    catch (Exception ex)
                    {
                        Log("Erro setting speed: " + ex.Message.ToString());
                    }
                }

                i++;
                lblSSP.Invoke(new MethodInvoker(delegate { pbSSP.Value = i; lblSSP.Text = i + "/" + Dg.Rows.Count.ToString(); }));
            }
        }
        private void treeView_AnStandards_AfterCheck(object sender, TreeViewEventArgs e)
        {
            if (e.Action != TreeViewAction.Unknown)
            {
                if (e.Node.Nodes.Count > 0)
                {
                    this.CheckAllChildNodes(e.Node, e.Node.Checked);
                }
            }
        }
        private void btNSR_Do_Click(object sender, EventArgs e)
        {
            if (txtNSR_TSID.Text != "")
            {
                StringBuilder SQL = new StringBuilder();


                string Acc = Registration.RegistrationId(_analystDB, "SKFCM_ASPT_Acc").ToString();
                string Vel = Registration.RegistrationId(_analystDB, "SKFCM_ASPT_AccToVel").ToString();
                string Env = Registration.RegistrationId(_analystDB, "SKFCM_ASPT_AccEnvelope").ToString();
                string DAD = Registration.RegistrationId(_analystDB, "SKFCM_ASDD_MicrologDAD").ToString();
                string CAST = "";

                //if (_analystDB.DBType == DBType.MSSQL)
                //{
                //    CAST = " and CAST(PT1.VALUESTRING AS numeric(17,2)) > " + float.Parse(Acel_GTRPM.Text) / (float)60;
                //}
                //else
                //{
                //    CAST = " and TO_NUMBER(PT1.VALUESTRING) > " + float.Parse(Acel_GTRPM.Text) / (float)60;
                //}

                SQL.Clear();
                SQL.Append(" select ");
                SQL.Append(" 	TB.TBLSETID as TBLSETID ");
                SQL.Append(" 	, TB.TBLSETNAME as Hierarchy_Name ");
                SQL.Append(" 	, TE3.NAME as Parent_Node ");
                SQL.Append(" 	, TE2.TREEELEMID as Asset_Id ");
                SQL.Append(" 	, TE2.NAME as Asset_Name ");
                SQL.Append(" 	, TE1.TREEELEMID as Point_Id ");
                SQL.Append(" 	, TE1.NAME as Point_Name ");
                SQL.Append(" from  ");
                SQL.Append(" 	treeelem TE1 ");
                SQL.Append(" 	, TABLESET TB ");
                SQL.Append(" 	, TREEELEM TE2 ");
                SQL.Append("    , TREEELEM TE3 ");
                SQL.Append(" 	, POINT PT ");
                SQL.Append(" 	, POINT PT2 ");
                SQL.Append(" 	, REGISTRATION RG1 ");
                SQL.Append(" where  ");
                SQL.Append(" 	TE1.CONTAINERTYPE=4 ");
                SQL.Append("    AND TE1.TBLSETID= " + txtNSR_TSID.Text);
                SQL.Append(" 	and TE1.parentid!=2147000000 ");
                SQL.Append(" 	and TE1.HIERARCHYTYPE=1 ");
                SQL.Append(" 	and TE1.ELEMENTENABLE=1 ");
                SQL.Append(" 	AND TE1.TBLSETID = TB.TBLSETID ");
                SQL.Append(" 	AND TE2.TREEELEMID=TE1.PARENTID ");
                SQL.Append(" 	AND TE3.TREEELEMID = TE2.PARENTID ");
                SQL.Append("	AND PT.ELEMENTID = TE1.TREEELEMID ");
                SQL.Append("   	AND PT.VALUESTRING IN ('" + Acc + "', '" + Vel + "', '" + Env + "') ");
                SQL.Append("   	AND PT2.VALUESTRING ='" + DAD + "' ");
                SQL.Append(" 	AND PT2.ELEMENTID = PT.ELEMENTID ");
                SQL.Append(" 	AND RG1.SIGNATURE = 'SKFCM_ASPF_Dad_Id' ");
                SQL.Append(" 	AND PT2.FIELDID	= RG1.REGISTRATIONID ");
                //SQL.Append(CAST);
                SQL.Append(" order by  ");
                SQL.Append(" 	tb.TBLSETID ");
                SQL.Append("	, te2.NAME ");
                SQL.Append(" 	, te1.SLOTNUMBER ");
                SQL.Append(" 	, te1.NAME ");

                Log(SQL.ToString());

                DataTable NSR_DT = _analystDB.DataTable(SQL.ToString());
                Log("Rows Point: " + NSR_DT.Rows.Count.ToString());

                lblNSR.Invoke(new MethodInvoker(delegate
                {
                    lblNSR.Text = "0/" + NSR_DT.Rows.Count.ToString();
                    pbNSR.Maximum = NSR_DT.Rows.Count;
                    pbNSR.Value = 0;
                }));

                Log("Starting Thread");
                Thread LoadIt = new Thread(() => { NSR_Apply_Ajust(_analystDB, NSR_DT, txtNSR_TSID.Text); });
                LoadIt.Start();
            }
            else
            {
                MessageBox.Show("You MUST input a TBLSETID!!!");
            }
        }
        private void NSR_DropAndCreate_WS(AnalystConnection cn, string hierarchyid)
        {
            int WS_Treeelem = cn.SQLtoInt("SELECT TREEELEMID FROM TREEELEM WHERE NAME='Out-Of-Compliance' AND HIERARCHYTYPE=3 AND PARENTID!=2147000000 AND TBLSETID= " + hierarchyid);

            if (WS_Treeelem != 0)
            {
                cn.SQLExec("DELETE FROM TREEELEM WHERE TREEELEMID=" + WS_Treeelem);
                SKF.RS.STB.Analyst.WorkSpace WorkSpace = new SKF.RS.STB.Analyst.WorkSpace(_analystDB, uint.Parse(hierarchyid), "Out-Of-Compliance");
            }
            else
            {
                SKF.RS.STB.Analyst.WorkSpace WorkSpace = new SKF.RS.STB.Analyst.WorkSpace(_analystDB, uint.Parse(hierarchyid), "Out-Of-Compliance");
            }
        }
        private void NSR_Apply_Ajust(AnalystConnection cn,  DataTable Dg, string hierarchyid)
        {
            if (chkWorkSpace.Checked == true)
                NSR_DropAndCreate_WS(cn, hierarchyid);

            string TBLSETID = "";
            if (hierarchyid != "")
                TBLSETID = " AND TBLSETID=" + hierarchyid + " ";

            //if (cn.SQLtoInt("SELECT TREEELEMID FROM TREEELEM WHERE NAME='Out-Of-Compliance' AND HIERARCHYTYPE=3 AND PARENTID!=2147000000 " + TBLSETID) != 0)
            //{
                Dg.Columns.Add("Detection", typeof(string));
                Dg.Columns.Add("New_Detection", typeof(string));
                Dg.Columns.Add("Lines", typeof(string));
                Dg.Columns.Add("New_Lines", typeof(string));
                Dg.Columns.Add("Start_Freq", typeof(string));
                Dg.Columns.Add("New_Start_Freq", typeof(string));
                Dg.Columns.Add("End_Freq", typeof(string));
                Dg.Columns.Add("New_End_Freq", typeof(string));
                Dg.Columns.Add("Low_Freq_CutOff", typeof(string));
                Dg.Columns.Add("New_Low_Freq_CutOff", typeof(string));
                Dg.Columns.Add("Averages", typeof(string));
                Dg.Columns.Add("New_Averages", typeof(string));
                int i = 0;
                foreach (DataRow dr in Dg.Rows)
                {
                    SKF.RS.STB.Analyst.Point Pt = new SKF.RS.STB.Analyst.Point(_analystDB, uint.Parse(dr["Point_Id"].ToString()));
                    dr["Detection"] = cn.SQLtoString("VALUESTRING", "POINT", "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Detection").ToString() + " AND ELEMENTID=" + dr["Point_Id"].ToString());
                    dr["Start_Freq"] = float.Parse(cn.SQLtoString("VALUESTRING", "POINT", "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Start_Freq").ToString() + " AND ELEMENTID=" + dr["Point_Id"].ToString()));
                    dr["End_Freq"] = float.Parse(cn.SQLtoString("VALUESTRING", "POINT", "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_End_Freq").ToString() + " AND ELEMENTID=" + dr["Point_Id"].ToString()));
                    dr["Low_Freq_CutOff"] = float.Parse(cn.SQLtoString("VALUESTRING", "POINT", "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Low_Freq_Cutoff").ToString() + " AND ELEMENTID=" + dr["Point_Id"].ToString()));
                    dr["Lines"] = cn.SQLtoString("VALUESTRING", "POINT", "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Lines").ToString() + " AND ELEMENTID=" + dr["Point_Id"].ToString());
                    dr["Averages"] = cn.SQLtoString("VALUESTRING", "POINT", "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Averages").ToString() + " AND ELEMENTID=" + dr["Point_Id"].ToString());
                    
                    if (Pt.PointType == PointType.Acc)
                    {
                        bool AccCheck = ACC_SSCP_Check(_analystDB, dr["Point_Id"].ToString());
                        if (AccCheck == true)
                        {
                            dr.Delete();
                        }
                        else
                        {
                            #region Acc
                            if (chkWorkSpace.Checked == true)
                            {
                                SKF.RS.STB.Analyst.WorkSpace WorkSpace = new SKF.RS.STB.Analyst.WorkSpace(_analystDB, uint.Parse(dr["TBLSETID"].ToString()), "Out-Of-Compliance");
                                Machine Maq = WorkSpace.FindMachine(uint.Parse(dr["Asset_Id"].ToString()));
                                
                                try
                                {
                                    if (Maq == null)
                                        WorkSpace.AddMachine(new Machine(cn, Convert.ToUInt32(uint.Parse(dr["Asset_Id"].ToString()))));
                                }
                                catch (Exception ex)
                                {
                                    GenericTools.DebugMsg("ProcessMachine(" + dr["Asset_Id"].ToString() + ") error: " + ex.Message);
                                }
                            }
                            
                            dr["New_Detection"] = ((int)Acc.DetectionType).ToString();
                            dr["New_Start_Freq"] = ((int)Acc.StartFreq).ToString();
                            dr["New_End_Freq"] = ((int)Acc.StopFreq).ToString();
                            dr["New_Low_Freq_CutOff"] = ((int)Acc.LowCutOff).ToString();
                            dr["New_Lines"] = ((int)Acc.Lines).ToString();
                            dr["New_Averages"] = ((int)Acc.Averages).ToString();
                            #endregion
                        }
                    }
                    else if (Pt.PointType == PointType.AccEnvelope)
                    {
                        bool EnvCheck = ENV_SSCP_Check(_analystDB, dr["Point_Id"].ToString());
                        if (EnvCheck == true)
                        {
                            dr.Delete();
                        }
                        else
                        {
                            #region Envelope
                            float RPM = GetPointRPM(cn, uint.Parse(dr["Point_Id"].ToString()));
                            float FTF = Get_FTF(cn, uint.Parse(dr["Point_Id"].ToString()));
                            float BPFI = Get_Great_BPFI(cn, uint.Parse(dr["Point_Id"].ToString()));
                            uint env_lines = Env_Lines(FTF, BPFI, RPM);
                            float env_stopfreq = (float)Math.Ceiling(Convert.ToDecimal(Env_Range(BPFI, RPM)));
                            float env_lowcutoff = (float)Math.Ceiling(Convert.ToDecimal(Env_LowCutOff(env_stopfreq, env_lines, RPM)));
                            dr["New_Detection"] = ((int)Env.DetectionType).ToString();
                            dr["New_Start_Freq"] = ((int)Env.StartFreq).ToString();
                            dr["New_End_Freq"] = env_stopfreq;
                            dr["New_Low_Freq_CutOff"] = env_lowcutoff;
                            dr["New_Lines"] = env_lines;
                            dr["New_Averages"] = ((int)Env.Averages).ToString();
                            #endregion
                        }
                    }
                    else if (Pt.PointType == PointType.AccToVel)
                    {
                        bool VelCheck = VEL_SSCP_Check(_analystDB, dr["Point_Id"].ToString());
                        if (VelCheck == true)
                        {
                            dr.Delete();
                        }
                        else
                        {
                            #region AccToVel
                            float RPM = GetPointRPM(cn, uint.Parse(dr["Point_Id"].ToString()));
                            float FTF = Get_FTF(cn, uint.Parse(dr["Point_Id"].ToString()));
                            float BPFI = Get_Great_BPFI(cn, uint.Parse(dr["Point_Id"].ToString()));
                            dr["New_Detection"] = ((int)Vel.DetectionType).ToString();
                            dr["New_Start_Freq"] = ((int)Vel.StartFreq).ToString();
                            dr["New_End_Freq"] = (float)Math.Ceiling(Convert.ToDecimal(Get_GMF(cn, uint.Parse(dr["Point_Id"].ToString()))));
                            dr["New_Lines"] = ((int)Vel.Lines).ToString();
                            dr["New_Low_Freq_CutOff"] = (float)Math.Ceiling(Convert.ToDecimal(Vel_LowCutOff(BPFI, uint.Parse(dr["Lines"].ToString()), RPM)));
                            dr["New_Averages"] = ((int)Vel.Averages).ToString();
                            #endregion
                        }
                    }
                    i++;
                    lblNSR.Invoke(new MethodInvoker(delegate { pbNSR.Value = i; lblNSR.Text = i + "/" + Dg.Rows.Count.ToString(); }));
                }
                dgNSR.Invoke(new MethodInvoker(delegate { dgNSR.DataSource = Dg; }));
            //}
        }
        private bool ACC_SSCP_Check(AnalystConnection cn, string Ponto)
        {
            bool _ACC_SSCP_Check = false;


            string Detection = cn.SQLtoString("VALUESTRING", "POINT", "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Detection").ToString() + " AND ELEMENTID=" + Ponto);
            float StartFreq = float.Parse(cn.SQLtoString("VALUESTRING", "POINT", "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Start_Freq").ToString() + " AND ELEMENTID=" + Ponto));
            float StopFreq = float.Parse(cn.SQLtoString("VALUESTRING", "POINT", "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_End_Freq").ToString() + " AND ELEMENTID=" + Ponto));
            float LowCutOff = float.Parse(cn.SQLtoString("VALUESTRING", "POINT", "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Low_Freq_Cutoff").ToString() + " AND ELEMENTID=" + Ponto));
            string Lines = cn.SQLtoString("VALUESTRING", "POINT", "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Lines").ToString() + " AND ELEMENTID=" + Ponto);
            string Averages = cn.SQLtoString("VALUESTRING", "POINT", "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Averages").ToString() + " AND ELEMENTID=" + Ponto);
            
            uint SSCP_Detection = (int)Acc.DetectionType;
            uint SSCP_Lines = (int)Acc.Lines;
            float SSCP_StartFreq = (int)Acc.StartFreq;
            float SSCP_StopFreq = (int)Acc.StopFreq;
            float SSCP_LowCutOff = (int)Acc.LowCutOff;
            uint SSCP_Averages = (int)Acc.Averages;
            uint SSCP_SaveDate = 0;

            SKF.RS.STB.Analyst.Point pt = new SKF.RS.STB.Analyst.Point(cn, uint.Parse(Ponto));
            
            float RPM = GetPointRPM(cn, uint.Parse(Ponto));
            float FTF = Get_FTF(cn, uint.Parse(Ponto));
            float BPFI = Get_Great_BPFI(cn, uint.Parse(Ponto));

            AccParam(pt, RPM, FTF, BPFI, out SSCP_Detection, out SSCP_SaveDate, out SSCP_Lines, out SSCP_Averages, out SSCP_StartFreq, out SSCP_StopFreq, out SSCP_LowCutOff);

            if (
                Detection != SSCP_Detection.ToString()
                || StartFreq != SSCP_StartFreq
                || Lines != SSCP_Lines.ToString()
                || StopFreq != SSCP_StopFreq
                || LowCutOff != SSCP_LowCutOff
                || Averages != SSCP_Averages.ToString()
                )
                _ACC_SSCP_Check = false;
            else
                _ACC_SSCP_Check = true;

            

            return _ACC_SSCP_Check;

        }
        private bool ENV_SSCP_Check(AnalystConnection cn, string Ponto)
        {
            bool _ENV_SSCP_Check = false;


            float RPM = GetPointRPM(cn, uint.Parse(Ponto));
            float FTF = Get_FTF(cn, uint.Parse(Ponto));
            float BPFI = Get_Great_BPFI(cn, uint.Parse(Ponto));


            string Detection = cn.SQLtoString("VALUESTRING", "POINT", "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Detection").ToString() + " AND ELEMENTID=" + Ponto);
            
            float StartFreq = float.Parse(cn.SQLtoString("VALUESTRING", "POINT", "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Start_Freq").ToString() + " AND ELEMENTID=" + Ponto));
            float StopFreq = float.Parse(cn.SQLtoString("VALUESTRING", "POINT", "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_End_Freq").ToString() + " AND ELEMENTID=" + Ponto));
            float LowCutOff = float.Parse(cn.SQLtoString("VALUESTRING", "POINT", "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Low_Freq_Cutoff").ToString() + " AND ELEMENTID=" + Ponto));
            string Lines = cn.SQLtoString("VALUESTRING", "POINT", "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Lines").ToString() + " AND ELEMENTID=" + Ponto);
            string Averages = cn.SQLtoString("VALUESTRING", "POINT", "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Averages").ToString() + " AND ELEMENTID=" + Ponto);

            uint env_lines = Env_Lines(FTF, BPFI, RPM);
            float env_stopfreq = (float)Math.Ceiling(Convert.ToDecimal(Env_Range(BPFI, RPM)));
            float env_lowcutoff = (float)Math.Ceiling(Convert.ToDecimal(Env_LowCutOff(env_stopfreq, env_lines, RPM)));


            uint SSCP_Detection = (int)Env.DetectionType;
            uint SSCP_Lines = env_lines;
            float SSCP_StartFreq = (int)Env.StartFreq;
            float SSCP_StopFreq = env_stopfreq;
            float SSCP_LowCutOff = env_lowcutoff;
            uint SSCP_Averages = (int)Env.Averages;
            uint SSCP_SaveDate = 0;

            SKF.RS.STB.Analyst.Point pt = new SKF.RS.STB.Analyst.Point(cn, uint.Parse(Ponto));
            EnvParam(pt, RPM, FTF, BPFI, out SSCP_Detection, out SSCP_SaveDate, out SSCP_Lines, out SSCP_Averages, out SSCP_StartFreq, out SSCP_StopFreq, out SSCP_LowCutOff);

            if (
                Detection != SSCP_Detection.ToString()
                || StartFreq != SSCP_StartFreq
                || Lines != SSCP_Lines.ToString()
                || StopFreq != SSCP_StopFreq
                || LowCutOff != SSCP_LowCutOff
                || Averages != SSCP_Averages.ToString()
                )
                _ENV_SSCP_Check = false;
            else
                _ENV_SSCP_Check = true;

            return _ENV_SSCP_Check;

        }
        private bool VEL_SSCP_Check(AnalystConnection cn, string Ponto)
        {
            bool _VEL_SSCP_Check = false;

            float RPM = GetPointRPM(cn, uint.Parse(Ponto));
            float FTF = Get_FTF(cn, uint.Parse(Ponto));
            float BPFI = Get_Great_BPFI(cn, uint.Parse(Ponto));

            string Detection = cn.SQLtoString("VALUESTRING", "POINT", "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Detection").ToString() + " AND ELEMENTID=" + Ponto);
            float StartFreq = float.Parse(cn.SQLtoString("VALUESTRING", "POINT", "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Start_Freq").ToString() + " AND ELEMENTID=" + Ponto));
            float StopFreq = float.Parse(cn.SQLtoString("VALUESTRING", "POINT", "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_End_Freq").ToString() + " AND ELEMENTID=" + Ponto));
            float LowCutOff = float.Parse(cn.SQLtoString("VALUESTRING", "POINT", "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Low_Freq_Cutoff").ToString() + " AND ELEMENTID=" + Ponto));
            string Lines = cn.SQLtoString("VALUESTRING", "POINT", "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Lines").ToString() + " AND ELEMENTID=" + Ponto);
            string Averages = cn.SQLtoString("VALUESTRING", "POINT", "FIELDID=" + Registration.RegistrationId(cn, "SKFCM_ASPF_Averages").ToString() + " AND ELEMENTID=" + Ponto);

            uint SSCP_Detection = (int)Vel.DetectionType;
            uint SSCP_Lines = (uint)Vel.Lines;
            float SSCP_StartFreq = (int)Env.StartFreq;
            float SSCP_StopFreq = (float)Math.Ceiling(Convert.ToDecimal(Get_GMF(cn, uint.Parse(Ponto))));
            float SSCP_LowCutOff = (float)Math.Ceiling(Convert.ToDecimal(Vel_LowCutOff(BPFI, SSCP_Lines, RPM)));
            uint SSCP_Averages = (int)Vel.Averages;
            uint SSCP_SaveDate = 0;

            SKF.RS.STB.Analyst.Point pt = new SKF.RS.STB.Analyst.Point(cn, uint.Parse(Ponto));
            VelParam(pt, RPM, FTF, BPFI, out SSCP_Detection, out SSCP_SaveDate, out SSCP_Lines, out SSCP_Averages, out SSCP_StartFreq, out SSCP_StopFreq, out SSCP_LowCutOff);


            if (SSCP_StopFreq == 0)
                StopFreq = SSCP_StopFreq;

            if (
                Detection != SSCP_Detection.ToString()
                || StartFreq != SSCP_StartFreq
                || Lines != SSCP_Lines.ToString()
                || StopFreq != SSCP_StopFreq
                || LowCutOff != SSCP_LowCutOff
                || Averages != SSCP_Averages.ToString()
                )
                _VEL_SSCP_Check = false;
            else
                _VEL_SSCP_Check = true;

            return _VEL_SSCP_Check;

        }
        public void AddMachine(WorkSpace WorkSpace, SKF.RS.STB.Analyst.Machine AnMachine, string Priority)
        {
            GenericTools.DebugMsg("AddMachine(" + WorkSpace.TreeElem.TreeElemId.ToString() + ", " + AnMachine.TreeElemId.ToString() + ", " + Priority + "): Starting...");
            try
            {
                Set WorkSpaceSet = new Set(WorkSpace.TreeElem);
                Set ParentSet = WorkSpace.AddSet(WorkSpaceSet, new Set(AnMachine.TreeElem.Parent), false);
                WorkSpace.AddMachine(ParentSet, AnMachine);
            }
            catch (Exception ex)
            {
                GenericTools.DebugMsg("AddMachine(" + WorkSpace.TreeElem.TreeElemId.ToString() + ", " + AnMachine.TreeElemId.ToString() + ", " + Priority + ") error: " + ex.Message);
            }
            GenericTools.DebugMsg("AddMachine(" + WorkSpace.TreeElem.TreeElemId.ToString() + ", " + AnMachine.TreeElemId.ToString() + ", " + Priority + "): Finished!");
        }
        private void treeView_AnStandards_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }
        private void label7_Click(object sender, EventArgs e)
        {

        }
        private void txtNSR_RPM_TextChanged(object sender, EventArgs e)
        {

        }
        private void label6_Click(object sender, EventArgs e)
        {

        }
        private void button2_Click_1(object sender, EventArgs e)
        {
            if (txtboxPrevMudTblSetId.Text != "")
            {
                string SQL_WS = "";
                if (checkBox1.Checked == true)
                {
                    SQL_WS = "SELECT referenceid from TREEELEM where HIERARCHYTYPE=3 and CONTAINERTYPE=4 and tblsetid=" + txtboxPrevMudTblSetId.Text +
                                " and Parentid!=2147000000 and hierarchyid in (" +
                                " SELECT TREEELEMID FROM TREEELEM WHERE HIERARCHYTYPE=3 and Parentid!=2147000000 and NAME LIKE '" + textBox1.Text + "' " +
                                " or NAME LIKE '" + textBox1.Text + "')";


                }


                StringBuilder SQL = new StringBuilder();

                SQL.Clear();
                SQL.Append(" select ");
                SQL.Append(" 	[Set].NAME as [Parent_Node],  ");
                SQL.Append(" 	Ativo.NAME as Asset_Name,  ");
                SQL.Append(" 	Ponto.TREEELEMID as Point_Id,  ");
                SQL.Append(" 	Ponto.NAME as Point_Name,  ");
                SQL.Append(" 	pt2.VALUESTRING as Sensor_Type, ");
                SQL.Append(" 	CASE pt3.VALUESTRING ");
                SQL.Append(" 		when '20500' then 'Peak' ");
                SQL.Append(" 		when '20501' then 'Peak to Peak' ");
                SQL.Append(" 		when '20502' then 'RMS' ");
                SQL.Append(" 	else 'None' end as [Detection], ");
                SQL.Append(" 	CASE PT10.VALUESTRING  ");
                SQL.Append(" 		WHEN '20400' then 'Uniform' ");
                SQL.Append(" 		WHEN '20401' then 'Hanning' ");
                SQL.Append(" 		WHEN '20402' then 'Flattop' ");
                SQL.Append(" 	end as [Window], ");
                SQL.Append(" 	ROUND((CAST(pt4.valuestring as float) * 60),2) as Speed_RPM, ");
                SQL.Append(" 	CASE pt5.VALUESTRING ");
                SQL.Append(" 		when '20200' then 'FFT' ");
                SQL.Append(" 		when '20201' then 'Time' ");
                SQL.Append(" 		when '20202' then 'FFT and Time' ");
                SQL.Append(" 	end as [Save_Data], ");
                SQL.Append(" 	(((CAST(pt11.valuestring as float) * 0.5) + 0.5) * CAST(pt9.valuestring as float) / CAST(pt7.valuestring as float)) as Process_Time, ");
                // SQL.Append(" 	ROUND((CAST(pt9.valuestring as float) / CAST(pt7.valuestring as float) * CAST(pt11.valuestring as float) * 0.5),3) as Process_Time, ");
                // SQL.Append(" 	(CAST(pt9.valuestring as float) / (pt7.valuestring as float)) as Process_Time, ");
                SQL.Append(" 	pt6.VALUESTRING + ' Hz' as [Start_Freq], ");
                SQL.Append(" 	pt8.VALUESTRING + ' Hz' as [Low_Freq_Cutoff], ");
                SQL.Append(" 	pt7.VALUESTRING + ' Hz' as [End_Freq], ");
                SQL.Append(" 	pt9.VALUESTRING as Lines, ");
                SQL.Append(" 	pt11.VALUESTRING as [Averages] ");
                SQL.Append(" 	, REPLACE( ");
                SQL.Append(" 		REPLACE( ");
                SQL.Append(" 		STUFF(( ");
                SQL.Append(" 				SELECT  ");
                SQL.Append(" 					 FE.NAME ");
                SQL.Append(" 				FROM  ");
                SQL.Append(" 					TREEELEM TE ");
                SQL.Append(" 					, POINT PT ");
                SQL.Append(" 					, FREQASSIGN FA ");
                SQL.Append(" 					, FREQENTRIES FE ");
                SQL.Append(" 				WHERE ");
                SQL.Append(" 					TE.TREEELEMID=PT.ELEMENTID ");
                SQL.Append(" 					AND CAST(PT.VALUESTRING as  nvarchar)=CAST((select REGISTRATIONID FROM REGISTRATION WHERE SIGNATURE='SKFCM_ASDD_MicrologDAD') as nvarchar) ");
                SQL.Append(" 					AND fa.ELEMENTID=PT.ELEMENTID ");
                SQL.Append(" 					AND Fa.FSID=FE.FSID ");
                SQL.Append(" 					AND PT.ELEMENTID=Ponto.TREEELEMID ");
                SQL.Append(" 				 FOR XML PATH ('') ");
                SQL.Append(" 			 ) ");
                SQL.Append(" 			 , 1, 0, '' ");
                SQL.Append(" 			) ");
                SQL.Append(" 		,'<NAME>','') ");
                SQL.Append(" 		,'</NAME>',', ') AS Asset_Data ");
                SQL.Append(" 	, '' as 'New_Save_Data' ");
                SQL.Append(" 	, '' as 'New_Process_Time' ");
                SQL.Append(" 	, '' as 'New_Start_Freq' ");
                SQL.Append(" 	, '' as 'New_Low_Freq_Cutoff' ");
                SQL.Append(" 	, '' as 'New_End_Freq' ");
                SQL.Append(" 	, '' as 'New_Lines' ");
                SQL.Append(" 	, '' as 'New_Averages' ");
                SQL.Append(" from ");
                SQL.Append(" 	TREEELEM Ponto, ");
                SQL.Append(" 	TREEELEM Ativo, ");
                SQL.Append(" 	TREEELEM [Set], ");
                SQL.Append(" 	POINT pt, ");
                SQL.Append(" 	POINT pt2, ");
                SQL.Append(" 	POINT pt3, ");
                SQL.Append(" 	POINT pt4, ");
                SQL.Append(" 	POINT pt5, ");
                SQL.Append(" 	POINT pt6, ");
                SQL.Append(" 	POINT pt7, ");
                SQL.Append(" 	POINT pt8, ");
                SQL.Append(" 	POINT pt9, ");
                SQL.Append(" 	POINT pt10, ");
                SQL.Append(" 	POINT pt11 ");
                SQL.Append(" where ");
                SQL.Append(" 	ativo.HIERARCHYTYPE=1 ");
                SQL.Append(" 	and ativo.CONTAINERTYPE=3 ");
                SQL.Append(" 	and ativo.ELEMENTENABLE=1 ");
                SQL.Append(" 	and ativo.PARENTENABLE=0 ");
                SQL.Append(" 	and ativo.tblsetid=" + txtboxPrevMudTblSetId.Text);
                SQL.Append(" 	and ativo.TREEELEMID=ponto.PARENTID ");
                if (SQL_WS != "")
                {
                    SQL.Append(" and ponto.TREEELEMID in (" + SQL_WS + ")");
                }
                SQL.Append(" 	and pt.VALUESTRING in ( ");
                SQL.Append(" 		CAST((select REGISTRATIONID from REGISTRATION where SIGNATURE='SKFCM_ASDD_MicrologDAD')AS nvarchar) ");
                SQL.Append(" 		, CAST((select REGISTRATIONID from REGISTRATION where SIGNATURE='SKFCM_ASDD_ImxDAD') AS nvarchar) ");
                SQL.Append(" 		)  ");
                SQL.Append(" 	and pt.ELEMENTID=pt2.ELEMENTID ");
                SQL.Append(" 	and pt2.FIELDID=(select REGISTRATIONID from REGISTRATION where SIGNATURE='SKFCM_ASPF_Sensor') ");
                SQL.Append(" 	and pt2.ELEMENTID=pt3.ELEMENTID ");
                SQL.Append(" 	and pt3.FIELDID=(select REGISTRATIONID from REGISTRATION where SIGNATURE='SKFCM_ASPF_Detection') ");
                SQL.Append(" 	and pt3.ELEMENTID=pt4.ELEMENTID ");
                SQL.Append(" 	and pt4.FIELDID=(select REGISTRATIONID from REGISTRATION where SIGNATURE='SKFCM_ASPF_Speed') ");
                SQL.Append(" 	and pt4.ELEMENTID=pt5.ELEMENTID ");
                SQL.Append(" 	and pt5.FIELDID=(select REGISTRATIONID from REGISTRATION where SIGNATURE='SKFCM_ASPF_Save_Data') ");
                SQL.Append(" 	and pt5.ELEMENTID=pt6.ELEMENTID ");
                SQL.Append(" 	and pt6.FIELDID=(select REGISTRATIONID from REGISTRATION where SIGNATURE='SKFCM_ASPF_Start_Freq') ");
                SQL.Append(" 	and pt6.ELEMENTID=pt7.ELEMENTID ");
                SQL.Append(" 	and pt7.FIELDID=(select REGISTRATIONID from REGISTRATION where SIGNATURE='SKFCM_ASPF_End_Freq') ");
                SQL.Append(" 	and pt7.ELEMENTID=pt8.ELEMENTID ");
                SQL.Append(" 	and pt8.FIELDID=(select REGISTRATIONID from REGISTRATION where SIGNATURE='SKFCM_ASPF_Low_Freq_Cutoff') ");
                SQL.Append(" 	and pt8.ELEMENTID=pt9.ELEMENTID ");
                SQL.Append(" 	and pt9.FIELDID=(select REGISTRATIONID from REGISTRATION where SIGNATURE='SKFCM_ASPF_Lines') ");
                SQL.Append(" 	and pt9.ELEMENTID=pt10.ELEMENTID ");
                SQL.Append(" 	and pt10.FIELDID=(select REGISTRATIONID from REGISTRATION where SIGNATURE='SKFCM_ASPF_Window') ");
                SQL.Append(" 	and pt10.ELEMENTID=pt11.ELEMENTID ");
                SQL.Append(" 	and pt11.FIELDID=(select REGISTRATIONID from REGISTRATION where SIGNATURE='SKFCM_ASPF_Averages') ");
                SQL.Append(" 	and Ponto.TREEELEMID=pt.ELEMENTID ");
                SQL.Append(" 	and Ativo.PARENTID=[Set].TREEELEMID ");

                Log(SQL.ToString());

                DataTable prevMedicao = _analystDB.DataTable(SQL.ToString());
                Log("Rows Point: " + prevMedicao.Rows.Count.ToString());

                uint SSCP_Detection = 0;
                uint SSCP_Lines = 0;
                float SSCP_StartFreq = 0;
                float SSCP_StopFreq = 0;
                float SSCP_LowCutOff = 0;
                uint SSCP_Averages = 0;
                uint SSCP_SaveDate = 0;

                for (int i = prevMedicao.Rows.Count - 1; i >= 0; i--)
                {
                    uint point = uint.Parse(prevMedicao.Rows[i]["Point_Id"].ToString());
                    SKF.RS.STB.Analyst.Point Pt = new SKF.RS.STB.Analyst.Point(_analystDB, point);

                    if (Pt.PointType == PointType.Displacement)
                    {
                        prevMedicao.Rows[i].Delete();
                    }
                }
                prevMedicao.AcceptChanges();

                foreach(DataRow dr in prevMedicao.Rows)
                {
                    SKF.RS.STB.Analyst.Point Pt = new SKF.RS.STB.Analyst.Point(_analystDB, uint.Parse(dr["Point_Id"].ToString()));


                    uint Ponto = uint.Parse(dr["Point_Id"].ToString());
                    float RPM = GetPointRPM(_analystDB, Ponto);
                    float FTF = Get_FTF(_analystDB, Ponto);
                    float BPFI = Get_Great_BPFI(_analystDB, Ponto);
                    

                    if (Pt.PointType == PointType.Acc)
                    {
                        SSCP_Detection = (int)Acc.DetectionType;
                        SSCP_Lines = (int)Acc.Lines;
                        SSCP_StartFreq = (int)Acc.StartFreq;
                        SSCP_StopFreq = (int)Acc.StopFreq;
                        SSCP_LowCutOff = (int)Acc.LowCutOff;
                        SSCP_Averages = (int)Acc.Averages;
                        SSCP_SaveDate = 0;
                        try
                        {
                            AccParam(Pt, RPM, FTF, BPFI, out SSCP_Detection, out SSCP_SaveDate, out SSCP_Lines, out SSCP_Averages, out SSCP_StartFreq, out SSCP_StopFreq, out SSCP_LowCutOff);
                        }
                        catch(Exception ex)
                        {
                            GenericTools.WriteLog("Calculation Error Acceleration: " + ex.Message);
                        }
                       
                    }
                    else if (Pt.PointType == PointType.AccEnvelope)
                    {
                        uint env_lines = Env_Lines(FTF, BPFI, RPM);
                        float env_stopfreq = (float)Math.Ceiling(Convert.ToDecimal(Env_Range(BPFI, RPM)));
                        float env_lowcutoff = (float)Math.Ceiling(Convert.ToDecimal(Env_LowCutOff(env_stopfreq, env_lines, RPM)));

                        SSCP_Detection = (int)Env.DetectionType;
                        SSCP_Lines = env_lines;
                        SSCP_StartFreq = (int)Env.StartFreq;
                        SSCP_StopFreq = env_stopfreq;
                        SSCP_LowCutOff = env_lowcutoff;
                        SSCP_Averages = (int)Env.Averages;
                        SSCP_SaveDate = 0;
                        try
                        {                        
                            EnvParam(Pt, RPM, FTF, BPFI, out SSCP_Detection, out SSCP_SaveDate, out SSCP_Lines, out SSCP_Averages, out SSCP_StartFreq, out SSCP_StopFreq, out SSCP_LowCutOff);
                        }
                        catch (Exception ex)
                        {
                            GenericTools.WriteLog("Calculation Error Envelope: " + ex.Message);
                        }
                    }
                    else if (Pt.PointType == PointType.AccToVel)
                    {
                        SSCP_Detection = (int)Vel.DetectionType;
                        SSCP_Lines = (uint)Vel.Lines;
                        SSCP_StartFreq = (int)Env.StartFreq;
                        SSCP_StopFreq = (float)Math.Ceiling(Convert.ToDecimal(Get_GMF(_analystDB, Ponto)));
                        SSCP_LowCutOff = (float)Math.Ceiling(Convert.ToDecimal(Vel_LowCutOff(BPFI, SSCP_Lines, RPM)));
                        SSCP_Averages = (int)Vel.Averages;
                        SSCP_SaveDate = 0;
                        try
                        {
                            VelParam(Pt, RPM, FTF, BPFI, out SSCP_Detection, out SSCP_SaveDate, out SSCP_Lines, out SSCP_Averages, out SSCP_StartFreq, out SSCP_StopFreq, out SSCP_LowCutOff);
                        }
                        catch (Exception ex)
                        {
                            GenericTools.WriteLog("Calculation Error Velocity: " + ex.Message);
                        }
                    }
                    if (Pt.PointId == 625775)
                    {
                        string parou = "";
                    }
                   
                    string salva_dados = "";
                    switch (SSCP_SaveDate)
                    {
                        case 20200:
                            salva_dados = "FFT";
                            break;
                        case 20201:
                            salva_dados = "Time";
                            break;
                        case 20202:
                            salva_dados = "FFT and Time";
                            break;
                    }

                    dr[16] = salva_dados;
                    dr[17] = SSCP_Lines / SSCP_StopFreq;
                    dr[18] = SSCP_StartFreq + " Hz";
                    dr[20] = SSCP_StopFreq + " Hz";
                    dr[19] = SSCP_LowCutOff + " Hz";
                    dr[21] = SSCP_Lines;
                    dr[22] = SSCP_Averages;
                }

                dgPrevisaoSSCP.DataSource = prevMedicao;
            }
            else
            {
                MessageBox.Show("You MUST input a TBLSETID!!!");
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (dgPrevisaoSSCP.Rows.Count > 0)
            {
                        // Creating a Excel object. 
                Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application(); 
                Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing); 
                Microsoft.Office.Interop.Excel._Worksheet worksheet = null; 
 
                try 
                { 
 
                    worksheet = workbook.ActiveSheet; 
 
                    worksheet.Name = "Export"; 
 
                    int cellRowIndex = 1; 
                    int cellColumnIndex = 1; 
 
                    //Loop through each row and read value from each column. 
                    for (int i = 0; i < dgPrevisaoSSCP.Rows.Count - 1; i++) 
                    {
                        for (int j = 0; j < dgPrevisaoSSCP.Columns.Count; j++) 
                        { 
                            // Excel index starts from 1,1. As first Row would have the Column headers, adding a condition check. 
                            if (cellRowIndex == 1) 
                            {
                                worksheet.Cells[cellRowIndex, cellColumnIndex] = dgPrevisaoSSCP.Columns[j].HeaderText; 
                            } 
                            else 
                            {
                                worksheet.Cells[cellRowIndex, cellColumnIndex] = dgPrevisaoSSCP.Rows[i].Cells[j].Value.ToString(); 
                            } 
                            cellColumnIndex++; 
                        } 
                        cellColumnIndex = 1; 
                        cellRowIndex++; 
                    } 
 
                    //Getting the location and file name of the excel to save from user. 
                    SaveFileDialog saveDialog = new SaveFileDialog(); 
                    saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*"; 
                    saveDialog.FilterIndex = 2; 
 
                    if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK) 
                    { 
                        workbook.SaveAs(saveDialog.FileName); 
                        MessageBox.Show("Export Successful"); 
                    } 
                } 
                catch (System.Exception ex) 
                { 
                    MessageBox.Show(ex.Message); 
                } 
                finally 
                { 
                    excel.Quit(); 
                    workbook = null; 
                    excel = null; 
                } 
            }
        }
        private void button4_Click(object sender, EventArgs e)
        {
             //Getting the location and file name of the excel to save from user. 
            SaveFileDialog saveDialog = new SaveFileDialog(); 
            saveDialog.Filter = "CSV files (*.csv)|*.csv"; 
            saveDialog.FilterIndex = 2; 
 
            if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK) 
            {
                CSVExport.ExportToCSV(dgPrevisaoSSCP, saveDialog.FileName);
            }
        }
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                cbRRS_Workspace.Enabled = true;
                if (cbRRS_Workspace.Items.Count > 0)
                {
                    cbRRS_Workspace.SelectedIndex = 0;
                    button2.Enabled = true;
                }
                else
                    button2.Enabled = false;
            }
            else
            {
                cbRRS_Workspace.Text = "";
                cbRRS_Workspace.Enabled = false;
                button2.Enabled = true;
            }
        }

        private void cbASS_Hierarchy_SelectedIndexChanged(object sender, EventArgs e)
        {
            dynamic test = this.cbASS_Hierarchy.Items[cbASS_Hierarchy.SelectedIndex];
            if (test.Value.ToString() != "0")
            {
                updateWorkspace(int.Parse(test.Value.ToString()), cboASS_Workspace);
                cboASS_Workspace.Enabled = false;
                chkWorkSpace.Enabled = true;
                chkWorkSpace.Checked = false;
                cboASS_Workspace.Text = "";
                txtTBLSETID.Text = test.Value.ToString();
            }
            else
            {
                cboASS_Workspace.Items.Clear();
                chkWorkSpace.Enabled = false;
                cboASS_Workspace.Enabled = false;
                txtTBLSETID.Text = "";
            }
        }

        private void chkWorkSpace_CheckedChanged(object sender, EventArgs e)
        {
            if (chkWorkSpace.Checked)
            {
                cboASS_Workspace.Enabled = true;
                if (cboASS_Workspace.Items.Count > 0)
                {
                    cboASS_Workspace.SelectedIndex = 0;
                    button1.Enabled = true;
                }
                else
                    button1.Enabled = false;
            }
            else
            {
                cboASS_Workspace.Text = "";
                cboASS_Workspace.Enabled = false;
                button1.Enabled = true;
            }
        }

        private void cboASS_Workspace_SelectedIndexChanged(object sender, EventArgs e)
        {
            dynamic test = this.cboASS_Workspace.Items[cboASS_Workspace.SelectedIndex];
             if (test.Value.ToString() != "0")
             {
                 txtWS.Text = test.Text;
             }
        }

        private void cbCSS_Hierarchy_SelectedIndexChanged(object sender, EventArgs e)
        {
            dynamic test = this.cbCSS_Hierarchy.Items[cbCSS_Hierarchy.SelectedIndex];
            if (test.Value.ToString() != "0")
            {
                updateWorkspace(int.Parse(test.Value.ToString()), cbCSS_Workspace);
                cbCSS_Workspace.Enabled = false;
                chkCSS_Workspace.Enabled = true;
                chkCSS_Workspace.Checked = false;
                cbCSS_Workspace.Text = "";

                txtAES_TBLSETID.Text = test.Value.ToString();
            }
            else
            {
                cbCSS_Workspace.Items.Clear();
                chkCSS_Workspace.Enabled = false;
                cbCSS_Workspace.Enabled = false;
                txtAES_TBLSETID.Text = "";
            }
        }

        private void cbCSS_Workspace_SelectedIndexChanged(object sender, EventArgs e)
        {
            dynamic test = this.cbCSS_Workspace.Items[cbCSS_Workspace.SelectedIndex];
            if (test.Value.ToString() != "0")
            {
                txtAES_WS.Text = test.Text;
            }
        }

        private void chkCSS_Workspace_CheckedChanged(object sender, EventArgs e)
        {
            if (chkCSS_Workspace.Checked)
            {
                cbCSS_Workspace.Enabled = true;
                if (cbCSS_Workspace.Items.Count > 0)
                {
                    cbCSS_Workspace.SelectedIndex = 0;
                    btSSP_Load.Enabled = true;
                }
                else
                    btSSP_Load.Enabled = false;
            }
            else
            {
                cbCSS_Workspace.Text = "";
                cbCSS_Workspace.Enabled = false;
                btSSP_Load.Enabled = true;
            }
        }

        private void cbCOC_Hierarchy_SelectedIndexChanged(object sender, EventArgs e)
        {
            dynamic test = this.cbCOC_Hierarchy.Items[cbCOC_Hierarchy.SelectedIndex];

            txtNSR_TSID.Text = test.Value.ToString() != "0" ? (string) test.Value.ToString() : "";
        }

        private void cbRRS_Hierarchy_SelectedIndexChanged(object sender, EventArgs e)
        {
            dynamic test = this.cbRRS_Hierarchy.Items[cbRRS_Hierarchy.SelectedIndex];
            if (test.Value.ToString() != "0")
            {
                updateWorkspace(int.Parse(test.Value.ToString()), cbRRS_Workspace);
                cbRRS_Workspace.Enabled = false;
                checkBox1.Enabled = true;
                checkBox1.Checked = false;
                cbRRS_Workspace.Text = "";
                txtboxPrevMudTblSetId.Text = test.Value.ToString();
            }
            else
            {
                cbRRS_Workspace.Items.Clear();
                checkBox1.Enabled = false;
                cbRRS_Workspace.Enabled = false;
                txtboxPrevMudTblSetId.Text = "";
            }
        }

        private void cbRRS_Workspace_SelectedIndexChanged(object sender, EventArgs e)
        {
            dynamic test = this.cbRRS_Workspace.Items[cbRRS_Workspace.SelectedIndex];
            if (test.Value.ToString() != "0")
            {
                textBox1.Text = test.Text;
            }
        }


        private void button5_Click(object sender, EventArgs e)
        {
            
            button5.Enabled = false;

            dynamic selected = tbl_setCombobox.Items[tbl_setCombobox.SelectedIndex];
            uint tblSetId = UInt32.Parse(selected.Value.ToString());

            _standardPoint = new StandardPointName(_analystDB);



            if (workspace_checkbox.Checked && txtWS.Text != string.Empty)
            {
                dynamic selectedWorkSpace = workspace_comboBox.Items[workspace_comboBox.SelectedIndex];
                uint workSpaceId = UInt32.Parse(selectedWorkSpace.Value.ToString());

                var workSpacePoints = _standardPoint.GetWorkSpacePoints(workSpaceId, tblSetId);
                //_newNamesToUpdateList = _standardPoint.ParseStandardNames(workSpacePoints);

                UpdateGridView(workSpacePoints);
            }
            else
            {
                var measPoints = _standardPoint.GetPointsByTableSetId(tblSetId);
                //_newNamesToUpdateList = _standardPoint.ParseStandardNames(measPoints);
                UpdateGridView(measPoints);
            }

            selectAllButton.Enabled = true;
            unselectAllButton.Enabled = true;
        }

        private void UpdateGridView(IEnumerable<TreeElem> treeElem)
        {
            _newNamesToUpdateList = new List<EstimateName>();

            DataTable dataTable = new DataTable();
            dataTable.Columns.AddRange(new DataColumn[6]
            {
                new DataColumn("ElementID", typeof(uint)),
                new DataColumn("Old Name", typeof(string)),
                new DataColumn("New Estimated Name", typeof(string)),
                new DataColumn("HierarchyId", typeof(uint)),
                new DataColumn("Select", typeof(bool)),
                new DataColumn("RefId", typeof(uint)),
            });

            // workspace already handled in button5_Click method

            if (autoEstimateCheckBox.Checked)
            {
                _newNamesToUpdateList = _standardPoint.ParseStandardNames(treeElem);
                // we have 2 cases hierarchy ID might be 0 if not checked
                // not 0 if checked
                foreach (var pt in _newNamesToUpdateList)
                {
                    dataTable.Rows.Add(pt.PointId, pt.OldName,
                        pt.NewName, pt.HierarchyId, false, pt.ReferenceId);
                }
            }
            else
            {
                foreach (var pt in treeElem)
                {
                    dataTable.Rows.Add(pt.TREEELEMID, pt.NAME,
                        "", pt.HIERARCHYID, false, pt.REFERENCEID);
                }
            }


            estimation_dataGridView.DataSource = dataTable;
            estimation_dataGridView.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            estimation_dataGridView.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            estimation_dataGridView.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader;

            estimation_dataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            estimation_dataGridView.MultiSelect = false;
            
            estimation_dataGridView.Columns[0].ReadOnly = true;
            estimation_dataGridView.Columns[0].Visible = false;

            estimation_dataGridView.Columns[1].ReadOnly = true;
            estimation_dataGridView.Columns[2].ReadOnly = false;
            estimation_dataGridView.Columns[3].Visible = false;
            estimation_dataGridView.Columns[5].Visible = false;
            
            estimation_dataGridView.AllowUserToAddRows = false;
            estimation_dataGridView.AllowUserToDeleteRows = false;
        }
        
        
        private void tbl_setCombobox_SelectedIndexChanged(object sender, EventArgs e)
        {
            dynamic test = this.tbl_setCombobox.Items[tbl_setCombobox.SelectedIndex];
            if (test.Value.ToString() != "0")
            {
                updateWorkspace(int.Parse(test.Value.ToString()), workspace_comboBox);
                workspace_comboBox.Enabled = false;
                workspace_checkbox.Enabled = true;
                workspace_checkbox.Checked = false;
                workspace_comboBox.Text = "";
                txtTBLSETID.Text = test.Value.ToString();
                button5.Enabled = true;
            }
            else
            {
                workspace_comboBox.Items.Clear();
                workspace_checkbox.Enabled = false;
                workspace_comboBox.Enabled = false;               
                txtTBLSETID.Text = "";
            }
        }

        private void UpdateEditMeasurementGroupBox()
        {
            editMeas_groupBox.Visible = true;

            var ints = new List<ComboBoxItem>(){
                new ComboBoxItem(){Id = 0, Text = "--SELECT--"}
            };
            for (int i = 1; i < 100; i++)
            {
                
                ints.Add(new ComboBoxItem(){Id = i, Text = i.ToString("00")});
            }

            // ----  bearingNumberComboBox
            bearingNumberComboBox.DataSource = ints;
            bearingNumberComboBox.DisplayMember = "Text";
            bearingNumberComboBox.ValueMember = "Id";
            
            // ----  angularOrientationComboBox
            angularOrientationComboBox.DataSource = AngularOrientation.Codes;
            angularOrientationComboBox.DisplayMember = "Text";
            angularOrientationComboBox.ValueMember = "Id";
            angularOrientationComboBox.SelectedItem = AngularOrientation.Codes[0];

            // ---- mearsurementTypeComboBox
            mearsurementTypeComboBox.DataSource = MeasurementType.Codes;
            mearsurementTypeComboBox.DisplayMember = "Text";
            mearsurementTypeComboBox.ValueMember = "Id";
            mearsurementTypeComboBox.SelectedItem = MeasurementType.Codes[0];


            // ---- machineSideComboBox
            machineSideComboBox.DataSource = MachineSide.Codes;
            machineSideComboBox.DisplayMember = "Text";
            machineSideComboBox.ValueMember = "Id";
            machineSideComboBox.SelectedItem = MachineSide.Codes[0];

            // ---- dataSourceCombobox
            dataSourceCombobox.DataSource = DataSource.Codes;
            dataSourceCombobox.DisplayMember = "Text";
            dataSourceCombobox.ValueMember = "Id";
            dataSourceCombobox.SelectedItem = DataSource.Codes[0];

            // ----- machineShaftComboBox
            machineShaftComboBox.DataSource = MachineShaft.Codes;
            machineShaftComboBox.DisplayMember = "Text";
            machineShaftComboBox.ValueMember = "Id";
            machineShaftComboBox.SelectedItem = MachineShaft.Codes[0];

            // -----measurementAttributeComboBox
            measurementAttributeComboBox.DataSource = MeasurementAttribute.Codes;
            measurementAttributeComboBox.DisplayMember = "Text";
            measurementAttributeComboBox.ValueMember = "Id";
            measurementAttributeComboBox.SelectedItem = MeasurementAttribute.Codes[0];
        }


        private void button6_Click(object sender, EventArgs e)
        {
            if (estimation_dataGridView.SelectedRows.Count == 0) return;

            var row = estimation_dataGridView.SelectedRows[0];
            var elementId = row.Cells["ElementID"].Value;
            var elementName = row.Cells["Old Name"].Value;

            if (bearingNumberComboBox.SelectedIndex <= 0 || angularOrientationComboBox.SelectedIndex <= 0 ||
                mearsurementTypeComboBox.SelectedIndex <= 0 || machineSideComboBox.SelectedIndex <= 0)
            {
                MessageBox.Show($"Please Select Required Fields", "Information");
                return;
            }

            var dataSourceItem = (dataSourceCombobox.SelectedIndex<=0) ?  null : (ComboBoxItem)dataSourceCombobox.SelectedItem;
            var bearingNumberItem = (ComboBoxItem)bearingNumberComboBox.SelectedItem;
            var angularOrientationItem =  (ComboBoxItem)angularOrientationComboBox.SelectedItem;
            var mearsurementTypeItem =  (ComboBoxItem)mearsurementTypeComboBox.SelectedItem;
            var machineShaftItem = machineShaftComboBox.SelectedIndex<=0 ? null : (ComboBoxItem)machineShaftComboBox.SelectedItem;
            var machineSideItem = (ComboBoxItem)machineSideComboBox.SelectedItem;
            var measAttributeItem = measurementAttributeComboBox.SelectedIndex<= 0 ? null : (ComboBoxItem)measurementAttributeComboBox.SelectedItem;
            
            var newName =
                $" {dataSourceItem?.Text} {bearingNumberItem.Text}{angularOrientationItem.Text}{mearsurementTypeItem.Text} {machineShaftItem?.Text}" +
                $" {machineSideItem.Text} {measAttributeItem?.Text}".TrimEnd();

            
            var est = new EstimateName()
            {
                PointId = UInt32.Parse(elementId.ToString()),
                OldName = elementName.ToString(),
                NewName = newName.TrimStart()
            };

            var estObject = _newNamesToUpdateList.FirstOrDefault(x => x.PointId == est.PointId);

            if (estObject != null)
            {
                estObject.NewName = est.NewName;
            }
            else
            {
                _newNamesToUpdateList.Add(est);
            }
            
            editMeas_groupBox.Visible = false;


            var cells = estimation_dataGridView.SelectedRows[0].Cells;

            _editProgrammatically = true;
            cells["New Estimated Name"].Value = est.NewName;
            _editProgrammatically = false;
           
        }

        private void spn_ChkWorkSpace_CheckedChanged(object sender, EventArgs e)
        {
            if (workspace_checkbox.Checked)
            {
                workspace_comboBox.Enabled = true;
                if (workspace_comboBox.Items.Count > 0)
                {
                    workspace_comboBox.SelectedIndex = 0;
                    button5.Enabled = true;
                }
                else
                    button5.Enabled = false;
            }
            else
            {
                workspace_comboBox.Text = "";
                workspace_comboBox.Enabled = false;
                button5.Enabled = true;
            }
        }

        private void workspace_comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            dynamic test = this.workspace_comboBox.Items[workspace_comboBox.SelectedIndex];
            if (test.Value.ToString() == "0") return;
            txtWS.Text = test.Text;
            button5.Enabled = true;
        }

        private void SelectRow_Changed(object sender, EventArgs e)
        {
            if(estimation_dataGridView.MultiSelect) return;
            // 0 because we are only selecton one row
            if (estimation_dataGridView.SelectedRows.Count != 0)
            {
                var row = estimation_dataGridView.SelectedRows[0];
                var elementId = row.Cells["ElementID"].Value;
                var elementName = row.Cells["Old Name"].Value;
                var hierarchyId = row.Cells["HierarchyId"].Value;
                var refId = row.Cells["RefId"].Value;

                lableToUpdate.Text = elementName.ToString();
                lableToUpdate.Visible = true;
            }

            editMeas_groupBox.Visible = true;
            UpdateEditMeasurementGroupBox();
        }

        private void newNameCell_Edit(object sender, DataGridViewCellEventArgs e)
        {
            if (_editProgrammatically)
            {
                return;
            }
            ElementAlreadyExist();

            var est = GetGridRowInfo();
            if (est == null) return;
            _newNamesToUpdateList.Add(est); 
        }


        private EstimateName GetGridRowInfo()
        {
            if (estimation_dataGridView.SelectedRows.Count == 0) return null;

            var cells = estimation_dataGridView.SelectedRows[0].Cells;
            var elementId = uint.Parse(cells["ElementID"].Value.ToString()) ;
            var elementName = cells["Old Name"].Value;
            var elementNewName = cells["New Estimated Name"].Value;
            var hierarchyId = uint.Parse(cells["HierarchyId"].Value.ToString());
            bool select = (bool) cells["Select"].Value;
            var refId = uint.Parse(cells["RefId"].Value.ToString());


            return new EstimateName()
            {
                PointId = elementId,
                OldName = elementName.ToString(),
                NewName = elementNewName.ToString(),
                HierarchyId = hierarchyId,
                ReferenceId = refId,
            };
        }

        // check if element already exist
        private void ElementAlreadyExist()
        {
            if (_newNamesToUpdateList.Count <= 0) return;
            var element = GetGridRowInfo();
            if (element ==null) return;
            var elementExist = _newNamesToUpdateList.Exists(x => x.PointId == element.PointId);
            if (!elementExist) return;

            // Not important more check!
            var elem = _newNamesToUpdateList.First(x => x.PointId == element.PointId);
            
            _newNamesToUpdateList.Remove(elem);
            
        }

        private void confirm_update_Click(object sender, EventArgs e)
        {
            var selectedListToUpdate = new List<EstimateName>();

            var gridRows = estimation_dataGridView.Rows.Cast<DataGridViewRow>();

            foreach (var dg in gridRows)
            {
                //var selected = Convert.ToBoolean(dg.Cells["Select"].Value);
                DataGridViewCheckBoxCell cellCheckBox = (DataGridViewCheckBoxCell) dg.Cells["Select"];
                if (!(bool) cellCheckBox.Value) continue;
                selectedListToUpdate.Add(new EstimateName()
                {
                    PointId = uint.Parse(dg.Cells["ElementID"].Value.ToString()),
                    OldName = dg.Cells["Old Name"].Value.ToString(),
                    NewName = dg.Cells["New Estimated Name"].Value.ToString(),
                    HierarchyId = uint.Parse(dg.Cells["HierarchyId"].Value.ToString()),
                    ReferenceId = uint.Parse(dg.Cells["RefId"].Value.ToString())
                });
            }

            foreach (var pt in selectedListToUpdate)
            {
                if (pt.NewName!=string.Empty)
                {
                    _standardPoint.UpdatePoint(pt);
                }
            }
        }

        private void autoEstimateCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (button5.Enabled == false)
            {
                button5.Enabled = true;
            }
        }

        private void selectAllButton_Click(object sender, EventArgs e)
        {
            
            RowSelectionHandler(true);
        }

        private void unselectAllButton_Click(object sender, EventArgs e)
        {
            RowSelectionHandler(false);
        }

        private void RowSelectionHandler(bool s)
        {
            estimation_dataGridView.ClearSelection();
            var gridRows = estimation_dataGridView.Rows.Cast<DataGridViewRow>();

            foreach (var dg in gridRows)
            {
                DataGridViewCheckBoxCell cellCheckBox = (DataGridViewCheckBoxCell) dg.Cells["Select"];
                cellCheckBox.Value = s;
            }
        }

      
    }
}
 


