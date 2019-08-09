using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace _A_Standard_Point___SCP
{
    static class Program
    {
        public static string[] args;

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args0)
        {
            args = args0;
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}

//private void TreeView_Check(TreeNode treeNode)
//        {
//            if (treeNode.Checked == true && treeNode.ImageIndex == 3)
//            {
//                uint PointId;

//                if (treeNode.Tag != null)
//                    PointId = Convert.ToUInt32(treeNode.Tag);
//                else
//                    PointId = Convert.ToUInt32(treeNode.Name);

//                SKF.RS.STB.Analyst.Point AnPoint = new SKF.RS.STB.Analyst.Point(AnalystDB, PointId);

//                float RPM = GetPointRPM(AnalystDB, PointId);
//                float FTF = Get_FTF(AnalystDB, PointId);
//                float BPFI = Get_Great_BPFI(AnalystDB, PointId);

//                if (AnPoint.PointType == PointType.Acc)
//                {
//                    // ACELERATION

//                    string acel_detection = null;
//                    string acel_savedate = null;
//                    string Detection = "Peak";
//                    string SaveData = "FFT and Time";
//                    uint acel_lines;
//                    uint acel_averages;
//                    float acel_startfreq;
//                    float acel_lowcutoff;
//                    float acel_stopfreq;

                    
//                    acel_lines = 3200;
//                    acel_averages = 4;
//                    acel_lowcutoff = 20;
//                    acel_startfreq = 0;
//                    acel_stopfreq = 7000;


//                    if (Detection == "Peak")
//                        acel_detection = "20500";
//                    else if (Detection == "Peak to Peak")
//                        acel_detection = "20501";
//                    else if (Detection == "RMS")
//                        acel_detection = "20502";

//                    if (SaveData == "FFT")
//                        acel_savedate = "20200";
//                    else if (SaveData == "Time")
//                        acel_savedate = "20201";
//                    else if (SaveData == "FFT and Time")
//                        acel_savedate = "20202";

//                    Reply_Aceleration(AnalystDB, PointId, acel_lines, acel_averages, acel_startfreq, acel_stopfreq, acel_lowcutoff, acel_savedate, acel_detection);
//                }

               

//                if (AnPoint.PointType == PointType.AccEnvelope)
//                {

//                    // ENVELOPE

//                    string env_detection = null;
//                    string env_savedate = null;
//                    string Detection = "Peak to Peak";
//                    string SaveData = "FFT and Time";
//                    uint env_lines = 2;
//                    uint env_averages;
//                    float env_lowcutoff;
//                    float env_startfreq = 0;
//                    float env_stopfreq;

//                    env_lines = Env_Lines(FTF, BPFI, RPM);
//                    env_averages = 2;
//                    env_stopfreq = (float)Math.Ceiling(Convert.ToDecimal(Env_Range(BPFI, RPM)));
//                    env_lowcutoff = (float)Math.Ceiling(Convert.ToDecimal(Env_LowCutOff(env_stopfreq, env_lines, RPM)));

//                    if (Detection == "Peak")
//                        env_detection = "20500";
//                    else if (Detection == "Peak to Peak")
//                        env_detection = "20501";
//                    else if (Detection == "RMS")
//                        env_detection = "20502";

//                    if (SaveData == "FFT")
//                        env_savedate = "20200";
//                    else if (SaveData == "Time")
//                        env_savedate = "20201";
//                    else if (SaveData == "FFT and Time")
//                        env_savedate = "20202";


//                    Reply_Envelope(AnalystDB, PointId, env_lines, env_averages, env_startfreq, env_stopfreq, env_lowcutoff, env_savedate, env_detection);
//                }
                
//                if (AnPoint.PointType == PointType.AccToVel)
//                {
//                    string vel_detection = null;
//                    string vel_savedate = null;
//                    string Detection = "RMS";
//                    string SaveData = "FFT and Time";
//                    uint vel_lines;
//                    uint vel_averages;
//                    float vel_startfreq;
//                    float vel_lowcutoff;
//                    float vel_stopfreq = 0;

//                    vel_lines = 1600;
//                    vel_averages = 3;
//                    vel_lowcutoff = (float)Math.Ceiling(Convert.ToDecimal(Vel_LowCutOff(BPFI, vel_lines, RPM)));
//                    vel_startfreq = 0;

//                    if (treeNode.Text.Substring(0, Math.Max(0, treeNode.Text.Length - 2)).ToUpper().Trim() != "MI 01" && treeNode.Text.Substring(0, Math.Max(0, treeNode.Text.Length - 2)).ToUpper().Trim() != "MI 02")
//                        vel_stopfreq = (float)Math.Ceiling(Convert.ToDecimal(Get_GMF(AnalystDB, PointId)));


//                    if (Detection == "Peak")
//                        vel_detection = "20500";
//                    else if (Detection == "Peak to Peak")
//                        vel_detection = "20501";
//                    else if (Detection == "RMS")
//                        vel_detection = "20502";

//                    if (SaveData == "FFT")
//                        vel_savedate = "20200";
//                    else if (SaveData == "Time")
//                        vel_savedate = "20201";
//                    else if (SaveData == "FFT and Time")
//                        vel_savedate = "20202";


//                    if (treeNode.Text.ToString() != "MI 01HV Speed" 
//                       && treeNode.Text.ToString() != "MI 02HV" 
//                       && treeNode.Text.ToString() != "MI 01HV"
//                       && treeNode.Text.Substring(Math.Max(0, treeNode.Text.Length - 4), 4).ToUpper().Trim() != "ZOOM")
//                        Reply_Velocity(AnalystDB, PointId, vel_lines, vel_averages, vel_startfreq, vel_stopfreq, vel_lowcutoff, vel_savedate, vel_detection);


//                    if (treeNode.Text.Substring(Math.Max(0, treeNode.Text.Length - 4), 4).ToUpper().Trim() == "ZOOM")
//                    {

//                        uint velz_lines;
//                        uint velz_averages;
//                        float velz_stopfreq;

//                        velz_stopfreq = 400;
//                        velz_lines = 1600;
//                        velz_averages = 2;

//                        Reply_Velocity_ZOOM(AnalystDB, PointId, velz_lines, velz_averages, velz_stopfreq);
//                    }
//                }
//                treeView_AnStandards.Invoke(new MethodInvoker(delegate
//               {
//                   progressBar_Points.Value = progressBar_Points.Value + 1;
//                   label_Machine_Left.Text = progressBar_Points.Value + "/" + progressBar_Points.Maximum.ToString();
//               }));
//            }

//            foreach (TreeNode tn in treeNode.Nodes)
//            {
//                TreeView_Check(tn);
//            }
//        }


        //public void Tree_Analyst_3(AnalystConnection cn, TreeView Tree)
        //{
        //    if (cn.IsConnected)
        //    {
        //        Hashtable Nodes_Tree = new Hashtable();

        //        Tree.Invoke(new MethodInvoker(delegate
        //        {
        //            Log("@A Stand - Conectado");
        //            Nodes_Tree.Clear();

        //            Nodes_Tree.Add("00321", Tree.Nodes.Add("00322", "Error TAGS"));

        //            Tree.Enabled = true;
        //            Tree.ImageList = imageList1;
        //        }));

        //        DataTable AnTblHierarchy;
        //        string TBLSETID = "";

        //        if (txtTBLSETID.Text != "")
        //            TBLSETID = " AND TBLSETID= " + txtTBLSETID.Text + " ";

        //        AnTblHierarchy = cn.DataTable("*", "TREEELEM", "REFERENCEID=0 and HIERARCHYTYPE=1 " + TBLSETID + " and Parentid!=2147000000 ORDER BY tblsetid, containertype, parentid, branchlevel, SLOTNUMBER");
               
        //        // DataTable AnTblHierarchy = cn.DataTable("*", "TREEELEM", "REFERENCEID=0 and HIERARCHYTYPE=1 and Parentid!=2147000000 ORDER BY tblsetid, containertype, parentid, branchlevel, SLOTNUMBER");

        //        Log("@A Stand - Get Hierarchy List");

        //        TreeNode NewNode = null;
        //        TreeNode ParentNode = null;
        //        Log("Monta Arvore");

        //        pbHierarchy.Invoke(new MethodInvoker(delegate
        //        {
        //            pbHierarchy.Value = 0;
        //            pbHierarchy.Maximum = AnTblHierarchy.Rows.Count;
        //        }));

        //        lblPbHierarchy.Invoke(new MethodInvoker(delegate { lblPbHierarchy.Text = "0/" + AnTblHierarchy.Rows.Count; }));

        //        SKF.RS.STB.Analyst.Point AnPoint;

        //        for (Int32 i = 0; i < AnTblHierarchy.Rows.Count; i++)
        //        {
        //            Log(" Inicio do processo");
        //            if (Convert.ToInt64(AnTblHierarchy.Rows[i]["Parentid"]) == 0)
        //            {
        //                Log("@A Stand - Add Hierarchy: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString());
        //                Tree.Invoke(new MethodInvoker(delegate
        //                {
        //                    Nodes_Tree.Add(Convert.ToInt64(AnTblHierarchy.Rows[i]["TreeelemId"])
        //                        , Tree.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString()
        //                        , AnTblHierarchy.Rows[i]["Name"].ToString()));

        //                    Log("Arvore Add: " + AnTblHierarchy.Rows[i]["Name"].ToString());
        //                }));


        //            }
        //            else
        //            {
        //                NewNode = null;
        //                ParentNode = (TreeNode)Nodes_Tree[Convert.ToInt64(AnTblHierarchy.Rows[i]["Parentid"])];


        //                if ((Convert.ToInt64(AnTblHierarchy.Rows[i]["ContainerType"])) == 4)
        //                {
        //                    Log(" Ponto ");
        //                    if (Convert.ToUInt32(AnTblHierarchy.Rows[i]["ReferenceID"]) == 0)
        //                    {
        //                        AnPoint = new SKF.RS.STB.Analyst.Point(cn, Convert.ToUInt32(AnTblHierarchy.Rows[i]["TreeElemId"]));
        //                        Log("Treeelem Point");
        //                    }
        //                    else
        //                    {
        //                        AnPoint = new SKF.RS.STB.Analyst.Point(cn, Convert.ToUInt32(AnTblHierarchy.Rows[i]["ReferenceID"]));
        //                        Log("Referenced Point");
        //                    }
        //                    float RPM = float.Parse(AnTblHierarchy.Rows[i]["RPM"].ToString()); //GetPointRPM(cn, Convert.ToUInt32(AnTblHierarchy.Rows[i]["TreeElemId"]));
        //                    float tmp1 = float.Parse(Acel_GTRPM.Text);

        //                    if (RPM > (tmp1 / 60))
        //                    {
        //                        Log("RPM Dentro");
        //                        Tree.Invoke(new MethodInvoker(delegate
        //                        {
        //                            try
        //                            {

        //                                Log("@A Stand - Add Point: " + AnPoint.Name + " - " + AnPoint.TreeElemId);
        //                                NewNode = ParentNode.Nodes.Add(AnPoint.TreeElemId.ToString(), AnPoint.Name.ToString(), 3, 3);
        //                                Log(" Tag: " + AnTblHierarchy.Rows[i]["ReferenceId"].ToString());
        //                                if (AnTblHierarchy.Rows[i]["ReferenceId"].ToString() != "0")
        //                                    NewNode.Tag = AnTblHierarchy.Rows[i]["ReferenceId"].ToString();

        //                            }
        //                            catch (Exception ex)
        //                            {
        //                                //if (ParentNode == null)
        //                                //    ParentNode = (TreeNode)Nodes_Tree[1];

        //                                //NewNode = ParentNode.Nodes.Add(AnPoint.TreeElemId.ToString(), "ParentId: " + AnTblHierarchy.Rows[i]["ParentId"] + " - " + AnPoint.Name.ToString(), 3, 3);
        //                                //Log(" Tag: " + AnTblHierarchy.Rows[i]["ReferenceId"].ToString());
        //                                //if (AnTblHierarchy.Rows[i]["ReferenceId"].ToString() != "0")
        //                                //    NewNode.Tag = AnTblHierarchy.Rows[i]["ReferenceId"].ToString();

        //                                Log("ELEMENTO DE REFERÊNCIA NÃO ENCONTRADO: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " " + AnPoint.Name.ToString() + " - Exceção: " + ex.ToString());
        //                            }
        //                        }));
        //                    }
        //                    else if (RPM == 0 && chkExZero.Checked == true)
        //                    {
        //                        Tree.Invoke(new MethodInvoker(delegate
        //                        {
        //                            try
        //                            {
        //                                NewNode = ParentNode.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString(), AnTblHierarchy.Rows[i]["Name"].ToString(), 3, 3);
        //                                Log(" Tag: " + AnTblHierarchy.Rows[i]["ReferenceId"].ToString());
        //                                if (AnTblHierarchy.Rows[i]["ReferenceId"].ToString() != "0")
        //                                    NewNode.Tag = AnTblHierarchy.Rows[i]["ReferenceId"].ToString();

        //                                Log("@A Stand - Add Point RPM=0 " + AnTblHierarchy.Rows[i]["TreeElemId"].ToString());
        //                            }
        //                            catch (Exception ex)
        //                            {
        //                              //  if (ParentNode == null)
        //                             //       ParentNode = (TreeNode)Nodes_Tree[1];

        //                              //  NewNode = ParentNode.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString(), "Parentid: " + AnTblHierarchy.Rows[i]["ParentId"].ToString() + " - " + AnTblHierarchy.Rows[i]["Name"].ToString(), 3, 3);
        //                                //Log(" Tag: " + AnTblHierarchy.Rows[i]["ReferenceId"].ToString());
        //                                //if (AnTblHierarchy.Rows[i]["ReferenceId"].ToString() != "0")
        //                                //    NewNode.Tag = AnTblHierarchy.Rows[i]["ReferenceId"].ToString();

        //                                Log("ELEMENTO DE REFERÊNCIA NÃO ENCONTRADO: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " " + AnPoint.Name.ToString() + " - Exceção: " + ex.ToString());
        //                            }
        //                        }));
        //                    }
        //                }
        //                else if ((Convert.ToInt64(AnTblHierarchy.Rows[i]["ContainerType"])) == 3)
        //                {
        //                    Tree.Invoke(new MethodInvoker(delegate
        //                    {
        //                        try
        //                        {
        //                            Log("@A Stand - Add Machine: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString());
        //                            NewNode = ParentNode.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString(), AnTblHierarchy.Rows[i]["Name"].ToString(), 2, 2);
        //                            Log(" Machine Added: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString());
        //                        }
        //                        catch (Exception ex)
        //                        {
        //                            //if (ParentNode == null)
        //                            //    ParentNode = (TreeNode)Nodes_Tree[1];

        //                            //NewNode = ParentNode.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString(), "Parentid: " + AnTblHierarchy.Rows[i]["ParentId"].ToString() + " - " + AnTblHierarchy.Rows[i]["Name"].ToString(), 2, 2);
        //                            Log("ELEMENTO DE REFERÊNCIA NÃO ENCONTRADO: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " - Exceção: " + ex.ToString());
        //                        }
        //                    }));
        //                }

        //                else
        //                {
        //                    Tree.Invoke(new MethodInvoker(delegate
        //                     {
        //                         try
        //                         {
        //                             Log("@A Stand - Add SET: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " - " + AnTblHierarchy.Rows[i]["Name"].ToString());

        //                             NewNode = ParentNode.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString(), AnTblHierarchy.Rows[i]["Name"].ToString(), 1, 1);
        //                             NewNode.StateImageIndex = 0;

        //                             Log(" SET Added: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString());

        //                         }
        //                         catch (Exception ex)
        //                         {
        //                             if (ParentNode == null)
        //                                 ParentNode = (TreeNode)Nodes_Tree["00321"];

        //                             Log("ELEMENTO DE REFERÊNCIA NÃO ENCONTRADO: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " - Exceção: " + ex.ToString());
                                     
        //                             NewNode = ParentNode.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString(), "Parentid: " + AnTblHierarchy.Rows[i]["ParentId"].ToString() + " - " + AnTblHierarchy.Rows[i]["Name"].ToString(), 1, 1);
        //                             NewNode.StateImageIndex = 0;

        //                             ParentNode = null;
        //                         }
        //                     }));

        //                }
        //                if (NewNode != null)
        //                {
        //                    Tree.Invoke(new MethodInvoker(delegate
        //                    {
        //                        try
        //                        {
        //                            Log("@A Stand - Add To Tree: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " - " + AnTblHierarchy.Rows[i]["Name"].ToString());
        //                            Nodes_Tree.Add(Convert.ToInt64(AnTblHierarchy.Rows[i]["TreeelemId"]), NewNode);
        //                            Log(" Added: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString());
        //                        }
        //                        catch (Exception ex)
        //                        {
        //                            Log("ELEMENTO DE REFERÊNCIA NÃO ENCONTRADO: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " - Exceção: " + ex.ToString());
        //                        }
        //                    }));
        //                }
        //            }

        //            pbHierarchy.Invoke(new MethodInvoker(delegate { pbHierarchy.Value = i; }));
        //            lblPbHierarchy.Invoke(new MethodInvoker(delegate { lblPbHierarchy.Text = i + "/" + AnTblHierarchy.Rows.Count; }));
        //        }

        //        AnTblHierarchy.Dispose();

        //        pbHierarchy.Invoke(new MethodInvoker(delegate { pbHierarchy.Value = pbHierarchy.Maximum; }));
        //        lblPbHierarchy.Invoke(new MethodInvoker(delegate { lblPbHierarchy.Text = pbHierarchy.Maximum + "/" + AnTblHierarchy.Rows.Count; }));
        //        Tree.Invoke(new MethodInvoker(delegate { Tree.Enabled = true; }));
        //        Tree.Invoke(new MethodInvoker(delegate { Tree.SelectedNode = Tree.TopNode; }));

                
        //    }
        //}

//public void Tree_Analyst_WorkSpace(AnalystConnection cn, TreeView Tree, string WSFilter)
//{
//    if (cn.IsConnected)
//    {
//        Log("@A Stand - Conectado");
//        Hashtable Nodes_Tree = new Hashtable();
//        Tree.Invoke(new MethodInvoker(delegate
//                   {
//                       Nodes_Tree.Clear();

//                       Tree.Enabled = true;
//                       Tree.ImageList = imageList1;
//                   }));
//        TreeNode NewNode;
//        TreeNode ParentNode;
//        DataTable AnTblHierarchy;

//        if (WSFilter == "")
//            AnTblHierarchy = cn.DataTable("*", "TREEELEM", "HIERARCHYTYPE=3 and Parentid!=2147000000 and NAME!='Analise pendente' ORDER BY tblsetid, parentid, containertype, branchlevel, SLOTNUMBER");
//        else
//            AnTblHierarchy = cn.DataTable("*", "TREEELEM", "HIERARCHYTYPE=3 and Parentid!=2147000000 and NAME!='Analise pendente' and hierarchyid in (SELECT TREEELEMID FROM TREEELEM WHERE HIERARCHYTYPE=3 and Parentid!=2147000000 and NAME LIKE '" + WSFilter + "') or NAME LIKE '" + WSFilter + "' ORDER BY tblsetid, parentid, containertype, branchlevel, SLOTNUMBER");

//        Log("@A Stand - Get Hierarchy List");
//        pbHierarchy.Invoke(new MethodInvoker(delegate
//        {
//            pbHierarchy.Value = 0;
//            pbHierarchy.Maximum = AnTblHierarchy.Rows.Count;
//        }));

//        lblPbHierarchy.Invoke(new MethodInvoker(delegate { lblPbHierarchy.Text = "0/" + AnTblHierarchy.Rows.Count; }));

//        SKF.RS.STB.Analyst.Point AnPoint;

//        for (Int32 i = 0; i < AnTblHierarchy.Rows.Count; i++)
//        {

//            if (Convert.ToInt32(AnTblHierarchy.Rows[i]["Parentid"]) == 0)
//            {
//                Log("@A Stand - Add Hierarchy: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString());
//                Tree.Invoke(new MethodInvoker(delegate
//                   {
//                Nodes_Tree.Add(Convert.ToInt32(AnTblHierarchy.Rows[i]["TreeelemId"])
//                    , Tree.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString()
//                    , AnTblHierarchy.Rows[i]["Name"].ToString()));
//                   }));
//            }
//            else
//            {
//                NewNode = null;
//                ParentNode = (TreeNode)Nodes_Tree[Convert.ToInt32(AnTblHierarchy.Rows[i]["Parentid"])];
//                if ((Convert.ToInt32(AnTblHierarchy.Rows[i]["ContainerType"])) == 4)
//                {
//                    AnPoint = new SKF.RS.STB.Analyst.Point(cn, Convert.ToUInt32(AnTblHierarchy.Rows[i]["ReferenceId"]));

//                    float RPM = GetPointRPM(cn, Convert.ToUInt32(AnTblHierarchy.Rows[i]["ReferenceId"]));
//                    float tmp1 = float.Parse(Acel_GTRPM.Text);

//                    if (RPM > tmp1)
//                    {
//                        if (AnPoint.PointType == PointType.AccEnvelope || AnPoint.PointType == PointType.Acc || AnPoint.PointType == PointType.AccToVel)
//                        {
//                            Tree.Invoke(new MethodInvoker(delegate
//                            {
//                                try
//                                {
//                                    Log("@A Stand - Add Point: " + AnPoint.Name + " - " + AnPoint.TreeElemId);
//                                    NewNode = ParentNode.Nodes.Add(AnPoint.TreeElemId.ToString(), AnPoint.Name.ToString(), 3, 3);
//                                    NewNode.Tag = AnTblHierarchy.Rows[i]["ReferenceId"].ToString();

//                                }
//                                catch (Exception ex)
//                                {
//                                    if (ParentNode == null)
//                                        ParentNode = (TreeNode)Nodes_Tree[1];

//                                    NewNode = ParentNode.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString(), "Parentid: " + AnTblHierarchy.Rows[i]["ParentId"].ToString() + " - " + AnTblHierarchy.Rows[i]["Name"].ToString(), 3, 3);
//                                    NewNode.Tag = AnTblHierarchy.Rows[i]["ReferenceId"].ToString();
//                                    Log("ELEMENTO DE REFERÊNCIA NÃO ENCONTRADO: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " " + AnPoint.Name.ToString() + " - Exceção: " + ex.ToString());

//                                }
//                            }));
//                        }
//                    }
//                    else if (RPM == 0 && chkExZero.Checked == true)
//                    {
//                        if (AnPoint.PointType == PointType.AccEnvelope || AnPoint.PointType == PointType.Acc || AnPoint.PointType == PointType.AccToVel)
//                        {
//                        Tree.Invoke(new MethodInvoker(delegate
//                            {
//                                try 
//                                {
//                                    NewNode = ParentNode.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString(), AnTblHierarchy.Rows[i]["Name"].ToString(), 3, 3);
//                                    NewNode.Tag = AnTblHierarchy.Rows[i]["ReferenceId"].ToString();
//                                    Log("@A Stand - Add Point RPM=0 " + AnTblHierarchy.Rows[i]["ReferenceId"].ToString());
//                                }
//                                catch (Exception ex)
//                                {
//                                    if (ParentNode == null)
//                                        ParentNode = (TreeNode)Nodes_Tree[1];

//                                    NewNode = ParentNode.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString(), "Parentid: " + AnTblHierarchy.Rows[i]["ParentId"].ToString() + " - " + AnTblHierarchy.Rows[i]["Name"].ToString(), 3, 3);
//                                    NewNode.Tag = AnTblHierarchy.Rows[i]["ReferenceId"].ToString();
//                                    Log("ELEMENTO DE REFERÊNCIA NÃO ENCONTRADO: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " " + AnPoint.Name.ToString() + " - Exceção: " + ex.ToString());
//                                }
//                            }));
//                        }
//                    }
//                }
//                else if ((Convert.ToInt32(AnTblHierarchy.Rows[i]["ContainerType"])) == 3)
//                {
//                    Tree.Invoke(new MethodInvoker(delegate
//                        {
//                            try
//                            {
//                                Log("@A Stand - Add Machine: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString());
//                                NewNode = ParentNode.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString(), AnTblHierarchy.Rows[i]["Name"].ToString(), 2, 2);
//                            }
//                            catch (Exception ex)
//                            {
//                                if (ParentNode == null)
//                                    ParentNode = (TreeNode)Nodes_Tree[1];

//                                NewNode = ParentNode.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString(), "Parentid: " + AnTblHierarchy.Rows[i]["ParentId"].ToString() + " - " + AnTblHierarchy.Rows[i]["Name"].ToString(), 2, 2);
//                                Log("ELEMENTO DE REFERÊNCIA NÃO ENCONTRADO: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " - Exceção: " + ex.ToString());
//                            }
//                    }));
//                }

//                else
//                {
//                Tree.Invoke(new MethodInvoker(delegate
//                    {
//                        try
//                        {
//                            Log("@A Stand - Add SET: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " - " + AnTblHierarchy.Rows[i]["Name"].ToString());
//                            NewNode = ParentNode.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString(), AnTblHierarchy.Rows[i]["Name"].ToString(), 1, 1);
//                            NewNode.StateImageIndex = 0;
//                        }
//                        catch (Exception ex)
//                        {
//                            if (ParentNode == null)
//                                ParentNode = (TreeNode)Nodes_Tree[1];

//                            NewNode = ParentNode.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString(), "Parentid: " + AnTblHierarchy.Rows[i]["ParentId"].ToString() + " - " + AnTblHierarchy.Rows[i]["Name"].ToString(), 1, 1);
//                            NewNode.StateImageIndex = 0;
//                            Log("ELEMENTO DE REFERÊNCIA NÃO ENCONTRADO: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " - Exceção: " + ex.ToString());
//                        }
//                    }));
//                }
//                if (NewNode != null)
//                {
//                Tree.Invoke(new MethodInvoker(delegate
//                    {
//                        try
//                        {
//                            Log("@A Stand - Add To Tree: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " - " + AnTblHierarchy.Rows[i]["Name"].ToString());
//                            Nodes_Tree.Add(Convert.ToInt32(AnTblHierarchy.Rows[i]["TreeelemId"]), NewNode);
//                        }
//                        catch (Exception ex)
//                        {
//                            Log("ELEMENTO DE REFERÊNCIA NÃO ENCONTRADO: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " - Exceção: " + ex.ToString() );
//                        }
//                    }));
//                }
//            }

//            pbHierarchy.Invoke(new MethodInvoker(delegate { pbHierarchy.Value = i; }));
//            lblPbHierarchy.Invoke(new MethodInvoker(delegate { lblPbHierarchy.Text = i + "/" + AnTblHierarchy.Rows.Count; }));
//        }
//        AnTblHierarchy.Dispose();

//        pbHierarchy.Invoke(new MethodInvoker(delegate { pbHierarchy.Value = pbHierarchy.Maximum; }));
//        lblPbHierarchy.Invoke(new MethodInvoker(delegate { lblPbHierarchy.Text = pbHierarchy.Maximum + "/" + AnTblHierarchy.Rows.Count; }));
//        Tree.Invoke(new MethodInvoker(delegate { Tree.Enabled = true; }));
//        Tree.Invoke(new MethodInvoker(delegate { Tree.SelectedNode = Tree.TopNode; }));
//    }
//}




//public void Tree_Analyst(AnalystConnection cn, TreeView Tree)
//{
//    if (cn.IsConnected)
//    {
//        Hashtable Nodes_Tree = new Hashtable();

//        Tree.Invoke(new MethodInvoker(delegate
//        {
//            Log("@A Stand - Conectado");
//            Nodes_Tree.Clear();

//            Tree.Enabled = true;
//            Tree.ImageList = imageList1;
//        }));

//        TreeNode NewNode;
//        TreeNode ParentNode;

//        string Acc = Registration.RegistrationId(AnalystDB, "SKFCM_ASPT_Acc").ToString();
//        string Vel = Registration.RegistrationId(AnalystDB, "SKFCM_ASPT_AccToVel").ToString();
//        string Env = Registration.RegistrationId(AnalystDB, "SKFCM_ASPT_AccEnvelope").ToString();
//        string DAD = Registration.RegistrationId(AnalystDB, "SKFCM_ASDD_MicrologDAD").ToString();

//        StringBuilder SQL = new StringBuilder();
//        SQL.Append("   select ");
//        SQL.Append("   	TE1.*");
//        SQL.Append("   	, PT1.VALUESTRING as RPM ");
//        SQL.Append("   from  ");
//        SQL.Append("   	treeelem TE1 ");
//        SQL.Append("   	, TABLESET TB ");
//        SQL.Append("   	, TREEELEM TE2 ");
//        SQL.Append("   	, TREEELEM TE3 ");
//        SQL.Append("   	, POINT PT ");
//        SQL.Append("   	, POINT PT1 ");
//        SQL.Append("   	, POINT PT2 ");
//        SQL.Append("   	, REGISTRATION RG ");
//        SQL.Append("   	, REGISTRATION RG1 ");
//        SQL.Append("   where  ");
//        SQL.Append("   	TE1.CONTAINERTYPE=4 ");
//        SQL.Append("   	and TE1.parentid!=2147000000 ");
//        SQL.Append("   	and TE1.HIERARCHYTYPE=1 ");
//        SQL.Append("   	and TE1.ELEMENTENABLE=1 ");
//        SQL.Append("   	AND TE1.TBLSETID = TB.TBLSETID ");
//        SQL.Append("   	AND TE2.TREEELEMID=TE1.PARENTID ");
//        SQL.Append("   	AND TE3.TREEELEMID = TE2.PARENTID ");
//        SQL.Append("  	AND PT.ELEMENTID = TE1.TREEELEMID ");
//        SQL.Append("   	AND PT.VALUESTRING IN ('" + Acc + "', '" + Vel + "', '" + Env + "') ");
//        SQL.Append("   	AND PT1.ELEMENTID = PT.ELEMENTID ");
//        SQL.Append("   	AND RG.SIGNATURE = 'SKFCM_ASPF_Speed' ");
//        SQL.Append("   	AND PT1.FIELDID	= RG.REGISTRATIONID ");
//        SQL.Append("   	AND PT2.VALUESTRING ='" + DAD + "' ");
//        SQL.Append("   	AND PT2.ELEMENTID = PT1.ELEMENTID ");
//        SQL.Append("   	AND RG1.SIGNATURE = 'SKFCM_ASPF_Dad_Id' ");
//        SQL.Append("   	AND PT2.FIELDID	= RG1.REGISTRATIONID ");
//        SQL.Append("   order by  ");
//        SQL.Append("   	  TE1.TBLSETID");
//        SQL.Append("   	  ,TE1.containertype");
//        SQL.Append("   	  ,TE1.parentid");
//        SQL.Append("   	  ,TE1.branchlevel");
//        SQL.Append("   	  ,TE1.SLOTNUMBER");

//        DataTable AnTblHierarchy = cn.DataTable(SQL.ToString());

//        // DataTable AnTblHierarchy = cn.DataTable("*", "TREEELEM", "REFERENCEID=0 and HIERARCHYTYPE=1 and Parentid!=2147000000 ORDER BY tblsetid, containertype, parentid, branchlevel, SLOTNUMBER");
//        //DataTable AnTblHierarchy = cn.DataTable("*", "TREEELEM", "TREEELEMID IN (1,2,3,4,5,6,7,8,9,10) AND REFERENCEID=0 and HIERARCHYTYPE=1 and Parentid!=2147000000 ORDER BY tblsetid, containertype, parentid, branchlevel, SLOTNUMBER");

//        Log("@A Stand - Get Hierarchy List");
//        pbHierarchy.Invoke(new MethodInvoker(delegate
//        {
//            pbHierarchy.Value = 0;
//            pbHierarchy.Maximum = AnTblHierarchy.Rows.Count;
//        }));

//        lblPbHierarchy.Invoke(new MethodInvoker(delegate { lblPbHierarchy.Text = "0/" + AnTblHierarchy.Rows.Count; }));

//        SKF.RS.STB.Analyst.Point AnPoint;

//        for (Int32 i = 0; i < AnTblHierarchy.Rows.Count; i++)
//        {

//            if (Convert.ToInt32(AnTblHierarchy.Rows[i]["Parentid"]) == 0)
//            {
//                Log("@A Stand - Add Hierarchy: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString());
//                Tree.Invoke(new MethodInvoker(delegate
//                {
//                    Nodes_Tree.Add(Convert.ToInt32(AnTblHierarchy.Rows[i]["TreeelemId"])
//                        , Tree.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString()
//                        , AnTblHierarchy.Rows[i]["Name"].ToString()));
//                }));
//            }
//            else
//            {
//                NewNode = null;
//                ParentNode = (TreeNode)Nodes_Tree[Convert.ToInt32(AnTblHierarchy.Rows[i]["Parentid"])];


//                if ((Convert.ToInt32(AnTblHierarchy.Rows[i]["ContainerType"])) == 4)
//                {


//                    AnPoint = new SKF.RS.STB.Analyst.Point(cn, Convert.ToUInt32(AnTblHierarchy.Rows[i]["TreeElemId"]));

//                    float RPM = float.Parse(AnTblHierarchy.Rows[i]["RPM"].ToString()); //GetPointRPM(cn, Convert.ToUInt32(AnTblHierarchy.Rows[i]["TreeElemId"]));
//                    float tmp1 = float.Parse(Acel_GTRPM.Text);

//                    if (RPM > (tmp1 / 60))
//                    {
//                        if (AnPoint.PointType == PointType.AccEnvelope || AnPoint.PointType == PointType.Acc || AnPoint.PointType == PointType.AccToVel)
//                        {
//                            Tree.Invoke(new MethodInvoker(delegate
//                            {
//                                try
//                                {

//                                    Log("@A Stand - Add Point: " + AnPoint.Name + " - " + AnPoint.TreeElemId);
//                                    NewNode = ParentNode.Nodes.Add(AnPoint.TreeElemId.ToString(), AnPoint.Name.ToString(), 3, 3);

//                                }
//                                catch (Exception ex)
//                                {
//                                    if (ParentNode == null)
//                                        ParentNode = (TreeNode)Nodes_Tree[1];

//                                    NewNode = ParentNode.Nodes.Add(AnPoint.TreeElemId.ToString(), "ParentId: " + AnTblHierarchy.Rows[i]["ParentId"] + " - " + AnPoint.Name.ToString(), 3, 3);
//                                    Log("ELEMENTO DE REFERÊNCIA NÃO ENCONTRADO: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " " + AnPoint.Name.ToString() + " - Exceção: " + ex.ToString());
//                                }
//                            }));
//                        }
//                    }
//                    else if (RPM == 0 && chkExZero.Checked == true)
//                    {
//                        if (AnPoint.PointType == PointType.AccEnvelope || AnPoint.PointType == PointType.Acc || AnPoint.PointType == PointType.AccToVel)
//                        {
//                            Tree.Invoke(new MethodInvoker(delegate
//                            {
//                                try
//                                {
//                                    NewNode = ParentNode.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString(), AnTblHierarchy.Rows[i]["Name"].ToString(), 3, 3);
//                                    Log("@A Stand - Add Point RPM=0 " + AnTblHierarchy.Rows[i]["TreeElemId"].ToString());
//                                }
//                                catch (Exception ex)
//                                {
//                                    if (ParentNode == null)
//                                        ParentNode = (TreeNode)Nodes_Tree[1];

//                                    NewNode = ParentNode.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString(), "Parentid: " + AnTblHierarchy.Rows[i]["ParentId"].ToString() + " - " + AnTblHierarchy.Rows[i]["Name"].ToString(), 3, 3);
//                                    Log("ELEMENTO DE REFERÊNCIA NÃO ENCONTRADO: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " " + AnPoint.Name.ToString() + " - Exceção: " + ex.ToString());
//                                }
//                            }));
//                        }
//                    }
//                }
//                else if ((Convert.ToInt32(AnTblHierarchy.Rows[i]["ContainerType"])) == 3)
//                {
//                    Tree.Invoke(new MethodInvoker(delegate
//                    {
//                        try
//                        {
//                            Log("@A Stand - Add Machine: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString());
//                            NewNode = ParentNode.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString(), AnTblHierarchy.Rows[i]["Name"].ToString(), 2, 2);

//                        }
//                        catch (Exception ex)
//                        {
//                            if (ParentNode == null)
//                                ParentNode = (TreeNode)Nodes_Tree[1];

//                            NewNode = ParentNode.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString(), "Parentid: " + AnTblHierarchy.Rows[i]["ParentId"].ToString() + " - " + AnTblHierarchy.Rows[i]["Name"].ToString(), 2, 2);
//                            Log("ELEMENTO DE REFERÊNCIA NÃO ENCONTRADO: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " - Exceção: " + ex.ToString());
//                        }
//                    }));
//                }

//                else
//                {
//                    Tree.Invoke(new MethodInvoker(delegate
//                    {
//                        try
//                        {
//                            Log("@A Stand - Add SET: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " - " + AnTblHierarchy.Rows[i]["Name"].ToString());

//                            NewNode = ParentNode.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString(), AnTblHierarchy.Rows[i]["Name"].ToString(), 1, 1);
//                            NewNode.StateImageIndex = 0;

//                        }
//                        catch (Exception ex)
//                        {
//                            if (ParentNode == null)
//                                ParentNode = (TreeNode)Nodes_Tree[1];

//                            NewNode = ParentNode.Nodes.Add(AnTblHierarchy.Rows[i]["TreeelemId"].ToString(), "Parentid: " + AnTblHierarchy.Rows[i]["ParentId"].ToString() + " - " + AnTblHierarchy.Rows[i]["Name"].ToString(), 1, 1);
//                            NewNode.StateImageIndex = 0;
//                            Log("ELEMENTO DE REFERÊNCIA NÃO ENCONTRADO: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " - Exceção: " + ex.ToString());
//                        }
//                    }));
//                }
//                if (NewNode != null)
//                {
//                    Tree.Invoke(new MethodInvoker(delegate
//                    {
//                        try
//                        {
//                            Log("@A Stand - Add To Tree: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " - " + AnTblHierarchy.Rows[i]["Name"].ToString());

//                            Nodes_Tree.Add(Convert.ToInt32(AnTblHierarchy.Rows[i]["TreeelemId"]), NewNode);

//                        }
//                        catch (Exception ex)
//                        {
//                            Log("ELEMENTO DE REFERÊNCIA NÃO ENCONTRADO: " + AnTblHierarchy.Rows[i]["TreeelemId"].ToString() + " - Exceção: " + ex.ToString());
//                        }
//                    }));
//                }
//            }

//            pbHierarchy.Invoke(new MethodInvoker(delegate { pbHierarchy.Value = i; }));
//            lblPbHierarchy.Invoke(new MethodInvoker(delegate { lblPbHierarchy.Text = i + "/" + AnTblHierarchy.Rows.Count; }));
//        }
//        AnTblHierarchy.Dispose();

//        pbHierarchy.Invoke(new MethodInvoker(delegate { pbHierarchy.Value = pbHierarchy.Maximum; }));
//        lblPbHierarchy.Invoke(new MethodInvoker(delegate { lblPbHierarchy.Text = pbHierarchy.Maximum + "/" + AnTblHierarchy.Rows.Count; }));
//        Tree.Invoke(new MethodInvoker(delegate { Tree.Enabled = true; }));
//        Tree.Invoke(new MethodInvoker(delegate { Tree.SelectedNode = Tree.TopNode; }));
//    }
//}


