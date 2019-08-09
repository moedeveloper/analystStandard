using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using SKF.RS.AddOns.STB.Analyst;
using SKF.RS.AddOns.STB.DB;
using SKF.RS.AddOns.STB.Generic;

namespace SKF.RS.AddOns.STB.Analyst
{
    public partial class AnUserLogon : Form
    {
        public AnalystConnection AnalystConnection;

        private bool _IsAuthenticated = false;
        /// <summary>
        /// Returns if user was authenticated into @Analyst Database
        /// </summary>
        public bool IsAuthenticated { get { return _IsAuthenticated; } }

        private Int32 _UserId = 0;
        /// <summary>
        /// Returns user id if authenticated
        /// </summary>
        public Int32 UserId { get { return _UserId; } }
        
        public AnUserLogon()
        {
            InitializeComponent();
        }

        private string button_OK_FirstPass;
        private bool button_OK_FirstClick = true;
        private void button_OK_Click(object sender, EventArgs e)
        {
            bool Authenticated = AnalystConnection.Login(textBox_UserName.Text, textBox_UserPass.Text);

            if (Authenticated)
            {
                label_InfoPanel.Text = "Conectando ao SKF @ptitude Analyst...";
                this.Update();
/*
                Process AnalystExe = new Process();
                AnalystExe.StartInfo.FileName = AppFileName;
                AnalystExe.StartInfo.Arguments = AppParams;
                AnalystExe.Start();

                INIFile.Write(GenericTools.WindowsGetUserName(), "LastProcessId", AnalystExe.Id);
                */
                this.Close();
            }
            else
                if (AnalystConnection.SQLtoString("Passwd", "UserTbl", "LoginName='" + textBox_UserName.Text + "'") == AnalystConnection.NoPassword)
                    if (button_OK_FirstClick)
                    {
                        button_OK_FirstClick = false;
                        textBox_UserName.Enabled = false;
                        button_OK_FirstPass = textBox_UserPass.Text;
                        textBox_UserPass.Text = string.Empty;
                        label_InfoPanel.Text = "Configurando nova senha para usuário." + Environment.NewLine + "Repita a nova senha, por favor.";
                    }
                    else
                    {
                        button_OK_FirstClick = true;
                        textBox_UserName.Enabled = true;
                        label_InfoPanel.Text = "Digite suas informações de login.";
                        if (button_OK_FirstPass != textBox_UserPass.Text)
                            MessageBox.Show("As senhas digitadas não conferem.");
                        else
                        {
                            AnalystConnection.SQLUpdate("UserTbl", "Passwd", GenericTools.PassEncode(textBox_UserPass.Text), "LoginName='" + textBox_UserName.Text + "'");
                            button_OK_Click(sender, e);
                        }
                    }
                else
                    label_InfoPanel.Text = "Usuário ou senha da aplicação incorretos." + Environment.NewLine + "Verifique seus dados de login e tente novamente.";
            this.Update();
        }
    }
}
