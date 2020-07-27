using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ComTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            textBox1.Text = "HRPGzhchx.StartSun";
        }

        private void button1_Click(object sender, EventArgs e)
        {

            try
            {
                if (string.IsNullOrEmpty(textBox1.Text))
                {
                    MessageBox.Show("请输入内容");
                    return;
                }

                Type tmpType = null;
                if (radioButton1.Checked)
                    tmpType = Type.GetTypeFromProgID(textBox1.Text);
                if (radioButton2.Checked)
                    tmpType = Type.GetTypeFromCLSID(new Guid(textBox1.Text));
                if (tmpType == null)
                {
                    MessageBox.Show("未检测到COM组件");
                    return;
                }
                Object tmpClass = Activator.CreateInstance(tmpType);
                if (tmpClass == null)
                {
                    MessageBox.Show("未检测到COM组件");
                }
                else
                {
                    MessageBox.Show("已检测到COM组件");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("未检测到COM组件\r\n" + ex.Message);
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
                label1.Text = radioButton1.Text;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
                label1.Text = radioButton2.Text;
        }
    }
}
