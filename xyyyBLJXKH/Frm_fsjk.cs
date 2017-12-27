using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using dbbase;
using System.Diagnostics;

namespace pathnetfsjk
{
    public partial class Frm_fsjk : Form
    {
        public Frm_fsjk()
        {
            InitializeComponent();
            this.Hide();
        }
        dbbase.odbcdb aa = new odbcdb("DSN=pathnet;UID=pathnet;PWD=4s3c2a1p", "", "");
        IniFiles f = new IniFiles(Application.StartupPath + "\\sz.ini");
        int fscs =5;
      
        private void Form1_Load(object sender, EventArgs e)
        {
            this.Hide();
            cmbfszt.Text = "未处理";
            cmbbgzt.Text = "全部";
              fscs =f.ReadInteger("fsjk", "fscs", 5);
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string sql = "select F_id,F_blh as 病理号,F_jssj as 时间,F_FSZt as 发送状态,F_fscs as 发送次数,f_bz as 备注,f_bglx as 报告类型,f_bgxh as 报告序号 from T_JXKH_FS where   F_jssj>='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'  and F_jssj<'" + dateTimePicker2.Value.AddDays(1).ToString("yyyy-MM-dd") + "'  ";
            if(cmbfszt.Text.Trim()!="" &&cmbfszt.Text.Trim()!="全部")
                sql=sql +" and  F_FSZT='"+cmbfszt.Text.Trim()+"'";

            if (cmbbgzt.Text.Trim() != "" && cmbbgzt.Text.Trim() != "全部")
                sql = sql + " and  F_BGZT='" + cmbbgzt.Text.Trim() + "'";

              if (textBox1.Text.Trim()!="" )
                sql = sql + " and  F_blh='" + textBox1.Text.Trim() + "'";

            DataTable dt1 = aa.GetDataTable(sql, "fsbg");
            dataGridView1.DataSource = dt1;
            label6.Text ="行数："+ dataGridView1.RowCount.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string sql = "select F_id,F_blh as 病理号,F_jssj as 时间,F_FSZt as 发送状态,F_fscs as 发送次数,f_bz as 备注,f_bglx as 报告类型,f_bgxh as 报告序号 from T_JXKH_FS where  (F_fszt='未处理' or F_fszt='') and F_jssj< getdate()-(" + f.ReadInteger("fsjk", "delay", 1) + "/24.0) and F_jssj>='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'  and F_jssj<'" + dateTimePicker2.Value.AddDays(1).ToString("yyyy-MM-dd") + "'  ";
            if (cmbfszt.Text.Trim() != "" && cmbfszt.Text.Trim() != "全部")
                sql = sql + " and  F_FSZT='" + cmbfszt.Text.Trim() + "'";

            if (cmbbgzt.Text.Trim() != "" && cmbbgzt.Text.Trim() != "全部")
                sql = sql + " and  F_BGZT='" + cmbbgzt.Text.Trim() + "'";

            DataTable dt1 = aa.GetDataTable(sql, "fsbg");
            dataGridView1.DataSource = dt1;
            label6.Text = dataGridView1.RowCount.ToString();
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            timer1.Interval = f.ReadInteger("fsjk", "waittime", 5) * 1000 + 5000;
            timer1.Start();
        }
        private bool autosend = true;
        private void timer1_Tick(object sender, EventArgs e)
        {
            if (autosend)
            {
                label8.Text = "自动模式\r\n运行中。。。";
                jzfsjk(""); 
            }
        }

        public void jzfsjk(string F_id)
        {
            string F_blh = "";
            string F_BGZT = "SH";
            string F_bglx;
            string F_bgxh;
            dbbase.odbcdb aa = new odbcdb("DSN=pathnet;UID=pathnet;PWD=4s3c2a1p", "", "");
            string SQLSTR = f.ReadString("fsjk", "sqlstr", "");
            //不检查当前病理号，取触发器库里的数据进行上传处理。

            DataTable jcxx = new DataTable();
            string sqlstr = "";
            try
            {
                if (F_id != "")
                {
                    sqlstr = "select top 1 * from T_JXKH_FS where F_id='" + F_id + "'";
                }
                else
                {
                    sqlstr = "select top 1 * from T_JXKH_FS where (F_fszt='未处理' or F_FSZT='' or F_FSZT is null ) and F_FSCS<=" + fscs
                       + " and F_jssj<= CONVERT(varchar(20),getdate()-(" + f.ReadInteger("fsjk", "delay", 12) + "/24.0),20) and (F_FSSJ<=CONVERT(varchar(20),getdate()-(" + f.ReadInteger("fsjk", "jgsj", 2) + "/24.0),20)  or F_FSSJ='' or f_fssj is null )  order by F_jssj";

                   if (SQLSTR.Trim() != "")
                       sqlstr = SQLSTR;

                }

                jcxx = aa.GetDataTable(sqlstr,"tab");
            }
            catch (Exception ex)
            {
                log.WriteMyLog(F_blh+'-'+ex.Message.ToString());
                return;
            }

            if (f.ReadString("fsjk", "debug", "").Trim() == "1")
            {
                log.WriteMyLog(sqlstr + "\r\n" + "查询到数据条数" + jcxx.Rows.Count.ToString());
            }
            try
            {

                if (jcxx.Rows.Count > 0)
                {
                    F_blh = jcxx.Rows[0]["F_blh"].ToString().Trim();
                    F_id = jcxx.Rows[0]["F_id"].ToString();
                    F_bglx = jcxx.Rows[0]["F_bglx"].ToString();
                    F_bgxh = jcxx.Rows[0]["F_bgxh"].ToString();

                    if (jcxx.Rows[0]["F_bgzt"].ToString() == "取消审核")
                    {
                        F_BGZT = "QXSH";
                    }
                    if (jcxx.Rows[0]["F_bgzt"].ToString() == "已审核")
                    {
                        F_BGZT = "SH";
                    }
                }
                else
                {
                    return;
                }
            }
            catch
            {
                log.WriteMyLog("查询数据异常：" + sqlstr); return;
            }
            DataTable db_jcxx = new DataTable();
            try
            {
                db_jcxx = aa.GetDataTable("select * from T_jcxx where F_blh='" + F_blh + "'", "jcxx");
            }
            catch (Exception ex)
            {
               log.WriteMyLog(F_blh+"-"+ex.Message.ToString());
                return;
            }

            if (db_jcxx.Rows.Count < 1)
            {
                aa.ExecuteSQL("update T_JXKH_FS set F_fszt='不处理',f_bz='无此病理号' where F_id='" + F_id + "'");               
                return;
            }
            aa.ExecuteSQL("update T_JXKH_FS set F_fscs=F_fscs + 1,F_fssj='" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "' where F_id='" + F_id + "'");
           
            try
            {
                XyyyJXKH jxkh = new XyyyJXKH();

                jxkh.pathtohis(F_blh, F_BGZT, F_id,F_bglx,F_bgxh);
            }
            catch (Exception ex)
            {
                log.WriteMyLog(ex.Message);
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            label7.Text="每条数据发送间隔30秒,请耐心等待。。。。";
            if (dataGridView1.Rows.Count > 0)
            {
                label8.Text = "手动模式\r\n运行中。。。";
                autosend = false;
            }
            else
            {
                return;
            }
            try
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    jzfsjk(dataGridView1.Rows[i].Cells["F_id"].Value.ToString());
                    System.Threading.Thread.Sleep(1000);
 
                }
                autosend = true;
                MessageBox.Show("发送完成");
            }
            catch
            {
                autosend = true;
            }
        }

        private void notifyIcon1_DoubleClick(object sender, EventArgs e)
        {
            //if( f.ReadString("fsjk", "isshow", "1")=="1")
            //this.Show();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {

            switch (e.CloseReason)
            {
                case CloseReason.ApplicationExitCall:
                case CloseReason.TaskManagerClosing:
                case CloseReason.WindowsShutDown:
                    e.Cancel = false;
                    break;
                default:
                    {e.Cancel = true; this.Hide();}
                    break;
            } 

         
           
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string sql = "select F_id,F_blh as 病理号,F_bglx as 报告类型,F_bgxh as 报告序号,f_jssj as 时间,F_FSZt as 发送状态,F_fscs as 发送次数,f_bz as 备注,f_bglx as 报告类型,f_bgxh as 报告序号 from T_JXKH_FS where  F_fscs>=" + fscs + "  and F_jssj>='" + dateTimePicker1.Value.ToString("yyyy-MM-dd") + "'  and F_jssj<'" + dateTimePicker2.Value.AddDays(1).ToString("yyyy-MM-dd") + "'";
            if (cmbfszt.Text.Trim() != "" && cmbfszt.Text.Trim() != "全部")
                sql = sql + " and  F_FSZT='" + cmbfszt.Text.Trim() + "'";

            if (cmbbgzt.Text.Trim() != "" && cmbbgzt.Text.Trim() != "全部")
                sql = sql + " and  F_BGZT='" + cmbbgzt.Text.Trim() + "'";

            DataTable dt1 = aa.GetDataTable(sql, "fsbg");
            dataGridView1.DataSource = dt1;
            label6.Text = dataGridView1.RowCount.ToString();
        }

        private void 退出ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Environment.Exit(0);   
        }

        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            //this.Show();
            //this.WindowState = System.Windows.Forms.FormWindowState.Normal;
        }

        private void 显示ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (f.ReadString("fsjk", "isshow", "1") == "1")
                this.Show();
           // this.Show();
            this.WindowState = System.Windows.Forms.FormWindowState.Normal;
        }

        private void Frm_fsjk_Load(object sender, EventArgs e)
        {
          
            cmbfszt.Text = "未处理";
            cmbbgzt.Text = "全部";
            fscs = f.ReadInteger("fsjk", "fscs", 5);
        }
    }
}