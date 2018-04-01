using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AchievementManage
{
    public partial class frmAchievementAdd : Form
    {
        public frmAchievementAdd()
        {
            InitializeComponent();
        }

        private string txt_achievement_name = string.Empty;//存储数据库内的主键的值，用于删除记录

        private void frmAchievementAdd_Load(object sender, EventArgs e)//窗体载入时初始化
        {
            txt_achievement_name = string.Empty;
            txtAchievementType.Visible = false;//成果类型编辑框不可见
            btnDelete.Enabled = false;//删除此项成果按钮不可点击
            btnOutToExcel.Text = "导出此项成果到Excel";//导出此项成果到Excel按钮显示内容
            btnOutToExcel.Enabled = false;//导出此项成果到Excel按钮不可点击
            this.dgvData.DataSource = null;//DataGridView控件显示数据

            #region 成果类型组合框初始化
            if (MyDatabase.TestMyDatabaseConnect() == false)//数据库连接失败
            {
                MessageBox.Show("数据库连接失败！", "提示");
                return;
            }
            try
            {
                DataSet ds = MyDatabase.GetDataSetBySql("select distinct AchievementType from datainfo;");
                if (ds == null)//数据库连接失败
                {
                    MessageBox.Show("数据库连接失败！", "提示");
                    return;
                }
                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    this.cboAchievementType.Items.Add(ds.Tables[0].Rows[i][0].ToString().Trim());
                }
                this.cboAchievementType.Items.Add("(其他)");
                this.cboAchievementType.SelectedIndex = 0;
                //防止数据库中没有任何成果类型而使得成果类型中只有"(其他)"而使得永远无法触发cboAchievementType_SelectedIndexChanged函数使得成果类型编辑框永远不可见，所以加入以下判断代码在初始化代码中
                if (string.Compare(this.cboAchievementType.SelectedItem.ToString(), "(其他)") == 0)
                {
                    txtAchievementType.Visible = true;//成果类型编辑框可见
                }
                else
                {
                    txtAchievementType.Visible = false;//成果类型编辑框不可见
                }
            }
            catch (Exception ex)//数据库可能连接不上，MyDatabase.GetDataSetBySql函数会出错，返回的DataSet值为null，导致下面引用DataSet具体值时会抛出异常(后经改进在DataSet具体值使用之前加入了值是否为null的判断，后又经过改进在每次事件处理前均加入了检测能否成功连接的函数)
            {
                MessageBox.Show(ex.Message);
            }
            #endregion
        }

        private void btnAdd_Click(object sender, EventArgs e)//录入此项成果
        {
            if (MyDatabase.TestMyDatabaseConnect() == false)//数据库连接失败
            {
                MessageBox.Show("数据库连接失败！", "提示");
                return;
            }
            if (this.txtAchievementName.Text.ToString().Trim() == string.Empty || this.dtpAchievementDate.Text.ToString().Trim() == string.Empty || this.txtAchievementAuthor.Text.ToString().Trim() == string.Empty || this.txtAchievementCompany.Text.ToString().Trim() == string.Empty || this.txtAchievementMoney.Text.ToString().Trim() == string.Empty)//数据是否为空检测
            {
                MessageBox.Show("信息输入不完整，请重新检查之后再重试！", "提示");
                return;
            }
            if (string.Compare(this.cboAchievementType.SelectedItem.ToString(), "(其他)") == 0)//对成果类型数据是否为空进行单独判断
            {
                if (this.txtAchievementType.Text.ToString().Trim() == string.Empty)
                {
                    MessageBox.Show("信息输入不完整，请重新检查之后再重试！", "提示");
                    return;
                }
            }
            try//对成果时间和成果支撑基金的数据格式进行检测(如数据有误会throw异常，程序进入catch捕捉代码中)
            {
                Convert.ToDateTime(this.dtpAchievementDate.Text.ToString().Trim()).ToString("yyyy-MM-dd");//最后加入ToString("yyyy-MM-dd")是为了去除时间只留下日期
                Convert.ToDouble(this.txtAchievementMoney.Text.ToString().Trim());
            }
            catch (Exception ex)
            {
                MessageBox.Show("信息输入格式有误，请重新检查之后再重试！", "提示");
                //throw new Exception(ex.Message);
                return;
            }
            //所有数据的合法性均检测完毕
            string sql_1 = string.Empty;
            if(string.Compare(this.cboAchievementType.SelectedItem.ToString(), "(其他)") == 0)//对成果类型数据进行判断
            {
                sql_1 = string.Format("insert into datainfo values('{0}', '{1}', '{2}', '{3}', '{4}', {5});", this.txtAchievementName.Text.ToString().Trim(), this.txtAchievementType.Text.ToString().Trim(), Convert.ToDateTime(this.dtpAchievementDate.Text.ToString().Trim()).ToString("yyyy-MM-dd"), this.txtAchievementAuthor.Text.ToString().Trim(), this.txtAchievementCompany.Text.ToString().Trim(), Convert.ToDouble(this.txtAchievementMoney.Text.ToString().Trim()));//sql语句                
            }
            else
            {
                sql_1 = string.Format("insert into datainfo values('{0}', '{1}', '{2}', '{3}', '{4}', {5});", this.txtAchievementName.Text.ToString().Trim(), this.cboAchievementType.SelectedItem.ToString().Trim(), Convert.ToDateTime(this.dtpAchievementDate.Text.ToString().Trim()).ToString("yyyy-MM-dd"), this.txtAchievementAuthor.Text.ToString().Trim(), this.txtAchievementCompany.Text.ToString().Trim(), Convert.ToDouble(this.txtAchievementMoney.Text.ToString().Trim()));//sql语句
            }
            if (MyDatabase.UpdateDataBySql(sql_1))//添加成功
            {
                btnDelete.Enabled = true;//删除此项成果按钮可点击
                btnOutToExcel.Text = "导出此项成果到Excel";//导出此项成果到Excel按钮显示内容
                btnOutToExcel.Enabled = true;//导出此项成果到Excel按钮可点击
                txt_achievement_name = this.txtAchievementName.Text.ToString().Trim();//记录数据库内主键的值
                string sql_2 = string.Format("select datainfo.AchievementName as '成果名称', datainfo.AchievementType as '成果类型', datainfo.AchievementDate as '时间', datainfo.AchievementAuthor as '作者', datainfo.AchievementCompany as '单位', datainfo.AchievementMoney as '支撑基金' from datainfo where AchievementName='{0}';", this.txtAchievementName.Text.ToString().Trim());
                DataSet ds = MyDatabase.GetDataSetBySql(sql_2);
                if (ds == null)//数据库连接失败
                {
                    btnDelete.Enabled = false;//删除此项成果按钮不可点击
                    btnOutToExcel.Text = "导出此项成果到Excel";//导出此项成果到Excel按钮显示内容
                    btnOutToExcel.Enabled = false;//导出此项成果到Excel按钮不可点击
                    txt_achievement_name = string.Empty;//记录数据库内主键的值为空
                    this.dgvData.DataSource = null;//DataGridView控件显示数据
                    MessageBox.Show("数据库连接失败！", "提示");
                    return;
                }
                this.dgvData.AutoGenerateColumns = true;//自动
                this.dgvData.DataSource = ds.Tables[0];//DataGridView控件显示数据
                this.dgvData.Columns[0].ReadOnly = true;//设为只读
                this.dgvData.Columns[1].ReadOnly = true;//设为只读
                this.dgvData.Columns[2].ReadOnly = true;//设为只读
                this.dgvData.Columns[3].ReadOnly = true;//设为只读
                this.dgvData.Columns[4].ReadOnly = true;//设为只读
                this.dgvData.Columns[5].ReadOnly = true;//设为只读
                MessageBox.Show("添加成功！", "提示");
            }
            else//添加失败
            {
                btnDelete.Enabled = false;//删除此项成果按钮不可点击
                btnOutToExcel.Text = "导出此项成果到Excel";//导出此项成果到Excel按钮显示内容
                btnOutToExcel.Enabled = false;//导出此项成果到Excel按钮不可点击
                txt_achievement_name = string.Empty;//记录数据库内主键的值为空
                this.dgvData.DataSource = null;//DataGridView控件显示数据
                MessageBox.Show("添加失败！", "提示");
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)//删除此项成果
        {
            if (MyDatabase.TestMyDatabaseConnect() == false)//数据库连接失败
            {
                MessageBox.Show("数据库连接失败！", "提示");
                return;
            }
            if (txt_achievement_name != string.Empty)
            {
                string sql = string.Format("delete from datainfo where datainfo.AchievementName='{0}';", txt_achievement_name);//sql语句
                if (MyDatabase.UpdateDataBySql(sql))//删除成功
                {
                    btnDelete.Enabled = false;//删除此项成果按钮不可点击
                    btnOutToExcel.Text = "导出此项成果到Excel";//导出此项成果到Excel按钮显示内容
                    btnOutToExcel.Enabled = false;//导出此项成果到Excel按钮不可点击
                    txt_achievement_name = string.Empty;//记录数据库内主键的值为空
                    this.dgvData.DataSource = null;//DataGridView控件显示数据
                    MessageBox.Show("删除成功！", "提示");
                    return;
                }
            }
            MessageBox.Show("删除失败！", "提示");
        }

        private void btnOutToExcel_Click(object sender, EventArgs e)//导出此项成果到Excel
        {
            if (MyDatabase.TestMyDatabaseConnect() == false)//数据库连接失败
            {
                MessageBox.Show("数据库连接失败！", "提示");
                return;
            }
            if (txt_achievement_name != string.Empty)
            {
                try
                {
                    DataTable dt = MyExcel.GetDgvToTable(dgvData);
                    this.sfdOutToExcel.Filter = "Excel 工作簿(*.xlsx)|*.xlsx";//设置保存类型
                    DialogResult result = this.sfdOutToExcel.ShowDialog();
                    if (result == DialogResult.OK)//点击了保存
                    {
                        btnOutToExcel.Text = "正在导出中。。。";//导出此项成果到Excel按钮显示内容
                        btnOutToExcel.Enabled = false;//导出此项成果到Excel按钮显示不可点击
                        if (MyExcel.DateTimeRemoveTime(dt) == false)//对dt内的数据进行单独处理，去除时间(默认值00:00:00)，只保留日期
                        {
                            btnOutToExcel.Text = "导出此项成果到Excel";//设置导出此项成果到Excel按钮显示内容
                            btnOutToExcel.Enabled = true;//设置导出此项成果到Excel按钮显示可点击
                            MessageBox.Show("导出失败！", "提示");
                            return;
                        }
                        MyExcel.SaveDataToExcel(dt, sfdOutToExcel.FileName);//此过程很慢
                        btnOutToExcel.Text = "导出此项成果到Excel";//导出此项成果到Excel按钮显示内容
                        btnOutToExcel.Enabled = true;//导出此项成果到Excel按钮显示可点击
                        MessageBox.Show("导出成功！", "提示");
                        return;
                    }
                    else if (result == DialogResult.Cancel)//点击了取消
                    {
                        btnOutToExcel.Text = "导出此项成果到Excel";//设置导出此项成果到Excel按钮显示内容
                        btnOutToExcel.Enabled = true;//设置导出此项成果到Excel按钮显示可点击
                        return;
                    }
                }
                catch (Exception ex)
                {
                    btnOutToExcel.Text = "导出此项成果到Excel";//设置导出此项成果到Excel按钮显示内容
                    btnOutToExcel.Enabled = true;//设置导出此项成果到Excel按钮显示可点击
                    MessageBox.Show("导出失败！", "提示");
                    return;
                }
            }
            btnOutToExcel.Text = "导出此项成果到Excel";//设置导出此项成果到Excel按钮显示内容
            btnOutToExcel.Enabled = true;//设置导出此项成果到Excel按钮显示可点击
            MessageBox.Show("导出失败！", "提示");
        }

        private void btnExit_Click(object sender, EventArgs e)//退出
        {
            this.Close();//关闭窗体
        }

        private void cboAchievementType_SelectedIndexChanged(object sender, EventArgs e)//combox内容变化时
        {
            if (string.Compare(this.cboAchievementType.SelectedItem.ToString(), "(其他)") == 0)
            {
                txtAchievementType.Visible = true;//成果类型编辑框可见
            }
            else
            {
                txtAchievementType.Visible = false;//成果类型编辑框不可见
            }
        }
    }
}
