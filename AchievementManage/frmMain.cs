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
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private bool checkchildfrm(string childfrmname)//查询子窗体是否存在
        {
            foreach (Form childFrm in this.MdiChildren)//遍历子窗体
            {
                if (childFrm.Name == childfrmname)//判断子窗体的名称
                {
                    if (childFrm.WindowState == FormWindowState.Minimized)//如果子窗体处于最小化的状态
                    {
                        childFrm.WindowState = FormWindowState.Normal;//恢复正常显示
                    }
                    childFrm.Activate();//激活窗体
                    return true;//返回真值
                }
            }
            return false;//返回假值
        }

        private void tsmniAchievementAdd_Click(object sender, EventArgs e)//成果录入
        {
            if (this.checkchildfrm("frmAchievementAdd") == true)//检测该窗体是否处于打开状态
            {
                return;//窗口已打开，返回
            }
            frmAchievementAdd achievement_add = new frmAchievementAdd();//实例化成果录入窗体
            achievement_add.MdiParent = this;//设置为当前窗体的子窗体
            achievement_add.Show();//成果录入窗体以非模式对话框的方式打开
        }

        private void tsmniAchievementSearchEasy_Click(object sender, EventArgs e)//简单成果检索
        {
            if (this.checkchildfrm("frmAchievementSearchEasy") == true)//检测该窗体是否处于打开状态
            {
                return;//窗口已打开，返回
            }
            frmAchievementSearchEasy achievement_search_easy = new frmAchievementSearchEasy();//实例化简单成果检索窗体
            achievement_search_easy.MdiParent = this;//设置为当前窗体的子窗体
            achievement_search_easy.Show();//简单成果检索窗体以非模式对话框的方式打开
        }

        private void tsmniAchievementSearchComplex_Click(object sender, EventArgs e)//高级成果检索
        {
            if (this.checkchildfrm("frmAchievementSearchComplex") == true)//检测该窗体是否处于打开状态
            {
                return;//窗口已打开，返回
            }
            frmAchievementSearchComplex achievement_search_complex = new frmAchievementSearchComplex();//实例化高级成果检索窗体
            achievement_search_complex.MdiParent = this;//设置为当前窗体的子窗体
            achievement_search_complex.Show();//高级成果检索窗体以非模式对话框的方式打开
        }

        private void tsmniExit_Click(object sender, EventArgs e)//退出
        {
            Application.Exit();//退出应用程序
        }
    }
}
