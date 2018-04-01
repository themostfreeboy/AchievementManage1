namespace AchievementManage
{
    partial class frmMain
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.mnsMain = new System.Windows.Forms.MenuStrip();
            this.tsmniAchievementManage = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmniAchievementAdd = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmniAchievementSearch = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmniAchievementSearchEasy = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmniAchievementSearchComplex = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmniExit = new System.Windows.Forms.ToolStripMenuItem();
            this.mnsMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // mnsMain
            // 
            this.mnsMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmniAchievementManage,
            this.tsmniExit});
            this.mnsMain.Location = new System.Drawing.Point(0, 0);
            this.mnsMain.Name = "mnsMain";
            this.mnsMain.Size = new System.Drawing.Size(884, 25);
            this.mnsMain.TabIndex = 1;
            this.mnsMain.Text = "menuStrip1";
            // 
            // tsmniAchievementManage
            // 
            this.tsmniAchievementManage.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmniAchievementAdd,
            this.tsmniAchievementSearch});
            this.tsmniAchievementManage.Name = "tsmniAchievementManage";
            this.tsmniAchievementManage.Size = new System.Drawing.Size(68, 21);
            this.tsmniAchievementManage.Text = "成果管理";
            // 
            // tsmniAchievementAdd
            // 
            this.tsmniAchievementAdd.Name = "tsmniAchievementAdd";
            this.tsmniAchievementAdd.Size = new System.Drawing.Size(124, 22);
            this.tsmniAchievementAdd.Text = "成果录入";
            this.tsmniAchievementAdd.Click += new System.EventHandler(this.tsmniAchievementAdd_Click);
            // 
            // tsmniAchievementSearch
            // 
            this.tsmniAchievementSearch.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmniAchievementSearchEasy,
            this.tsmniAchievementSearchComplex});
            this.tsmniAchievementSearch.Name = "tsmniAchievementSearch";
            this.tsmniAchievementSearch.Size = new System.Drawing.Size(124, 22);
            this.tsmniAchievementSearch.Text = "成果检索";
            // 
            // tsmniAchievementSearchEasy
            // 
            this.tsmniAchievementSearchEasy.Name = "tsmniAchievementSearchEasy";
            this.tsmniAchievementSearchEasy.Size = new System.Drawing.Size(244, 22);
            this.tsmniAchievementSearchEasy.Text = "通过成果类型进行简单成果检索";
            this.tsmniAchievementSearchEasy.Click += new System.EventHandler(this.tsmniAchievementSearchEasy_Click);
            // 
            // tsmniAchievementSearchComplex
            // 
            this.tsmniAchievementSearchComplex.Name = "tsmniAchievementSearchComplex";
            this.tsmniAchievementSearchComplex.Size = new System.Drawing.Size(244, 22);
            this.tsmniAchievementSearchComplex.Text = "高级成果检索";
            this.tsmniAchievementSearchComplex.Click += new System.EventHandler(this.tsmniAchievementSearchComplex_Click);
            // 
            // tsmniExit
            // 
            this.tsmniExit.Name = "tsmniExit";
            this.tsmniExit.Size = new System.Drawing.Size(44, 21);
            this.tsmniExit.Text = "退出";
            this.tsmniExit.Click += new System.EventHandler(this.tsmniExit_Click);
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(884, 562);
            this.Controls.Add(this.mnsMain);
            this.IsMdiContainer = true;
            this.MainMenuStrip = this.mnsMain;
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "成果管理系统";
            this.mnsMain.ResumeLayout(false);
            this.mnsMain.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip mnsMain;
        private System.Windows.Forms.ToolStripMenuItem tsmniAchievementManage;
        private System.Windows.Forms.ToolStripMenuItem tsmniAchievementAdd;
        private System.Windows.Forms.ToolStripMenuItem tsmniAchievementSearch;
        private System.Windows.Forms.ToolStripMenuItem tsmniAchievementSearchEasy;
        private System.Windows.Forms.ToolStripMenuItem tsmniAchievementSearchComplex;
        private System.Windows.Forms.ToolStripMenuItem tsmniExit;
    }
}

