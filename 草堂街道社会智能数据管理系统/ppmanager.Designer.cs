using System.Windows.Forms;

namespace 草堂街道社会智能数据管理系统
{
    partial class ppmanager
    {
        public Form main = null; 
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.checkBox15 = new System.Windows.Forms.CheckBox();
            this.checkBox14 = new System.Windows.Forms.CheckBox();
            this.checkBox13 = new System.Windows.Forms.CheckBox();
            this.checkBox12 = new System.Windows.Forms.CheckBox();
            this.button1 = new System.Windows.Forms.Button();
            this.search = new System.Windows.Forms.Button();
            this.checkBox11 = new System.Windows.Forms.CheckBox();
            this.checkBox10 = new System.Windows.Forms.CheckBox();
            this.checkBox9 = new System.Windows.Forms.CheckBox();
            this.checkBox8 = new System.Windows.Forms.CheckBox();
            this.checkBox7 = new System.Windows.Forms.CheckBox();
            this.checkBox6 = new System.Windows.Forms.CheckBox();
            this.checkBox5 = new System.Windows.Forms.CheckBox();
            this.checkBox4 = new System.Windows.Forms.CheckBox();
            this.checkBox3 = new System.Windows.Forms.CheckBox();
            this.checkBox2 = new System.Windows.Forms.CheckBox();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.ssyl_cb = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.sswg_cb = new System.Windows.Forms.ComboBox();
            this.sssq_cb = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.dgv = new System.Windows.Forms.DataGridView();
            this.name = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cardid = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.district = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.grid = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.block = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ad = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.member = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.resident = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.handicapped = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.old = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.alone = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.jwry = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.syry = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.vip = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.clear = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.correction = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.released = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.poor = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dope = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.添加人员ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.button2 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Controls.Add(this.checkBox15);
            this.groupBox1.Controls.Add(this.checkBox14);
            this.groupBox1.Controls.Add(this.checkBox13);
            this.groupBox1.Controls.Add(this.checkBox12);
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.search);
            this.groupBox1.Controls.Add(this.checkBox11);
            this.groupBox1.Controls.Add(this.checkBox10);
            this.groupBox1.Controls.Add(this.checkBox9);
            this.groupBox1.Controls.Add(this.checkBox8);
            this.groupBox1.Controls.Add(this.checkBox7);
            this.groupBox1.Controls.Add(this.checkBox6);
            this.groupBox1.Controls.Add(this.checkBox5);
            this.groupBox1.Controls.Add(this.checkBox4);
            this.groupBox1.Controls.Add(this.checkBox3);
            this.groupBox1.Controls.Add(this.checkBox2);
            this.groupBox1.Controls.Add(this.checkBox1);
            this.groupBox1.Controls.Add(this.ssyl_cb);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.sswg_cb);
            this.groupBox1.Controls.Add(this.sssq_cb);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.textBox2);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(12, 28);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(984, 141);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "人员管理";
            // 
            // checkBox15
            // 
            this.checkBox15.AutoSize = true;
            this.checkBox15.Location = new System.Drawing.Point(698, 87);
            this.checkBox15.Name = "checkBox15";
            this.checkBox15.Size = new System.Drawing.Size(48, 16);
            this.checkBox15.TabIndex = 26;
            this.checkBox15.Text = "党员";
            this.checkBox15.UseVisualStyleBackColor = true;
            // 
            // checkBox14
            // 
            this.checkBox14.AutoSize = true;
            this.checkBox14.Location = new System.Drawing.Point(698, 65);
            this.checkBox14.Name = "checkBox14";
            this.checkBox14.Size = new System.Drawing.Size(72, 16);
            this.checkBox14.TabIndex = 25;
            this.checkBox14.Text = "重要人员";
            this.checkBox14.UseVisualStyleBackColor = true;
            // 
            // checkBox13
            // 
            this.checkBox13.AutoSize = true;
            this.checkBox13.Location = new System.Drawing.Point(698, 42);
            this.checkBox13.Name = "checkBox13";
            this.checkBox13.Size = new System.Drawing.Size(72, 16);
            this.checkBox13.TabIndex = 24;
            this.checkBox13.Text = "失业人员";
            this.checkBox13.UseVisualStyleBackColor = true;
            // 
            // checkBox12
            // 
            this.checkBox12.AutoSize = true;
            this.checkBox12.Location = new System.Drawing.Point(698, 20);
            this.checkBox12.Name = "checkBox12";
            this.checkBox12.Size = new System.Drawing.Size(72, 16);
            this.checkBox12.TabIndex = 23;
            this.checkBox12.Text = "清洁人员";
            this.checkBox12.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(903, 112);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 22;
            this.button1.Text = "取消";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // search
            // 
            this.search.Location = new System.Drawing.Point(819, 112);
            this.search.Name = "search";
            this.search.Size = new System.Drawing.Size(75, 23);
            this.search.TabIndex = 21;
            this.search.Text = "筛选";
            this.search.UseVisualStyleBackColor = true;
            // 
            // checkBox11
            // 
            this.checkBox11.AutoSize = true;
            this.checkBox11.Enabled = false;
            this.checkBox11.Location = new System.Drawing.Point(500, 86);
            this.checkBox11.Name = "checkBox11";
            this.checkBox11.Size = new System.Drawing.Size(48, 16);
            this.checkBox11.TabIndex = 20;
            this.checkBox11.Text = "空巢";
            this.checkBox11.UseVisualStyleBackColor = true;
            // 
            // checkBox10
            // 
            this.checkBox10.AutoSize = true;
            this.checkBox10.Enabled = false;
            this.checkBox10.Location = new System.Drawing.Point(500, 64);
            this.checkBox10.Name = "checkBox10";
            this.checkBox10.Size = new System.Drawing.Size(48, 16);
            this.checkBox10.TabIndex = 19;
            this.checkBox10.Text = "精神";
            this.checkBox10.UseVisualStyleBackColor = true;
            // 
            // checkBox9
            // 
            this.checkBox9.AutoSize = true;
            this.checkBox9.Location = new System.Drawing.Point(596, 87);
            this.checkBox9.Name = "checkBox9";
            this.checkBox9.Size = new System.Drawing.Size(72, 16);
            this.checkBox9.TabIndex = 18;
            this.checkBox9.Text = "吸毒人员";
            this.checkBox9.UseVisualStyleBackColor = true;
            // 
            // checkBox8
            // 
            this.checkBox8.AutoSize = true;
            this.checkBox8.Location = new System.Drawing.Point(434, 108);
            this.checkBox8.Name = "checkBox8";
            this.checkBox8.Size = new System.Drawing.Size(72, 16);
            this.checkBox8.TabIndex = 17;
            this.checkBox8.Text = "境外人员";
            this.checkBox8.UseVisualStyleBackColor = true;
            // 
            // checkBox7
            // 
            this.checkBox7.AutoSize = true;
            this.checkBox7.Location = new System.Drawing.Point(596, 42);
            this.checkBox7.Name = "checkBox7";
            this.checkBox7.Size = new System.Drawing.Size(72, 16);
            this.checkBox7.TabIndex = 16;
            this.checkBox7.Text = "刑满人员";
            this.checkBox7.UseVisualStyleBackColor = true;
            // 
            // checkBox6
            // 
            this.checkBox6.AutoSize = true;
            this.checkBox6.Location = new System.Drawing.Point(596, 20);
            this.checkBox6.Name = "checkBox6";
            this.checkBox6.Size = new System.Drawing.Size(72, 16);
            this.checkBox6.TabIndex = 15;
            this.checkBox6.Text = "矫正人员";
            this.checkBox6.UseVisualStyleBackColor = true;
            // 
            // checkBox5
            // 
            this.checkBox5.AutoSize = true;
            this.checkBox5.Location = new System.Drawing.Point(596, 65);
            this.checkBox5.Name = "checkBox5";
            this.checkBox5.Size = new System.Drawing.Size(72, 16);
            this.checkBox5.TabIndex = 14;
            this.checkBox5.Text = "低保人员";
            this.checkBox5.UseVisualStyleBackColor = true;
            // 
            // checkBox4
            // 
            this.checkBox4.AutoSize = true;
            this.checkBox4.Location = new System.Drawing.Point(434, 86);
            this.checkBox4.Name = "checkBox4";
            this.checkBox4.Size = new System.Drawing.Size(60, 16);
            this.checkBox4.TabIndex = 13;
            this.checkBox4.Text = "老年人";
            this.checkBox4.UseVisualStyleBackColor = true;
            // 
            // checkBox3
            // 
            this.checkBox3.AutoSize = true;
            this.checkBox3.Location = new System.Drawing.Point(434, 64);
            this.checkBox3.Name = "checkBox3";
            this.checkBox3.Size = new System.Drawing.Size(60, 16);
            this.checkBox3.TabIndex = 12;
            this.checkBox3.Text = "残疾人";
            this.checkBox3.UseVisualStyleBackColor = true;
            // 
            // checkBox2
            // 
            this.checkBox2.AutoSize = true;
            this.checkBox2.Location = new System.Drawing.Point(434, 42);
            this.checkBox2.Name = "checkBox2";
            this.checkBox2.Size = new System.Drawing.Size(72, 16);
            this.checkBox2.TabIndex = 11;
            this.checkBox2.Text = "常住人员";
            this.checkBox2.UseVisualStyleBackColor = true;
            // 
            // checkBox1
            // 
            this.checkBox1.AutoSize = true;
            this.checkBox1.Location = new System.Drawing.Point(434, 20);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(72, 16);
            this.checkBox1.TabIndex = 10;
            this.checkBox1.Text = "流动人员";
            this.checkBox1.UseVisualStyleBackColor = true;
            // 
            // ssyl_cb
            // 
            this.ssyl_cb.FormattingEnabled = true;
            this.ssyl_cb.Location = new System.Drawing.Point(233, 95);
            this.ssyl_cb.Name = "ssyl_cb";
            this.ssyl_cb.Size = new System.Drawing.Size(121, 20);
            this.ssyl_cb.TabIndex = 9;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(174, 98);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(53, 12);
            this.label5.TabIndex = 8;
            this.label5.Text = "所属院落";
            // 
            // sswg_cb
            // 
            this.sswg_cb.FormattingEnabled = true;
            this.sswg_cb.Location = new System.Drawing.Point(233, 60);
            this.sswg_cb.Name = "sswg_cb";
            this.sswg_cb.Size = new System.Drawing.Size(121, 20);
            this.sswg_cb.TabIndex = 7;
            this.sswg_cb.SelectedIndexChanged += new System.EventHandler(this.sswg_cb_SelectedIndexChanged);
            // 
            // sssq_cb
            // 
            this.sssq_cb.FormattingEnabled = true;
            this.sssq_cb.Location = new System.Drawing.Point(233, 22);
            this.sssq_cb.Name = "sssq_cb";
            this.sssq_cb.Size = new System.Drawing.Size(121, 20);
            this.sssq_cb.TabIndex = 6;
            this.sssq_cb.SelectedIndexChanged += new System.EventHandler(this.sssq_cb_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(174, 63);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 12);
            this.label4.TabIndex = 5;
            this.label4.Text = "所属网格";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(174, 25);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 4;
            this.label3.Text = "所属社区";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(57, 60);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(100, 21);
            this.textBox2.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(10, 63);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "身份证";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(57, 22);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(100, 21);
            this.textBox1.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "姓名";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.dgv);
            this.groupBox2.Location = new System.Drawing.Point(12, 175);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(984, 542);
            this.groupBox2.TabIndex = 4;
            this.groupBox2.TabStop = false;
            // 
            // dgv
            // 
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.dgv.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle2;
            this.dgv.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dgv.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.name,
            this.cardid,
            this.district,
            this.grid,
            this.block,
            this.ad,
            this.member,
            this.resident,
            this.handicapped,
            this.old,
            this.alone,
            this.jwry,
            this.syry,
            this.vip,
            this.clear,
            this.correction,
            this.released,
            this.poor,
            this.dope});
            this.dgv.Dock = System.Windows.Forms.DockStyle.Left;
            this.dgv.Location = new System.Drawing.Point(3, 17);
            this.dgv.MultiSelect = false;
            this.dgv.Name = "dgv";
            this.dgv.ReadOnly = true;
            this.dgv.RowHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.Single;
            this.dgv.RowHeadersVisible = false;
            this.dgv.RowTemplate.Height = 23;
            this.dgv.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgv.Size = new System.Drawing.Size(978, 522);
            this.dgv.TabIndex = 0;
            // 
            // name
            // 
            this.name.DataPropertyName = "na";
            this.name.HeaderText = "姓名";
            this.name.Name = "name";
            this.name.ReadOnly = true;
            // 
            // cardid
            // 
            this.cardid.DataPropertyName = "cid";
            this.cardid.HeaderText = "身份证";
            this.cardid.Name = "cardid";
            this.cardid.ReadOnly = true;
            // 
            // district
            // 
            this.district.DataPropertyName = "nd";
            this.district.HeaderText = "所属社区";
            this.district.Name = "district";
            this.district.ReadOnly = true;
            // 
            // grid
            // 
            this.grid.DataPropertyName = "ng";
            this.grid.HeaderText = "所属网格";
            this.grid.Name = "grid";
            this.grid.ReadOnly = true;
            // 
            // block
            // 
            this.block.DataPropertyName = "nb";
            this.block.HeaderText = "所属院落";
            this.block.Name = "block";
            this.block.ReadOnly = true;
            // 
            // ad
            // 
            this.ad.DataPropertyName = "ad";
            this.ad.HeaderText = "居住地址";
            this.ad.Name = "ad";
            this.ad.ReadOnly = true;
            // 
            // member
            // 
            this.member.DataPropertyName = "edu";
            this.member.HeaderText = "党员";
            this.member.Name = "member";
            this.member.ReadOnly = true;
            // 
            // resident
            // 
            this.resident.DataPropertyName = "res";
            this.resident.HeaderText = "常住人员";
            this.resident.Name = "resident";
            this.resident.ReadOnly = true;
            this.resident.Visible = false;
            // 
            // handicapped
            // 
            this.handicapped.DataPropertyName = "hand";
            this.handicapped.HeaderText = "残疾人";
            this.handicapped.Name = "handicapped";
            this.handicapped.ReadOnly = true;
            this.handicapped.Visible = false;
            // 
            // old
            // 
            this.old.DataPropertyName = "old";
            this.old.HeaderText = "老年人";
            this.old.Name = "old";
            this.old.ReadOnly = true;
            this.old.Visible = false;
            // 
            // alone
            // 
            this.alone.DataPropertyName = "olda";
            this.alone.HeaderText = "空巢老人";
            this.alone.Name = "alone";
            this.alone.ReadOnly = true;
            // 
            // jwry
            // 
            this.jwry.DataPropertyName = "fore";
            this.jwry.HeaderText = "境外人员";
            this.jwry.Name = "jwry";
            this.jwry.ReadOnly = true;
            // 
            // syry
            // 
            this.syry.DataPropertyName = "unjob";
            this.syry.HeaderText = "失业人员";
            this.syry.Name = "syry";
            this.syry.ReadOnly = true;
            // 
            // vip
            // 
            this.vip.DataPropertyName = "vip";
            this.vip.HeaderText = "重要人员";
            this.vip.Name = "vip";
            this.vip.ReadOnly = true;
            // 
            // clear
            // 
            this.clear.DataPropertyName = "cle";
            this.clear.HeaderText = "清洁人员 ";
            this.clear.Name = "clear";
            this.clear.ReadOnly = true;
            // 
            // correction
            // 
            this.correction.DataPropertyName = "cor";
            this.correction.HeaderText = "矫正人员";
            this.correction.Name = "correction";
            this.correction.ReadOnly = true;
            this.correction.Visible = false;
            // 
            // released
            // 
            this.released.DataPropertyName = "rel";
            this.released.HeaderText = "刑满人员";
            this.released.Name = "released";
            this.released.ReadOnly = true;
            this.released.Visible = false;
            // 
            // poor
            // 
            this.poor.DataPropertyName = "poor";
            this.poor.HeaderText = "低保人员";
            this.poor.Name = "poor";
            this.poor.ReadOnly = true;
            this.poor.Visible = false;
            // 
            // dope
            // 
            this.dope.DataPropertyName = "dope";
            this.dope.HeaderText = "吸毒人员";
            this.dope.Name = "dope";
            this.dope.ReadOnly = true;
            this.dope.Visible = false;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.添加人员ToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1008, 25);
            this.menuStrip1.TabIndex = 5;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // 添加人员ToolStripMenuItem
            // 
            this.添加人员ToolStripMenuItem.Name = "添加人员ToolStripMenuItem";
            this.添加人员ToolStripMenuItem.Size = new System.Drawing.Size(68, 21);
            this.添加人员ToolStripMenuItem.Tag = "add";
            this.添加人员ToolStripMenuItem.Text = "添加人员";
            this.添加人员ToolStripMenuItem.Click += new System.EventHandler(this.添加人员ToolStripMenuItem_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(735, 112);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 27;
            this.button2.Text = "导出结果";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // ppmanager
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1008, 729);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "ppmanager";
            this.Text = "人员管理";
            this.Load += new System.EventHandler(this.ppmanager_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgv)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox checkBox8;
        private System.Windows.Forms.CheckBox checkBox7;
        private System.Windows.Forms.CheckBox checkBox6;
        private System.Windows.Forms.CheckBox checkBox5;
        private System.Windows.Forms.CheckBox checkBox4;
        private System.Windows.Forms.CheckBox checkBox3;
        private System.Windows.Forms.CheckBox checkBox2;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.ComboBox ssyl_cb;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox sswg_cb;
        private System.Windows.Forms.ComboBox sssq_cb;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox checkBox10;
        private System.Windows.Forms.CheckBox checkBox9;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox checkBox11;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button search;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem 添加人员ToolStripMenuItem;
        private System.Windows.Forms.DataGridView dgv;
        private CheckBox checkBox14;
        private CheckBox checkBox13;
        private CheckBox checkBox12;
        private CheckBox checkBox15;
        private DataGridViewTextBoxColumn name;
        private DataGridViewTextBoxColumn cardid;
        private DataGridViewTextBoxColumn district;
        private DataGridViewTextBoxColumn grid;
        private DataGridViewTextBoxColumn block;
        private DataGridViewTextBoxColumn ad;
        private DataGridViewTextBoxColumn member;
        private DataGridViewTextBoxColumn resident;
        private DataGridViewTextBoxColumn handicapped;
        private DataGridViewTextBoxColumn old;
        private DataGridViewTextBoxColumn alone;
        private DataGridViewTextBoxColumn jwry;
        private DataGridViewTextBoxColumn syry;
        private DataGridViewTextBoxColumn vip;
        private DataGridViewTextBoxColumn clear;
        private DataGridViewTextBoxColumn correction;
        private DataGridViewTextBoxColumn released;
        private DataGridViewTextBoxColumn poor;
        private DataGridViewTextBoxColumn dope;
        private Button button2;
    }
}