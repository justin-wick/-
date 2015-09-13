using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using 草堂街道社会智能数据管理系统.ComClass;
using 草堂街道社会智能数据管理系统.dataClass;

namespace 草堂街道社会智能数据管理系统
{
    public partial class ppmanager : Form
    {
        private readonly CommonUse commUse = new CommonUse();
        private readonly DataBase db = new DataBase();
 

        public ppmanager()
        {
            InitializeComponent();
        

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void 添加人员ToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
            CommonUse commUse = new CommonUse();
            var x = (ToolStripMenuItem)sender;

            commUse.ShowForm(x.Tag.ToString(), this.main);
        }

        private void ppmanager_Load(object sender, EventArgs e)
        {
            string sqlcmd = "SELECT district.`name` AS nd, grid.`name` AS ng, block.`name` AS nb, population.card_id AS cid, population.`name` AS na,population.address AS ad, "
           + "CASE population.educational WHEN '党员' THEN '是' ELSE '否' end AS edu,"
           + "CASE features.vip when true then '是' ELSE '否'end AS vip,"
            + "CASE features.cleaner when true then '是' ELSE '否'end AS cle,"
            + "CASE features.old when true then '是' ELSE '否'end AS old,"
            + "CASE features.old_alone when true then '是' ELSE '否'end AS olda,"
            + "CASE features.poor when true then '是' ELSE '否'end AS poor,"
            + "CASE features.handicapped when true then '是' ELSE '否'end AS hand,"
            + "CASE features.resident when true then '是' ELSE '否'end AS res,"
            + "CASE features.unjob when true then '是' ELSE '否'end AS unjob,"
            + "CASE features.dope when true then '是' ELSE '否'end AS dope,"
            + "CASE features.correction when true then '是' ELSE '否'end AS cor,"
            + "CASE features.released when true then '是' ELSE '否'end AS rel,"
            + "CASE features.foreigner when true then '是' ELSE '否'end AS fore FROM district INNER JOIN grid ON grid.district = district.id INNER JOIN block ON block.grid = grid.id INNER JOIN population ON population.block = block.id INNER JOIN features ON population.features = features.id ";
            List<item> items = new List<item>();
            MySqlDataReader sdr;
            sdr = db.GetDataReader("SELECT district.`name`,district.`id` FROM district");
            while (sdr.Read())
            {
                item it = new item(sdr[0].ToString(), sdr[1].ToString());
                items.Add(it);
            }
            sssq_cb.DataSource = items;
            //   items.Clear();
            sdr.Close();
            sdr = db.GetDataReader("SELECT 	grid.`name`,grid.id FROM district INNER JOIN grid ON grid.district = district.id WHERE district.id = 1");
            while (sdr.Read())
            {
                item it = new item(sdr[0].ToString(), sdr[1].ToString());
                items.Add(it);
            }
            sswg_cb.DataSource = items;
            //  items.Clear();
            sdr.Close();
            sdr = db.GetDataReader("SELECT 	block.`name`,block.id FROM grid INNER JOIN block ON block.grid = grid.id WHERE grid.id = 1");
            while (sdr.Read())
            {
                item it = new item(sdr[0].ToString(), sdr[1].ToString());
                items.Add(it);
            }
            ssyl_cb.DataSource = items;
            //  items.Clear();
            sdr.Close();
            
            if (this.Tag.ToString() != null)
            {
                switch (this.Tag.ToString())
                {
                    case "m4":
                        {
                            sqlcmd += "WHERE population.educational = '党员'" ;
                        }
                        break;
                    
                    default:
                        break;
                }
                dgv.DataSource = db.GetDataSet(sqlcmd, "t").Tables["t"];
            }
        }

        private void sssq_cb_SelectedIndexChanged(object sender, EventArgs e)
        {
            commUse.district_gird_block(sssq_cb, sswg_cb, "district", "grid");
        }

        private void sswg_cb_SelectedIndexChanged(object sender, EventArgs e)
        {
            commUse.district_gird_block(sswg_cb, ssyl_cb, "grid", "block");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Excel.ExportExcel excel = new Excel.ExportExcel();

            excel.CreateExcel();
            excel.CreateWorkSheet("人员导出信息");
            // 第一行加粗
            excel.FontStyle(1, 1, 1, 10, true, false, Excel.UnderlineStyle.无下划线);
            //p.CellsUnite(1, 1, 1, 5);
            // 表格数据

            // Excel第一行数据
            excel.WriteData("姓名", 1, 1);
            excel.WriteData("所属社区", 1, 2);
            excel.WriteData("所属网络", 1, 3);
            excel.WriteData("身份证", 1, 4);
            excel.WriteData("所属院落", 1, 5);
            excel.WriteData("居住地址", 1, 6);
            excel.WriteData("是否党员", 1, 7);
            excel.WriteData("重要人员", 1, 8);
            excel.WriteData("清洁人员", 1, 9);
            excel.WriteData("境外人员", 1, 10);

            DataTable outDatatable = new DataTable();
            outDatatable.Columns.Add("name");
            outDatatable.Columns.Add("district");
            outDatatable.Columns.Add("grid");
            outDatatable.Columns.Add("cardid");
            outDatatable.Columns.Add("ad");
            outDatatable.Columns.Add("member");
            outDatatable.Columns.Add("vip");
            outDatatable.Columns.Add("clear");
            outDatatable.Columns.Add("jwry");

            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                
                DataRow newRow = outDatatable.NewRow();
                newRow["name"] = dgv.Rows[i].Cells["name"].Value;
                newRow["district"] = dgv.Rows[i].Cells["district"].Value;
                newRow["grid"] = dgv.Rows[i].Cells["grid"].Value;
                newRow["cardid"] = dgv.Rows[i].Cells["cardid"].Value;
                newRow["ad"] = dgv.Rows[i].Cells["ad"].Value;
                newRow["member"] = dgv.Rows[i].Cells["member"].Value;
                newRow["vip"] = dgv.Rows[i].Cells["vip"].Value;
                newRow["clear"] = dgv.Rows[i].Cells["clear"].Value;
                newRow["jwry"] = dgv.Rows[i].Cells["jwry"].Value;

                outDatatable.Rows.Add(newRow);
            }

            excel.WriteData(outDatatable, 2, 1);
            excel.Close(true, "bbb.xlsx");
        }
    }
}
