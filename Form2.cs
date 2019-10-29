using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PgmMainMonitor
{
    public partial class Form2 : Form
    {
        GroupBox[] group;
        Label[] label_info;

        public Form2()
        {
            InitializeComponent();

			group = new GroupBox [] {groupBox1, groupBox2, groupBox3, groupBox4, groupBox5, groupBox7, groupBox8, groupBox9, groupBox29, groupBox10,
									 groupBox11, groupBox12, groupBox13, groupBox14, groupBox15, groupBox16, groupBox17, groupBox18, groupBox25, groupBox26,
									 groupBox19, groupBox20, groupBox21, groupBox22, groupBox23, groupBox24
									};

            label_info = new Label[] {label1, label2, label3, label4, label5, label7, label8, label27, label9, label10,
                                      label11, label12, label13, label14, label15, label16, label17, label18, label25, label26,
                                      label19, label20, label21, label22, label23, label24
                                      };

            this.MaximizeBox = false;         // 最大化ボタン
            this.MinimizeBox = false;         // 最小化ボタン

			self_term_timer.Enabled = true;
        }

        public void SetInfo(string info_str)
		{
			//情報をlabelに表示
            string[] data = info_str.Split(',');
			for(int i = 0; i < label_info.Length; i++)
			{
				label_info[i].Text = data[i];
			}

			//注意事項は背景色を変更する
			for(int i = 0; i < group.Length; i++)
			{
				if(i == 0)//成型機
				{
					if(label_info[i].Text.IndexOf("LS") >= 0)
					{
						group[i].BackColor = Color.RoyalBlue;
					}
					else if(label_info[i].Text.IndexOf("NQD") >= 0)
					{
						group[i].BackColor = Color.Orange;
					}
					else if(label_info[i].Text.IndexOf("HS") >= 0)
					{
						group[i].BackColor = Color.Violet;
					}
				}
				else if(i == 1)//稼働状況
				{
					if(label_info[i].Text == "準備中")
					{
						group[i].BackColor = Color.Yellow;
					}
					else if(label_info[i].Text == "停止中")
					{
						group[i].BackColor = Color.Tomato;
					}
					else
					{
						group[i].BackColor = Color.Lime;
					}
				}
				else if(i == 2)//ﾀｲﾌﾟ
				{
					if(label_info[i].Text.IndexOf("ｼﾝｸﾞﾙ") >= 0)
					{
						group[i].BackColor = Color.Khaki;
					}
					else if(label_info[i].Text.IndexOf("ﾏﾙﾁ") >= 0)
					{
						group[i].BackColor = Color.YellowGreen;
					}
				}
				else if(i == 3)//稼働率
				{
					string kadou = label_info[i].Text.Substring(0, (label_info[i].Text.Length - 1));
					double kadou_double = double.Parse(kadou);
					if(kadou_double < Form1.SETDATA.kadouYellow)
					{
						group[i].BackColor = Color.Tomato;
					}
					else if(Form1.SETDATA.kadouYellow <= kadou_double && kadou_double < Form1.SETDATA.kadouRed)
					{
						group[i].BackColor = Color.Yellow;
					}
					else if(Form1.SETDATA.kadouRed <= kadou_double)
					{
						group[i].BackColor = Color.Lime;
					}
				}
				else if(i == 6)//加圧時間
				{
					if(label_info[i].Text.IndexOf("注意") >= 0)
					{
						group[i].BackColor = Color.Tomato;
					}
					else
					{
						group[i].BackColor = Color.Lime;
					}
				}
				else if(i == 10 || i == 11 || i == 12 || i == 13 || i == 14 || i == 15 || i == 16 || i == 17 || i == 18 || i == 19)//スリーブ毎のショット数
				{
					if(label_info[i].Text.IndexOf("要注意") >= 0)
					{
						group[i].BackColor = Color.Tomato;
					}
					else if(label_info[i].Text.IndexOf("注意") >= 0)
					{
						group[i].BackColor = Color.Yellow;
					}
					else if(label_info[i].Text == "-")
					{
						group[i].BackColor = Color.White;
					}
					else
					{
						group[i].BackColor = Color.Lime;
					}
				}
				else
				{
					if(label_info[i].Text == "-")
					{
						group[i].BackColor = Color.White;
					}
					else
					{
						group[i].BackColor = Color.Lime;
					}
				}
			}

			timer1.Enabled = true;
		}

        private void timer1_Tick(object sender, EventArgs e)
        {
			for(int i = 0; i < group.Length; i++)
			{
				if(i == 1 && (label_info[i].Text == "準備中" || label_info[i].Text == "停止中"))//稼働状況
				{
					Color color_fore = group[i].ForeColor;
					group[i].ForeColor = Color.FromArgb(color_fore.ToArgb() ^ 0xFFFFFF);//反転

					Color color_back = group[i].BackColor;
					group[i].BackColor = Color.FromArgb(color_back.ToArgb() ^ 0xFFFFFF);//反転
				}
				else if(i == 3)
				{
					string kadou = label_info[i].Text.Substring(0, (label_info[i].Text.Length - 1));
					double kadou_double = double.Parse(kadou);
					if(kadou_double < Form1.SETDATA.kadouYellow)
					{
						Color color_fore = group[i].ForeColor;
						group[i].ForeColor = Color.FromArgb(color_fore.ToArgb() ^ 0xFFFFFF);//反転

						Color color_back = group[i].BackColor;
						group[i].BackColor = Color.FromArgb(color_back.ToArgb() ^ 0xFFFFFF);//反転
					}
					else if(Form1.SETDATA.kadouYellow <= kadou_double && kadou_double < Form1.SETDATA.kadouRed)
					{
						Color color_fore = group[i].ForeColor;
						group[i].ForeColor = Color.FromArgb(color_fore.ToArgb() ^ 0xFFFFFF);//反転

						Color color_back = group[i].BackColor;
						group[i].BackColor = Color.FromArgb(color_back.ToArgb() ^ 0xFFFFFF);//反転
					}
				}
				else if(i == 6 && label_info[i].Text.IndexOf("注意") >= 0)//加圧時間
				{
					Color color_fore = group[i].ForeColor;
					group[i].ForeColor = Color.FromArgb(color_fore.ToArgb() ^ 0xFFFFFF);//反転

					Color color_back = group[i].BackColor;
					group[i].BackColor = Color.FromArgb(color_back.ToArgb() ^ 0xFFFFFF);//反転
				}
				else if(i == 10 || i == 11 || i == 12 || i == 13 || i == 14 || i == 15 || i == 16 || i == 17 || i == 18 || i == 19)//スリーブ毎のショット数
				{
					if(label_info[i].Text.IndexOf("要注意") >= 0 || label_info[i].Text.IndexOf("注意") >= 0)
					{
						Color color_fore = group[i].ForeColor;
						group[i].ForeColor = Color.FromArgb(color_fore.ToArgb() ^ 0xFFFFFF);//反転

						Color color_back = group[i].BackColor;
						group[i].BackColor = Color.FromArgb(color_back.ToArgb() ^ 0xFFFFFF);//反転
					}
				}
			}
        }

        private void self_term_timer_Tick(object sender, EventArgs e)
        {
			self_term_timer.Enabled = false;
			this.Close();
        }
    }
}
