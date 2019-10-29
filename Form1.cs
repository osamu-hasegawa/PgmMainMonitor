using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using System.Xml.Serialization;
using System.Diagnostics;
using System.Collections;

namespace PgmMainMonitor
{
    public partial class Form1 : Form
    {
//PCがスリープ状態に入らないようする start
        #region Win32 API
        [FlagsAttribute]
        public enum ExecutionState : uint
        {
            // 関数が失敗した時の戻り値
            Null = 0,
            // スタンバイを抑止(Vista以降は効かない？)
            SystemRequired = 1,
            // 画面OFFを抑止
            DisplayRequired = 2,
            // 効果を永続させる。ほかオプションと併用する。
            Continuous = 0x80000000,
        }

        [DllImport("user32.dll")]
        extern static uint SendInput(
            uint nInputs,   // INPUT 構造体の数(イベント数)
            INPUT[] pInputs,   // INPUT 構造体
            int cbSize     // INPUT 構造体のサイズ
            );

        [StructLayout(LayoutKind.Sequential)]  // アンマネージ DLL 対応用 struct 記述宣言
        struct INPUT
        {
            public int type;  // 0 = INPUT_MOUSE(デフォルト), 1 = INPUT_KEYBOARD
            public MOUSEINPUT mi;
            // Note: struct の場合、デフォルト(パラメータなしの)コンストラクタは、
            //       言語側で定義済みで、フィールドを 0 に初期化する。
        }

        [StructLayout(LayoutKind.Sequential)]  // アンマネージ DLL 対応用 struct 記述宣言
        struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public int mouseData;  // amount of wheel movement
            public int dwFlags;
            public int time;  // time stamp for the event
            public IntPtr dwExtraInfo;
            // Note: struct の場合、デフォルト(パラメータなしの)コンストラクタは、
            //       言語側で定義済みで、フィールドを 0 に初期化する。
        }

        // dwFlags
        const int MOUSEEVENTF_MOVED = 0x0001;
        const int MOUSEEVENTF_LEFTDOWN = 0x0002;  // 左ボタン Down
        const int MOUSEEVENTF_LEFTUP = 0x0004;  // 左ボタン Up
        const int MOUSEEVENTF_RIGHTDOWN = 0x0008;  // 右ボタン Down
        const int MOUSEEVENTF_RIGHTUP = 0x0010;  // 右ボタン Up
        const int MOUSEEVENTF_MIDDLEDOWN = 0x0020;  // 中ボタン Down
        const int MOUSEEVENTF_MIDDLEUP = 0x0040;  // 中ボタン Up
        const int MOUSEEVENTF_WHEEL = 0x0080;
        const int MOUSEEVENTF_XDOWN = 0x0100;
        const int MOUSEEVENTF_XUP = 0x0200;
        const int MOUSEEVENTF_ABSOLUTE = 0x8000;

        const int screen_length = 0x10000;  // for MOUSEEVENTF_ABSOLUTE
        [DllImport("kernel32.dll")]
        static extern ExecutionState SetThreadExecutionState(ExecutionState esFlags);
        #endregion
//PCがスリープ状態に入らないようする end

        public struct SleeveList
        {
            public string sleeveNumber;
            public string shotCount;
            public DateTime workDt;
        }
		List<SleeveList> list = new List<SleeveList>();

        private DataTable table;
		string seikeiki, kadou_jyoukyou, seikei_type, kadou_ritsu, seihin_mei, niku_up, niku_dn, shime_jyoukyou, budomari, seikei_su, seikei_nichiji, kaatsu_jikan, housha_ritsu, tantousha = "";

        string[] headerName = {"成\r\n型\r\n機","状況","ﾀｲﾌﾟ","稼\r\n働\r\n率","型式名","最終成型日時","加圧時間\r\n[上限-下限]","ﾀｸﾄ","放\r\n射\r\n率","最大\r\nｼｮｯﾄ\r\n数","SL①\r\n(ｼｮｯﾄ数)","SL②\r\n(ｼｮｯﾄ数)","SL③\r\n(ｼｮｯﾄ数)","SL④\r\n(ｼｮｯﾄ数)","SL⑤\r\n(ｼｮｯﾄ数)","SL⑥\r\n(ｼｮｯﾄ数)","SL⑦\r\n(ｼｮｯﾄ数)","SL⑧\r\n(ｼｮｯﾄ数)","SL⑨\r\n(ｼｮｯﾄ数)","SL⑩\r\n(ｼｮｯﾄ数)","担\r\n当","締\r\nめ","歩\r\n留","成\r\n型","良\r\n品","不\r\n良"};

        string[] machineName = 
		{"LS1","LS2","LS3","LS4","LS5","LS6","LS7","LS8","LS9","LS10","LS11","LS12","LS14","LS15","LS16","LS17","LS18","LS19","LS20","LS21","LS22","LS23",
		"NQD1","NQD2","NQD3","NQD5","NQD6","NQD7","NQD8","NQD9","NQD10",
		"HS1","HS2","HS3","HS4","HS5","HS6","HS7","HS8","HS9"};

        string[] machineIpAddress = {"192.168.0.3","192.168.0.4","192.168.0.5","192.168.0.6","192.168.0.7","192.168.0.8","192.168.0.9","192.168.0.10",
		"192.168.0.11","192.168.0.12","192.168.0.13","192.168.0.14","192.168.0.15","192.168.0.16","192.168.0.17","192.168.0.18",
		"192.168.0.19","192.168.0.20","192.168.0.21","192.168.0.22","192.168.0.23","192.168.0.24",
		"192.168.0.25","192.168.0.26","192.168.0.27","192.168.0.28","192.168.0.29","192.168.0.30","192.168.0.31","192.168.0.32","192.168.0.33",
		"192.168.0.34","192.168.0.35","192.168.0.36","192.168.0.37","192.168.0.38","192.168.0.39","192.168.0.40","192.168.0.41","192.168.0.42"};

		public DateTime backFileTime;
		public string backCsvFile = "";

		public DateTime backDailyTime;
		public string backDailyFile = "";

		bool [] isShimeta = {false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false,
    						 false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false};

        Form2 form2 = null;
        static public SYSSET SETDATA = new SYSSET();

        public bool isRemote = true;
        public bool isFirstRead = false;
		public int max_sleeve_count = 0;

        public Form1()
        {
            InitializeComponent();
			ReadDataFromXml();

			//DataGridViewの画面ちらつきをおさえるため、DoubleBufferedを有効にする
			//DataGirdViewのTypeを取得
			System.Type dgvtype = typeof(DataGridView);
			//プロパティ設定の取得
			System.Reflection.PropertyInfo dgvPropertyInfo = dgvtype.GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
			//対象のDataGridViewにtrueをセットする
			dgvPropertyInfo.SetValue(dataGridView1, true, null);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
			//Windowサイズ確認
            int wd = this.Width;
            int ht = this.Height;
            int dgv_wd = this.dataGridView1.Width;
            int dgv_ht = this.dataGridView1.Height;

            this.Width = SETDATA.windowWidth;
            this.Height = SETDATA.windowHeight;
            this.dataGridView1.Width = SETDATA.dataGridViewWidth;
            this.dataGridView1.Height = SETDATA.dataGridViewHeight;

			//IPアドレスから現場LOCALでの接続か、社内LANかを判断する
			//ホスト名を取得
			string hostname = System.Net.Dns.GetHostName();
			//ホスト名からIPアドレスを取得
			System.Net.IPAddress[] addr_arr = System.Net.Dns.GetHostAddresses(hostname);
			foreach(System.Net.IPAddress addr in addr_arr)
			{
				string addr_str = addr.ToString();
				//IPv4 && "10."で始まれば社内LAN
				if ( addr_str.IndexOf( "." ) > 0 && addr_str.StartsWith("10.") )
				{
					isRemote = false;
					break;
				}
			}

            table = new DataTable("Table");
 
			for(int i = 0; i < headerName.Length; i++)
			{
				table.Columns.Add(headerName[i]);
			}

			for(int i = 0; i < machineName.Length; i++)
			{
				table.Rows.Add(machineName[i], "", "", "", "", "", "", "", "", "", "", "", "");
			}
 
            dataGridView1.DataSource = table;

            dataGridView1.DefaultCellStyle.Font = new Font("メイリオ", SETDATA.dataGridViewFontSize, FontStyle.Bold);
            dataGridView1.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("メイリオ", SETDATA.dataGridViewFontSize, FontStyle.Bold);

			dataGridView1.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
    
			//左端行を成型機タイプ毎に色分け
			for(int i = 0; i < machineName.Length; i++)
			{
				if(0 <= i && i < 22)
				{
					dataGridView1[0, i].Style.BackColor = Color.RoyalBlue;
				}
				else if(22 <= i && i < 31)
				{
					dataGridView1[0, i].Style.BackColor = Color.Orange;
				}
				else if(31 <= i && i < 40)
				{
					dataGridView1[0, i].Style.BackColor = Color.Violet;
				}
			}

            //列ヘッダの背景色の変更
			dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Aqua;

            //ソートさせない
            foreach (DataGridViewColumn c in dataGridView1.Columns)
			{
			    c.SortMode = DataGridViewColumnSortMode.NotSortable;
			}

			//初期化
			for(int i = 0; i < dataGridView1.Rows.Count; i++)
			{
				if(i == 3)//LS4(Windows ME PC)
				{
					dataGridView1.Rows[i].Cells[1].Value = "PC対象外";
					dataGridView1.Rows[i].Cells[1].Style.BackColor = Color.Gray;
					continue;
				}
	
				dataGridView1.Rows[i].Cells[21].Value = "未";
				dataGridView1.Rows[i].Cells[21].Style.BackColor = Color.Pink;
			}

			timer1.Interval = 10;
			timer1.Enabled = true;
			daily_timer.Enabled = true;
			kadou_timer.Enabled = true;
			alive_timer.Enabled = true;
			all_read_timer.Enabled = true;
        }

        static private void ReadDataFromXml()
		{
            SETDATA.load(ref Form1.SETDATA);
		}

		static public void WriteDataToXml()
		{
            SETDATA.save(Form1.SETDATA);
		}

		public class SYSSET:System.ICloneable
		{
			public int windowWidth;
			public int windowHeight;
            public int dataGridViewWidth;
            public int dataGridViewHeight;
            public int dataGridViewFontSize;
			public int kadouDelay;
			public int stopDelay;
			public int kadouRed;
			public int kadouYellow;
			public int shotRed;
			public int shotYellow;

            public bool load(ref SYSSET ss)
			{
                string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
                string path = stCurrentDir + "\\PgmMainMonitorSettingData.xml";
				bool ret = false;
				try {
					XmlSerializer sz = new XmlSerializer(typeof(SYSSET));
					System.IO.StreamReader fs = new System.IO.StreamReader(path, System.Text.Encoding.Default);
					SYSSET obj;
					obj = (SYSSET)sz.Deserialize(fs);
					fs.Close();
					obj = (SYSSET)obj.Clone();
					ss = obj;
					ret = true;
				}
				catch (Exception /*ex*/) {
				}
				return(ret);
			}

			public Object Clone()
			{
				SYSSET cln = (SYSSET)this.MemberwiseClone();
				return (cln);
			}

			public bool save(SYSSET ss)
			{
                string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
                string path = stCurrentDir + "\\PgmMainMonitorSettingData.xml";
				bool ret = false;
				try {
					XmlSerializer sz = new XmlSerializer(typeof(SYSSET));
					System.IO.StreamWriter fs = new System.IO.StreamWriter(path, false, System.Text.Encoding.Default);
					sz.Serialize(fs, ss);
					fs.Close();
					ret = true;
				}
				catch (Exception /*ex*/) {
				}
				return (ret);
			}
		}

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            string mes = "アプリを終了すると成型機からのデータを受け取れなくなります！" + "\r\n" + "本当に終了しますか？";
            DialogResult result = MessageBox.Show(mes, "PGMメインモニター", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
            if (result == DialogResult.Yes)
            {
            }
            else
            {
                e.Cancel = true;
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
			WriteDataToXml();
        }

		public void LogFileOut(string logMessage)
		{
            string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
			string path = stCurrentDir + "\\PgmMainMonitor.log";
			
			using(var sw = new System.IO.StreamWriter(path, true, System.Text.Encoding.Default))
			{
				sw.WriteLine($"{DateTime.Now.ToLongDateString()} {DateTime.Now.ToLongTimeString()}");
				sw.WriteLine($"  {logMessage}");
				sw.WriteLine ("--------------------------------------------------------------");
			}
		}

        private void dataGridView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            DataGridView.HitTestInfo info = ((DataGridView)sender).HitTest(e.X, e.Y);
            int pos = info.RowIndex;
            int col = info.ColumnIndex;

            if(pos == -1 || pos == 3)//header or Windows ME PC
            {
                return;
            }

            if ((string)dataGridView1.Rows[pos].Cells[1].Value == "")
            {
                return;
            }

			if(isRemote)
			{
				if(col == 0 && pos >= 0)//左端の「成型機名」をクリックした時
				{
					string ipaddr = machineIpAddress[pos];
                    string arg = "/password pgm " + ipaddr + ":5900";

					Process.Start("C:\\Program Files\\uvnc bvba\\UltraVNC\\vncviewer.exe", arg);
	                return;
				}
			}

            if (form2 == null || form2.IsDisposed)
			{
				form2 = new Form2();

				string info_str = "";
				for(int i = 0; i < 26; i++)
				{
					string data = (string)dataGridView1.Rows[pos].Cells[i].Value;
					if(i == 0)
					{
						info_str = data;
						continue;
					}

					info_str += "," + data;
				}
				form2.SetInfo(info_str);
                form2.ShowDialog();
                form2 = null;
            }

        }

        private void daily_timer_Tick(object sender, EventArgs e)
        {
			string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
			string csv_path = stCurrentDir + "\\work\\daily";

            DateTime dt = DateTime.Now;
            string today_str = dt.ToString("yyyyMMdd");
            string day_str = dt.ToString("MMdd");
            csv_path += "\\" + today_str;

			if(isRemote)
			{
				string net_path = "\\192.168.0.2\\Public\\work\\daily";
//	            string net_path = "\\ts-xhl5A9\\share\\ﾊﾞｯｸｱｯﾌﾟ臨時\\永田";
	            net_path = "\\" + net_path;
	            net_path += "\\" + today_str;

				if(!Directory.Exists(net_path))//サーバーに日付フォルダが存在するか
				{
					return;
				}

				//コピー先のディレクトリがないときは作る
				if(!System.IO.Directory.Exists(csv_path))
				{
					System.IO.Directory.CreateDirectory(csv_path);
				}

				try
				{
					DirectoryInfo sourceDirectory = new DirectoryInfo(net_path);
					DirectoryInfo destinationDirectory = new DirectoryInfo(csv_path);
					//サーバーのCSVファイルを全てコピーする
					foreach(FileInfo fileInfo in sourceDirectory.GetFiles()) 
					{
						//同じファイルが存在していたら、常に上書きする
						fileInfo.CopyTo(destinationDirectory.FullName + @"\" + fileInfo.Name, true);
					}
				}
	            catch(System.IO.IOException ex)
	            {
	                string errorStr = "他のアプリがdailyのCSVファイルを開いていてコピーに失敗した可能性があります";
	                System.Console.WriteLine(errorStr);
	                System.Console.WriteLine(ex.Message);
					LogFileOut(errorStr);
					return;
	            }

	        }

			if(!Directory.Exists(csv_path))//日付フォルダが存在するか
			{
				return;
			}

            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(csv_path);
            IEnumerable<System.IO.FileInfo> files = di.EnumerateFiles("*.csv", System.IO.SearchOption.TopDirectoryOnly);

			//最近更新されたファイルと更新時間の探索
			string strTime = "2019/01/01 01:23:45";
			DateTime lastDt = DateTime.Parse(strTime);
			DateTime destDt;
			string lastFileName = "";

            foreach(System.IO.FileInfo f in files)
            {
                destDt = System.IO.File.GetLastWriteTime(f.FullName);
				if(lastDt < destDt)
				{
					lastDt = destDt;
					lastFileName = f.Name;
				}
			}

			string str_back = backDailyTime.ToString();
			string str_last = lastDt.ToString();
            if (backDailyFile == lastFileName && str_back == str_last)//最新ファイル名が同一かつ更新日時が同一の場合
			{
				return;//更新なしと判断
			}
			
			//更新あればファイル名、時間を更新
			backDailyFile = lastFileName;
			backDailyTime = lastDt;

			//更新があった場合、全CSVを読み込む
			StreamReader sr = null;
            string srcFile = "";
            string dstFile = "";

			string sort_path = csv_path + "\\";

			try
			{
				foreach (string file in Directory.EnumerateFiles(sort_path).OrderBy(f => File.GetLastWriteTime(f)))
				{
					string file_name = Path.GetFileName(file);

//	            foreach(System.IO.FileInfo f in files)
//	            {
					//成型機の入力位置を決定
					int pos = 0;
					for(int i = 0; i < machineName.Length; i++)
					{
						string search_word = machineName[i].ToString() + "号機";
						if(file_name.IndexOf(search_word) >= 0)
						{
							pos = i;
							break;
						}
					}

		            //ファイルを開く
					FileStream fp;
		            srcFile = file;
		            dstFile = file + ".tmp";
					try
					{
						//最新のCSVファイルをWORK用一時ファイルにコピーし、以降は一時ファイルを扱う。
						//最新CSVファイルは成型機VBアプリもアクセスするため、競合区間を最小限にするため。
						File.Copy(@srcFile, @dstFile);
			            fp = new FileStream(dstFile, FileMode.Open, FileAccess.Read);
			            sr = new StreamReader(fp, System.Text.Encoding.GetEncoding("Shift_JIS"));

						string line = "";
						int allCount = 0;
			            while(sr.EndOfStream == false)
			            {
			                line = sr.ReadLine();

							if(allCount == 0)
							{
								allCount++;
								continue;
							}
							
				            string[] strline = line.Split(',');

//							if(strline.Length > 7)//ガード：項目数チェック
//							{
//								break;
//							}

							//DataGridViewに入力
							dataGridView1.Rows[pos].Cells[22].Value = strline[3];
							dataGridView1.Rows[pos].Cells[23].Value = strline[4];
							dataGridView1.Rows[pos].Cells[24].Value = strline[5];
							dataGridView1.Rows[pos].Cells[25].Value = strline[6];

							dataGridView1.Rows[pos].Cells[21].Value = day_str + "済";
							dataGridView1.Rows[pos].Cells[21].Style.BackColor = Color.Lime;
							isShimeta[pos] = true;
							shime_release_timer.Enabled = true;

							allCount++;
				        }

						if(sr != null)
						{
							sr.Close();
						}
						File.Delete(@dstFile);//一時ファイルは削除
					}
					catch (System.IO.IOException ex)
					{
						string errorStr = "DailyのCSVファイルコピー失敗またはCSVファイルが開けなかった可能性があります";
					    System.Console.WriteLine(errorStr);
					    System.Console.WriteLine(ex.Message);
						
						if(sr != null)
						{
							sr.Close();
						}
						File.Delete(@dstFile);//一時ファイルは削除
						LogFileOut(errorStr);
					    continue;
					}
				}

            }
            catch (System.IO.IOException ex)
            {
                string errorStr = "他のアプリがDailyのCSVファイルを開いている可能性があります";
                System.Console.WriteLine(errorStr);
                System.Console.WriteLine(ex.Message);
				LogFileOut(errorStr);
                return;
            }
            finally
            {
			}

        }

        private void shime_release_timer_Tick(object sender, EventArgs e)
        {
			string strTime = "15:00:00";
			DateTime limit_dt = DateTime.Parse(strTime);
			DateTime today_dt = DateTime.Now;

			if(limit_dt < today_dt)//本日の15時を過ぎたか
			{
				//初期化
				for(int i = 0; i < dataGridView1.Rows.Count; i++)
				{
					if(i == 3)//LS4(Windows ME PC)
					{
						continue;
					}

		            if ((string)dataGridView1.Rows[i].Cells[1].Value == "")
		            {
		                continue;
		            }

					dataGridView1.Rows[i].Cells[21].Value = "未";
					dataGridView1.Rows[i].Cells[21].Style.BackColor = Color.Pink;
					isShimeta[i] = false;
				}
				shime_release_timer.Enabled = false;
			}

        }

        private void alive_timer_Tick(object sender, EventArgs e)
        {
            //画面暗転阻止
            SetThreadExecutionState(ExecutionState.DisplayRequired);

            // ドラッグ操作の準備 (struct 配列の宣言)
            INPUT[] input = new INPUT[1];  // イベントを格納

            // ドラッグ操作の準備 (イベントの定義 = 相対座標へ移動)
            input[0].mi.dx = 0;  // 相対座標で0　つまり動かさない
            input[0].mi.dy = 0;  // 相対座標で0 つまり動かさない
            input[0].mi.dwFlags = MOUSEEVENTF_MOVED;

            // ドラッグ操作の実行 (イベントの生成)
            SendInput(1, input, Marshal.SizeOf(input[0]));
        }

        private void kadou_timer_Tick(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
				//稼働状況
				string kadou = (string)dataGridView1.Rows[i].Cells[1].Value;
                if (kadou == "準備中" || kadou == "停止中")
                {
                    Color current_color = dataGridView1.Rows[i].Cells[1].Style.BackColor;
                    dataGridView1.Rows[i].Cells[1].Style.BackColor = Color.FromArgb(current_color.ToArgb() ^ 0xFFFFFF);//反転

                    Color current_forec = dataGridView1.Rows[i].Cells[1].Style.ForeColor;
                    dataGridView1.Rows[i].Cells[1].Style.ForeColor = Color.FromArgb(current_forec.ToArgb() ^ 0xFFFFFF);//反転
                }

				//稼働率
				string kadou_ritsu = (string)dataGridView1.Rows[i].Cells[3].Value;
                if(kadou == "成型中")
                {
					kadou_ritsu = kadou_ritsu.Substring(0, (kadou_ritsu.Length - 1));
					double cal_kadou = double.Parse(kadou_ritsu);
					if((cal_kadou < SETDATA.kadouRed) || (SETDATA.kadouRed <= cal_kadou && cal_kadou < SETDATA.kadouYellow))
					{
	                    Color current_color = dataGridView1.Rows[i].Cells[3].Style.BackColor;
	                    dataGridView1.Rows[i].Cells[3].Style.BackColor = Color.FromArgb(current_color.ToArgb() ^ 0xFFFFFF);//反転

	                    Color current_forec = dataGridView1.Rows[i].Cells[3].Style.ForeColor;
	                    dataGridView1.Rows[i].Cells[3].Style.ForeColor = Color.FromArgb(current_forec.ToArgb() ^ 0xFFFFFF);//反転
	                }
	            }

				//加圧時間
				string kaatsu = (string)dataGridView1.Rows[i].Cells[6].Value;
                if (kaatsu.IndexOf("注意") >= 0)
                {
                    Color current_color = dataGridView1.Rows[i].Cells[6].Style.BackColor;
                    dataGridView1.Rows[i].Cells[6].Style.BackColor = Color.FromArgb(current_color.ToArgb() ^ 0xFFFFFF);//反転

                    Color current_forec = dataGridView1.Rows[i].Cells[6].Style.ForeColor;
                    dataGridView1.Rows[i].Cells[6].Style.ForeColor = Color.FromArgb(current_forec.ToArgb() ^ 0xFFFFFF);//反転
                }

				//スリーブとｼｮｯﾄ数
				if(kadou == "" || kadou == "PC対象外")
				{
					continue;
				}
				for(int j = 0; j < 10; j++)
				{
					string sl = (string)dataGridView1.Rows[i].Cells[10 + j].Value;
					if(sl != "" && sl != "-")
					{
						if(sl.IndexOf("注意") >= 0 || sl.IndexOf("要注意") >= 0)
						{
							Color current_color = dataGridView1.Rows[i].Cells[10 + j].Style.BackColor;
							dataGridView1.Rows[i].Cells[10 + j].Style.BackColor = Color.FromArgb(current_color.ToArgb() ^ 0xFFFFFF);//反転

							Color current_forec = dataGridView1.Rows[i].Cells[10 + j].Style.ForeColor;
							dataGridView1.Rows[i].Cells[10 + j].Style.ForeColor = Color.FromArgb(current_forec.ToArgb() ^ 0xFFFFFF);//反転
						}
					}
				}

            }
        }

        private void all_read_timer_Tick(object sender, EventArgs e)
        {
			isFirstRead = false;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
			string stCurrentDir = System.IO.Directory.GetCurrentDirectory();
			string csv_path = stCurrentDir + "\\work\\period";
			if(isRemote)
			{
	            string net_path = "\\192.168.0.2\\Public\\work\\period";
//	            string net_path = "\\ts-xhl5A9\\share\\ﾊﾞｯｸｱｯﾌﾟ臨時\\永田";
	            net_path = "\\" + net_path;

				//コピー先のディレクトリがないときは作る
				if(!System.IO.Directory.Exists(csv_path))
				{
					System.IO.Directory.CreateDirectory(csv_path);
				}

				try
				{
					DirectoryInfo sourceDirectory = new DirectoryInfo(net_path);
					DirectoryInfo destinationDirectory = new DirectoryInfo(csv_path);
					//サーバーのCSVファイルを全てコピーする
					foreach(FileInfo fileInfo in sourceDirectory.GetFiles()) 
					{
						//同じファイルが存在していたら、常に上書きする
						fileInfo.CopyTo(destinationDirectory.FullName + @"\" + fileInfo.Name, true);
					}
				}
	            catch(System.IO.IOException ex)
	            {
	                string errorStr = "他のアプリがCSVファイルを開いていてコピーに失敗した可能性があります";
	                System.Console.WriteLine(errorStr);
	                System.Console.WriteLine(ex.Message);
					LogFileOut(errorStr);
					return;
	            }
			}

			//最近更新されたファイルと更新時間の探索
			string strTime = "2019/01/01 01:23:45";
			DateTime lastDt = DateTime.Parse(strTime);
			DateTime destDt;
			string lastFileName = "";

            System.IO.DirectoryInfo di = new System.IO.DirectoryInfo(csv_path);
            IEnumerable<System.IO.FileInfo> files = di.EnumerateFiles("*.csv", System.IO.SearchOption.TopDirectoryOnly);

            SortedList slist = new SortedList();//ログファイル名と更新時間のペア

            foreach (System.IO.FileInfo f in files)
            {
                destDt = System.IO.File.GetLastWriteTime(f.FullName);
				if(lastDt < destDt)
				{
					lastDt = destDt;
					lastFileName = f.Name;
				}

                //ログファイルのリストに入れる(＆昇順に並べ替えてくれている)⇒高速化
                if (!slist.ContainsKey(destDt))
                {
                    slist.Add(destDt, f.FullName);
                }
            }

            string str_back = backFileTime.ToString();
			string str_last = lastDt.ToString();
            if (backCsvFile == lastFileName && str_back == str_last)//最新ファイル名が同一かつ更新日時が同一の場合
			{
				return;//更新なしと判断
			}

			timer1.Enabled = false;//timerを一時停止

            //行や列を追加したり、セルに値を設定するときは、自動サイズ設定しない。
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
			
			//更新あればファイル名、時間を更新
			backCsvFile = lastFileName;
			backFileTime = lastDt;


			//更新があった場合、全CSVを読み込む
			StreamReader sr = null;
            string srcFile = "";
            string dstFile = "";

			//全ファイル(CSV)を取得し最終書き込み日時順に並べ替える
			//(型式変更があった場合、同一成型機名でも型式名が異なるCSVがある為、同一成型機名でも更新日時の新しい方に書き換える)
			string sort_path = csv_path + "\\";

			try
			{
                //foreach (string file in Directory.EnumerateFiles(sort_path).OrderBy(f => File.GetLastWriteTime(f)))//⇒OrderByは遅い？
                foreach(object item in slist.Values)
				{
					string file = item.ToString();
					if(isFirstRead)
					{
	                    DateTime fd = System.IO.File.GetLastWriteTime(file);
						DateTime nd = DateTime.Now;
	                    nd = nd.AddMinutes(-5);
	                    if(nd > fd)//現在よりも基準時間以内のファイルのみ更新する。それ以外は弾く
	                    {
	                        continue;
	                    }
	                }

                    //成型機を判定
                    bool isMulti = false;
					bool isLS_NQD = false;
					bool isHS = false;
					string file_name = Path.GetFileName(file);
					if(file_name.IndexOf("_マルチCav") >= 0)//マルチか？
					{
						isMulti = true;
					}
					if(file_name.IndexOf("_LS") >= 0)
					{
						isLS_NQD = true;
					}
					else if(file_name.IndexOf("_NQD") >= 0)
					{
						isLS_NQD = true;
					}
					else if(file_name.IndexOf("_HS") >= 0)
					{
						isHS = true;
					}

		            //ファイルを開く
					FileStream fp;
		            srcFile = file;
		            dstFile = file + ".tmp";
					try
					{
						//最新のCSVファイルをWORK用一時ファイルにコピーし、以降は一時ファイルを扱う。
						//最新CSVファイルは成型機VBアプリもアクセスするため、競合区間を最小限にするため。
						File.Copy(@srcFile, @dstFile);
			            fp = new FileStream(dstFile, FileMode.Open, FileAccess.Read);
			            sr = new StreamReader(fp, System.Text.Encoding.GetEncoding("Shift_JIS"));
					}
					catch (System.IO.IOException ex)
					{
						string errorStr = "CSVファイルコピー失敗またはCSVファイルが開けなかった可能性があります";
					    System.Console.WriteLine(errorStr);
					    System.Console.WriteLine(ex.Message);
						
						if(sr != null)
						{
							sr.Close();
						}
						File.Delete(@dstFile);//一時ファイルは削除
						LogFileOut(errorStr);
					    continue;
					}

					//締めサインの範囲を探す
					DateTime dt_now = DateTime.Now;
					string now_sta = dt_now.ToString("MMdd");
					now_sta += "開始";
					string now_end = dt_now.ToString("MMdd");
					now_end += "終了";
					string sta_time = "";
					string end_time = "";
					string text = File.ReadAllText(dstFile, Encoding.GetEncoding("Shift_JIS"));
					string[] rows = text.Trim().Replace("\r","").Split('\n');
					string[] strline = null;
					for(int i = 0; i < rows.Length; i++)
					{
						if(i == 0)
						{
							continue;
						}

			            strline = rows[i].Split(',');

//						if(strline.Length > 33)//ガード：項目数チェック
//						{
//							break;
//						}

						if(strline[5] == now_sta)
						{
							if(isLS_NQD)
							{
								sta_time = strline[20];
							}
							if(isHS)
							{
								sta_time = strline[6];
							}
						}
						else if(strline[5] == now_end)
						{
							if(isLS_NQD)
							{
								end_time = strline[20];
							}
							if(isHS)
							{
								end_time = strline[6];
							}
						}

					}

					DateTime start;
					DateTime end;
					if(sta_time == "")
					{
						string ss = "00:00:00";
						start = DateTime.Parse(ss);
					}
					else
					{
						start = DateTime.Parse(sta_time);
					}

					if(end_time == "")
					{
						string se = "23:59:59";
	                    end = DateTime.Parse(se);
	                }
					else
					{
						end = DateTime.Parse(end_time);
					}

					int allCount = 0;
					int slCount = 0;
					int seikei_ok_count = 0;
					int seikei_ng_count = 0;
					seikei_su = "0";
					list.Clear();

					for(int ll = 0; ll < rows.Length; ll++)
					{
						if(allCount == 0)
						{
							allCount++;
							continue;
						}

			            strline = rows[ll].Split(',');

						//稼働率算出：当日のスリーブがダミー以外の比率として出す
						if(isLS_NQD)
						{
							string sl = strline[6];//スリーブ
							DateTime tar = DateTime.Parse(strline[20]);
							if(start <= tar && tar <= end)
							{
								if(sl != "0" && sl != "" && sl != "D" && sl != "SD" && sl != "先" && sl != "空")
								{
                                    slCount++;

									if(!isMulti)
									{

//										if(strline.Length > 33)//ガード：項目数チェック
//										{
//											break;
//										}

										//スリーブの種類をリストに登録(重複しない)
					                    if(list.Find(m => m.sleeveNumber == sl).sleeveNumber != sl)
						                {
					                        SleeveList ls = new SleeveList();
					                        ls.sleeveNumber = sl;
										    ls.shotCount = strline[28];
										    ls.workDt = tar;
	                                        list.Add(ls);
	                                    }
						                else//重複していたら、ショット数の大きい方へ入れ替え
						                {
											for(int j = 0; j < list.Count; j++)
											{
												if(list[j].sleeveNumber == sl)
												{
													if(strline[28] != "")
													{
														int sc = int.Parse(strline[28]);
														int ls_sc = int.Parse(list[j].shotCount);
	//													if(ls_sc < sc)
														if(list[j].workDt < tar)
														{
															SleeveList tmpList = list[j];
															tmpList.shotCount = sc.ToString();
															list[j] = tmpList;
														}
													}
												}
											}
										}
										
										if(strline[20] == "OK")
										{
											seikei_ok_count++;
										}
										else if(strline[20] == "NG")
										{
											seikei_ng_count++;
										}
									}
									else
									{

//										if(strline.Length > 48)//ガード：項目数チェック
//										{
//											break;
//										}

										//スリーブの種類をリストに登録(重複しない)
					                    if(list.Find(m => m.sleeveNumber == sl).sleeveNumber != sl)
						                {
					                        SleeveList ls = new SleeveList();
					                        ls.sleeveNumber = sl;
										    ls.shotCount = strline[43];
										    ls.workDt = tar;
	                                        list.Add(ls);
	                                    }
						                else//重複していたら、ショット数の大きい方へ入れ替え
						                {
											for(int j = 0; j < list.Count; j++)
											{
												if(list[j].sleeveNumber == sl)
												{
													if(strline[43] != "")
													{
														int sc = int.Parse(strline[43]);
														int ls_sc = int.Parse(list[j].shotCount);
	//													if(ls_sc < sc)
														if(list[j].workDt < tar)
														{
															SleeveList tmpList = list[j];
															tmpList.shotCount = sc.ToString();
															list[j] = tmpList;
														}
													}
												}
											}
										}

										for(int jj = 0; jj < 16; jj+=3)
										{
											if(strline[24 + jj] == "OK")
											{
												seikei_ok_count++;
											}
											else if(strline[24 + jj] == "NG")
											{
												seikei_ng_count++;
											}
										}

									}

								}
					            allCount++;
							}
						}
						else if(isHS)
						{
							string sl = strline[11];//スリーブ
							DateTime tar = DateTime.Parse(strline[6]);
							if(start <= tar && tar <= end)
							{
								if(sl != "0" && sl != "")
								{
									slCount++;

									if(!isMulti)
									{

//										if(strline.Length > 70)//ガード：項目数チェック
//										{
//											break;
//										}

										//スリーブの種類をリストに登録(重複しない)
					                    if(list.Find(m => m.sleeveNumber == sl).sleeveNumber != sl)
						                {
					                        SleeveList ls = new SleeveList();
					                        ls.sleeveNumber = sl;
										    ls.shotCount = strline[65];
										    ls.workDt = tar;
	                                        list.Add(ls);
	                                    }
						                else//重複していたら、ショット数の大きい方へ入れ替え
						                {
											for(int j = 0; j < list.Count; j++)
											{
												if(list[j].sleeveNumber == sl)
												{
													if(strline[65] != "")
													{
														int sc = int.Parse(strline[65]);
														int ls_sc = int.Parse(list[j].shotCount);
	//													if(ls_sc < sc)
														if(list[j].workDt < tar)
														{
															SleeveList tmpList = list[j];
															tmpList.shotCount = sc.ToString();
															list[j] = tmpList;
														}
													}
												}
											}
										}

										if(strline[60] == "OK")
										{
											seikei_ok_count++;
										}
										else if(strline[60] == "NG")
										{
											seikei_ng_count++;
										}
									}
									else
									{

//										if(strline.Length > 85)//ガード：項目数チェック
//										{
//											break;
//										}

										//スリーブの種類をリストに登録(重複しない)
					                    if(list.Find(m => m.sleeveNumber == sl).sleeveNumber != sl)
						                {
					                        SleeveList ls = new SleeveList();
					                        ls.sleeveNumber = sl;
										    ls.shotCount = strline[80];
										    ls.workDt = tar;
	                                        list.Add(ls);
	                                    }
						                else//重複していたら、ショット数の大きい方へ入れ替え
						                {
											for(int j = 0; j < list.Count; j++)
											{
												if(list[j].sleeveNumber == sl)
												{
													if(strline[80] != "")
													{
														int sc = int.Parse(strline[80]);
														int ls_sc = int.Parse(list[j].shotCount);
	//													if(ls_sc < sc)
														if(list[j].workDt < tar)
														{
															SleeveList tmpList = list[j];
															tmpList.shotCount = sc.ToString();
															list[j] = tmpList;
														}
													}
												}
											}
										}

										for(int jj = 0; jj < 16; jj+=3)
										{
											if(strline[60 + jj] == "OK")
											{
												seikei_ok_count++;
											}
											else if(strline[60 + jj] == "NG")
											{
												seikei_ng_count++;
											}
										}
									}


								}
					            allCount++;
							}
			            }
					}

					//最新のログを加工
		            string[] strResults = rows[rows.Length - 1].Split(',');

					//成型機に応じた処理に分岐
					seikeiki = strResults[2];//成型機名
//					seikeiki = seikeiki.Substring(0, (seikeiki.Length - 2));
					
					seihin_mei = strResults[3];//製品名
					shime_jyoukyou = strResults[5];//締めサイン

					bool isKaatsuNG = false;

					double kaatsu_double = 0;
					string kaatsu_up = "";
					string kaatsu_dn = "";
					int kaatsu_up_value = 0;
					int kaatsu_dn_value = 0;
					string limit_shot = "";
					string tact = "";
					
					if(isLS_NQD)
					{
						seikei_nichiji = strResults[20];//成型最終日時

						kaatsu_jikan = strResults[13];//加圧時間
						kaatsu_jikan = kaatsu_jikan.Substring(0, (kaatsu_jikan.Length - 1));//'s'を消去
						int zero_pos = kaatsu_jikan.IndexOf("0");
						if(zero_pos >= 0)
						{
							if(zero_pos == 0)
							{
								kaatsu_jikan = kaatsu_jikan.Substring(1, 2);//先頭の"0"を消去 "012"→"12"
								zero_pos = kaatsu_jikan.IndexOf("0");
								if(zero_pos == 0)
								{
									kaatsu_jikan = kaatsu_jikan.Substring(1, 1);//先頭の"0"を消去 "02"→"2"
								}
							}
						}
						kaatsu_jikan += ".0";
						kaatsu_double = double.Parse(kaatsu_jikan);
						kaatsu_jikan += "秒";
						
						tact = strResults[15];//ﾀｸﾄ

						if(!isMulti)//シングルLSかシングルNQD
						{

//							if(strResults.Length > 33)//ガード：項目数チェック
//							{
//								break;
//							}

							niku_up = strResults[21];//肉厚上限
							niku_dn = strResults[23];//肉厚下限
							housha_ritsu = strResults[26];//放射率
							tantousha = strResults[27];//担当者

							kaatsu_up = strResults[30];//加圧時間上限
							kaatsu_dn = strResults[31];//加圧時間下限
							limit_shot = strResults[32];//限界ｼｮｯﾄ数
						}
						else//マルチLSかマルチNQD
						{

//							if(strResults.Length > 48)//ガード：項目数チェック
//							{
//								break;
//							}

							niku_up = strResults[21];//肉厚上限
							niku_dn = strResults[22];//肉厚下限
							housha_ritsu = strResults[41];//放射率
							tantousha = strResults[42];//担当者

							kaatsu_up = strResults[45];//加圧時間上限
							kaatsu_dn = strResults[46];//加圧時間下限
							limit_shot = strResults[47];//限界ｼｮｯﾄ数
						}
					}
					else if(isHS)
					{
						seikei_nichiji = strResults[6];//成型最終日時

						kaatsu_jikan = strResults[18];//加圧時間

						kaatsu_double = double.Parse(kaatsu_jikan);
						kaatsu_jikan += "秒";
						housha_ritsu = "-";

						tact = strResults[10];//ﾀｸﾄ
#if false//HSは秒で出力されるので分に換算
						double cal_tact = double.Parse(tact);
						int fun = (int)(cal_tact /= 60);
						int byo = (int)(cal_tact - (fun * 60));
						tact = fun + ":" + byo;
#endif

						if(!isMulti)//シングルHS
						{

//							if(strResults.Length > 70)//ガード：項目数チェック
//							{
//								break;
//							}

							niku_up = strResults[57];//肉厚上限
							niku_dn = strResults[59];//肉厚下限
							tantousha = strResults[63];//担当者

							kaatsu_up = strResults[67];//加圧時間上限
							kaatsu_dn = strResults[68];//加圧時間下限
							limit_shot = strResults[69];//限界ｼｮｯﾄ数
						}
						else//マルチHS
						{

//							if(strResults.Length > 85)//ガード：項目数チェック
//							{
//								break;
//							}

							niku_up = strResults[57];//肉厚上限
							niku_dn = strResults[58];//肉厚下限
							tantousha = strResults[78];//担当者

							kaatsu_up = strResults[82];//加圧時間上限
							kaatsu_dn = strResults[83];//加圧時間下限
							limit_shot = strResults[84];//限界ｼｮｯﾄ数
						}
					}

					if(kaatsu_up != "")
					{
						kaatsu_up_value = int.Parse(kaatsu_up);
					}
					if(kaatsu_dn != "")
					{
						kaatsu_dn_value = int.Parse(kaatsu_dn);
					}
					if(kaatsu_double < kaatsu_dn_value || kaatsu_up_value < kaatsu_double)
					{
						isKaatsuNG = true;
						kaatsu_jikan = "注意" + kaatsu_jikan;
					}

					kaatsu_jikan = kaatsu_jikan + "[" + kaatsu_up + "-" + kaatsu_dn + "]";
					kaatsu_up += "秒";
					kaatsu_dn += "秒";
//					tantousha += "さん";
					string niku_hani = "[" + niku_up + "-" + niku_dn + "]";


					//成型機の入力位置を決定
					int pos = 0;
					for(int i = 0; i < machineName.Length; i++)
					{
						if((machineName[i] + "号機") == seikeiki)
						{
							pos = i;
							break;
						}
					}

					seikeiki = seikeiki.Substring(0, (seikeiki.Length - 2));
					dataGridView1.Rows[pos].Cells[0].Value = seikeiki;

					DateTime dt = DateTime.Now;
					DateTime sn = DateTime.Parse(seikei_nichiji);

					DateTime cp = sn.AddHours(SETDATA.kadouDelay);
					DateTime sd = sn.AddHours(SETDATA.stopDelay);
					
					dataGridView1.Rows[pos].Cells[1].Style.ForeColor = Color.Black;

					if(dt < cp)//現在日時の方が最終成型日時＋準備中の境界値より新しい
					{
						kadou_jyoukyou = "成型中";
						dataGridView1.Rows[pos].Cells[1].Value = kadou_jyoukyou;
						dataGridView1.Rows[pos].Cells[1].Style.BackColor = Color.Lime;
					}
					else
					{
						if(dt > sd)//現在日時の方が最終成型日時＋停止中の境界値より新しい
						{

							kadou_jyoukyou = "停止中";
							dataGridView1.Rows[pos].Cells[1].Value = kadou_jyoukyou;
							dataGridView1.Rows[pos].Cells[1].Style.BackColor = Color.Tomato;
						}
						else
						{
							kadou_jyoukyou = "準備中";
							dataGridView1.Rows[pos].Cells[1].Value = kadou_jyoukyou;
							dataGridView1.Rows[pos].Cells[1].Style.BackColor = Color.Yellow;
						}
					}


					//マルチ／シングル
					Color type_color;
					if(!isMulti)
					{
						seikei_type = "ｼﾝｸﾞﾙ";
						type_color = Color.Khaki;
					}
					else
					{
						seikei_type = "ﾏﾙﾁ";
						type_color = Color.YellowGreen;
					}
					dataGridView1.Rows[pos].Cells[2].Value = seikei_type;
					dataGridView1.Rows[pos].Cells[2].Style.BackColor = type_color;

					//金型の稼働率
					double kadou = (double)((double)slCount / (double)allCount);
					double cal_kadou = kadou * 100;
					kadou_ritsu = string.Format("{0:p1}", kadou);
					dataGridView1.Rows[pos].Cells[3].Value = kadou_ritsu;

					//成型中の成型機で稼働率がXX%未満のものは背景色を変える
					dataGridView1.Rows[pos].Cells[3].Style.ForeColor = Color.Black;
					if(kadou_jyoukyou == "成型中")
					{
						if(cal_kadou < SETDATA.kadouRed)
						{
							dataGridView1.Rows[pos].Cells[3].Style.BackColor = Color.Tomato;
						}
						else if(SETDATA.kadouRed <= cal_kadou && cal_kadou < SETDATA.kadouYellow)
						{
							dataGridView1.Rows[pos].Cells[3].Style.BackColor = Color.Yellow;
						}
						else
						{
							dataGridView1.Rows[pos].Cells[3].Style.BackColor = Color.Lime;
						}
					}

					//歩留り
					double budo_calc = 0;
					if(slCount != 0)
					{
						budo_calc = (double)((double)seikei_ok_count / (double)slCount);
					}
					
					budomari = budo_calc.ToString();
                    budomari += "％";
					seikei_su = slCount.ToString();
					string sei_ok = seikei_ok_count.ToString();
					string sei_ng = seikei_ng_count.ToString();

					seikei_nichiji = seikei_nichiji.Substring(2, (seikei_nichiji.Length - 2));//2019 -> 19
					seikei_nichiji = seikei_nichiji.Substring(0, (seikei_nichiji.Length - 3));//12:23:45 -> 12:23

					//加圧時間
					dataGridView1.Rows[pos].Cells[6].Style.ForeColor = Color.Black;
					if(isKaatsuNG)
					{
						dataGridView1.Rows[pos].Cells[6].Style.BackColor = Color.Yellow;
					}
					else
					{
						dataGridView1.Rows[pos].Cells[6].Style.BackColor = Color.Lime;
					}

					//放射率
					if(housha_ritsu == "-")
					{
						dataGridView1.Rows[pos].Cells[8].Style.BackColor = Color.Gray;
					}

					//スリーブリスト
					for(int j = 0; j < list.Count; j++)
					{
						string shot = (string)list[j].shotCount;
						if(shot == "")
						{
							shot = "0";
						}
						string sl_shot = list[j].sleeveNumber + "(" + shot + "回)";

						if(limit_shot == "")
						{
							limit_shot = "0";
						}
						
						int tar_shot = int.Parse(shot);
						int lmt_shot = int.Parse(limit_shot);

						dataGridView1.Rows[pos].Cells[10 + j].Style.ForeColor = Color.Black;
						if((lmt_shot - SETDATA.shotRed) < tar_shot)//限界ショット数を超えたら
						{
							sl_shot = "要注意" + sl_shot;
							dataGridView1.Rows[pos].Cells[10 + j].Style.BackColor = Color.Tomato;
						}
						else if((lmt_shot - SETDATA.shotYellow) < tar_shot &&  tar_shot <= (lmt_shot - SETDATA.shotRed))//限界ショット数を超えたら
						{
							sl_shot = "注意" + sl_shot;
							dataGridView1.Rows[pos].Cells[10 + j].Style.BackColor = Color.Yellow;
						}
						else
						{
							dataGridView1.Rows[pos].Cells[10 + j].Style.BackColor = Color.Lime;
						}
						dataGridView1.Rows[pos].Cells[10 + j].Value = sl_shot;
					}
					//空きのスリーブはハイフォン表示
					if(list.Count < 10)
					{
						int df = 10 - list.Count;
						for(int j = 10; j > (10 - df); j--)
						{
							dataGridView1.Rows[pos].Cells[9 + j].Value = "-";
							dataGridView1.Rows[pos].Cells[9 + j].Style.BackColor = Color.White;
						}
					}
					
					if(max_sleeve_count < list.Count)//最大値を更新
					{
						max_sleeve_count = list.Count;
					}

					//DataGridViewに入力
					dataGridView1.Rows[pos].Cells[4].Value = seihin_mei;
//                    dataGridView1.Rows[pos].Cells[5].Value = niku_hani;
					dataGridView1.Rows[pos].Cells[5].Value = seikei_nichiji;
					dataGridView1.Rows[pos].Cells[6].Value = kaatsu_jikan;
					dataGridView1.Rows[pos].Cells[7].Value = tact;
					dataGridView1.Rows[pos].Cells[8].Value = housha_ritsu;
					dataGridView1.Rows[pos].Cells[9].Value = limit_shot;

					dataGridView1.Rows[pos].Cells[20].Value = tantousha;

					if(!isShimeta[pos])//まだ締めていなければ
					{
	                    dataGridView1.Rows[pos].Cells[22].Value = budomari;
						dataGridView1.Rows[pos].Cells[23].Value = seikei_su;
						dataGridView1.Rows[pos].Cells[24].Value = sei_ok;
						dataGridView1.Rows[pos].Cells[25].Value = sei_ng;
					}

					//一時ファイルを閉じる
					if(sr != null)
					{
						sr.Close();
					}
					File.Delete(@dstFile);//一時ファイルは削除

                }

				//全成型機を通して一つも入っていないスリーブは大きい方から列を非表示にする
				for(int j = 10; j < 20; j++)
				{
					dataGridView1.Columns[j].Visible = true;
				}
				int nv = 10 - max_sleeve_count;
				for(int j = 10; j > (10 - nv); j--)
				{
					dataGridView1.Columns[9 + j].Visible = false;
				}

				timer1.Interval = 60000;
                isFirstRead = true;
            }
            catch (System.IO.IOException ex)
            {
                string errorStr = "他のアプリがCSVファイルを開いている可能性があります";
                System.Console.WriteLine(errorStr);
                System.Console.WriteLine(ex.Message);
				LogFileOut(errorStr);
            }
            finally
            {
			}

			// 自動でサイズを設定するのは、行や列を追加したり、セルに値を設定した後にする。
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;

			timer1.Enabled = true;//timerを再開
        }

    }
}
