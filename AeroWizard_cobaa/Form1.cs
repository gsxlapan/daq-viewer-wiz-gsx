using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

using AeroWizard;
using ZedGraph;

/*
 * 
 * Tes Commit first
 * Tes fork & commit dari PC Kantor <akun: gsxlapan>
 */

namespace AeroWizard_cobaa
{
    public partial class Form1 : Form
    {
        baca_csv bc;
        Color[] warnaKurva = new Color[10]
        {
            Color.Crimson,          // sebisa mungkin untuk igniter signal
            Color.Gold,
            Color.DeepSkyBlue,
            Color.LimeGreen,
            Color.MediumOrchid,
            Color.DarkGray,
            Color.Pink,
            Color.Aqua,
            Color.DarkOrange,
            Color.YellowGreen
        };

        public Form1()
        {
            InitializeComponent();
            _show_panel(false);
            textBox1.Visible = false;

            //wz.NextButtonText = "Selanjutnya";
           //wz.CancelButtonText = "Batal";
            wz.FinishButtonText = "Selesai";
            wz.BackButtonToolTipText = "Kembali ke halaman sebelumnys";
            
        }

        //variabel data inisial pakai struct
        struct data_us
        {
            public DataTable dtSensor;
            public string[] chName;
            public string[] chUnit;

            public int chJum;
            public int rowJum;
            public int sampelRate;

            public string scopeId;
            public string motorId;
            public string dateId;
            public string timeId;
            public string hari;
            public string tanggal;
        }
        data_us daq = new data_us();

        string strPathCsv = String.Empty;
        string strRoketId = String.Empty;

        int _thrust_idx = 0;        //letak, urutan unit kgf untuk sync antara Tabel, Chart (YAxis), Hitung2an (Isp)
        int _pressure_idx = 0;      //juga untuk cek apakah ada Kgf, ada Bar, ada Ign? ==> untuk ngisi Hitungan & FormUS
        int _ignition_idx = 0;
        private bool[] pageOpenState = new bool[9]; //var untuk deteksi page sudah dibuka sebelumnya atau belum; flag utk first time opened
        
        //untuk eksekusi fungsi di setiap page yg terbuka, 'first executed'
        private void wizardControl1_SelectedPageChanged(object sender, EventArgs e)
        {
            if (wz.SelectedPage == wzPage1)
            {
                pageOpenState[2] = false;   //tabel daq
                pageOpenState[4] = false;   //chart plotting
            }
            if (wz.SelectedPage == wzPage2)
            {
                if (!pageOpenState[2])          // kalau baru pertama kali dibuka
                {
                    isi_tabel_daq();
                    pageOpenState[2] = true;
                }
                else                            // kalau sudah pernah dibuka pageOpenState=false
                    dgv1.Refresh();
            }
            if (wz.SelectedPage == wzPage4)
            {
                if (!pageOpenState[4])
                {
                    init_chart0();
                    cek_index_unit();
                    create_chart0_us();
                    isi_tabel_cursor();

                    pageOpenState[4] = true;
                }
                else
                {
                    zg0.Refresh();
                    zg0.AxisChange();
                    dgv2.Refresh();

                }
            }
            if (wz.SelectedPage == wzPage5)
            {
                for(int i=0; i<100; i++)
                {
                    progressBar1.Value = i;
                    progressBar1.Update();
                    Thread.Sleep(1);    //5 ms
                }

            }
            /*
             * Bukak form windows baru
            if (wz.SelectedPage == wzPage6)
            {
                Form2 f2 = new Form2();
                f2.Show();
            }
            */



        }




    #region Page1 : LOAD CSV
        private void ambil_csv_daq()
        {
            Stream mystream = null;
            OpenFileDialog ofd = new OpenFileDialog();

            ofd.Filter = "CSV Files (*.csv)|*.csv";
            ofd.FilterIndex = 1;
            ofd.RestoreDirectory = true;
            DialogResult dr = ofd.ShowDialog();
            if (dr == DialogResult.OK)
            {
                try
                {
                    if ((mystream = ofd.OpenFile()) != null)
                    {
                        using (mystream)
                        {
                            bc = null;
                            bc = new baca_csv(mystream);
                            mystream.Close();
                        }
                    }
                    strPathCsv = Path.GetFullPath(ofd.FileName);
                    strRoketId = Path.GetFileNameWithoutExtension(ofd.FileName);    //untuk ambil nama
                }
                catch (Exception err)
                {
                    MessageBox.Show("(0) Gagal mengambil CSV, mohon diulangi, broo ! " + err.Message);
                    //Application.Exit();     //restart program, exit dulu
                    //Environment.Exit(0);
                }
            }

            if (dr == DialogResult.Cancel)
            {
                MessageBox.Show("(1) Gagal mengambil CSV, mohon diulangi, broo ! ");
            }

        }
        


        //pakai nama tipe roket
        private String refine_roket_id(string s)
        {
            //cari index huruf R > hapus index sebelumnya > copy ke string baru
            try
            {
                int idx = 0;
                idx = s.IndexOf("R", StringComparison.InvariantCultureIgnoreCase);
                s = s.Remove(0, idx);
            }
            catch (Exception err)
            {
                MessageBox.Show("(2)Nama roket (Nama file Csv) salah, mohon diperbaiki, bro!", "", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                Application.Exit();
                Environment.Exit(0);
            }
            return s;
        }

        CultureInfo lokalID = new System.Globalization.CultureInfo("id-ID");
        private void ambil_hari_ID(string s)
        {
            DateTime dayTime = DateTime.ParseExact(s, "yyyy/MM/dd", CultureInfo.InvariantCulture);
            daq.hari = lokalID.DateTimeFormat.GetDayName(dayTime.DayOfWeek).ToString();
            daq.tanggal = dayTime.ToString("dd MMM yyyy");
        }
        private void Isi_data_inisial()
        {
            try
            {
                if (bc != null)
                {
                    //Ambil DataTable
                    daq.dtSensor = bc.get_dataRekap().Copy(); ;

                    //HEADER
                    daq.scopeId = bc.get_daqID();               //scope corder
                    daq.motorId = refine_roket_id(strRoketId);  //motor roket kode
                    daq.dateId = bc.get_dateID();               //date
                    daq.timeId = bc.get_timeID();               //time

                    ambil_hari_ID(daq.dateId); //dapat hari dan tanggal dalam format Indonesia, string

                    //Ambil DATA_HEADER
                    daq.chName = bc.get_channelID();        //array channel name
                    daq.chUnit = bc.get_channelUNIT();      //array channel unit

                    daq.rowJum = bc.get_jumlahData() - 1;   //jumlah row / data sensor
                    daq.chJum = bc.get_jumlahCh();          //jumlah channel
                    daq.sampelRate = int.Parse(bc.get_Sps());//sample rate
                }
            }

            catch (Exception) //err
            {
                MessageBox.Show("(1) TIDAK BERHASIL MENGAMBIL FILE ! <ERROR_STREAM>", "Wassalam..", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Application.Exit();
                Environment.Exit(0);
            }
        }




        DataTable dtSens = new DataTable();
        private void isi_tabel_pembacaan()
        {
            //Inisialisasi tabel
            //isi kolom nomer
            dtSens = daq.dtSensor.Copy();

            dtSens.Columns.Add("Data", typeof(int)).SetOrdinal(0);
            for (int a = 0; a < daq.rowJum; a++)
                dtSens.Rows[a][0] = a;

            //init dgv : clear first
            if (dgv0 != null)
            {
                dgv0.DataSource = null;
                dgv0.Rows.Clear();
                dgv0.Columns.Clear();
            }

            dgv0.DataSource = dtSens;
            dgv0.Columns[0].Width = 55;

            //style
            foreach (DataGridViewColumn col in dgv0.Columns)
            {
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                col.HeaderCell.Style.Font = new System.Drawing.Font("HP Simplified", 14F, FontStyle.Bold, GraphicsUnit.Pixel);
                col.SortMode = DataGridViewColumnSortMode.NotSortable;

                col.ReadOnly = true;        //semua readonly..
            }

            dgv0.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgv0.Refresh();
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            ambil_csv_daq();
            Isi_data_inisial();
           
            textBox1.Text = strPathCsv;
            textBox2.Text = daq.motorId;
            textBox3.Text = daq.hari+", "+daq.tanggal+", "+ daq.timeId;
            textBox4.Text = daq.chJum.ToString();
            textBox5.Text = String.Join(", ", daq.chName);
            

            isi_tabel_pembacaan();
            _show_panel(true);
            textBox1.Visible = true;
            wzPage1.AllowNext = true;
        }

        private void _show_panel(bool b)
        {
            panel3.Visible = b;
            panel2.Visible = b;
        }
        #endregion Page1

    #region Page2 : ISI DAQ PARAMETER
        //tabel parameter
        DataGridViewComboBoxColumn cbSensorType = new DataGridViewComboBoxColumn();
        DataGridViewComboBoxColumn cbSensorSign = new DataGridViewComboBoxColumn();

        private void create_combobox()
        {

            cbSensorType.HeaderText = "Sensor Type";
            cbSensorType.Items.Add("TML LoadCell");
            cbSensorType.Items.Add("PT750 Pressure");
            cbSensorType.Items.Add("HPT200 Pressure");
            cbSensorType.Items.Add("HPT902 Pressure");
            cbSensorType.Items.Add("lain-lain");
            cbSensorType.Name = "cbSenType";

            cbSensorSign.HeaderText = "Sensor Signal";
            cbSensorSign.Items.Add("Straingauge");
            cbSensorSign.Items.Add("Themocouple");
            cbSensorSign.Items.Add("4-20 mAmp (via ADC)");
            cbSensorSign.Items.Add("DC Voltage");
            cbSensorSign.Name = "cbSenSign";


            //dataGridView2.Columns.Add(dgv2ColorColumn); // dipakai di tabel parameter chart saja

        }
        private void isi_tabel_daq()
        {
            create_combobox();
            try
            {
                //clear dgv1
                dgv1.DataSource = null;
                dgv1.Rows.Clear();
                dgv1.Columns.Clear();
                dgv1.Refresh();

                //bikin column header
                dgv1.Columns.Add("nM", "No");            //0
                dgv1.Columns.Add("chName", "Channel Name");        //1
                dgv1.Columns.Add("chNo", "Channel No.");           //2
                dgv1.Columns.Add("unit", "Unit Satuan");      //3
                dgv1.Columns.Add(cbSensorType);                //4
                dgv1.Columns.Add("sn", "Sensor Serial No.");   //5
                dgv1.Columns.Add("fs", "Sensor Full Scale");   //6
                dgv1.Columns.Add(cbSensorSign);            //7
                dgv1.Columns.Add("date", "Last Calibration Date");     //8

                //style
                dgv1.Columns[0].Width = 30;
                dgv1.Columns[0].ReadOnly = true;
                dgv1.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgv1.Columns[1].ReadOnly = true;
                dgv1.Columns[1].Width = 100;
                dgv1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgv1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgv1.Columns[3].Width = 65;
                dgv1.Columns[3].ReadOnly = true;
                dgv1.Columns[5].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dgv1.Columns[6].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                ((DataGridViewTextBoxColumn)dgv1.Columns[2]).MaxInputLength = 2;    //channel
                ((DataGridViewTextBoxColumn)dgv1.Columns[6]).MaxInputLength = 5;    //scale

                //bikin style setiap Header Column sama identik
                foreach (DataGridViewColumn col in dgv1.Columns)
                {
                    col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    col.HeaderCell.Style.Font = new System.Drawing.Font("HP Simplified", 14F, FontStyle.Bold, GraphicsUnit.Pixel);
                    col.SortMode = DataGridViewColumnSortMode.NotSortable;
                }

                //masukkan data tiap rows
                for (int a = 0; a < daq.chJum; a++)
                    dgv1.Rows.Add((a + 1), daq.chName[a], 0, daq.chUnit[a + 1]);

                dgv1.Refresh();
            }
            catch (Exception)
            {
                MessageBox.Show("Tabel Param Error");
            }
        }

        //button click untuk save parameter
        bool[] setupSaved = new bool[2];
        private void btnSaveSFiring_Click(object sender, EventArgs e)
        {
            setupSaved[0] = true;
            if (setupSaved[0] && setupSaved[1])
                wzPage2.AllowNext = true;
        }
        private void btnSaveSNotes_Click(object sender, EventArgs e)
        {

        }
        private void btnSaveSDaq_Click(object sender, EventArgs e)
        {
            setupSaved[1] = true;
            if (setupSaved[0] && setupSaved[1])
                wzPage2.AllowNext = true;
        }

        private void save_setup_firing()
        {
            
        }
        
        private void save_setup_notes()
        {

        }

        private void save_setup_daq()
        {

        }

        #endregion Page2

    #region Page4 : PLOTTING CHART

        /*----------------------------------------------------------------- CHART AREA --------------------------*/
        GraphPane mychart0 = new GraphPane();
        LineItem kurvachx0;
        private string[] datachx0;

        List<string[]> datachx_list = new List<string[]>();

        private void init_chart0()
        {
            //mychart0.YAxisList.Clear();   errorr

            /*---------------------------------------------------------------------------CHART_SETUP-------*/
            //inisiasi grafik plot area
            mychart0 = zg0.GraphPane;

            mychart0.CurveList.Clear();

            //setup background
            mychart0.Legend.IsVisible = true;
            mychart0.Legend.Position = ZedGraph.LegendPos.Top;
            mychart0.IsFontsScaled = false;

            //judul grafik
            mychart0.Title.Text = "Grafik Hasil Uji Statis " + daq.motorId;
            mychart0.Title.FontSpec.Family = "HP Simplified";
            mychart0.Title.FontSpec.Size = 15;
            mychart0.Title.FontSpec.IsBold = true;

            /*---------------------------------------------------------------------------AXIS_SETUP-------*/
            //setup X-Axis
            mychart0.XAxis.Title.Text = "Waktu (Detik)";
            mychart0.XAxis.Title.FontSpec = new FontSpec("HP Simplified", 15, Color.Black, true, false, false);
            mychart0.XAxis.Title.FontSpec.Border = new Border(false, Color.Black, 0);

            mychart0.XAxis.Scale.Max = daq.rowJum / daq.sampelRate;
            mychart0.XAxis.Scale.Min = 0;
            mychart0.XAxis.MajorGrid.IsVisible = true;
            mychart0.XAxis.MajorGrid.Color = Color.Black;
            mychart0.XAxis.MinorGrid.IsVisible = true;

            //setup Y-Axis
            mychart0.YAxis.Title.Text = "Gaya dorong (Kgf)";
            mychart0.YAxis.Title.FontSpec = new FontSpec("HP Simplified", 15, Color.Black, true, false, false);
            mychart0.YAxis.Title.FontSpec.Border = new Border(false, Color.Black, 0);
            mychart0.YAxis.Title.FontSpec.Angle = 180;

            mychart0.YAxis.Scale.MaxAuto = true;
            mychart0.YAxis.Scale.Min = 0;
            mychart0.YAxis.MajorGrid.IsVisible = true;
            mychart0.YAxis.MinorGrid.IsVisible = false;

            mychart0.YAxis.MajorTic.IsOpposite = false;  //biar Tic tidak muncul di semua YAxis-Y2
            mychart0.YAxis.MinorTic.IsOpposite = false;
            mychart0.YAxis.MajorTic.IsInside = false;
            mychart0.YAxis.MinorTic.IsInside = false;
            mychart0.YAxis.Scale.Align = AlignP.Inside;
            mychart0.YAxis.Cross = 0.0;

            //setup Y2-AXIS
            mychart0.Y2Axis.Title.Text = "Ignition (Volt)";
            mychart0.Y2Axis.Title.FontSpec = new FontSpec("HP Simplified", 15, Color.Black, true, false, false);
            mychart0.Y2Axis.Title.FontSpec.Border = new Border(false, Color.Black, 0);
            mychart0.Y2Axis.Title.FontSpec.Angle = 180;

            mychart0.Y2Axis.Scale.Max = 30;      //30 volt...
            mychart0.Y2Axis.Scale.Min = 0;

            mychart0.Y2Axis.MajorTic.IsInside = false;
            mychart0.Y2Axis.MajorTic.IsInside = false;
            mychart0.Y2Axis.MinorTic.IsInside = false;
            mychart0.Y2Axis.MajorTic.IsOpposite = false;
            mychart0.Y2Axis.MinorTic.IsOpposite = false;
            mychart0.Y2Axis.Scale.Align = AlignP.Inside;

            //setup Y-Axis 1
            YAxis YAxis0 = new YAxis("Bar");    //YAxisIndex nya = 1...
            mychart0.YAxisList.Add(YAxis0);

            mychart0.YAxisList[1].Title.Text = "Tekanan (Bar)";
            mychart0.YAxisList[1].Title.FontSpec = new FontSpec("HP Simplified", 15, Color.Black, true, false, false);
            mychart0.YAxisList[1].Title.FontSpec.Border = new Border(false, Color.Black, 0);
            mychart0.YAxisList[1].Title.FontSpec.Angle = 180;

            mychart0.YAxisList[1].Scale.Max = 150;       //bar max scale view
            mychart0.YAxisList[1].Scale.Min = 0;
            mychart0.YAxisList[1].MajorGrid.IsVisible = false;
            mychart0.YAxisList[1].MinorGrid.IsVisible = false;

            mychart0.YAxisList[1].MajorTic.IsOpposite = false;  //biar Tic tidak muncul di semua YAxis-Y2
            mychart0.YAxisList[1].MinorTic.IsOpposite = false;
            mychart0.YAxisList[1].MajorTic.IsInside = false;
            mychart0.YAxisList[1].MinorTic.IsInside = false;
            mychart0.YAxisList[1].Scale.Align = AlignP.Inside;
            mychart0.YAxisList[1].Cross = 0.0;

            /*---------------------------------------------------------------------------ZGC_CONTROL_SETUP-------*/
            // setup draw plotter
            zg0.IsShowPointValues = true;
            zg0.IsShowHScrollBar = true;
            zg0.IsShowVScrollBar = true;
            zg0.IsAutoScrollRange = true;

            zg0.AxisChange();
        }

        private void cek_index_unit()
        {
            for(int a=0; a<daq.chJum; a++)
            {
                if (String.Format(daq.chUnit[a + 1], StringComparison.InvariantCultureIgnoreCase) == "bar")
                    _pressure_idx = a;
                if (String.Format(daq.chUnit[a + 1], StringComparison.InvariantCultureIgnoreCase) == "kgf")
                    _thrust_idx = a;
                if (String.Format(daq.chUnit[a + 1], StringComparison.InvariantCultureIgnoreCase) == "volt" || String.Format(daq.chUnit[a + 1], StringComparison.InvariantCultureIgnoreCase) == "V")
                    _ignition_idx = a;
            }
        }

        private void create_chart0_us()
        {
            //clear chart
            zg0.GraphPane.CurveList.Clear();
            zg0.GraphPane.GraphObjList.Clear();
            mychart0.CurveList.Clear();
            mychart0.GraphObjList.Clear();

            //pointpair
            double[] xVal = new double[daq.rowJum];
            double[] yVal = new double[daq.rowJum];

            int resolusi = 10000;

            //isi kurva (a) tiap Channel (a) ==> tiap kurva
            for(int a=0; a<daq.chJum; a++)
            {
                //isi data pointpair untuk kurva
                datachx0 = dtSens.AsEnumerable().Select(r => r.Field<string>(a+1)).ToArray();
                datachx_list.Add(datachx0);

                float b = 0;
                for(int c=0;c<daq.rowJum;c++)
                {
                    xVal[c] = b / daq.sampelRate;
                    yVal[c] = float.Parse(datachx0[c]) / resolusi;
                    b++;
                }

                //isi kurva line & style
                kurvachx0 = new LineItem(daq.chName[a + 1], xVal, yVal, warnaKurva[a], SymbolType.None, 3.0f);
                kurvachx0.Line.IsAntiAlias = true;
                kurvachx0.Line.IsSmooth = true;

                if (_ignition_idx == a)
                {
                    kurvachx0.IsY2Axis = true;
                    mychart0.Y2Axis.IsVisible = true;
                }
                if(_pressure_idx == a)
                {
                    kurvachx0.YAxisIndex = 1;
                    mychart0.YAxisList[1].IsVisible = true;
                }
                //if(_thrust_idx == a)

                mychart0.CurveList.Add(kurvachx0);
                zg0.AxisChange();
                zg0.Refresh();
            }




        }

        /*----------------------------------------------------------------- TABEL ANALITIK AREA ------------------*/
        private void isi_tabel_cursor()
        {
            DataTable dtAnalytic = new DataTable();
            //clear datatabel
            dtAnalytic.Clear();
            dtAnalytic.Rows.Clear();
            dtAnalytic.Columns.Clear();
            //clear dgv2
            dgv2.DataSource = null;
            dgv2.Rows.Clear();
            dgv2.Columns.Clear();
            dgv2.Refresh();

            try
            {
                //isi header column dulu
                dtAnalytic.Columns.Add("NO", typeof(int));      //0
                dtAnalytic.Columns.Add("KURVA", typeof(bool));    //1
                dtAnalytic.Columns.Add("CHANNEL", typeof(string));  //2
                dtAnalytic.Columns.Add("NILAI (detik-X)", typeof(string));       //3
                dtAnalytic.Columns.Add("UNIT", typeof(string)); //4

                //isi auto idChannel
                for (int a = 0; a < daq.chJum; a++)
                {
                    dtAnalytic.Rows.Add((a + 1), true, daq.chName[a + 1], 0, daq.chUnit[a + 1]);
                }

                //binding dgv1 dengan dtAnalytic yg telah dibuat
                dgv2.DataSource = dtAnalytic;

                //style centang curve default 
                for (int a = 0; a < daq.chJum; a++)
                {
                    //aksi_centang(a, true);
                }

                //style
                dgv2.Columns[0].Width = 25;
                dgv2.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                //dataGridView1.Columns[1].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dgv2.Columns[1].Width = 50;
                dgv2.Columns[1].ReadOnly = false;
                dgv2.Columns[2].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                dgv2.Columns[3].Width = 75;
                dgv2.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;

                //style header column
                foreach (DataGridViewColumn col in dgv2.Columns)
                {
                    col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    col.HeaderCell.Style.Font = new System.Drawing.Font("HP Simplified", 14F, FontStyle.Bold, GraphicsUnit.Pixel);
                    col.SortMode = DataGridViewColumnSortMode.NotSortable;

                    col.ReadOnly = true;        //semua readonly..
                }
                dgv2.Columns[1].ReadOnly = false;  //..kecuali column centang


             
                //style rows color & isi nomor rows
                foreach (DataGridViewRow row in dgv2.Rows)
                {
                    row.DefaultCellStyle.Font = new System.Drawing.Font("HP Simplified", 12F, FontStyle.Bold, GraphicsUnit.Pixel);

                    //isi row number & style cell 
                    row.Cells[3].Style.Alignment = DataGridViewContentAlignment.MiddleRight;
                    row.Cells[4].Style.Alignment = DataGridViewContentAlignment.MiddleCenter;

                    //set color cell
                    for (int k = 0; k < daq.chJum; k++)  // cara 1
                    {
                        dgv2.Rows[k].Cells[0].Style.BackColor = warnaKurva[k]; // kolom 2 (ID)
                        dgv2.Rows[k].Cells[1].Style.BackColor = warnaKurva[k]; // kolom 2 (ID)
                        dgv2.Rows[k].Cells[2].Style.BackColor = warnaKurva[k]; // kolom 3 (channelName)
                        dgv2.Rows[k].Cells[1].Style.SelectionBackColor = warnaKurva[k];
                    }
                }

                dgv2.Refresh();
            }
            catch (Exception errz)
            {
                MessageBox.Show("Gagal menampilkan Tabel Analitik! " + Environment.NewLine + errz.Message);
            }

        }


    #endregion Page4



    #region DGV1 EVENT CLICK

        //sub-fungsi untuk FORCE HURUF kAPITAL & NUMERIC ACCEPTANCE DI SPESIFIK KOLOM*/
        private void dgv1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {

            // Force untuk Capslock
            if (dgv1.CurrentCell.ColumnIndex == 5)     // KOLOM 6 input force to capslock
            {
                if (e.Control is TextBox)
                {
                    ((TextBox)(e.Control)).CharacterCasing = CharacterCasing.Upper;
                }
            }

            // Hanya terima input numeric
            e.Control.KeyPress -= new KeyPressEventHandler(col_keypress);       //bikin event baru

            if (dgv1.CurrentCell.ColumnIndex == 2 || dgv1.CurrentCell.ColumnIndex == 6) //kolom yg diinginkan
            {
                TextBox tb = e.Control as TextBox;
                if (tb != null)
                {
                    tb.KeyPress += new KeyPressEventHandler(col_keypress);
                }
            }
        }

        private void col_keypress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }

        }

        //sub-fungsi untuk otomatis klik cell combobox
        private void dgv1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            bool validClick = (e.RowIndex != -1 && e.ColumnIndex != -1); //Make sure the clicked row/column is valid.
            var datagridview = sender as DataGridView;

            // Check to make sure the cell clicked is the cell containing the combobox 
            if (datagridview.Columns[e.ColumnIndex] is DataGridViewComboBoxColumn && validClick)
            {
                datagridview.BeginEdit(true);
                ((ComboBox)datagridview.EditingControl).DroppedDown = true;
            }
        }
        private void dgv1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit);
        }
        
        //calendar picker, tabel kolom terakhir
        DateTimePicker calendarPicker = new DateTimePicker();
        private void dgv1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // If any cell is clicked on the Second column which is our date Column  
            if (e.ColumnIndex ==8 && e.RowIndex >= 0)
            {
                //Adding DateTimePicker control into DataGridView   
                dgv1.Controls.Add(calendarPicker);

                // Setting the format (i.e. 2014-10-10)  
                calendarPicker.Format = DateTimePickerFormat.Short;

                // It returns the retangular area that represents the Display area for a cell  
                Rectangle calendarBox = dgv1.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);

                //Setting area for DateTimePicker Control  
                calendarPicker.Size = new Size(calendarBox.Width, calendarBox.Height);

                // Setting Location  
                calendarPicker.Location = new Point(calendarBox.X, calendarBox.Y);

                // An event attached to dateTimePicker Control which is fired when DateTimeControl is closed  
                calendarPicker.CloseUp += new EventHandler(calendarPicker_CloseUp);

                // An event attached to dateTimePicker Control which is fired when any date is selected  
                calendarPicker.TextChanged += new EventHandler(calendarPicker_OnTextChange);

                // Now make it visible  
                calendarPicker.Visible = true;
            }
        }
        private void calendarPicker_OnTextChange(object sender, EventArgs e)
        {
            // Saving the 'Selected Date on Calendar' into DataGridView current cell  
            dgv1.CurrentCell.Value = calendarPicker.Text.ToString();
        }

        private void calendarPicker_CloseUp(object sender, EventArgs e)
        {
            calendarPicker.Visible = false;
        }














        #endregion DGV1

    }
}
