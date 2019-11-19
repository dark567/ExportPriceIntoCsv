﻿using FirebirdSql.Data.FirebirdClient;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ExportPriceFor1C
{
    public partial class MainForm : Form
    {
        FbConnection fb; //fb ссылается на соединение с нашей базой данных, по-этому она должна быть доступна всем методам нашего класса
        public string path_db;
        public string FileName;

        public MainForm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// выбор пути
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void label8_Click(object sender, EventArgs e)
        {
            SaveFileDialog OPF = new SaveFileDialog();
            OPF.Filter = "Файлы csv|*.csv";
            OPF.FileName = FileName;

            if (OPF.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = OPF.FileName;

                //MessageBox.Show(OPF.FileName);
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private FbConnection GetConnection()
        {
            string connectionString =
                "User=SYSDBA;" +
                "Password=masterkey;" +
                @"Database=" + path_db + ";" +
                "Charset=UTF8;" +
                "Pooling=true;" +
                "ServerType=0;";

            FbConnection conn = new FbConnection(connectionString.ToString());

            conn.Open();

            return conn;
        }

        /// <summary>
        /// выгрузка 2
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            int res = 0;
            fb = GetConnection();
            try
            {
                PriceTabel(fb, null, null, null, null, 2);

                //if (!string.IsNullOrEmpty(textBox2.Text))
                //{
                //    res = WorkWithReport2();
                //    textBox2 = null;
                //}
                //else
                //{
                    FileName = $"Report_2_{DateTime.Now.ToString("dd-mm-yyyy")}";
                    label8_Click(sender, e);
                    if (!string.IsNullOrEmpty(textBox2.Text))
                    {
                        res = WorkWithReport2();
                    }
                    
               // }
                if (res != 0) MessageBox.Show($"Выгрузка успешно {res}");
                else MessageBox.Show($"Выгрузка отмена {res}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Выгрузка неудачно {ex.Message}");
            }
        }

        /// <summary>
        /// выгрузка 3
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            int res = 0;
            fb = GetConnection();
            try
            {
                PriceTabel(fb, null, null, null, null, 3);

                //if (!string.IsNullOrEmpty(textBox2.Text))
                //{
                //    res = WorkWithReport3();
                //    textBox2 = null;
                //}
                //else
                //{
                    FileName = $"Report_3_{DateTime.Now.ToString("dd-mm-yyyy")}";
                    label8_Click(sender, e);
                    if (!string.IsNullOrEmpty(textBox2.Text))
                    {
                        res = WorkWithReport3();
                    }
                    
               // }
                if (res != 0) MessageBox.Show($"Выгрузка успешно {res}");
                else MessageBox.Show($"Выгрузка отмена {res}");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Выгрузка неудачно {ex.Message}");
            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            try
            {
                //Создание объекта, для работы с файлом
                INIManager manager = new INIManager(Application.StartupPath + @"\set.ini");
                //Получить значение по ключу name из секции main
                path_db = manager.GetPrivateString("connection", "db");
                db_puth.Value = path_db;

                File.AppendAllText(Application.StartupPath + @"\Event.log", "путь к db:" + db_puth.Value + "\n");
                //Записать значение по ключу age в секции main
                // manager.WritePrivateString("main", "age", "21");

                OnUserNameMessage(path_db);
            }
            catch (Exception ex)
            {
                MessageBox.Show("ini не прочтен" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OnUserNameMessage(string path_db)
        {
            if (string.IsNullOrEmpty(path_db))
                this.Text = "Экспорт Прайсов";
            else
                this.Text = "Экспорт Прайсов - (" + path_db + ")";
        }

        //public DataTable ORGTabel(FbConnection conex)
        //{
        //    string query = "SELECT id, Name, EGRPOU FROM Dic_org";
        //    FbCommand comando = new FbCommand(query, conex);
        //    try
        //    {
        //        conex.Open();
        //        FbDataAdapter datareader = new FbDataAdapter(comando);
        //        DataTable usuarios = new DataTable();
        //        datareader.Fill(usuarios);
        //        return usuarios;
        //    }
        //    catch (Exception err)
        //    {
        //        throw err;
        //    }
        //    finally
        //    {
        //        conex.Close();
        //    }
        //}

        public DataTable PriceTabel(FbConnection conn, string GRP_ID, string FILTER_, string REFRESH_ID, string IN_ORG_ID, int type)
        {
            if (string.IsNullOrEmpty(GRP_ID)) { GRP_ID = "null"; }
            if (string.IsNullOrEmpty(FILTER_)) { FILTER_ = "null"; }
            if (string.IsNullOrEmpty(REFRESH_ID)) { REFRESH_ID = "null"; }
            if (string.IsNullOrEmpty(IN_ORG_ID)) { IN_ORG_ID = "null"; }

            string query2 = "select  dgg.ID as ID_GROUP /*id группы*/ " +
                        ", dgg.name as NAME_GROUP   /*имя группы*/ " +
                        ",dg.ID as ID_GOODS      /*id goods*/ " +
                        ",dg.NAME as NAME_GOODS    /*имя услуги*/ " +
                        ",case  /*IS_NORMS_EXISTS and IS_CALCS_EXISTS*/ " +
                        "when not exists(select first 1 1 " +
                        "from DIC_GOODS_LAB_NORMS N " +
                        "where N.GOODS_ID = dg.ID) and " +
                        "exists(select first 1 1 " +
                        "from DIC_CALCULATIONS C " +
                        "where C.HD_ID = DG.ID) then 1 " +
                        "when exists(select first 1 1 " +
                        "from DIC_GOODS_LAB_NORMS N " +
                        "where N.GOODS_ID = dg.ID) and " +
                        "exists(select first 1 1 " +
                        "from DIC_CALCULATIONS C " +
                        "where C.HD_ID = DG.ID) " +
                        "and dg.id = '65' then 1 " +
                        "else 0 " +
                        "end as TYPE_GOODS " +
                        ",case when(dg.PRICE_OUT is null) then 0 else dg.PRICE_OUT end as PRICE_GOODS " +
                        ",0 as ID_SOST " +
                        ",dg.code as CODE_GOODS " +
                        "from dic_goods dg " +
                        "join DIC_GOODS_GRP dgg on dg.GRP_ID = dgg.ID " +
                        "where DG.IS_SERVICE = 1 and DG.IS_ACTIVE = 1 and dgg.id <> '189' " +
                        "union " +
                        "/*part 2*/ " +
                        "select dgg1.ID as ID_GROUP      /*id группы*/ " +
                        ",dgg1.name as NAME_GROUP    /*имя группы*/ " +
                        ",dg1.ID as ID_GOODS       /*id goods*/ " +
                        ",dg1.NAME as NAME_GOODS     /*имя услуги*/ " +
                        ",case  /*IS_NORMS_EXISTS and IS_CALCS_EXISTS*/ " +
                        "when not exists(select first 1 1 " +
                        "from DIC_GOODS_LAB_NORMS N " +
                        "where N.GOODS_ID = dg.ID) and " +
                        "exists(select first 1 1 " +
                        "from DIC_CALCULATIONS C " +
                        "where C.HD_ID = dg.ID) then 2 " +
                        "when exists(select first 1 1 " +
                        "from DIC_GOODS_LAB_NORMS N " +
                        "where N.GOODS_ID = dg.ID) and " +
                        "exists(select first 1 1 " +
                        "from DIC_CALCULATIONS C " +
                        "where C.HD_ID = DG.ID) " +
                        "and dg.id = '65' then 2 " +
                        "else 5 " +
                        "end as TYPE_GOODS " +
                        ",case when(dg1.PRICE_OUT is null) then 0 else dg1.PRICE_OUT end as PRICE_GOODS " +
                        ",case  /*IS_NORMS_EXISTS and IS_CALCS_EXISTS*/ " +
                        "when not exists(select first 1 1 " +
                        "from DIC_GOODS_LAB_NORMS N " +
                        "where N.GOODS_ID = dg.ID) " +
                        "and exists(select first 1 1 " +
                        "from DIC_CALCULATIONS C " +
                        "where C.HD_ID = dg.ID) then dg.ID " +
                        "when exists(select first 1 1 " +
                        "from DIC_GOODS_LAB_NORMS N " +
                        "where N.GOODS_ID = dg.ID) and " +
                        "exists(select first 1 1 " +
                        "from DIC_CALCULATIONS C " +
                        "where C.HD_ID = DG.ID) " +
                        "and dg.id = '65' then DG.ID " +
                        "else 0 " +
                        "end as ID_SOST " +
                        ",dg1.code as CODE_GOODS " +
                        "from dic_goods dg " +
                        "join DIC_GOODS_GRP dgg on dg.GRP_ID = dgg.ID " +
                        "left " +
                        "join DIC_CALCULATIONS dc on dc.hd_id = dg.id and dc.IS_AUTO_ADD = 0 " +
                        "left join dic_goods dg1 on dg1.id = dc.goods_id " +
                        "left join DIC_GOODS_GRP dgg1 on dg1.GRP_ID = dgg1.ID " +
                        "where DG.IS_SERVICE = 1 and DG.IS_ACTIVE = 1 and dgg.id <> '189' " +
                        "and((case " +
                        "when dc.IS_AUTO_ADD = 1 then(select RESULT " +
                        "                  from translate('IS_AUTO_ADD')) " +
                        "when((dg1.IS_SERVICE = 1) and(dc.IS_AUTO_ADD = 0)) then(select RESULT " +
                        "                                      from translate('IS_COMPLEX')) " +
                        "else (select RESULT " +
                        "from translate('IS_CALCULATION')) " +
                        "end) = '#_IS_COMPLEX' ) " +
                        "and((not exists(select first 1 1 " +
                        "from DIC_GOODS_LAB_NORMS N " +
                        "where N.GOODS_ID = dg.ID) " +
                        "and exists(select first 1 1 " +
                        "from DIC_CALCULATIONS C " +
                        "where C.HD_ID = DG.ID)) or(exists(select first 1 1 " +
                        "from DIC_GOODS_LAB_NORMS N " +
                        "where N.GOODS_ID = dg.ID) " +
                        "and exists(select first 1 1 " +
                        "from DIC_CALCULATIONS C " +
                        "where C.HD_ID = DG.ID) and dg.id = '65')) " +
                        "order by 7,3";

            string query3 = "select dtp.id as ID_TYPEPRICE, dtp.name as NAME_TYPEPRICE, S.ID as ID_GOODS, S.NAME as NAME_GOODS, S.PRICE_OUT as PRICE_GOODS,/* P.DISCONT_PRC,*/ " +
                           "P.PRICE_OUT_DISC as PRICE_TYPEPRICE " +
                           "from DIC_GOODS S " +
                           "join DIC_PRICE_LIST P on P.GOOD_ID = S.ID " +
                           "join dic_type_prices dtp on dtp.id = P.TYPE_PRICE_ID and exists(select first 1 DOO.Id " +
                           "from dic_org DOO " +
                           "where doo.type_price_id = dtp.id) " +
                           "where S.PRICE_OUT <> 0 and s.is_service = 1 and s.is_active = 1";

            // MessageBox.Show($"{query}");
            FbCommand cmd = new FbCommand(type == 2 ? query2 : query3, conn);

            try
            {
                //conn.Open();
                FbDataAdapter datareader = new FbDataAdapter(cmd);
                DataTable usuarios = new DataTable();

                datareader.Fill(usuarios);
                return usuarios;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
            }
        }


        private void label7_Click(object sender, EventArgs e)
        {
            textBox2.Text = "";
        }



        /// <summary>
        /// 
        /// </summary>
        private int WorkWithReport2()
        {
            int count = 0;
            using (var w = new StreamWriter(textBox2.Text))
            {
                foreach (DataColumn column in PriceTabel(fb, null, null, null, null, 2).Columns)
                {
                    w.Write($"{column.ColumnName};");
                    count++;
                }
                w.Write("\n");

                foreach (DataRow dataRow in PriceTabel(fb, null, null, null, null, 2).AsEnumerable().ToList())
                {
                    var first = dataRow[0].ToString(); // 
                    var second = dataRow[1].ToString(); //
                    var third = dataRow[2].ToString(); //
                    var fourth = dataRow[3].ToString(); //
                    var line = string.Format($"{first};{second};{third};{fourth}");
                    w.WriteLine(line);
                    w.Flush();
                    count++;
                }

            }
            return count;
        }

        /// <summary>
        /// 
        /// </summary>
        private int WorkWithReport3()
        {
            int count = 0;
            using (var w = new StreamWriter(textBox2.Text))
            {
                foreach (DataColumn column in PriceTabel(fb, null, null, null, null, 3).Columns)
                {
                    w.Write($"{column.ColumnName};");
                    count++;
                }
                w.Write("\n");

                foreach (DataRow dataRow in PriceTabel(fb, null, null, null, null, 3).AsEnumerable().ToList())
                {
                    var first = dataRow[0].ToString(); // 
                    var second = dataRow[1].ToString(); //
                    var third = dataRow[2].ToString(); //
                    var fourth = dataRow[3].ToString(); //
                    var five = dataRow[4].ToString(); //

                    var line = string.Format($"{first};{second};{third};{fourth};{five}");
                    w.WriteLine(line);
                    w.Flush();
                    count++;
                }
            }
            return count;
        }
    }
}