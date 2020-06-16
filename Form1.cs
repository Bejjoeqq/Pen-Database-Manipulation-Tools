using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.Data.SqlClient;
using System.Threading;
using System.Speech;
using System.Speech.Synthesis;
using System.IO;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;

namespace Pen
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        SpeechSynthesizer rd = new SpeechSynthesizer();

        MySqlConnection conn;
        MySqlDataAdapter da;
        MySqlCommandBuilder cmb;

        SqlConnection conns;
        SqlDataAdapter das;
        SqlCommandBuilder cmbs;


        DataSet ds;
        public void splashss()
        {
            Application.Run(new Form4());
        }
        public void splashs()
        {
            Application.Run(new Form3());
        }
        public void splash()
        {
            Application.Run(new Form2());
        }
        public static DialogResult InputBox(string judul, string promptTeks, ref string nilai)
        {
            Form form = new Form();
            Label label = new Label();
            TextBox textBox = new TextBox();
            Button buttonOk = new Button();
            Button buttonCancel = new Button();
            form.Text = judul;
            label.Text = promptTeks;
            textBox.Text = nilai;
            buttonOk.Text = "OK";
            buttonCancel.Text = "Cancel";
            buttonOk.DialogResult = DialogResult.OK;
            buttonCancel.DialogResult = DialogResult.Cancel;
            label.SetBounds(9, 20, 372, 13);
            textBox.SetBounds(12, 36, 372, 20);
            buttonOk.SetBounds(228, 72, 75, 23);
            buttonCancel.SetBounds(309, 72, 75, 23);
            label.AutoSize = true;
            textBox.Anchor = textBox.Anchor | AnchorStyles.Right;
            buttonOk.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            buttonCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            form.ClientSize = new Size(396, 107);
            form.Controls.AddRange(new Control[] { label, textBox, buttonOk, buttonCancel });
            form.ClientSize = new Size(Math.Max(300, label.Right + 10), form.ClientSize.Height);
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.MinimizeBox = false;
            form.MaximizeBox = false;
            form.AcceptButton = buttonOk;
            form.CancelButton = buttonCancel;
            DialogResult dialogResult = form.ShowDialog();
            nilai = textBox.Text;
            return dialogResult;
        }
        private void FindAndReplace(Word.Application wordApp, object FindText, object ReplaceWith)
        {
            
            object MatchCase = true;
            object MatchWholeWord = true;
            object MatchWildcards = false;
            object MatchSoundsLike = false;
            object MatchAllWordForms = false;
            object Forward = true;
            object Format = false;
            object MatchKashida = false;
            object MatchDiacritics = false;
            object MatchAlefHamza = false;
            object MatchControl = false;
            object read_only = false;
            object visible = true;
            object Replace = 2;
            object Wrap = 1;

            wordApp.Selection.Find.Execute(ref FindText,
                ref MatchCase, ref MatchWholeWord,
                ref MatchWildcards, ref MatchSoundsLike,
                ref MatchAllWordForms, ref Forward,
                ref Wrap, ref Format, ref ReplaceWith,
                ref Replace, ref MatchKashida,
                ref MatchDiacritics, ref MatchAlefHamza,
                ref MatchControl);
        }
        private void CreateWord(object filename, object SaveAs)
        {
            Word.Application wordApp = new Word.Application();
            object missing = Missing.Value;
            Word.Document myWordDoc = null;
            if (File.Exists((string)filename))
            {
                object readOnly = false;
                object isVisible = false;
                wordApp.Visible = false;
                myWordDoc = wordApp.Documents.Open(ref filename, ref missing, ref readOnly,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing);
                myWordDoc.Activate();

                this.FindAndReplace(wordApp, "<nama>", name);
                this.FindAndReplace(wordApp, "<ktp>", ktp);
                this.FindAndReplace(wordApp, "<alamat>", alamat);
                this.FindAndReplace(wordApp, "<ttl>", ttl);
                this.FindAndReplace(wordApp, "<jk>", jk);
                this.FindAndReplace(wordApp, "<hp>", hp);
                this.FindAndReplace(wordApp, "<email>", email);
                this.FindAndReplace(wordApp, "<rek>", rek);
                this.FindAndReplace(wordApp, "<kode>", kode);
                this.FindAndReplace(wordApp, "<phari>", phari);
                this.FindAndReplace(wordApp, "<jbulan>", jbulan);
                this.FindAndReplace(wordApp, "<hphari>", hari);
                this.FindAndReplace(wordApp, "<akhir>", target);
                this.FindAndReplace(wordApp, "<now>", label6.Text);
            }
            else
            {
                MessageBox.Show("File not found","Try Again");
            }
            myWordDoc.SaveAs(ref SaveAs, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing);
            myWordDoc.Close();
            wordApp.Quit();
            MessageBox.Show("File Created","Success");
        }
        string speech;
        int hitung;
        private void Form1_Load(object sender, EventArgs e)
        {
            Thread t = new Thread(new ThreadStart(splash));
            t.Start();
            Thread.Sleep(2000);
            t.Abort();

            string fileint = File.ReadAllText(@"D:\Pen\Setting\log.txt");
            hitung = Convert.ToInt32(fileint);

            pictureBox10.Hide();
            pictureBox11.Hide();
            speech = "Welcome To Database";
            rd.Dispose();
            rd = new SpeechSynthesizer();
            rd.SpeakAsync(speech);

            panel1.Hide();
            panel2.Hide();
            comboBox1.Text = "MySql";
            comboBox2.Text = "Tb_BTC";
        }
        public void FillDGV(string find)
        {
            try
            {
                string data = @"Data Source=.\SQLEXPRESS;AttachDbFilename=" + hasil2 + ";Integrated Security=True;Connect Timeout=30;User Instance=True;";
                conns = new SqlConnection(data);
                das = new SqlDataAdapter("select * from " + tabel + " where " + comboBox4.Text + " like '%" + find + "%'", conns);
                ds = new System.Data.DataSet();
                das.Fill(ds, "Bejjo");
                dataGridView1.DataSource = ds.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed : " + ex.Message, "Result");
            }
        }
        public void FillDGVA(string find)
        {
            try
            {
                string data = "SERVER=localhost;" + "DATABASE=" + hasil3 + ";" + "UID=root;" + "PASSWORD=;";
                conn = new MySqlConnection(data);
                da = new MySqlDataAdapter("SELECT * FROM " + tabel + " where " + comboBox4.Text + " like '%" + find + "%'", conn);
                ds = new System.Data.DataSet();
                da.Fill(ds, "Bejjo");
                dataGridView1.DataSource = ds.Tables[0];
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed : " + ex.Message, "Result");
            }
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "MySql")
            {
                comboBox3.Enabled = true;
                comboBox3.Text = "pentechno1";
                button1.Enabled = false;
            }
            else if (comboBox1.Text == "SQL Server")
            {
                comboBox3.Enabled = false;
                comboBox3.Text = "";
                button1.Enabled = true;
            }
        }
        string hasil1, hasil2, hasil3, tabel;
        private void button1_Click(object sender, EventArgs e)
        {
            string contoh, lgh;
            int a;
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Microsoft SQL Server Databases Files (*.mdf)|*.mdf|All files (*.*)|*.*";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                contoh = dlg.SafeFileName.ToString();
                hasil2 = dlg.FileName.ToString();
                lgh = contoh.Length.ToString();
                a = Convert.ToInt32(lgh) - 4;
                hasil1 = contoh.Remove(a);
                comboBox3.Text = hasil1.ToString();

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            tabel = comboBox2.Text;
            if (comboBox1.Text == "MySql")
            {
                hasil3 = comboBox3.Text;
                try
                {
                    string data = "SERVER=localhost;" + "DATABASE="+hasil3+";" + "UID=root;" + "PASSWORD=;";
                    conn = new MySqlConnection(data);
                    da = new MySqlDataAdapter("select * from "+tabel, conn);
                    ds = new System.Data.DataSet();
                    da.Fill(ds, "Bejjo");
                    dataGridView1.DataSource = ds.Tables[0];
                    panel1.Show();
                    panel2.Hide();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed : " + ex.Message,"Result");
                }
            }
            else if (comboBox1.Text == "SQL Server")
            {
                if (textBox2.Text == "Bejjoeqq")
                {
                    try
                    {
                        string data = @"Data Source=.\SQLEXPRESS;AttachDbFilename=" + hasil2 + ";Integrated Security=True;Connect Timeout=30;User Instance=True;";
                        conns = new SqlConnection(data);
                        das = new SqlDataAdapter("select * from " + tabel, conns);
                        ds = new System.Data.DataSet();
                        das.Fill(ds, "Bejjo");
                        dataGridView1.DataSource = ds.Tables[0];
                        panel1.Show();
                        panel2.Hide();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Failed : " + ex.Message, "Result");
                    }
                }
                else
                {
                    panel2.Show();
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            comboBox2.Text = "";
            comboBox3.Text = "";
        }

        private void comboBox3_TextChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text != "" && comboBox3.Text != "")
            {
                button2.Enabled = true;
            }
            else
            {
                button2.Enabled = false;
            }
        }

        private void comboBox2_TextChanged(object sender, EventArgs e)
        {
            if (comboBox2.Text != "" && comboBox3.Text != "")
            {
                button2.Enabled = true;
            }
            else
            {
                button2.Enabled = false;
            }
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "MySql")
            {
                FillDGVA(textBox1.Text);
            }
            else if (comboBox1.Text == "SQL Server")
            {
                FillDGV(textBox1.Text);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == "MySql")
            {
                try
                {
                    cmb = new MySqlCommandBuilder(da);
                    da.Update(ds, "Bejjo");
                    MessageBox.Show("Success", "Result");

                    string data = "SERVER=localhost;" + "DATABASE=" + hasil3 + ";" + "UID=root;" + "PASSWORD=;";
                    conn = new MySqlConnection(data);
                    da = new MySqlDataAdapter("select * from " + tabel, conn);
                    ds = new System.Data.DataSet();
                    da.Fill(ds, "Bejjo");
                    dataGridView1.DataSource = ds.Tables[0];
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed : " + ex.Message, "Result");
                }
            }
            else if (comboBox1.Text == "SQL Server")
            {
                try
                {
                    cmbs = new SqlCommandBuilder(das);
                    das.Update(ds, "Bejjo");
                    MessageBox.Show("Success", "Result");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed : " + ex.Message, "Result");
                }
            }
        }
        string nama = "", name, ktp, alamat, ttl, jk, hp, email, rek;
        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            nama = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            name = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            alamat = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            ttl = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            rek = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            jk = dataGridView1.CurrentRow.Cells[7].Value.ToString();
            ktp = dataGridView1.CurrentRow.Cells[8].Value.ToString();
            hp = dataGridView1.CurrentRow.Cells[9].Value.ToString();
            email = dataGridView1.CurrentRow.Cells[10].Value.ToString();
        }

        private void dataGridView1_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (comboBox1.Text == "MySql")
            {
                string kata ="Are you sure you want to delete this ID : " + nama;
                DialogResult button = MessageBox.Show(kata, "Delete data!!!",
                                      MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                                      MessageBoxDefaultButton.Button2);
                if (button == DialogResult.Yes)
                {
                    string data = "SERVER=localhost;" + "DATABASE=" + hasil3 + ";" + "UID=root;" + "PASSWORD=;";
                    conn = new MySqlConnection(data); string dass = "delete from " + tabel + " where ID=@ID";
                    string das = "delete from " + tabel + " where ID=@ID";
                    MySqlCommand cmddb = new MySqlCommand(das, conn);
                    cmddb.Parameters.Add("@ID", MySqlDbType.VarChar).Value = nama;
                    MySqlDataReader myreader;
                    try
                    {
                        conn.Open();
                        myreader = cmddb.ExecuteReader();
                        MessageBox.Show("Success", "Result");
                        while (myreader.Read())
                        {
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Failed : " + ex.Message, "Result");
                    }
                    ds = new System.Data.DataSet();
                    da.Fill(ds, "Bejjo");
                    dataGridView1.DataSource = ds.Tables[0];
                }
            }
            else if (comboBox1.Text == "SQL Server")
            {
                string data = @"Data Source=.\SQLEXPRESS;AttachDbFilename=" + hasil2 + ";Integrated Security=True;Connect Timeout=30;User Instance=True;";
                conns = new SqlConnection(data);
                string dass = "delete from " + tabel + " where Username=@Username";
                SqlCommand cmddb = new SqlCommand(dass,conns);
                cmddb.Parameters.Add("@Username", SqlDbType.VarChar).Value = nama;
                SqlDataReader myreader;
                try
                {
                    conns.Open();
                    myreader = cmddb.ExecuteReader();
                    MessageBox.Show("Success", "Result");
                    while (myreader.Read())
                    {
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed : "+ex.Message, "Result");
                }
                ds = new System.Data.DataSet();
                das.Fill(ds, "Bejjo");
                dataGridView1.DataSource = ds.Tables[0];
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            FillDGVA("!@#$%^&*()))_+");
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (textBox2.Text == "Bejjoeqq")
            {
                panel2.Hide();
            }
        }

        private void panel2_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            panel2.Hide();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            panel1.Hide();
        }
        int ms = 1;
        int x = 0;
        private void timer1_Tick(object sender, EventArgs e)
        {
            ms--;
            if (ms == 0)
            {
                ms = 1;
                x++;
            }
            if (x == 1)
            {
                pictureBox3.Show();
                pictureBox4.Hide();
                pictureBox5.Hide();
                pictureBox6.Hide();
            }
            else if (x == 2)
            {
                pictureBox4.Show();
                pictureBox5.Hide();
                pictureBox6.Hide();
                pictureBox3.Hide();
            }
            else if (x == 3)
            {
                pictureBox5.Show();
                pictureBox4.Hide();
                pictureBox6.Hide();
                pictureBox3.Hide();
            }
            else if (x == 4)
            {
                pictureBox6.Show();
                pictureBox4.Hide();
                pictureBox5.Hide();
                pictureBox3.Hide();
                x = 0;
            }
        }
        private void pictureBox7_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.ShowDialog();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            speech = "See You Next Time";
            rd.Dispose();
            rd = new SpeechSynthesizer();
            rd.SpeakAsync(speech);
            this.Hide(); 
            Thread t = new Thread(new ThreadStart(splashss));
            t.Start();
            Thread.Sleep(1500);
            t.Abort();
        }
        string date, hari;
        private void timer2_Tick(object sender, EventArgs e)
        {
            string th, bln;
            label5.Text = DateTime.Now.ToLongTimeString();
            label6.Text = DateTime.Now.ToLongDateString();
            th = DateTime.Now.ToShortDateString();
            bln = DateTime.Now.ToShortDateString();
            th = th.Remove(0, 6);
            bln = bln.Remove(5);
            bln = bln.Remove(0, 3);
            date = bln + th;
        }
        string template = @"D:\Pen\Setting\Template.docx";
        string print = @"D:\Pen\Formulir\FormulirBTC ";
        string kode;
        string phari,jbulan,target;
        private void button5_Click(object sender, EventArgs e)
        {
            string lgh, lgh2;
            if (nama == "")
            {
                MessageBox.Show("Please select the data you want to generate", "Try Again");
            }
            else
            {
                try
                {
                    if (InputBox("Input", "Jangka Waktu : (Bulan)", ref jbulan) == DialogResult.OK)
                    {
                        if (InputBox("Input", "Mulai Perjanjian : (Hari Setelah Penandatangan)", ref phari) == DialogResult.OK)
                        {
                            int a, aa, aaa, aaaa;
                            DateTime dt = DateTime.Now;
                            aa = Convert.ToInt32(phari);
                            a = Convert.ToInt32(jbulan);
                            hari = dt.AddDays(aa).ToString();
                            target = dt.AddMonths(a).AddDays(aa).ToString();

                            lgh2 = target.Length.ToString();
                            lgh = hari.Length.ToString();
                            aaa = Convert.ToInt32(lgh) - 8;
                            aaaa = Convert.ToInt32(lgh2) - 8;
                            hari = hari.Remove(aaa);
                            target = target.Remove(aaaa);

                            hitung++;
                            kode = date + hitung.ToString();
                            StreamWriter file = new StreamWriter(@"D:\Pen\Setting\log.txt");
                            file.Write(hitung);
                            file.Close();
                            CreateWord(template, print + name + hitung.ToString());
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed : " + ex);
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Word Documents (*.docx)|*.docx|All files (*.*)|*.*";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                template = dlg.FileName.ToString();
                label8.Text = "Changed";
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            SaveFileDialog dlg = new SaveFileDialog();
            dlg.Filter = "Word Documents (*.docx)|*.docx|All files (*.*)|*.*";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                print = dlg.FileName.ToString();
                label9.Text = "Changed";
            }
        }

        private void label8_Click(object sender, EventArgs e)
        {
            if (label8.Text == "Changed")
            {
                label8.Text = "Default";
                template = @"D:\Pen\Setting\Template.docx";
            }
        }

        private void label9_Click(object sender, EventArgs e)
        {
            if (label9.Text == "Changed")
            {
                label9.Text = "Default";
                print = @"D:\Pen\Formulir\FormulirBTC ";
            }
        }

        private void pictureBox8_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Template Location : D:/Pen/Setting/Template.docx", "Default Setting");
            MessageBox.Show("Print Location : D:/Pen/Formulir/FormulirBTC", "Default Setting");
            MessageBox.Show("Click on 'Changed' to set default", "Information");
        }

        private void pictureBox9_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();
            f3.ShowDialog();
        }

        private void textBox2_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            panel2.Hide();
        }

        private void label10_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            panel2.Hide();
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
