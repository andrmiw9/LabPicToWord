using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using Xceed.Document.NET;
using Xceed.Words.NET;
using System.Xml.Serialization;
using System.Net;
using System.Collections.Specialized;
using System.Configuration;
using LabPicturesToWord.Properties;

namespace LabPicturesToWord
{
    public partial class Form1 : Form
    {
        #region defining
        readonly Dictionary<int, Button> Knopki = new Dictionary<int, Button>(6);   // Словарь для 6 кнопок
        string[] Names;                      // Главный массив, хранящий полные имена выбранных изображений
        string FolderPath = "";
        string TrueName = "";
        bool isIMGsReady = false;
        int counter = 0;
        Color Butclr = ColorTranslator.FromHtml("#FF0000");
        Color ButHoverClr = ColorTranslator.FromHtml("#FFFFFF");
        Color ButPresClr = ColorTranslator.FromHtml("#0");
        Color ButTextClr = ColorTranslator.FromHtml("#0");
        int CurrentTheme;
        bool DontAskConfirm;
        bool SvoiNameOrNot;
        string SvoiNameofFile;
        bool SvoiSettings;
        bool SaveOutputFolder;
        string OutputFolder;
        bool AddNumberInTheEnd;
        int CurrentLabNumber;
        bool FormatPicBeforePaste;
        #endregion
        public Form1()
        {
            InitializeComponent();
            Knopki.Add(1, button1);
            Knopki.Add(2, button2);
            Knopki.Add(3, button3);
            Knopki.Add(4, buttonSave);
            Knopki.Add(5, button5);
            Knopki.Add(6, button6);
            Knopki.Add(7, buttonLoad);
            LoadSettings();
            InitializeOpenFileDialog();
            radioButton1.Checked = true;
            string tt1 = "Введите имя файла БЕЗ расширения(типа файла). " +
                "Если оставить поле пустым, то файл будет иметь название Лаба";
            string tt2 = "Лаба Иванова Вани ";
            string tt3 = "По умолчанию - та же папка, где лежит прога";
            string tt4 = "Растянет фото до соотношения 16:9 и таких размеров, чтобы фото оказалось по центру";
            ToolTip t1 = new ToolTip();
            t1.SetToolTip(textBox2, tt1);
            ToolTip t2 = new ToolTip();
            t2.SetToolTip(textBox1, tt2);
            ToolTip t3 = new ToolTip();
            t3.SetToolTip(button2, tt3);
            ToolTip t4 = new ToolTip();
            t4.SetToolTip(checkBox5, tt4);
        }
        void LoadSettings()
        {
            Settings set = new Settings();
            SvoiSettings = set.SvoiSettings;
            CurrentTheme = set.CurrentTheme;
            DontAskConfirm = set.DontAskConfirm;
            SvoiNameOrNot = set.SvoiNameOrNot;
            SvoiNameofFile = set.SvoiNameofFile;
            SaveOutputFolder = set.SaveOutputFolder;
            OutputFolder = set.OutputFolder;
            AddNumberInTheEnd = set.AddNumberInTheEnd;
            CurrentLabNumber = set.CurrentLabNumber;
            FormatPicBeforePaste = set.FormatPicBeforePaste;
            //-------------------------------------------------
            if (SvoiNameOrNot)
            {
                textBox1.Text = SvoiNameofFile;
                textBox2.Enabled = false;
                textBox2.Text = SvoiNameofFile;
            }
            else
            {
                textBox1.Text = "Лаба Иванова Вани";
                textBox2.Text = "Лаба";
            }
            if (SaveOutputFolder)
            {
                FolderPath = OutputFolder;
                button2.Text = "<Сохраненный путь>";
                button2.Enabled = false;
                button2.BackColor = ButPresClr;
            }
            if (SvoiSettings)
            {
                ChangeColor(CurrentTheme);
            }
            else
            {
                ChangeColor(0);
            }
            DefaultSettings();
        }
        private void InitializeOpenFileDialog()
        {
            // Set the file dialog to filter for graphics files.
            this.openFileDialog1.Filter =
                "Изображения (*.BMP;*.JPG;*.PNG;)|*.BMP;*.JPG;*.PNG;|" + "All files (*.*)|*.*";
            // Allow the user to select multiple images.
            this.openFileDialog1.Multiselect = true;
            this.openFileDialog1.Title = "Выберите нужные изображения";
        }
        private void Button1_Click(object sender, EventArgs e)
        {
            DialogResult dr = this.openFileDialog1.ShowDialog();
            if (dr == DialogResult.OK)
            {
                button1.Enabled = false;
                button1.BackColor = ButPresClr;
                isIMGsReady = true;
                Names = new string[openFileDialog1.FileNames.Length];
                counter = 0;
                foreach (String file in openFileDialog1.FileNames)       // Для каждого названия файла из выбранных
                {
                    Names[counter] = file;
                    counter++;
                }
                button1.Text = "Выбрано " + counter + " фото";
            }
        }
        private void Button2_Click(object sender, EventArgs e)
        {
            if (!SaveOutputFolder)
            {
                FolderBrowserDialog FBD = new FolderBrowserDialog();
                if (FBD.ShowDialog() == DialogResult.OK)
                {
                    FolderPath = FBD.SelectedPath;
                }
            }
        }
        private void Button3_Click(object sender, EventArgs e)            // Run the Paste Process
        {
            if (isIMGsReady)
            {
                if (SvoiSettings && SvoiNameOrNot)                        // Проверка на пустое имя должна быть сделана раньше
                {
                    TrueName = SvoiNameofFile;
                }
                else
                {
                    if (!String.IsNullOrWhiteSpace(textBox2.Text))
                    {
                        TrueName = textBox2.Text;
                    }
                    else
                    {
                        TrueName = "Лаба";
                    }
                }
                if (String.IsNullOrWhiteSpace(FolderPath))
                {
                    FolderPath = AppDomain.CurrentDomain.BaseDirectory;                     // Текущее расположение программы
                }
                DialogResult result;
                if (DontAskConfirm)
                {
                    result = DialogResult.OK;
                }
                else
                {
                    string output = TrueName;
                    if (AddNumberInTheEnd)
                    {
                        output = output + " " + (CurrentLabNumber + 1).ToString();
                    }
                    result = MessageBox.Show(
                    "Путь: " + FolderPath + "\n" +
                    "Имя файла: " + output + "\n" +
                    "Кол-во фото: " + counter + "\n",
                    "Подтвердите",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.DefaultDesktopOnly);
                }
                //----------------------------------------
                this.TopMost = true;
                if (result == DialogResult.OK)
                {
                    if (File.Exists(FolderPath + "/" + TrueName + ".docx"))
                    {
                        DialogResult res = MessageBox.Show(
                            "Хотите заменить файл?",
                            "Ошибка: Файл уже существует",
                            MessageBoxButtons.OKCancel,
                            MessageBoxIcon.Error,
                            MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.DefaultDesktopOnly);
                        if (res == DialogResult.Cancel)
                        {
                            return;
                        }
                    }
                    using (var doc1 = DocX.Create("smth.docx"))
                    {
                        int counter = 0;
                        do
                        {
                            var img = doc1.AddImage(Names[counter]);
                            Picture pic;
                            if (FormatPicBeforePaste)
                            {
                                pic = img.CreatePicture(690, 450);
                                // (была) комфортная ширина картинки = 520, 292 = 520/1.777777 (16/9 = 1.777777)
                            }
                            else
                            {
                                System.Drawing.Image image = System.Drawing.Image.FromFile(Names[counter]);
                                // На что надо умножить image.Width чтобы получить заветные 390?
                                double x = (450.0 / Convert.ToDouble(image.Width));
                                pic = img.CreatePicture(Convert.ToInt32(image.Height * x), 450);
                            }
                            var p = doc1.InsertParagraph("");
                            p.AppendPicture(pic);
                            counter++;
                        } while (counter != Names.Length);
                        if (AddNumberInTheEnd)
                        {
                            CurrentLabNumber++;
                            doc1.SaveAs(FolderPath + "/" + TrueName + " " + CurrentLabNumber + ".docx");
                            Settings s = new Settings();
                            s.CurrentLabNumber = CurrentLabNumber;
                            s.Save();
                        }
                        else
                        {
                            doc1.SaveAs(FolderPath + "/" + TrueName + ".docx");
                        }
                        MessageBox.Show("Готово!");
                        Thread.Sleep(500);
                        //Environment.Exit(0);                              // Выход
                        this.Close();                                       // Выход
                        Application.Exit();                                 // Выход
                    }
                }
            }
            else
            {
                this.TopMost = true;
                MessageBox.Show(
                "Сначала выберите изображения!",
                "Ошибка",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1,
                MessageBoxOptions.DefaultDesktopOnly);
            }
        }
        public static int GetPercent(int b, int a)
        {
            if (b == 0) return 0;
            return (int)(a / (b / 100M));
        }
        private void RadioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked == true)
            {
                panel2.Enabled = true;
                ChangeColor(CurrentTheme);
            }
            else                               //
            {
                // проверка нужна
                panel2.Enabled = false;
                CurrentTheme = 0;
                ChangeColor(CurrentTheme);
                DefaultSettings();
            }
        }
        void DefaultSettings()
        {
            //Settings.Default.Reset();
            //Settings.Default.Save();
            if (SvoiNameOrNot)
                checkBox1.CheckState = CheckState.Checked;
            else
                checkBox1.CheckState = CheckState.Unchecked;
            if (AddNumberInTheEnd)
                checkBox2.CheckState = CheckState.Checked;
            else
                checkBox2.CheckState = CheckState.Unchecked;
            if (DontAskConfirm)
                checkBox3.CheckState = CheckState.Checked;
            else
                checkBox3.CheckState = CheckState.Unchecked;
            if (SaveOutputFolder)
                checkBox4.CheckState = CheckState.Checked;
            else
                checkBox4.CheckState = CheckState.Unchecked;
            if (FormatPicBeforePaste)
                checkBox5.CheckState = CheckState.Checked;
            else
                checkBox5.CheckState = CheckState.Unchecked;
            for (int i = 4; i < 8; i++)                    // Прозрачные кнопки на форме
            {
                Knopki[i].BackColor = Color.Transparent;
                Knopki[i].FlatStyle = FlatStyle.Flat;
                Knopki[i].FlatAppearance.MouseDownBackColor = Color.Transparent;
                Knopki[i].FlatAppearance.MouseOverBackColor = Color.Transparent;
            }
        }
        void ChangeColor(int numberoftheme)
        {   // clr = color
            string bclr = "#FFFFFF";            // Button
            string bhclr = "#FFFFFF";           // Hovered Button
            string bpclr = "#FFFFFF";           // Pressed Button
            string btclr = "#FFFFFF";           // Button Text color override (optional)
            string tclr = "#FFFFFF";            // Text
            Color Textclr;
            switch (numberoftheme)
            {
                case 1:                                                      // Бирюзово - персиковая
                    this.BackColor = ColorTranslator.FromHtml("#54dada");
                    textBox2.BackColor = ColorTranslator.FromHtml("#f5dfce");
                    groupBox1.BackColor = ColorTranslator.FromHtml("#f4d6bd");
                    groupBox2.BackColor = ColorTranslator.FromHtml("#f4d6bd");
                    groupBox3.BackColor = ColorTranslator.FromHtml("#eec4a1");
                    panel1.BackColor = ColorTranslator.FromHtml("#eec4a1");
                    bclr = "#FFB579";
                    bhclr = "#ffc594";
                    bpclr = "#f69d55";
                    tclr = "#0";
                    btclr = tclr;
                    break;
                case 2:                                                      // Голубоватая
                    this.BackColor = ColorTranslator.FromHtml("#e3fdfd");
                    textBox2.BackColor = ColorTranslator.FromHtml("#e0fafa");
                    groupBox1.BackColor = ColorTranslator.FromHtml("#cbf1f5");
                    groupBox2.BackColor = ColorTranslator.FromHtml("#cbf1f5");
                    groupBox3.BackColor = ColorTranslator.FromHtml("#b7e8ed");
                    panel1.BackColor = ColorTranslator.FromHtml("#b7e8ed");
                    textBox2.BackColor = Color.White;
                    bclr = "#a6e3e9";
                    bhclr = "#85d3db";
                    bpclr = "#70c6cf";
                    tclr = "#0";
                    btclr = tclr;
                    break;
                case 3:                                                      // Черно - Желтая
                    this.BackColor = ColorTranslator.FromHtml("#333c3d");
                    groupBox1.BackColor = ColorTranslator.FromHtml("#44494a");
                    groupBox2.BackColor = ColorTranslator.FromHtml("#44494a");
                    groupBox3.BackColor = ColorTranslator.FromHtml("#3e4242");
                    panel1.BackColor = ColorTranslator.FromHtml("#3e4242");
                    textBox2.BackColor = ColorTranslator.FromHtml("#44494a");
                    bclr = "#ffe616";
                    bhclr = "#fff242";
                    bpclr = "#deba00";
                    tclr = "#FFFFFF";
                    btclr = "#0";
                    break;
                case 4:                                                      // Черно - бирюзовая
                    this.BackColor = ColorTranslator.FromHtml("#222831");
                    groupBox1.BackColor = ColorTranslator.FromHtml("#393e46");
                    groupBox2.BackColor = ColorTranslator.FromHtml("#393e46");
                    groupBox3.BackColor = ColorTranslator.FromHtml("#434a55");
                    panel1.BackColor = ColorTranslator.FromHtml("#434a55");
                    textBox2.BackColor = ColorTranslator.FromHtml("#60656c");
                    bclr = "#00adb5";
                    bhclr = "#00a0a7";
                    bpclr = "#008389";
                    tclr = "#eeeeee";
                    btclr = tclr;
                    break;
                case 5:                                                      // Сине - Зеленая
                    this.BackColor = ColorTranslator.FromHtml("#3e5f8a");
                    groupBox1.BackColor = ColorTranslator.FromHtml("#3e5f8a");
                    groupBox2.BackColor = ColorTranslator.FromHtml("#3e5f8a");
                    groupBox3.BackColor = ColorTranslator.FromHtml("#3e5f8a");
                    panel1.BackColor = ColorTranslator.FromHtml("#304f78");
                    textBox2.BackColor = ColorTranslator.FromHtml("#304f78");
                    bclr = "#1fca95";
                    bhclr = "#1bb384";
                    bpclr = "#0ba273";
                    tclr = "#ffffff";
                    btclr = tclr;
                    break;
                case 0:
                default:                                                     // Серая
                    this.BackColor = ColorTranslator.FromHtml("#f0f0f0");
                    groupBox1.BackColor = ColorTranslator.FromHtml("#f0f0f0");
                    groupBox2.BackColor = ColorTranslator.FromHtml("#f0f0f0");
                    groupBox3.BackColor = ColorTranslator.FromHtml("#f0f0f0");
                    panel1.BackColor = ColorTranslator.FromHtml("#d7d6d6");
                    textBox2.BackColor = ColorTranslator.FromHtml("#e9e9e9");
                    bclr = "#e1e1e1";
                    bhclr = "#d6d5d5";
                    bpclr = "#c4c4c4";
                    tclr = "#0";
                    btclr = tclr;
                    break;
            }
            textBox1.BackColor = textBox2.BackColor;
            Textclr = ColorTranslator.FromHtml(tclr);
            this.ForeColor = Textclr;
            textBox2.ForeColor = Textclr;
            textBox1.ForeColor = Textclr;
            groupBox1.ForeColor = Textclr;
            groupBox2.ForeColor = Textclr;
            groupBox3.ForeColor = Textclr;
            Butclr = ColorTranslator.FromHtml(bclr);
            ButPresClr = ColorTranslator.FromHtml(bpclr);
            ButHoverClr = ColorTranslator.FromHtml(bhclr);
            ButTextClr = ColorTranslator.FromHtml(btclr);
            if (isIMGsReady)
            {
                button1.ForeColor = ButTextClr;
                button1.FlatStyle = FlatStyle.Standard;
                button1.BackColor = ButPresClr;
            }
            else
            {
                button1.ForeColor = ButTextClr;
                button1.FlatStyle = FlatStyle.Standard;
                button1.BackColor = Butclr;
                button1.MouseEnter += ButME;
                button1.MouseLeave += ButML;
                button1.MouseDown += ButMD;
            }
            for (int i = 2; i < 8; i++)
            {
                Knopki[i].ForeColor = ButTextClr;
                Knopki[i].FlatStyle = FlatStyle.Standard;
                Knopki[i].BackColor = Butclr;
                Knopki[i].MouseEnter += ButME;
                Knopki[i].MouseLeave += ButML;
                Knopki[i].MouseDown += ButMD;
            }
        }
        private void ButME(object sender, EventArgs e)                    // Button Mouse Enter
        {
            var button = sender as Button;
            button.BackColor = ButHoverClr;
        }
        private void ButML(object sender, EventArgs e)                    // Button Mouse Leave
        {
            var button = sender as Button;
            button.BackColor = Butclr;
        }
        private void ButMD(object sender, EventArgs e)                    // Button Mouse Down
        {
            var button = sender as Button;
            button.BackColor = ButPresClr;
        }
        private void Button5_Click(object sender, EventArgs e)            // Предыдущая тема
        {
            CurrentTheme--;
            if (CurrentTheme < 0)
                CurrentTheme = 5;
            ChangeColor(CurrentTheme);
        }        
        private void Button6_Click(object sender, EventArgs e)            // Следующая тема
        {
            CurrentTheme++;
            if (CurrentTheme > 5)
                CurrentTheme = 0;
            ChangeColor(CurrentTheme);
        }         
        private void TextBox1_Enter(object sender, EventArgs e)
        {
            textBox1.Text = null;
        }
        private void TextBox1_Leave(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(textBox1.Text))
            {
                textBox1.Text = "Лаба Иванова Вани ";
            }
        }
        private void TextBox2_Enter(object sender, EventArgs e)
        {
            textBox2.Text = null;
        }
        private void TextBox2_Leave(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(textBox2.Text))
            {
                textBox2.Text = "Лаба";
            }
        }
        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (!checkBox1.Checked)
            {
                textBox2.Enabled = true;
                textBox1.Text = "Лаба Иванова Вани";
                textBox2.Text = "Лаба";
                TrueName = "Лаба";
            }
            else
            {
                // textBox2.Text = textBox1.Text;
            }
        }
        private void ButtonSave_Click(object sender, EventArgs e)         // Save settings
        {
            this.TopMost = true;
            DialogResult result = MessageBox.Show(
                 "Хотите сохранить настройки?",
                 "Подтвердите",
                 MessageBoxButtons.OKCancel,
                 MessageBoxIcon.Information,
                 MessageBoxDefaultButton.Button1,
                 MessageBoxOptions.DefaultDesktopOnly);
            //----------------------------------------


            if (result == DialogResult.OK)
            {
                if (checkBox4.Checked && (String.IsNullOrWhiteSpace(FolderPath) || FolderPath == "/" || FolderPath == @"\"))    // Не бейте
                {
                    MessageBox.Show(
                        "Пустой путь до папки!",
                        "Ошибка",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error,
                        MessageBoxDefaultButton.Button1,
                        MessageBoxOptions.DefaultDesktopOnly);
                    SavingSettingsError();
                    return;
                }
                Settings Sett = new Settings();
                Sett.SvoiSettings = true;
                Sett.CurrentTheme = CurrentTheme;
                if (checkBox1.Checked)
                {
                    if (!String.IsNullOrWhiteSpace(textBox1.Text))
                    {
                        Sett.SvoiNameOrNot = true;
                        Sett.SvoiNameofFile = textBox1.Text;
                        textBox2.Text = textBox1.Text;
                        TrueName = textBox1.Text;
                        textBox2.Enabled = false;
                    }
                    else   // По сути ненужная проверка, тк в textBox1 у нас по-любому подсказка - текст 
                    {
                        MessageBox.Show(
                            "Пустое имя файла!",
                            "Ошибка",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error,
                            MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.DefaultDesktopOnly);
                    }
                }
                else
                {
                    Sett.SvoiNameOrNot = false;
                    Sett.SvoiNameofFile = "";
                    TrueName = "";
                }
                if (checkBox2.Checked)
                {
                    Sett.AddNumberInTheEnd = true;
                    AddNumberInTheEnd = true;
                }
                else
                {
                    Sett.AddNumberInTheEnd = false;
                    AddNumberInTheEnd = false;
                }
                if (checkBox3.Checked)
                {
                    Sett.DontAskConfirm = true;
                }
                else
                {
                    Sett.DontAskConfirm = false;
                }
                if (checkBox4.Checked)
                {
                    button2.Text = FolderPath;
                    button2.Enabled = false;
                    Sett.SaveOutputFolder = true;
                    Sett.OutputFolder = FolderPath + "/";
                }
                else
                {
                    Sett.SaveOutputFolder = false;
                    Sett.OutputFolder = "";
                }
                if (checkBox5.Checked)
                {
                    Sett.FormatPicBeforePaste = true;
                }
                else
                {
                    Sett.FormatPicBeforePaste = false;
                }
                Sett.Save();
                MessageBox.Show(
                    "Настройки сохранены",
                    "Готово",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.DefaultDesktopOnly);
            }

        }
        void SavingSettingsError()
        {
            MessageBox.Show(
                           "Настройки не сохранены",
                           "Ошибка",
                           MessageBoxButtons.OK,
                           MessageBoxIcon.Error,
                           MessageBoxDefaultButton.Button1,
                           MessageBoxOptions.DefaultDesktopOnly);
        }
        private void ButtonLoad_Click(object sender, EventArgs e)         // Load Settings
        {
            LoadSettings();                                               // Да, я в курсе, что это дибильно
        }
    }
}