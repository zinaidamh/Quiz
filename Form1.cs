using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using Quiz.Model;
using System.Configuration;
using SelectPdf;
using System.IO;
using LinqToExcel;
using LinqToExcel.Attributes;
using System.Threading;
using ExcelDataReader;
using HtmlControl;
using System.Text.RegularExpressions;
using System.Security.Authentication.ExtendedProtection;

namespace Quiz
{


    public partial class Form1 : Form
    {
               
        Font defaultFont = new Font(new FontFamily("Arial"), 24, FontStyle.Bold, GraphicsUnit.Pixel);
        Color defaultColor = Color.Red;
        int defaultSize = 5;
        HtmlControl.Editor editor = new HtmlControl.Editor();

        string liteDBPath = ConfigurationManager.AppSettings["DbPath"].ToString();
        Guid selectedQuestionItem = Guid.Empty;
        string app_directory = AppDomain.CurrentDomain.BaseDirectory;
     


        public Form1()
        {
            InitializeComponent();
            editor.Name = "txtAnswer";
            editor.Dock = DockStyle.Fill;

            pnlEditor.Controls.Add(editor);
           
            editor.SetupFont(defaultFont,defaultColor,defaultSize);

            QuestionRepository helper = new QuestionRepository(liteDBPath);
            filteredQuestions = helper.Get(0,0,chkActiveOnly.Checked);

            //chkActiveOnly.Checked = true;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            RefreshGridView(filteredQuestions);


            int binIndex = app_directory.IndexOf(@"\bin\");
            if (binIndex > -1)
            {
                app_directory = app_directory.Substring(0, binIndex+1); //project directory
            }
        }
        #region gridview

        private void RefreshGridView(IList<Question> items)
        {

            
            var bindingList = new BindingList<Question>(items);
            var source = new BindingSource(bindingList, null);
            //dataGridView1.ReadOnly = true;
            dataGridView1.DataSource = source; 
            dataGridView1.Columns[0].Visible = false; //QuestionId
            dataGridView1.Columns[1].Width=50;        //QuestionNumber
          //dataGridView1.Columns[2].Visible          //Question text
            dataGridView1.Columns[3].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; //AnswerText
            dataGridView1.Columns[4].Visible = false; //AnswerHTML
          //dataGridView1.Columns[5].Visible          //Participants
            dataGridView1.Columns[6].Width = 30;      //Active
          
      
            SelectRow();
            if (dataGridView1.SelectedRows.Count > 0)
            {
                dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.SelectedRows[0].Index;
                //find selected row
            }

        }


        private void FilterGridView()
        {
            var QuestionHelper = new QuestionRepository(liteDBPath);
            int num1, num2;
            if (!Int32.TryParse(this.txtFrom.Text, out num1))
                num1 = 0;
            if (!Int32.TryParse(this.txtTo.Text, out num2))
                num2 = 0;


            IList<Question> items = QuestionHelper.Get(num1, num2,chkActiveOnly.Checked);
            filteredQuestions = items;
            RefreshGridView(items);


        }

        private void dataGridView1_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (((DataGridView)sender).SelectedRows.Count > 0)
            {
                var row = ((DataGridView)sender).SelectedRows[0];
                selectedQuestionItem = Guid.Parse(row.Cells["QuestionId"].Value.ToString());
                EditSelectedRow(row);

            }
        }

        private DataGridViewRow SelectRow(int defaultIndex = 0)
        {
            //found current row by question number and select it
            if (dataGridView1.Rows.Count==0)
            {
                return null;
            }
            int rowIndex = -1;
            if (this.txtNumber.Text != "" && this.txtNumber.Text != "0")
            {
                rowIndex = MatchingRowIndex(dataGridView1, "QuestionNumber", txtNumber.Text);
            }
            if (rowIndex == -1)
            {
                rowIndex = defaultIndex;
            }
            else
            {
                dataGridView1.Rows[0].Selected = false;
            }
            if (rowIndex != -1)
            {
                dataGridView1.Rows[rowIndex].Selected = true;
                selectedQuestionItem = Guid.Parse(dataGridView1.Rows[rowIndex].Cells["QuestionId"].Value.ToString());
                return dataGridView1.Rows[rowIndex];
            }
            else
            {
                return null;
            }

        }
        public static int MatchingRowIndex(DataGridView dgv, string columnName, string searchValue)
        {
            //search row in gridview by column value
            int rowIndex = -1;

            if (dgv.Rows.Count > 0 && dgv.Columns.Count > 0 && dgv.Columns[columnName] != null)
            {
                DataGridViewRow row = dgv.Rows
                    .Cast<DataGridViewRow>()
                    .FirstOrDefault(r => r.Cells[columnName].Value.ToString().Equals(searchValue));
                if (row != null)
                {
                    rowIndex = row.Index;
                }
                else
                {
                    rowIndex = -1;
                }
            }

            return rowIndex;
        }

        public void EditSelectedRow(DataGridViewRow row)
        {
            if (row.Cells["AnswerHTML"].Value != null)
            {
                editor.Html = row.Cells["AnswerHTML"].Value.ToString();
            }
            else
            {
                return;
            }

            if (row.Cells["QuestionNumber"].Value != null)
            {
                txtNumber.Text = row.Cells["QuestionNumber"].Value.ToString();
            }
            else
            {
                txtNumber.Text = "0";
            }
            if (row.Cells["QuestionText"].Value != null)
            {
                txtQuestion.Text = row.Cells["QuestionText"].Value.ToString();
            }
            else
            {
                txtQuestion.Text = "";
            }
            if (row.Cells["Participants"].Value != null)
            {
                txtParticipants.Text = row.Cells["Participants"].Value.ToString();
            }
            else
            {
                txtParticipants.Text = "";
            }
            if (row.Cells["Active"].Value != null)
            {
                chkActive.Checked = bool.Parse(row.Cells["Active"].Value.ToString());
            }
            else
            {
                chkActive.Checked = true;
            }
            ValidateRow();

        }
        #endregion

        #region buttons
        private void btnLoad_Click(object sender, EventArgs e)
        {


            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.InitialDirectory = app_directory+"Excel";
            openFileDialog.Title = "Open Excel File";
            openFileDialog.Filter = "Excel files|*.xlsx";
         
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string file = openFileDialog.FileName;
                    try
                    {
                       LoadExcel(file);
                    }
                    catch (IOException ex)
                    {
                    MessageBox.Show(ex.Message);
                    }

                }
        }

        private void btnFilter_Click(object sender, EventArgs e)
        {
            FilterGridView();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            SaveRow(true);
          
        }

        private void btnSavePdf_Click(object sender, EventArgs e)
        {
            Question q = SaveRow(false);
            if (q != null)
            {
                string TxtHtmlCode = ReadFileToString();
                CreatePdf(q, TxtHtmlCode, true);
            }

        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            this.txtNumber.Text = "0";
            this.txtQuestion.Text = "";
            this.txtParticipants.Text = "";
            this.editor.BodyHtml = "";
            this.editor.BodyText = "";
            editor.SetupFont(defaultFont, defaultColor, defaultSize);
        }
        #endregion


        IList<Question> filteredQuestions
        {
            get; set;
        }

        #region importExcel

        public void LoadExcel(string fileName)
        {
          //  string fileName = app_directory + @"\Excel\Answers" + txtFrom.Text + "-" + txtTo.Text + ".xlsx";
            IExcelDataReader excelReader;
            try
            {
                FileStream stream = File.Open(fileName, FileMode.Open, FileAccess.Read);
                //2. Reading from a OpenXml Excel file (2007 format; *.xlsx)
                excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            }
            catch
            {
                MessageBox.Show("File not found or in use by another user and can not be open");
                return;
            }



            object strNumber;
            object strAnswer;
            object strQuestion;
            object strParticipants;
            int num;
            Dictionary<int, Question> dict = new Dictionary<int, Question>();
            while (excelReader.Read())
            {

                try
                {
                    //read from excel: number, question, answer, participant
                    strNumber = excelReader.GetValue(0);
                    if (strNumber == null)
                        strNumber = "0";
                    else
                        strNumber = Regex.Replace(strNumber.ToString(), "[^0-9]", "");


                    if (!Int32.TryParse(strNumber.ToString(), out num))
                    {
                        strNumber = "0";

                    }
                    strQuestion = excelReader.GetValue(1);
                    if (strQuestion == null)
                    {
                        strQuestion = " ";
                    }
                    strAnswer = excelReader.GetValue(2);
                    if (strAnswer == null)
                    {
                        strAnswer = " ";
                    }
                    strParticipants = excelReader.GetValue(3);
                    if (strParticipants == null)
                    {
                        strParticipants = " ";
                    }


                    if (num == 0) //question can include some rows, only first row include question number
                    {
                        if (strAnswer.ToString() != " " || strQuestion.ToString() != " " || strParticipants.ToString() != " ")
                        //next rows can include some or all of question/answer/participant next row
                        {
                            if (dict.Count > 0) //dictionary already have rows, so last row - is row for the same question
                            {
                                int key = dict.ElementAt(dict.Count - 1).Key;  //current question number
                                                                               //get next row for question/answer/participants for current 
                                                                               //question = current dictionary key

                                if (!string.IsNullOrWhiteSpace(strQuestion.ToString()))
                                {
                                    ((Question)dict[key]).QuestionText += " " + System.Environment.NewLine + strQuestion.ToString();
                                }
                                if (!string.IsNullOrWhiteSpace(strAnswer.ToString()))
                                {
                                    ((Question)dict[key]).AnswerText += " " + System.Environment.NewLine + strAnswer.ToString();
                                }

                                ((Question)dict[key]).AnswerHTML = ((Question)dict[key]).AnswerText;
                                if (!string.IsNullOrWhiteSpace(strParticipants.ToString()))
                                {
                                    ((Question)dict[key]).Participants += " " + System.Environment.NewLine + strParticipants.ToString();
                                }
                            }
                            else
                            {
                                string msgErr = System.Environment.NewLine + " For answer " + strAnswer.ToString() + " not defined number - not inserted";
                                MessageBox.Show(msgErr);
                                continue; //err msg
                            }
                        }
                    }
                    else
                    {
                        //add new question to dictionary
                        Question q = new Question();
                        q.QuestionId = Guid.NewGuid();
                        q.QuestionNumber = num;
                        q.QuestionText = strQuestion.ToString();
                        q.AnswerText = strAnswer.ToString();
                        q.AnswerHTML = strAnswer.ToString();
                        q.Participants = strParticipants.ToString();
                        q.Active = strAnswer.ToString().Trim() != "" ? true : false;
                        dict.Add(num, q);
                    }


                }
                catch
                {

                }


            }
            excelReader.Close();
            SaveImport(dict,fileName);


        }
        private string FormatAnswer(string text)
        {
            bool ul = false;
            if (text.IndexOf(System.Environment.NewLine) > -1)
            {
                if (text.IndexOf("1.") > -1 || text.IndexOf("а)") > -1 || text.IndexOf("1)") > -1) //list of items
                {
                    text = "<ul style='list-style-type:none;'><li>" +
                    text.Replace(System.Environment.NewLine, "<li>") + "</ul>";
                    ul = true;
                }
                else
                {
                    text = text.Replace(System.Environment.NewLine, "<BR>"); //many rows
                }
            }
            text = "<STRONG><FONT color=#ff0000 size=5 face=Arial>" + text + "</FONT></STRONG>"; //default text style
            if (!ul)
            {
                text = "<P style =\"MARGIN: auto; align-text: center\" align=center>" + text + "</P>"; //default ul style
            }
            return text;
        }

        private void SaveImport(Dictionary<int, Question> data, string fileName)
        {
            var QuestionHelper = new QuestionRepository(liteDBPath);
            int maxNum= data.Keys.Max();
            int minNum = data.Keys.Min();

            string msg = "";

           // QuestionHelper.DeleteAll();
            foreach (KeyValuePair<int, Question> item in data)
            {
                int txtNum = item.Key;
                item.Value.AnswerHTML = FormatAnswer(item.Value.AnswerHTML);

                Question oldQuestion = QuestionHelper.GetOne(txtNum); //if number not exists - add new, else update
                Question newQuestion = null;
                if (oldQuestion == null)
                {
                    item.Value.QuestionId = Guid.NewGuid();
                    QuestionHelper.Add(item.Value);
                }
                else
                {

                    newQuestion = oldQuestion;
                    newQuestion.AnswerText = item.Value.AnswerText;
                    newQuestion.AnswerHTML = item.Value.AnswerHTML;
                    newQuestion.QuestionText = item.Value.QuestionText;
                    newQuestion.Participants = item.Value.Participants;
                    newQuestion.Active = item.Value.Active;
                    QuestionHelper.Update(newQuestion);
                }
            }


            this.txtFrom.Text = minNum.ToString();
            this.txtTo.Text = maxNum.ToString();
            FilterGridView();
            //RefreshGridView(QuestionHelper.Get(0,0,chkActiveOnly.Checked));
            // Clear();
            if (msg != "")
                MessageBox.Show(msg);
            else
                MessageBox.Show(fileName + " Loaded OK");

        }


        #endregion
        #region exportPdf
        private void ExportPdfMulti()
        {

                
            string TxtHtmlCode = ReadFileToString();
            IList<Question> list = filteredQuestions;
            int filesCount = list.Count() + 1;
         
            int i = 0;
            int percentage = (i++ + 1) * 100 / filesCount;
            backgroundWorker1.ReportProgress(percentage);
           
         
            
            foreach (Question q in list)
            {
                CreatePdf(q, TxtHtmlCode);
                percentage = (i++ + 1) * 100 / filesCount;
                backgroundWorker1.ReportProgress(percentage);
            }
            string folderName = app_directory + @"\Pdf\";
            if (txtFrom.Text != "" && txtTo.Text != "")
            {

                folderName += txtFrom.Text + "-" + txtTo.Text;
                



            }
           
            //if (Directory.Exists(folderName))
            //{
            //    MessageBox.Show("Export to Pdf finished OK");
                
            //}
          }


        protected string ReadFileToString()
        {
            string path = app_directory+@"HtmlTemplate/templateNew.html";
            string Src=  app_directory+ @"HtmlTemplate/image1.png";
            try
            {
                string readText = File.ReadAllText(path);
                return readText.Replace("{src}", Src);
            }
            catch
            {
                return "";
            }

        }


        protected void CreatePdf(Question item, string TxtHtmlCode, bool openFile=false)
        {
            // read parameters from the webpage
            string Answer=item.AnswerHTML;
            if (Answer==null)
            {
                return;
            }
            if (Answer.Trim()=="")
            {
                return;
            }

          

            TxtHtmlCode = TxtHtmlCode.Replace("{answer}", Answer);
            TxtHtmlCode = TxtHtmlCode.Replace("{number}", item.QuestionNumber.ToString());
         

            string htmlString = TxtHtmlCode;
           

            string pdf_page_size = PdfPageSize.A4.ToString();
            PdfPageSize pageSize = (PdfPageSize)Enum.Parse(typeof(PdfPageSize),
                pdf_page_size, true);

            string pdf_orientation = PdfPageOrientation.Portrait.ToString();
            PdfPageOrientation pdfOrientation =
                (PdfPageOrientation)Enum.Parse(typeof(PdfPageOrientation),
                pdf_orientation, true);

            int webPageWidth = 1024;
            int webPageHeight = 0;
          

            // instantiate a html to pdf converter object
            HtmlToPdf converter = new HtmlToPdf();

            // set converter options
            converter.Options.PdfPageSize = pageSize;
            converter.Options.PdfPageOrientation = pdfOrientation;
            converter.Options.WebPageWidth = webPageWidth;
            converter.Options.WebPageHeight = webPageHeight;

          
            PdfDocument doc = converter.ConvertHtmlString(htmlString); //, baseUrl);

         
            //string folderName = @"D:\Quiz\Quiz\Pdf\";
           
            string folderName=app_directory+@"\Pdf\";
            if (txtFrom.Text != "" && txtTo.Text != "")
            {
               
                 folderName += txtFrom.Text + "-"+ txtTo.Text;
                if (!Directory.Exists(folderName))
                {
                    Directory.CreateDirectory(folderName); //for example PDF1-50
                   // MessageBox.Show("Created folder " + folderName);
                }

               

            }
            string fileName = folderName + @"\Pdf(" + item.QuestionNumber.ToString() + ") Правильный ответ.pdf";
            // doc.Save(@"D:\Quiz\Quiz\Pdf\Pdf(" + item.QuestionNumber.ToString() + ") Правильный ответ.pdf");
            doc.Save(fileName);
            doc.Close();
            if (openFile)
            {
                System.Diagnostics.Process.Start(fileName); //view pdf
            }
        }
   

        private void btnPdfMulti_Click(object sender, EventArgs e)
        {
          
            this.progressBar1.Visible = true;
            FilterGridView();
            backgroundWorker1.RunWorkerAsync();
            

        }

       

     

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            ExportPdfMulti();
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            // Set the text.
            this.Text = e.ProgressPercentage.ToString();
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.progressBar1.Visible = false;
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.InitialDirectory = app_directory + "Pdf";
            openFileDialog.Title = "Open Pdf File";
            openFileDialog.Filter = "Pdf files|*.pdf";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string file = openFileDialog.FileName;
                

            }
        }

        #endregion

        #region selectedRow
        public bool ValidateRow()
        {
            bool questionOK = true;
            string msg = "";
            if (txtQuestion.Text.Trim() == "")
            {

                questionOK = false;
                msg = " Question text is empty " + System.Environment.NewLine;
            }


            if (editor.BodyText == null || editor.BodyText.Trim() == "")
            {

                questionOK = false;
                msg += " Answer text is empty";
            }
            if (msg != "")
            {
                lblProblem.Text = msg;
                lblProblem.Visible = true;
                chkActive.Checked = false;
            }
            else
            {
                lblProblem.Text = "";
                lblProblem.Visible = false;
            }
            return questionOK;
        }
           public Question SaveRow(bool showMessage=false)
        {
            var QuestionHelper = new QuestionRepository(liteDBPath);


            if (!ValidateRow())
                chkActive.Checked = false;

            var Question = new Question
            {
                // DateTime = DateTime.Now,
                QuestionNumber = Int32.Parse(txtNumber.Text),
                QuestionText = txtQuestion.Text,
                Participants = txtParticipants.Text,
                AnswerText = editor.BodyText,
                AnswerHTML = editor.BodyHtml,
                Active = chkActive.Checked
                //QuestionType = cmbQuestionType.SelectedValue.ToString()
            };

            // Add or Update an Question Item
            if (selectedQuestionItem != Guid.Empty)
            {
                Question.QuestionId = selectedQuestionItem;
                QuestionHelper.Update(Question);
            }
            else
            {
                Question.QuestionId = Guid.NewGuid();
                selectedQuestionItem = Question.QuestionId;
                QuestionHelper.Add(Question);
            }
           
            FilterGridView();
            if (Question.Active==false)
            {
                chkActive.Checked = false;
            }
            if (showMessage)
            {
                MessageBox.Show("Saved OK");
            }
          
            return Question;

        }



        private void txtNumber_Leave(object sender, EventArgs e)
        {
            if (txtNumber.Text != "0")
            {
                DataGridViewRow row = SelectRow(-1);
                if (row != null)
                {
                    
                    EditSelectedRow(row);
                }
                else
                {
                    this.txtQuestion.Text = "";
                    this.txtParticipants.Text = "";
                    this.editor.BodyHtml = "";
                    this.editor.BodyText = "";
                }
            }
        }
        #endregion

        #region integer_only
        public bool IntegerOnly(char KeyChar)
        {
            return (char.IsDigit(KeyChar) || char.IsControl(KeyChar));

        }
        //input for numbers - integer digits only
        private void txtFrom_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !IntegerOnly(e.KeyChar);
        }

        private void txtTo_KeyPress(object sender, KeyPressEventArgs e)
        {
            e.Handled = !IntegerOnly(e.KeyChar);
        }
        #endregion

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            string liteDBPath = app_directory + ConfigurationManager.AppSettings["DbPath"].ToString();
            try
            {
                System.IO.File.Copy(liteDBPath, liteDBPath + DateTime.Now.ToString());
            }
            catch
            {

            }

        }
    }

}
