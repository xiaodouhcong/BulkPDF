using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace BulkPDF
{
    public partial class MainForm : Form
    {
        private bool customFont = false;
        private IDataSource dataSource;
        private bool finalize = false;
        private PDF pdf;
        private Dictionary<string, PDFField> pdfFields = new Dictionary<string, PDFField>();
        private ProgressForm progressForm;
        private int tempSelectedIndex;
        private bool unicode = false;

        public MainForm()
        {
            try
            {
                File.AppendAllText("debug.log", $"[{DateTime.Now}] MainForm构造函数开始\n");
                InitializeComponent();
                this.MinimumSize = new Size(500, 400);

                lVersion.Text = Application.ProductVersion.ToString();
                
                // 初始化语言选择
                InitializeLanguageSelection();
                
                File.AppendAllText("debug.log", $"[{DateTime.Now}] MainForm构造函数完成\n");
                File.AppendAllText("debug.log", $"[{DateTime.Now}] 窗口状态: Visible={this.Visible}, WindowState={this.WindowState}, Location={this.Location}\n");
            }
            catch (Exception ex)
            {
                string errorMsg = $"[{DateTime.Now}] MainForm构造函数错误: {ex.Message}\n{ex.StackTrace}\n";
                File.AppendAllText("debug.log", errorMsg);
                MessageBox.Show($"程序启动错误: {ex.Message}\n{ex.StackTrace}", "错误", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }

        private void InitializeLanguageSelection()
        {
            // 确保控件可见
            lblLanguage.Visible = true;
            cbLanguage.Visible = true;
            bLanguageSettings.Visible = true;
            
            // 设置按钮样式
            bLanguageSettings.BackColor = System.Drawing.SystemColors.Control;
            bLanguageSettings.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            bLanguageSettings.BringToFront();
            
            // 记录调试信息
            File.AppendAllText("debug.log", $"[{DateTime.Now}] InitializeLanguageSelection: bLanguageSettings.Visible={bLanguageSettings.Visible}, Location={bLanguageSettings.Location}, Size={bLanguageSettings.Size}\n");
            
            // 根据当前语言设置选择对应的选项
            string currentLang = OptionFileHandler.GetOptionValue("Language");
            switch (currentLang)
            {
                case "de":
                    cbLanguage.SelectedIndex = 1;
                    break;
                case "zh-CN":
                    cbLanguage.SelectedIndex = 2;
                    break;
                default:
                    cbLanguage.SelectedIndex = 0;
                    break;
            }
        }

        private void cbLanguage_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedLang = "";
            switch (cbLanguage.SelectedIndex)
            {
                case 0:
                    selectedLang = "en";
                    break;
                case 1:
                    selectedLang = "de";
                    break;
                case 2:
                    selectedLang = "zh-CN";
                    break;
            }

            if (!string.IsNullOrEmpty(selectedLang))
            {
                // 保存语言设置
                OptionFileHandler.SetOptionValue("Language", selectedLang);
                
                // 不再提示用户重启应用，直接应用语言设置
                // MessageBox.Show(Properties.Resources.MessageLanguageChanged, Properties.Resources.MessageInfo, 
                //     MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        
        private void bLanguageSettings_Click(object sender, EventArgs e)
        {
            // 创建语言设置对话框
            Form languageDialog = new Form();
            languageDialog.Text = "语言设置 / Language Settings";
            languageDialog.FormBorderStyle = FormBorderStyle.FixedDialog;
            languageDialog.ControlBox = true;
            languageDialog.MaximizeBox = false;
            languageDialog.MinimizeBox = false;
            languageDialog.StartPosition = FormStartPosition.CenterParent;
            languageDialog.Width = 350;
            languageDialog.Height = 200;
            
            // 添加标签
            Label label = new Label();
            label.Text = "请选择语言 / Please select language:";
            label.Location = new Point(20, 20);
            label.AutoSize = true;
            languageDialog.Controls.Add(label);
            
            // 添加下拉框
            ComboBox comboBox = new ComboBox();
            comboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox.Items.AddRange(new object[] { "English (en)", "Deutsch (de)", "中文 (zh-CN)" });
            comboBox.Location = new Point(20, 50);
            comboBox.Width = 200;
            
            // 设置当前选择的语言
            string currentLang = OptionFileHandler.GetOptionValue("Language");
            switch (currentLang)
            {
                case "de":
                    comboBox.SelectedIndex = 1;
                    break;
                case "zh-CN":
                    comboBox.SelectedIndex = 2;
                    break;
                default:
                    comboBox.SelectedIndex = 0;
                    break;
            }
            
            languageDialog.Controls.Add(comboBox);
            
            // 添加确定按钮
            Button okButton = new Button();
            okButton.Text = "确定 / OK";
            okButton.Location = new Point(20, 100);
            okButton.DialogResult = DialogResult.OK;
            languageDialog.Controls.Add(okButton);
            
            // 添加取消按钮
            Button cancelButton = new Button();
            cancelButton.Text = "取消 / Cancel";
            cancelButton.Location = new Point(120, 100);
            cancelButton.DialogResult = DialogResult.Cancel;
            languageDialog.Controls.Add(cancelButton);
            
            // 设置接受和取消按钮
            languageDialog.AcceptButton = okButton;
            languageDialog.CancelButton = cancelButton;
            
            // 显示对话框
            if (languageDialog.ShowDialog() == DialogResult.OK)
            {
                string selectedLang = "";
                switch (comboBox.SelectedIndex)
                {
                    case 0:
                        selectedLang = "en";
                        break;
                    case 1:
                        selectedLang = "de";
                        break;
                    case 2:
                        selectedLang = "zh-CN";
                        break;
                }

                if (!string.IsNullOrEmpty(selectedLang))
                {
                    // 保存语言设置
                    OptionFileHandler.SetOptionValue("Language", selectedLang);
                    
                    // 更新主界面的语言选择下拉框
                    cbLanguage.SelectedIndex = comboBox.SelectedIndex;
                    
                    // 提示用户需要重启应用才能完全应用新语言
                    MessageBox.Show("语言设置已保存。重启应用程序后新语言将完全生效。\n\nLanguage settings saved. Restart the application for the new language to take full effect.", 
                        "提示 / Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        /**************************************************/

        #region WizardPage

        /**************************************************/

        private void bBack_Click(object sender, EventArgs e)
        {
            this.SuspendLayout();
            if (wizardPages.SelectedIndex > 0)
                wizardPages.SelectedIndex -= 1;
            this.ResumeLayout();
        }

        private void bNext_Click(object sender, EventArgs e)
        {
            if (IsNextPageOk())
            {
                this.SuspendLayout();
                if (wizardPages.SelectedIndex < wizardPages.TabPages.Count)
                    wizardPages.SelectedIndex += 1;
                this.ResumeLayout();
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            File.AppendAllText("debug.log", $"[{DateTime.Now}] MainForm_Load事件开始\n");
            
            // Necessary hack to display the correct button(s) after program start
            wizardPages.SelectedIndex = 1;
            wizardPages.SelectedIndex = 0;
            
            // 确保窗口可见
            this.Visible = true;
            this.WindowState = FormWindowState.Normal;
            this.BringToFront();
            this.Activate();
            
            // 确保语言设置按钮可见
            bLanguageSettings.Visible = true;
            bLanguageSettings.BringToFront();
            
            File.AppendAllText("debug.log", $"[{DateTime.Now}] MainForm_Load事件完成 - 窗口应该可见\n");
        }

        private void wizardPages_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (wizardPages.SelectedIndex == 0)
            {
                bBack.Hide();
            }
            else
            {
                bBack.Show();
            }

            if (wizardPages.SelectedIndex != (wizardPages.TabPages.Count - 1))
            {
                bFinish.Hide();
                bNext.Show();
            }
            else
            {
                bNext.Show();
                bFinish.Show();
            }
        }

        /**************************************************/

        #endregion WizardPage

        /**************************************************/

        private void bSelectOwnFont_Click(object sender, EventArgs e)
        {
            // Select File
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            openFileDialog.Filter = "Font|*.ttf;";
            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = false;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                tbCustomFontPath.Text = openFileDialog.FileName;
            }
        }

        private void bShortcutCreator_Click(object sender, EventArgs e)
        {
            (new ShortcutCreator()).ShowDialog();
        }

        private bool IsNextPageOk()
        {
            switch (wizardPages.SelectedTab.Text)
            {
                case "SpreadsheetSelect":
                    if (dataSource == null)
                    {
                        MessageBox.Show(Properties.Resources.MessageSelectSpreadsheet);
                        return false;
                    }
                    if (dataSource.PossibleRows == 0)
                    {
                        MessageBox.Show(Properties.Resources.MessageNoUsableRows);
                        return false;
                    }
                    break;

                case "PDFSelect":
                    if (pdf == null)
                    {
                        MessageBox.Show(Properties.Resources.MessageNoPDFSelected);
                        return false;
                    }
                    break;
            }

            return true;
        }

        private void llBulkPDFde_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://opensource.bulkpdf.de/");
        }

        private void llDokumentation_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://opensource.bulkpdf.de/documentation");
        }

        private void llLicenses_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            var licenses = new Licenses();
            licenses.ShowDialog();
        }

        /**************************************************/

        #region Fill

        /**************************************************/

        private void backGroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                PDFFiller.CreateFiles(pdf, finalize, unicode, customFont, tbCustomFontPath.Text, dataSource, pdfFields, tbOutputDir.Text + @"\", ConcatFilename, progressForm.SetPercent, progressForm.GetIsAborted);
            }
            catch (Exception ex)
            {
                ExceptionHandler.Throw(Properties.Resources.ExceptionPDFFileAlreadyExistsAndInUse, ex.ToString());

                this.Invoke((MethodInvoker)delegate
                {
                    progressForm.Close();
                });
            }
        }

        private void bFinish_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(tbOutputDir.Text))
            {
                if (Directory.Exists(tbOutputDir.Text))
                {
                    if (IsFilenameUnique())
                    {
                        DialogResult dialogResult = MessageBox.Show(String.Format(Properties.Resources.MessageCreateNPDFFilesInDir, dataSource.PossibleRows, tbOutputDir.Text), Properties.Resources.MessageAreYouSure, MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.Yes)
                        {
                            this.SuspendLayout();

                            progressForm = new ProgressForm();
                            progressForm.Show();
                            finalize = cbFinalize.Checked;
                            unicode = cbUnicode.Checked;
                            customFont = cbCustomFont.Checked;

                            BackgroundWorker backGroundWorker = new BackgroundWorker();
                            backGroundWorker.DoWork += backGroundWorker_DoWork;
                            backGroundWorker.RunWorkerAsync();

                            this.ResumeLayout();
                        }
                    }
                    else
                    {
                        MessageBox.Show(Properties.Resources.MessageFilenameNotUnique);
                    }
                }
                else
                {
                    MessageBox.Show(Properties.Resources.MessageOutputDirNotExist);
                }
            }
            else
            {
                MessageBox.Show(Properties.Resources.MessageNoOutputDirSelected);
            }
        }

        private void bSelectOutputPath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                tbOutputDir.Text = folderBrowserDialog.SelectedPath;
        }

        private string ConcatFilename(int dataSourceRow)
        {
            string filename = "";
            filename += tbPrefix.Text;
            if (cbUseValueFromDataSource.Checked)
                filename += dataSource.GetField(tempSelectedIndex + 1);
            filename += tbSuffix.Text;
            if (cbRowNumberFilename.Checked)
                filename += dataSourceRow;
            filename += ".pdf";

            return filename;
        }

        private bool IsFilenameUnique()
        {
            if (cbRowNumberFilename.Checked)
                return true;

            // Is dataSource unique?
            if (cbUseValueFromDataSource.Checked)
            {
                var filenameFields = new List<string>();

                dataSource.ResetRowCounter();
                for (int dataSourceRow = 1; dataSourceRow <= dataSource.PossibleRows; dataSourceRow++)
                {
                    string field = dataSource.GetField(tempSelectedIndex + 1);

                    if (filenameFields.Contains(field))
                    {
                        return false;
                    }
                    else
                    {
                        filenameFields.Add(field);
                    }

                    dataSource.NextRow();
                }

                return true;
            }

            if (dataSource.PossibleRows == 1)
                return true;

            return false;
        }

        /**************************************************/

        #endregion Fill

        /**************************************************/

        /**************************************************/

        #region PDFSelect

        /**************************************************/

        private void bSelectPDF_Click(object sender, EventArgs e)
        {
            // Select File
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            openFileDialog.Filter = "PDF|*.pdf";
            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = false;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // PDF
                if (pdf != null)
                    pdf.Close();
                pdf = new PDF();
                OpenPDF(openFileDialog.FileName);
            }
        }

        private void dgvBulkPDF_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            var senderGrid = (DataGridView)sender;

            if (senderGrid.Columns[e.ColumnIndex] is DataGridViewButtonColumn && e.RowIndex >= 0)
            {
                var fieldOptionForm = new FieldOptionForm(new Point(this.Location.X + Convert.ToInt32(this.Width / 2.5), this.Location.Y + Convert.ToInt32(this.Height / 2.5))
                    , pdfFields[(string)dgvBulkPDF.Rows[e.RowIndex].Cells["ColField"].Value], dataSource.Columns);
                fieldOptionForm.ShowDialog();
                if (fieldOptionForm.ShouldBeSaved)
                {
                    pdfFields[fieldOptionForm.PDFField.Name] = fieldOptionForm.PDFField;
                    string value = pdfFields[fieldOptionForm.PDFField.Name].CurrentValue;

                    if (fieldOptionForm.PDFField.UseValueFromDataSource)
                        value = pdfFields[fieldOptionForm.PDFField.Name].DataSourceValue;
                    else if (fieldOptionForm.PDFField.UseFixedValue)
                        value = pdfFields[fieldOptionForm.PDFField.Name].FixedValue;
                    if (pdfFields[fieldOptionForm.PDFField.Name].MakeReadOnly)
                        value = "[#]" + value;

                    dgvBulkPDF.Rows[e.RowIndex].Cells["ColValue"].Value = value;
                }
            }
        }

        private bool OpenPDF(string pdfPath)
        {
            try
            {
                ResetPDF();
                pdf = new PDF();
                pdf.Open(pdfPath);

                // Fill DataGridView
                tbPDF.Text = pdfPath;
                if (pdf.IsXFA)
                {
                    tbFormTyp.Text = "XFA Form";
                }
                else
                {
                    tbFormTyp.Text = "Acroform";
                }
                foreach (PDFField pdfField in pdf.ListFields())
                {
                    dgvBulkPDF.Rows.Add();
                    int row = dgvBulkPDF.Rows.Count - 1;

                    dgvBulkPDF.Rows[row].Cells["ColField"].Value = pdfField.Name;
                    dgvBulkPDF.Rows[row].Cells["ColTyp"].Value = pdfField.Typ;
                    dgvBulkPDF.Rows[row].Cells["ColValue"].Value = pdfField.CurrentValue;
                    pdfFields.Add(pdfField.Name, pdfField);

                    var dgvButtonCell = new DataGridViewButtonCell();
                    dgvButtonCell.Value = Properties.Resources.CellButtonSelect;
                    dgvBulkPDF.Rows[row].Cells["ColOption"] = dgvButtonCell;
                }

                return true;
            }
            catch (Exception ex)
            {
                ExceptionHandler.Throw(String.Format(Properties.Resources.ExceptionPDFIsCorrupted, pdfPath), ex.ToString());
                ResetPDF();
            }
            return false;
        }

        private void ResetPDF()
        {
            pdf = null;
            dgvBulkPDF.Rows.Clear();
            pdfFields.Clear();
            tbPDF.Text = "";
            tbFormTyp.Text = "";
        }

        /**************************************************/

        #endregion PDFSelect

        /**************************************************/

        /**************************************************/

        #region Save & Load

        /**************************************************/

        private void bLoadConfiguration_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            openFileDialog.Filter = "BulkPDF Options|*.bulkpdf";
            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = false;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var fileString = File.ReadAllText(openFileDialog.FileName);
                    if (fileString.Contains("BulkPDF-Business")) throw new Exception("This is a BulkPDF-Business configuration file and is not compatible with the open source freeware software BulkPDF.");
                    XDocument xDocument = XDocument.Parse(fileString);

                    //// Options
                    var xmlOptions = xDocument.Root.Element("Options");
                    // DataSource
                    ResetDataSource();
                    ResetPDF();
                    dataSource = new Spreadsheet();
                    if (!OpenSpreadsheet(Environment.ExpandEnvironmentVariables(xmlOptions.Element("DataSource").Element("Parameter").Value)))
                    {
                        throw new Exception();
                    }
                    cbSpreadsheetTable.SelectedIndex = cbSpreadsheetTable.Items.IndexOf(xmlOptions.Element("Spreadsheet").Element("Table").Value);

                    // PDF
                    if (!OpenPDF(Environment.ExpandEnvironmentVariables(xmlOptions.Element("PDF").Element("Filepath").Value)))
                    {
                        throw new Exception();
                    }

                    //// PDFFieldValues
                    foreach (var node in xDocument.Root.Element("PDFFieldValues").Descendants("PDFFieldValue"))
                    {
                        var name = node.Element("Name").Value;

                        for (int row = 0; row < dgvBulkPDF.Rows.Count; row++)
                        {
                            if ((string)dgvBulkPDF.Rows[row].Cells[0].Value == name)
                            {
                                var pdfField = pdfFields[name];
                                string value = pdfFields[name].CurrentValue;

                                pdfField.UseValueFromDataSource = Convert.ToBoolean(node.Element("UseValueFromDataSource")?.Value ?? "False");
                                if (pdfField.UseValueFromDataSource)
                                    pdfField.DataSourceValue = node.Element("NewValue")?.Value;

                                pdfField.UseFixedValue = Convert.ToBoolean(node.Element("UseFixedValue")?.Value ?? "False");
                                if (pdfField.UseFixedValue)
                                    pdfField.FixedValue = node.Element("NewValue")?.Value;
                                
                                pdfField.MakeReadOnly = Convert.ToBoolean(node.Element("MakeReadOnly")?.Value ?? "False");
                                pdfFields[name] = pdfField;

                                if (pdfFields[name].UseValueFromDataSource)
                                    value = pdfFields[name].DataSourceValue;

                                if (pdfFields[name].UseFixedValue)
                                    value = pdfFields[name].FixedValue;

                                if (pdfFields[name].MakeReadOnly)
                                    value = "[#]" + value;

                                dgvBulkPDF.Rows[row].Cells["ColValue"].Value = value;
                            }
                        }
                    }

                    //// Filename
                    var xmlFilename = xmlOptions.Element("Filename");
                    tbPrefix.Text = xmlFilename.Element("Prefix").Value;
                    cbUseValueFromDataSource.Checked = Convert.ToBoolean(xmlFilename.Element("ValueFromDataSource").Value);
                    cbDataSourceColumnsFilename.SelectedIndex = cbDataSourceColumnsFilename.Items.IndexOf(xmlFilename.Element("DataSource").Value);
                    tbSuffix.Text = xmlFilename.Element("Suffix").Value;
                    cbRowNumberFilename.Checked = Convert.ToBoolean(xmlFilename.Element("RowNumber").Value);

                    //// Other
                    cbFinalize.Checked = Convert.ToBoolean(xmlOptions.Element("Finalize").Value);
                    try
                    {
                        cbUnicode.Checked = Convert.ToBoolean(xmlOptions.Element("Unicode").Value);
                    }
                    catch
                    {
                        // Ignore. Ugly but don't hurt anyone.
                    }
                    try
                    {
                        cbCustomFont.Checked = Convert.ToBoolean(xmlOptions.Element("CustomFont").Value);
                        tbCustomFontPath.Text = Environment.ExpandEnvironmentVariables(xmlOptions.Element("CustomFontPath").Value);
                    }
                    catch
                    {
                        // Ignore. Ugly but don't hurt anyone.
                    }
                    tbOutputDir.Text = Environment.ExpandEnvironmentVariables(xmlOptions.Element("OutputDir").Value);

                    wizardPages.SelectedIndex = wizardPages.TabPages.Count - 1;
                }
                catch (Exception ex)
                {
                    ExceptionHandler.Throw(Properties.Resources.ExceptionConfigurationFileIsCorrupted, ex.ToString());
                }
            }
        }

        private void bSaveConfiguration_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "BulkPDF Options|*.bulkpdf";
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                var xmlWriterSettings = new XmlWriterSettings();
                xmlWriterSettings.Indent = true;
                xmlWriterSettings.IndentChars = "  ";
                var xmlWriter = XmlWriter.Create(saveFileDialog.FileName, xmlWriterSettings);
                xmlWriter.WriteStartDocument();
                xmlWriter.WriteStartElement("BulkPDF"); // <BulkPDF>
                xmlWriter.WriteElementString("Version", Application.ProductVersion.ToString());
                xmlWriter.WriteStartElement("Options"); // <Options>
                xmlWriter.WriteStartElement("DataSource"); // <DataSource>
                xmlWriter.WriteElementString("Typ", "Spreadsheet");
                xmlWriter.WriteElementString("Parameter", dataSource.Parameter);
                xmlWriter.WriteEndElement(); // </DataSource>
                xmlWriter.WriteStartElement("PDF"); // <PDF>
                xmlWriter.WriteElementString("Filepath", tbPDF.Text);
                xmlWriter.WriteEndElement(); // </PDF>
                xmlWriter.WriteStartElement("Spreadsheet"); // <Spreadsheet>
                xmlWriter.WriteElementString("Table", (string)cbSpreadsheetTable.SelectedItem);
                xmlWriter.WriteEndElement();  // </Spreadsheet>
                xmlWriter.WriteStartElement("Filename"); // <Filename>
                xmlWriter.WriteElementString("Prefix", tbPrefix.Text);
                xmlWriter.WriteElementString("ValueFromDataSource", cbUseValueFromDataSource.Checked.ToString());
                xmlWriter.WriteElementString("DataSource", (string)cbDataSourceColumnsFilename.SelectedItem);
                xmlWriter.WriteElementString("Suffix", tbSuffix.Text);
                xmlWriter.WriteElementString("RowNumber", cbRowNumberFilename.Checked.ToString());
                xmlWriter.WriteEndElement(); // </Filename>
                xmlWriter.WriteElementString("Finalize", cbFinalize.Checked.ToString());
                xmlWriter.WriteElementString("Unicode", cbUnicode.Checked.ToString());
                xmlWriter.WriteElementString("CustomFont", cbCustomFont.Checked.ToString());
                xmlWriter.WriteElementString("CustomFontPath", tbCustomFontPath.Text);
                xmlWriter.WriteElementString("OutputDir", tbOutputDir.Text);
                xmlWriter.WriteEndElement(); // </Options>

                xmlWriter.WriteStartElement("PDFFieldValues"); // <PDFFieldValues>
                foreach (var pdfField in pdfFields)
                {
                    xmlWriter.WriteStartElement("PDFFieldValue"); // <PDFFieldValue>
                    xmlWriter.WriteElementString("Name", pdfField.Value.Name);
                    xmlWriter.WriteElementString("NewValue", pdfField.Value.UseValueFromDataSource ? pdfField.Value.DataSourceValue : pdfField.Value.FixedValue);
                    xmlWriter.WriteElementString("UseValueFromDataSource", pdfField.Value.UseValueFromDataSource.ToString());
                    xmlWriter.WriteElementString("UseFixedValue", pdfField.Value.UseFixedValue.ToString());
                    xmlWriter.WriteElementString("MakeReadOnly", pdfField.Value.MakeReadOnly.ToString());
                    xmlWriter.WriteEndElement(); // </PDFFieldValue>
                }
                xmlWriter.WriteEndElement(); // </PDFFieldValues>
                xmlWriter.WriteEndElement(); // </BulkPDF>
                xmlWriter.WriteEndDocument();
                xmlWriter.Close();
            }
        }

        /**************************************************/

        #endregion Save & Load

        /**************************************************/

        /**************************************************/

        #region DataSource

        /**************************************************/

        private void ResetDataSource()
        {
            tbSpreadsheet.Text = "";
            dataSource = null;
            lPossibleRowsValue.Text = "0";
            lPossibleColumnsValue.Text = "0";
            cbDataSourceColumnsExampleSpreadsheet.Items.Clear();
            cbDataSourceColumnsFilename.Items.Clear();
            cbSpreadsheetTable.Items.Clear();
        }

        private void UpdateDataSource()
        {
            // Textbox
            tbSpreadsheet.Text = dataSource.Parameter;
            lPossibleRowsValue.Text = dataSource.PossibleRows.ToString();
            lPossibleColumnsValue.Text = dataSource.Columns.Count.ToString();

            // Dropdown
            cbDataSourceColumnsExampleSpreadsheet.Items.Clear();
            cbDataSourceColumnsFilename.Items.Clear();
            foreach (var column in dataSource.Columns)
            {
                cbDataSourceColumnsExampleSpreadsheet.Items.Add(column);
                cbDataSourceColumnsFilename.Items.Add(column);
            }
            cbDataSourceColumnsExampleSpreadsheet.SelectedIndex = 0;
            cbDataSourceColumnsFilename.SelectedIndex = 0;
        }

        /**************************************************/

        #endregion DataSource

        /**************************************************/

        /**************************************************/

        #region ExampleFilename

        /**************************************************/

        private void cbDataSourceColumnsFilename_SelectedIndexChanged(object sender, EventArgs e)
        {
            tempSelectedIndex = cbDataSourceColumnsFilename.SelectedIndex;
            UpdateExampleFilename();
        }

        private void cbRowNumberFilename_CheckedChanged(object sender, EventArgs e)
        {
            UpdateExampleFilename();
        }

        private void cbUseValueFromDataSource_CheckedChanged(object sender, EventArgs e)
        {
            UpdateExampleFilename();
        }

        private void tbPrefix_TextChanged(object sender, EventArgs e)
        {
            UpdateExampleFilename();
        }

        private void tbSuffix_TextChanged(object sender, EventArgs e)
        {
            UpdateExampleFilename();
        }

        private void UpdateExampleFilename()
        {
            dataSource.ResetRowCounter();
            tbExampleFilename.Text = ConcatFilename(1);
        }

        /**************************************************/

        #endregion ExampleFilename

        /**************************************************/

        /**************************************************/

        #region SpreadsheetSelect

        /**************************************************/

        private void bSelectSpreadsheet_Click(object sender, EventArgs e)
        {
            // Select File
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            openFileDialog.Filter = "Spreadsheet|*.xlsx;*.xlsm;";
            openFileDialog.FilterIndex = 1;
            openFileDialog.Multiselect = false;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // DataSource
                if (dataSource != null)
                    dataSource.Close();
                ResetDataSource();
                ResetPDF();
                dataSource = new Spreadsheet();

                OpenSpreadsheet(openFileDialog.FileName);
            }
        }

        private void cbTable_SelectedIndexChanged(object sender, EventArgs e)
        {
            ((Spreadsheet)dataSource).SetSheet((string)cbSpreadsheetTable.SelectedItem);
            UpdateDataSource();
            ResetPDF();
        }

        private bool OpenSpreadsheet(string filePath)
        {
            try
            {
                dataSource.Open(filePath);

                // Sheet
                var sheetNames = ((Spreadsheet)dataSource).GetSheetNames();
                foreach (var sheet in sheetNames)
                    cbSpreadsheetTable.Items.Add(sheet);
                cbSpreadsheetTable.SelectedIndex = 0;

                UpdateDataSource();
                return true;
            }
            catch (IOException ex)
            {
                ExceptionHandler.Throw(Properties.Resources.ExceptionSpreadsheetAlreadyInUse, ex.ToString());
            }
            catch (FileFormatException ex)
            {
                ExceptionHandler.Throw(Properties.Resources.ExceptionSpreadsheetIsCorrupted, ex.ToString());
            }
            return false;
        }

        /**************************************************/

        #endregion SpreadsheetSelect

        /**************************************************/

        private void llSupport_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("mailto:support@bulkpdf.de");
        }

        private void MainForm_Shown(object sender, EventArgs e)
        {
            // 确保语言设置按钮可见
            bLanguageSettings.Visible = true;
            bLanguageSettings.BringToFront();
            
            // 记录调试信息
            File.AppendAllText("debug.log", $"[{DateTime.Now}] MainForm_Shown: bLanguageSettings.Visible={bLanguageSettings.Visible}, Location={bLanguageSettings.Location}, Size={bLanguageSettings.Size}\n");
        }
    }
}