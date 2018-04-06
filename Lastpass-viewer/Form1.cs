using System;
using System.Windows.Forms;
using LastPass;
using System.Data;
using System.Drawing;

namespace Lastpass_viewer
{
    public partial class Form1 : Form
    {
        // Fetch and create the vault from LastPass
        Vault vault = null;

        public Form1()
        {
            InitializeComponent();
        }

        // Input Box
        public static DialogResult InputBox(string title, string promptText, ref string value)
        {
            Form form = new Form();
            Label label = new Label();
            TextBox textBox = new TextBox();
            Button buttonOk = new Button();
            Button buttonCancel = new Button();

            form.Text = title;
            label.Text = promptText;
            textBox.Text = value;

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
            value = textBox.Text;
            return dialogResult;
        }

        /// <summary> 
        /// Exports the datagridview values to Excel. 
        /// </summary> 
        private void ExportToExcel()
        {
            // Creating a Excel object. 
            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            try
            {

                worksheet = workbook.ActiveSheet;

                worksheet.Name = "LastPass Accounts";

                int cellRowIndex = 1;
                int cellColumnIndex = 1;

                //Loop through each row and read value from each column. 
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        // Excel index starts from 1,1. As first Row would have the Column headers, adding a condition check. 
                        if (cellRowIndex == 1)
                        {
                            worksheet.Cells[cellRowIndex, cellColumnIndex] = dataGridView1.Columns[j].HeaderText;
                        }
                        else
                        {
                            worksheet.Cells[cellRowIndex, cellColumnIndex] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }
                        cellColumnIndex++;
                    }
                    cellColumnIndex = 1;
                    cellRowIndex++;
                }

                //Getting the location and file name of the excel to save from user. 
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveDialog.FilterIndex = 1;

                if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    workbook.SaveAs(saveDialog.FileName);
                    MessageBox.Show("Export Successful", "Export information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                excel.Quit();
                workbook = null;
                excel = null;
            }

        }
            

        private void btnExport_Click(object sender, EventArgs e)
        {
            ExportToExcel();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            //LastPass credentials
            string username = txtUsername.Text;
            string password = txtPassword.Text;

            try
            {
                // Frist try basic authentication
                vault = Vault.Create(username, password);
            }
            catch (LoginException ex)
            {
                switch (ex.Reason)
                {
                    case LoginException.FailureReason.LastPassIncorrectGoogleAuthenticatorCode:
                        {
                            // Request Google Authenticator code
                            var code = "";

                            // Input box
                            string value = "PIN";
                            if (InputBox("Google Authenticator", "Google Authenticator Code:", ref value) == DialogResult.OK)
                            {
                                code = value;
                            }

                            // Now try with GAuth code
                            vault = Vault.Create(username, password, code);

                            break;
                        }
                    case LoginException.FailureReason.LastPassIncorrectYubikeyPassword:
                        {
                            // Request Yubikey password
                            var yubikeyPassword = "";

                            // Input box
                            string value = "Password";
                            if (InputBox("Yubikey", "Yubikey Password:", ref value) == DialogResult.OK)
                            {
                                yubikeyPassword = value;
                            }

                            // Now try with Yubikey password
                            vault = Vault.Create(username, password, yubikeyPassword);

                            break;
                        }
                    default:
                        {
                            throw;
                        }
                }
            }

            // Dump all the accounts
            for (var i = 0; i < vault.Accounts.Length; ++i)
            {
                var account = vault.Accounts[i];
               
                int rowId = dataGridView1.Rows.Add();
                DataGridViewRow row = dataGridView1.Rows[rowId];

                row.Cells["AA"].Value = i + 1;
                row.Cells["ID"].Value = account.Id;
                row.Cells["Name1"].Value = account.Name;
                row.Cells["Username"].Value = account.Username;
                row.Cells["Password"].Value = account.Password;
                row.Cells["Url"].Value = account.Url;
                row.Cells["Group"].Value = account.Group;

            }
        }

       
    }
}
