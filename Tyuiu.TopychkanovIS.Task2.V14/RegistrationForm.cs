using MaterialSkin;
using MaterialSkin.Controls;
using System;
using System.Data.OleDb;
using System.Drawing;
using System.Windows.Forms;

namespace Tyuiu.TopychkanovIS.Task2.V14
{
    public partial class RegistrationForm : MaterialForm
    {
        private MaterialSingleLineTextField usernameTextBox;
        private MaterialSingleLineTextField passwordTextBox;
        private MaterialSingleLineTextField confirmPasswordTextBox;
        private MaterialRaisedButton registerButton;
        private MaterialCheckBox showPasswordCheckBox;
        private OleDbConnection connection;

        public RegistrationForm()
        {
            InitializeComponent();

            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(Primary.Blue400, Primary.Blue500, Primary.Blue500, Accent.LightBlue200, TextShade.WHITE);

            this.Size = new Size(500, 400);

            InitializeControls();
            connection = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.16.0;Data Source=lab2_TopychkanovIS_bd.accdb");
        }

        private void InitializeControls()
        {
            usernameTextBox = new MaterialSingleLineTextField
            {
                Hint = "Username",
                Size = new Size(250, 30),
                Location = new Point((this.ClientSize.Width - 250) / 2, 80),
            };

            passwordTextBox = new MaterialSingleLineTextField
            {
                Hint = "Password",
                Size = new Size(250, 30),
                Location = new Point((this.ClientSize.Width - 250) / 2, 130),
                PasswordChar = '*'
            };

            confirmPasswordTextBox = new MaterialSingleLineTextField
            {
                Hint = "Confirm Password",
                Size = new Size(250, 30),
                Location = new Point((this.ClientSize.Width - 250) / 2, 180),
                PasswordChar = '*'
            };

            showPasswordCheckBox = new MaterialCheckBox
            {
                Text = "Show Password",
                Location = new Point((this.ClientSize.Width - 250) / 2, 230),
                Size = new Size(250, 30)
            };
            showPasswordCheckBox.CheckedChanged += new EventHandler(showPasswordCheckBox_CheckedChanged);

            registerButton = new MaterialRaisedButton
            {
                Text = "Register",
                Size = new Size(250, 40),
                Location = new Point((this.ClientSize.Width - 250) / 2, 280),
            };
            registerButton.Click += new EventHandler(registerButton_Click);

            Controls.Add(usernameTextBox);
            Controls.Add(passwordTextBox);
            Controls.Add(confirmPasswordTextBox);
            Controls.Add(showPasswordCheckBox);
            Controls.Add(registerButton);
        }

        private void registerButton_Click(object sender, EventArgs e)
        {
            if (passwordTextBox.Text != confirmPasswordTextBox.Text)
            {
                MessageBox.Show("Passwords do not match.");
                return;
            }

            try
            {
                connection.Open();
                string query = "INSERT INTO Аккаунты ([имя пользователя], [пароль]) VALUES (@username, @password)";
                OleDbCommand command = new OleDbCommand(query, connection);
                command.Parameters.AddWithValue("@username", usernameTextBox.Text);
                command.Parameters.AddWithValue("@password", ComputeHash(passwordTextBox.Text));
                command.ExecuteNonQuery();
                MessageBox.Show("Registration successful.");
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                connection.Close();
            }
        }

        private void showPasswordCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            passwordTextBox.PasswordChar = showPasswordCheckBox.Checked ? '\0' : '*';
            confirmPasswordTextBox.PasswordChar = showPasswordCheckBox.Checked ? '\0' : '*';
        }

        private string ComputeHash(string input)
        {
            using (var sha256 = System.Security.Cryptography.SHA256.Create())
            {
                var bytes = System.Text.Encoding.UTF8.GetBytes(input);
                var hash = sha256.ComputeHash(bytes);
                return Convert.ToBase64String(hash);
            }
        }
    }
}
