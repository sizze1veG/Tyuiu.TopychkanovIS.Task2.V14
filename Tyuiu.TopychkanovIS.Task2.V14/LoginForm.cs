using MaterialSkin;
using MaterialSkin.Controls;
using System;
using System.Data.OleDb;
using System.Drawing;
using System.Windows.Forms;

namespace Tyuiu.TopychkanovIS.Task2.V14
{
    public partial class LoginForm : MaterialForm
    {
        private MaterialSingleLineTextField usernameTextBox;
        private MaterialSingleLineTextField passwordTextBox;
        private MaterialRaisedButton loginButton;
        private MaterialRaisedButton registerButton;
        private OleDbConnection connection;
        public LoginForm()
        {
            InitializeComponent();
            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(Primary.Blue400, Primary.Blue500, Primary.Blue500, Accent.LightBlue200, TextShade.WHITE);

            this.Size = new Size(400, 300);

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

            loginButton = new MaterialRaisedButton
            {
                Text = "Login",
                Size = new Size(250, 40),
                Location = new Point((this.ClientSize.Width - 250) / 2, 180),
            };
            loginButton.Click += new EventHandler(loginButton_Click);

            registerButton = new MaterialRaisedButton
            {
                Text = "Register",
                Size = new Size(250, 40),
                Location = new Point((this.ClientSize.Width - 250) / 2, 230),
            };
            registerButton.Click += new EventHandler(registerButton_Click);

            Controls.Add(usernameTextBox);
            Controls.Add(passwordTextBox);
            Controls.Add(loginButton);
            Controls.Add(registerButton);
        }

        private void loginButton_Click(object sender, EventArgs e)
        {
            try
            {
                connection.Open();
                string query = $"SELECT * FROM Аккаунты WHERE [имя пользователя]='{usernameTextBox.Text}' AND [пароль]='{ComputeHash(passwordTextBox.Text)}'";
                OleDbCommand command = new OleDbCommand(query, connection);
                OleDbDataReader reader = command.ExecuteReader();

                if (reader.HasRows)
                {
                    // Open main application form
                    //MainForm mainForm = new MainForm();
                    //mainForm.Show();
                    //this.Hide();
                }
                else
                {
                    MessageBox.Show("Invalid username or password.");
                }
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

        private void registerButton_Click(object sender, EventArgs e)
        {
            RegistrationForm registrationForm = new RegistrationForm();
            registrationForm.Show();
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
