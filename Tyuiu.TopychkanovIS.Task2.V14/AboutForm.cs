using System.Drawing;
using System.Windows.Forms;
using MaterialSkin;
using MaterialSkin.Controls;

namespace Tyuiu.TopychkanovIS.Task2.V14
{
    public partial class AboutForm : MaterialForm
    {
        private bool dragging = false;
        private Point dragCursorPoint;
        private Point dragFormPoint;

        public AboutForm()
        {
            InitializeComponent();

            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(Primary.Blue400, Primary.Blue500, Primary.Blue500, Accent.LightBlue200, TextShade.WHITE);
            this.ClientSize = new System.Drawing.Size(400, 300);
            this.Name = "AboutForm";
            this.Text = "О программе";
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MouseDown += new MouseEventHandler(AboutForm_MouseDown);
            this.MouseMove += new MouseEventHandler(AboutForm_MouseMove);
            this.MouseUp += new MouseEventHandler(AboutForm_MouseUp);
            this.ResumeLayout(false);

            InitializeControls();
        }

        private void InitializeControls()
        {
            var labelInfo = new MaterialLabel
            {
                Text = "Информация о предметной области и авторе:\n\n" +
                       "Предметная область: Турагентство\n" +
                       "Описание: Система управления турами, клиентами, заказами и услугами для турагентства. " +
                       "Позволяет эффективно управлять данными, предоставлять информацию о турах и услугах клиентам.\n\n" +
                       "Автор: Иван Топычканов\n" +
                       "Группа: АСОиУб-21-2\n" +
                       "Дата создания: Май 2024\n",
                Dock = DockStyle.Fill,
                Font = new System.Drawing.Font("Roboto", 12F)
            };

            Controls.Add(labelInfo);
        }

        private void AboutForm_MouseDown(object sender, MouseEventArgs e)
        {
            dragging = true;
            dragCursorPoint = Cursor.Position;
            dragFormPoint = this.Location;
        }

        private void AboutForm_MouseMove(object sender, MouseEventArgs e)
        {
            if (dragging)
            {
                Point diff = Point.Subtract(Cursor.Position, new Size(dragCursorPoint));
                this.Location = Point.Add(dragFormPoint, new Size(diff));
            }
        }

        private void AboutForm_MouseUp(object sender, MouseEventArgs e)
        {
            dragging = false;
        }
    }
}
