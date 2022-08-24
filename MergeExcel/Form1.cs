namespace MergeExcel
{
    public partial class Form1 : Form
    {
        OpenFileDialog dialog;

        public Form1()
        {
            InitializeComponent();

            textBox2.Text = "5";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dialog = new OpenFileDialog();
            dialog.Title = "����ഫ�ɮ�";
            dialog.InitialDirectory = ".\\";
            dialog.Filter = "xlsx files (*.*)|*.xlsx";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                //MessageBox.Show(dialog.FileName);

                textBox1.Text = dialog.FileName;

                textBox3.Text = String.Format("{0} �}�l�ഫ", dialog.FileName);
            }
        }

        private void button2_Click(object senders, EventArgs e)
        {
            int titleRow;


            if(dialog==null || string.IsNullOrEmpty(dialog.FileName))
            {
                textBox3.Text = textBox3.Text+ Environment.NewLine + "�٥����Excel�ɮ�";
                return;
            }

            if(!int.TryParse(textBox2.Text, out titleRow))
            {
                textBox3.Text = textBox3.Text + Environment.NewLine + "�L�Ī����D�C";
                return;
            }

            Service service = new Service();

            service.ReadExcel(dialog.FileName, titleRow);

            textBox3.Text = textBox3.Text + Environment.NewLine + "�X�֧����A���˵��̫�@�i�u�@��";
        }
    }
}