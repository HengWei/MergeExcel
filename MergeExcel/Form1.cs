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
            dialog.Title = "選擇轉換檔案";
            dialog.InitialDirectory = ".\\";
            dialog.Filter = "xlsx files (*.*)|*.xlsx";

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                //MessageBox.Show(dialog.FileName);

                textBox1.Text = dialog.FileName;

                textBox3.Text = String.Format("{0} 開始轉換", dialog.FileName);
            }
        }

        private void button2_Click(object senders, EventArgs e)
        {
            int titleRow;


            if(dialog==null || string.IsNullOrEmpty(dialog.FileName))
            {
                textBox3.Text = textBox3.Text+ Environment.NewLine + "還未選取Excel檔案";
                return;
            }

            if(!int.TryParse(textBox2.Text, out titleRow))
            {
                textBox3.Text = textBox3.Text + Environment.NewLine + "無效的標題列";
                return;
            }

            Service service = new Service();

            service.ReadExcel(dialog.FileName, titleRow);

            textBox3.Text = textBox3.Text + Environment.NewLine + "合併完畢，請檢視最後一張工作表";
        }
    }
}