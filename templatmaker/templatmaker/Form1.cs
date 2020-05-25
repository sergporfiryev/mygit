using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace templatmaker
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private readonly string TemplateFileName = @"C:\Users\ASUS\Desktop\DOCX Document.docx";

        private void ReplaceWordStub(string StubToReplace, string text, Word.Document wordDocument)
        {
            var range = wordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: StubToReplace, ReplaceWith: text);
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            var WordApp = new Word.Application
            {
                Visible = false
            };

            try
            {
                var wordDocument = WordApp.Documents.Open(TemplateFileName);
                ReplaceWordStub("{ОБЪЕКТ}", textBox1.Text, wordDocument);
                ReplaceWordStub("{ЗАКАЗЧИК}", textBox2.Text, wordDocument);
                ReplaceWordStub("{Раздел проекта 1}", textBox3.Text, wordDocument);
                ReplaceWordStub("{Раздел проекта 2}", textBox4.Text, wordDocument);
                ReplaceWordStub("{Полное наименование проекта}", textBox5.Text, wordDocument);
                ReplaceWordStub("{адрес объекта}", textBox6.Text, wordDocument);
                ReplaceWordStub("{должность}", textBox7.Text, wordDocument);
                ReplaceWordStub("{Ф.И.О.}", textBox8.Text, wordDocument);
                ReplaceWordStub("{год}", textBox9.Text, wordDocument);

                wordDocument.SaveAs(@"C:\Users\ASUS\Desktop\Титульник.docx");
                WordApp.Visible = true;
            }
            catch
            {
                MessageBox.Show("Произошла ошибка");
            }
        }
    }
}
