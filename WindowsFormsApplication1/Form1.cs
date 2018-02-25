/*using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
*/


using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
 
 
namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        private void Replace(Word._Document oDoc, Word._Application oWord, string toFind, string toReplace)
        {
            object toFindTxt = toFind;
            object toReplaceTxt = toReplace;
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object nmatchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object MatchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = Word.WdReplace.wdReplaceAll;
            object wrap = Word.WdFindWrap.wdFindContinue;

            foreach (Word.Range range in oDoc.StoryRanges)
            {
                range.Find.ClearFormatting();
                range.Find.Replacement.ClearFormatting();
                range.Find.Execute(
                   ref toFindTxt, ref matchCase, ref matchWholeWord, ref matchWildCards, ref matchSoundsLike,
                 ref nmatchAllWordForms, ref forward, ref wrap, ref format, ref toReplaceTxt, ref replace,
                ref matchKashida, ref MatchDiacritics, ref matchAlefHamza, ref matchControl
               );

                Word.Range tmpRange = range.NextStoryRange;

                while (tmpRange != null)
                {
                    tmpRange.Find.Execute(
                     ref toFindTxt, ref matchCase, ref matchWholeWord, ref matchWildCards, ref matchSoundsLike,
                   ref nmatchAllWordForms, ref forward, ref wrap, ref format, ref toReplaceTxt, ref replace,
                   ref matchKashida, ref MatchDiacritics, ref matchAlefHamza, ref matchControl
                 );
                    if (tmpRange.NextStoryRange != null)
                    { tmpRange = tmpRange.NextStoryRange; }
                    else { tmpRange = null; }
                }
            }
        }
        void Button1Click(object sender, EventArgs e)
        {


            var fname = textBox1.Text;
            var name = textBox2.Text;
            var otch = textBox3.Text;
            var year = textBox4.Text;



            var wordApp = new Word.Application();
            wordApp.Visible = false; // скрываем от пользователя окно Word'a
                                     // до завершения каких-либо действий с ним
            var wordDocument = wordApp.Documents.Open("c:\\C#\\12\\e12\\Attestat.docx");

            try
            {


                Replace(wordDocument, wordApp, "{Fam}", fname);
                Replace(wordDocument, wordApp, "{Nam}", name);
                Replace(wordDocument, wordApp, "{Otch}", otch);
                Replace(wordDocument, wordApp, "{Year}", year);

                wordDocument.SaveAs("c:\\C#\\12\\e12\\result.docx"); // сохраняем новую заявку

                //                wordDocument.Close(); // Перенести в finally, сделать проверку на открытие
            }
            catch (Exception Ситуация)
            {
                // Отчет обо всех возможных ошибках:
                MessageBox.Show(Ситуация.Message, "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            finally
            {
                //wordApp.Quit();
                wordApp.Visible = true;
            }


        }
        void Button2Click(object sender, EventArgs e)
        {
            Close();
        }













    }
}