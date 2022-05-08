using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace pdfSampleWGUI
{
    public partial class Form1 : Form
    {
        Aspose.Pdf.Document pdf;
        Aspose.Pdf.Text.TextFragmentCollection fragments;
        List<TextBox> boxes;
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            if(openFileDialog1.FileName != "" && openFileDialog1.FileName.EndsWith(".pdf"))
            {
                pdf = new Aspose.Pdf.Document(openFileDialog1.FileName);
                var textFragmentAbsorberAddress = new Aspose.Pdf.Text.TextFragmentAbsorber();
                pdf.Pages[1].Accept(textFragmentAbsorberAddress);
                fragments = textFragmentAbsorberAddress.TextFragments;
                int i = 0;
                boxes = new List<TextBox>();
                foreach(var f in fragments)
                {
                    var box = new TextBox() { Text = f.Text };
                    var under = new CheckBox() { Text = "下線",Checked = f.TextState.Underline, Name = i.ToString()};
                    var strikeout = new CheckBox() { Text = "打消し", Checked = f.TextState.StrikeOut, Name = i.ToString() };
                    this.Controls.Add(box);
                    flowLayoutPanel1.Controls.Add(box);
                    under.CheckedChanged += Under_CheckedChanged;
                    flowLayoutPanel1.Controls.Add(under);
                    strikeout.CheckedChanged += Strikeout_CheckedChanged;
                    flowLayoutPanel1.Controls.Add(strikeout);
                    boxes.Add(box);
                    i++;
                }
                button1.Enabled = false;
            }
            else
            {
                MessageBox.Show("ファイル名が不正です", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Under_CheckedChanged(object sender, EventArgs e)
        {
            var checkbox = (CheckBox)sender;
            var frag = fragments.ElementAt(Int32.Parse(checkbox.Name));
            frag.TextState.Underline = checkbox.Checked;
        }

        private void Strikeout_CheckedChanged(object sender, EventArgs e)
        {
            var checkbox = (CheckBox)sender;
            var frag = fragments.ElementAt(Int32.Parse(checkbox.Name));
            frag.TextState.StrikeOut = checkbox.Checked;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            saveFileDialog1.FileName = openFileDialog1.FileName.Replace(".pdf","_mod.pdf");
            saveFileDialog1.ShowDialog();
            if(saveFileDialog1.FileName != "" && saveFileDialog1.FileName.EndsWith(".pdf"))
            {
                try
                {
                    for(int i = 0; i < boxes.Count; i++)
                    {
                        var f = fragments.ElementAt(i);
                        f.Text = boxes.ElementAt(i).Text;
                    }
                    pdf.Save(saveFileDialog1.FileName);
                    MessageBox.Show("書き込み成功", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    this.Close();
                }
                catch (Exception)
                {
                    MessageBox.Show("対応していない文字が含まれています", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("ファイル名が不正です", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        
    }
}
