using Guna.UI2.WinForms;
using Microsoft.Office.Interop.Word;
using System.Drawing.Imaging;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using System.Windows.Forms;
using Spire.Doc;
using Spire.Pdf;
using System.IO;
using System.Drawing.Imaging;
using Document = Spire.Doc.Document;

namespace Convert_App
{
    public partial class Form1 : Form
    {
        private string docxPath;

        public Form1()
        {
            InitializeComponent();
        }
        private void label1_Click(object sender, EventArgs e)
        {

        }
        private void guna2Button6_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "PNG Files|*.png;",
                Title = "Bir PNG dosyasý seçin"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string inputFilePath = openFileDialog.FileName;

                guna2TextBox1.Text = "Seçilen Dosya: " + System.IO.Path.GetFileName(inputFilePath);
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "JPEG Files|*.jpg;*.jpeg",
                    Title = "Dosyayý Kaydet",
                    FileName = System.IO.Path.GetFileNameWithoutExtension(inputFilePath) + ".jpeg"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string outputFilePath = saveFileDialog.FileName;

                    try
                    {
                        using (Image image = Image.FromFile(inputFilePath))
                        {
                            image.Save(outputFilePath, System.Drawing.Imaging.ImageFormat.Jpeg);
                        }

                        MessageBox.Show("Dosya baþarýyla JPEG formatýna dönüþtürüldü!", "Baþarýlý", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Bir hata oluþtu: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void guna2Button3_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "JPEG Files|*.jpg;*.jpeg",
                Title = "Bir JPEG dosyasý seçin"
            };

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string inputFilePath = openFileDialog.FileName;

                guna2TextBox1.Text = "Seçilen Dosya: " + System.IO.Path.GetFileName(inputFilePath);
                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "PNG Files|*.png",
                    Title = "Dosyayý Kaydet",
                    FileName = System.IO.Path.GetFileNameWithoutExtension(inputFilePath) + ".png"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string outputFilePath = saveFileDialog.FileName;

                    try
                    {
                        using (Image image = Image.FromFile(inputFilePath))
                        {
                            image.Save(outputFilePath, System.Drawing.Imaging.ImageFormat.Png);
                        }

                        MessageBox.Show("Dosya baþarýyla PNG formatýna dönüþtürüldü!", "Baþarýlý", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Bir hata oluþtu: " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "JPG files (*.jpg)|*.jpg";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string jpgFilePath = openFileDialog.FileName;
                    guna2TextBox1.Text = jpgFilePath;
                    using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                    {
                        saveFileDialog.Filter = "ICO files (*.ico)|*.ico";
                        saveFileDialog.FileName = "image.ico"; 
                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            string icoFilePath = saveFileDialog.FileName;
                            using (Image jpgImage = Image.FromFile(jpgFilePath))
                            {
                                jpgImage.Save(icoFilePath, ImageFormat.Icon);
                                MessageBox.Show("Dönüþtürme iþlemi tamamlandý!", "Baþarýlý", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                    }
                }
            }
        }


        private void guna2Button5_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "ICO files (*.ico)|*.ico";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string icoFilePath = openFileDialog.FileName;
                    guna2TextBox1.Text = icoFilePath;
                    using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                    {
                        saveFileDialog.Filter = "JPEG files (*.jpg)|*.jpg";
                        saveFileDialog.FileName = "image.jpg";
                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            string jpegFilePath = saveFileDialog.FileName;
                            using (Image icoImage = Image.FromFile(icoFilePath))
                            {
                                icoImage.Save(jpegFilePath, ImageFormat.Jpeg);
                                MessageBox.Show("Dönüþtürme iþlemi tamamlandý!", "Baþarýlý", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                    }
                }
            }
        }

        private void guna2Button7_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "PNG files |*.png";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string pngFilePath = openFileDialog.FileName;
                    guna2TextBox1.Text = pngFilePath;
                    using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                    {
                        saveFileDialog.Filter = "ICO files (*.ico)|*.ico";
                        saveFileDialog.FileName = "image.ico"; 
                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            string icoFilePath = saveFileDialog.FileName;
                            using (Image pngImage = Image.FromFile(pngFilePath))
                            {
                                pngImage.Save(icoFilePath, ImageFormat.Icon);
                                MessageBox.Show("Dönüþtürme iþlemi tamamlandý!", "Baþarýlý", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                    }
                }
            }
        }
       

        private void guna2Button8_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "ICO files (*.ico)|*.ico";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string icoFilePath = openFileDialog.FileName;
                    guna2TextBox1.Text = icoFilePath;
                    using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                    {
                        saveFileDialog.Filter = "PNG files (*.png)|";
                        saveFileDialog.FileName = "image.png";
                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            string pngFilePath = saveFileDialog.FileName;
                            using (Image icoImage = Image.FromFile(icoFilePath))
                            {
                                icoImage.Save(pngFilePath, ImageFormat.Png);
                                MessageBox.Show("Dönüþtürme iþlemi tamamlandý!", "Baþarýlý", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                    }
                }
            }
        }
        

        private void guna2Button9_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Word Document (*.docx)|*.docx";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string inputFilePath = openFileDialog.FileName;
                guna2TextBox1.Text = "Seçilen Dosya: " + Path.GetFileName(inputFilePath);
                string docxFilePath = openFileDialog.FileName;

                try
                {
                    Document document = new Document();
                    document.LoadFromFile(docxFilePath);

                    string pdfFilePath = System.IO.Path.ChangeExtension(docxFilePath, ".pdf");

                    document.SaveToFile(pdfFilePath, Spire.Doc.FileFormat.PDF);

                    MessageBox.Show("Dönüþtürme Baþarýlý! Kayýt edildiði konum: " + pdfFilePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Hata: " + ex.Message);
                }
            }
        }

        private void guna2Button10_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "PDF Files (*.pdf)|*.pdf";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string inputFilePath = openFileDialog.FileName;
                guna2TextBox1.Text = "Seçilen Dosya: " + Path.GetFileName(inputFilePath);
                string pdfFilePath = openFileDialog.FileName;

                try
                {
                    Spire.Pdf.PdfDocument pdfDocument = new Spire.Pdf.PdfDocument();
                    pdfDocument.LoadFromFile(pdfFilePath);

                    string docxFilePath = System.IO.Path.ChangeExtension(pdfFilePath, ".docx");

                    pdfDocument.SaveToFile(docxFilePath, Spire.Pdf.FileFormat.DOCX);

                    MessageBox.Show("Dönüþtürme Baþarýlý! Kayýt edildiði konum: " + docxFilePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Hata: " + ex.Message);
                }
            }
        }

        private void guna2ControlBox3_Click(object sender, EventArgs e)
        {

        }

        private void guna2ControlBox3_MouseEnter(object sender, EventArgs e)
        {
            guna2ControlBox3.FillColor = System.Drawing.Color.Red;
        }

        private void guna2ControlBox3_MouseLeave(object sender, EventArgs e)
        {
            guna2ControlBox3.FillColor = System.Drawing.Color.White;
        }

        private void guna2Button3_MouseEnter(object sender, EventArgs e)
        {
            guna2Button3.FillColor = System.Drawing.Color.White;
        }

        private void guna2Button3_MouseLeave(object sender, EventArgs e)
        {
            guna2Button3.FillColor = System.Drawing.Color.CadetBlue;
        }

        private void guna2Button6_MouseEnter(object sender, EventArgs e)
        {
            guna2Button6.FillColor = System.Drawing.Color.White;
        }

        private void guna2Button6_MouseLeave(object sender, EventArgs e)
        {
            guna2Button6.FillColor = System.Drawing.Color.CadetBlue;
        }

        private void guna2Button2_MouseEnter(object sender, EventArgs e)
        {
            guna2Button2.FillColor = System.Drawing.Color.White;
        }

        private void guna2Button2_MouseLeave(object sender, EventArgs e)
        {
            guna2Button2.FillColor = System.Drawing.Color.CadetBlue;
        }

        private void guna2Button5_MouseEnter(object sender, EventArgs e)
        {
            guna2Button5.FillColor = System.Drawing.Color.White;

        }

        private void guna2Button5_MouseLeave(object sender, EventArgs e)
        {
            guna2Button5.FillColor = System.Drawing.Color.CadetBlue;
        }

        private void guna2Button7_MouseEnter(object sender, EventArgs e)
        {
            guna2Button7.FillColor = System.Drawing.Color.White;

        }

        private void guna2Button7_MouseLeave(object sender, EventArgs e)
        {
            guna2Button7.FillColor = System.Drawing.Color.CadetBlue;
        }

        private void guna2Button8_MouseEnter(object sender, EventArgs e)
        {
            guna2Button8.FillColor = System.Drawing.Color.White;
        }

        private void guna2Button8_MouseLeave(object sender, EventArgs e)
        {
            guna2Button8.FillColor = System.Drawing.Color.CadetBlue;
        }

        private void guna2Button9_MouseEnter(object sender, EventArgs e)
        {
            guna2Button9.FillColor = System.Drawing.Color.White;

        }

        private void guna2Button9_MouseLeave(object sender, EventArgs e)
        {
            guna2Button9.FillColor = System.Drawing.Color.CadetBlue;
        }

        private void guna2Button10_MouseEnter(object sender, EventArgs e)
        {
            guna2Button10.FillColor = System.Drawing.Color.White;
        }

        private void guna2Button10_MouseLeave(object sender, EventArgs e)
        {
            guna2Button10.FillColor = System.Drawing.Color.CadetBlue;
        }
    }
}



    


