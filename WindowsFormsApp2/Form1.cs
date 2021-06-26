using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using HtmlAgilityPack;
using System.IO;
using GemBox.Document;
using System.Threading;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        List<string> urls = new List<string>();
        int index = 1;
        List<string> urls_save = new List<string>();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string templateUrl = @"https://www.google.co.uk/search?q={0}&tbm=isch&site=imghp";

            //check that we have a term to search for.
            if (string.IsNullOrEmpty(textBox1.Text))
            {
                MessageBox.Show("Please supply a search term"); return;
            }
            else
            {
                using (WebClient wc = new WebClient())
                {
                    //lets pretend we are IE8 on Vista.
                    wc.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)");
                    string result = wc.DownloadString(String.Format(templateUrl, new object[] {textBox1.Text+" "+richTextBox2.Text}));
                    //Console.WriteLine("Our result {0}", result);

                    //we have valid markup, this will change from time to time as google updates.
                    //Console.WriteLine("before if");

                    if (result.Contains("t0fcAb"))
                    {
                        //Console.WriteLine("inside if");

                        HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                        doc.LoadHtml(result);

                        var htmlNodes = doc.DocumentNode.SelectNodes("//img");
                        foreach (var node in htmlNodes)
                        {
                            HtmlAttribute src = node.Attributes[@"src"];
                            //Console.WriteLine("src is {0}", src.Value);
                            urls.Add(src.Value);
                        }
                        urls.RemoveAt(0);
                        //lets create a linq query to find all the img's stored in that images_table class.
                        /*
                         * Essentially we get search for the table called images_table, and then get all images that have a valid src containing images?
                         * which is the string used by google
                        eg  https://encrypted-tbn3.gstatic.com/images?q=tbn:ANd9GcQmGxh15UUyzV_HGuGZXUxxnnc6LuqLMgHR9ssUu1uRwy0Oab9OeK1wCw
                         */

                        //Console.WriteLine("urls 1 = {0}", urls[1]);

                        byte[] downloadedData = wc.DownloadData(urls[0]);

                        if (downloadedData != null)
                        {
                            //Console.WriteLine("downloadeddata is not null");

                            //store the downloaded data in to a stream
                            System.IO.MemoryStream ms = new System.IO.MemoryStream(downloadedData, 0, downloadedData.Length);

                            //write to that stream the byte array
                            ms.Write(downloadedData, 0, downloadedData.Length);

                            //load an image from that stream.
                            pictureBox1.Image = Image.FromStream(ms);
                        }

                    }

                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            urls_save.Add(urls[index]);

            PopupWindowDoc popup = new PopupWindowDoc();

            popup.ShowDialog();

            popup.Dispose();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            using (WebClient wc = new WebClient())
            {
                //lets pretend we are IE8 on Vista.
                wc.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)");

                index++;

                if (index >= urls.Count)
                {
                    index = urls.Count-1;
                }

                if (index < 0)
                {
                    index = 1;
                }

                Console.WriteLine(index);

                byte[] downloadedData = wc.DownloadData(urls[index]);

                if (downloadedData != null)
                {
                    //Console.WriteLine("downloadeddata is not null");

                    //store the downloaded data in to a stream
                    System.IO.MemoryStream ms = new System.IO.MemoryStream(downloadedData, 0, downloadedData.Length);

                    //write to that stream the byte array
                    ms.Write(downloadedData, 0, downloadedData.Length);

                    //load an image from that stream.
                    pictureBox1.Image = Image.FromStream(ms);
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Console.WriteLine("Start Document Creation - {0}", urls_save[0]);

            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            var document = new DocumentModel();

            var section = new Section(document);
            document.Sections.Add(section);

            document.Sections.Add(
            new Section(document,
                new Paragraph(document,
                    new Run(document, "Title:"+textBox1.Text+" Content:"+richTextBox2.Text+" Content Bold:"+richTextBox2.Text))));

            var paragraph = new Paragraph(document);
            section.Blocks.Add(paragraph);

            // Create and add an inline picture with GIF image.
            for (int i = 0; i < urls_save.Count; i++)
            {
                Picture picture1 = new Picture(document, urls_save[i], 50, 50, LengthUnit.Pixel);
                paragraph.Inlines.Add(picture1);
            }

            // Create save options
            var saveOptions = new DocxSaveOptions();
            saveOptions.ProgressChanged += (eventSender, args) =>
            {
                Console.WriteLine($"Progress changed - {args.ProgressPercentage}%");
            };

            document.Save("Pictures.docx", saveOptions);

            PopupWindowDoc popup = new PopupWindowDoc();

            popup.ShowDialog();

            popup.Dispose();
        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            using (WebClient wc = new WebClient())
            {
                //lets pretend we are IE8 on Vista.
                wc.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 8.0; Windows NT 6.0; Trident/4.0)");
                
                index--;

                if (index < 0)
                {
                    index = 0;
                }
                else if (index > 19)
                {
                    index = 18;
                }

                Console.WriteLine(index);

                byte[] downloadedData = wc.DownloadData(urls[index]);

                if (downloadedData != null)
                {
                    //Console.WriteLine("downloadeddata is not null");

                    //store the downloaded data in to a stream
                    System.IO.MemoryStream ms = new System.IO.MemoryStream(downloadedData, 0, downloadedData.Length);

                    //write to that stream the byte array
                    ms.Write(downloadedData, 0, downloadedData.Length);

                    //load an image from that stream.
                    pictureBox1.Image = Image.FromStream(ms);
                }
            }
        }
    }
}
