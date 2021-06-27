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
        public List<string> urls = new List<string>();
        public int index = 1;
        public int count_search = 0;
        public List<string> urls_save = new List<string>();
        public List<string> titlebox = new List<string>();
        public List<string> contentbox = new List<string>();
        public List<string> contentboldbox = new List<string>();

        public Form1()
        {
            InitializeComponent();
        }

        private void Search_Button_Click(object sender, EventArgs e)
        {
            urls.Clear();
            index = 1;
            count_search++;

            titlebox.Add(textBox1.Text);
            contentbox.Add(richTextBox1.Text);
            contentboldbox.Add(richTextBox2.Text);

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

                            //Make Buttons Visible for Image and document manipulation
                            button3.Visible = true;
                            button4.Visible = true;
                            button5.Visible = true;
                            button2.Visible = true;
                        }

                    }

                }
            }
        }

        private void AddtoDocument_Button_Click(object sender, EventArgs e)
        {
            urls_save.Add(urls[index]);

            PopupWindowDoc popup = new PopupWindowDoc();

            popup.ShowDialog();

            popup.Dispose();
        }

        private void NextImage_Button_Click(object sender, EventArgs e)
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

        private void CreateDocument_Button_Click(object sender, EventArgs e)
        {
            Console.WriteLine("Start Document Creation - {0}", urls_save[0]);

            ComponentInfo.SetLicense("FREE-LIMITED-KEY");

            var document = new DocumentModel();

            var section = new Section(document);
            document.Sections.Add(section);

            var paragraph = new Paragraph(document);
            section.Blocks.Add(paragraph);

            // Create and add an inline picture with GIF image.
            for (int i = 0; i < urls_save.Count; i++)
            {
                Picture picture1 = new Picture(document, urls_save[i], 50, 50, LengthUnit.Pixel);
                paragraph.Inlines.Add(picture1);
            }

            for (int i = 0; i < count_search; i++)
            {
                Run run = new Run(document, "Title:" + titlebox[i] + " Content:" + contentbox[i] + " Content Bold:" + contentboldbox[i]+" ");
                paragraph.Inlines.Add(run);
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

        private void PreviousImage_Button_Click(object sender, EventArgs e)
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
