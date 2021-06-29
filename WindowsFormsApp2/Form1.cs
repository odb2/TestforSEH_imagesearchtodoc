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
//using GemBox.Document;
using GemBox.Presentation;
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
            string[] words = richTextBox2.Text.Split(',');
            foreach (string word in words)
            {
                int startindex = 0;
                while (startindex < richTextBox1.TextLength)
                {
                    int wordstartIndex = richTextBox1.Find(word, startindex, RichTextBoxFinds.None);
                    if (wordstartIndex != -1)
                    {
                        richTextBox1.SelectionStart = wordstartIndex;
                        richTextBox1.SelectionLength = word.Length;
                        richTextBox1.SelectionFont = new Font(richTextBox1.Font, FontStyle.Bold);
                    }
                    else
                        break;
                    startindex += wordstartIndex + word.Length;
                }
            }


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
            var presentation = new PresentationDocument();

            // Create new presentation slide.
            var slide = presentation.Slides.AddNew(SlideLayoutType.Custom);

            // Create and add an inline picture with GIF image.

            int[] arr = new int[] {2,4,6,8,10,12,14,16,18,20,22,24,26,28,30};
            for (int i = 0; i < urls_save.Count; i++)
            {
                // Create first picture from resource data.
                Picture picture = null;
                picture = slide.Content.AddPicture(urls_save[i], arr[i], 2,2,2, LengthUnit.Centimeter);

            }

            var slide2 = presentation.Slides.AddNew(SlideLayoutType.Custom);

            var textBox = slide2.Content.AddTextBox(ShapeGeometryType.Rectangle, 2, 2, 5, 4, LengthUnit.Centimeter);

            for (int i = 0; i < count_search; i++)
            {
                var paragraph = textBox.AddParagraph();

                paragraph.AddRun("Title:" + titlebox[i] + " Content:" + contentbox[i] + " Content Bold:" + contentboldbox[i] + " ");
            }

            presentation.Save("powerpoint.pptx");

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

        private void resetbutton_click(object sender, EventArgs e)
        {
            richTextBox1.SelectionStart = 0;
            richTextBox1.SelectAll();
            richTextBox1.SelectionFont = new Font(richTextBox1.Font, FontStyle.Regular);

        }
    }
}
