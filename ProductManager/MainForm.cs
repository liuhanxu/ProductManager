using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Threading;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using HtmlAgilityPack;
using System.Net;
using System.IO;

using Microsoft.Office.Interop;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace ProductManager
{
    public partial class MainForm : Form
    {
        private ExcelHelper excelHelper = new ExcelHelper();
        private Worksheet catagorySheet;
        private Worksheet listSheet;
        private Worksheet unavailableSheet;
        private string fileName;
        private bool loading = true;

        public MainForm()
        {
            InitializeComponent();
        }

        private void loadExcelFileBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Title = "Select excel file";
            dialog.Filter = "Excel Files(*.xls;*.xlsx;*.xlsm)|*.xls;*.xlsx;*.xlsm";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                fileName = dialog.FileName;
                Console.WriteLine(fileName);
                try
                {
                    excelHelper.Open(fileName);
                    catagorySheet = excelHelper.GetSheet("kk");
                    Console.WriteLine(catagorySheet.UsedRange.Rows.Count);
                    if (catagorySheet == null)
                        tipsText.Text = "Not found Sheet: kk";
                    else {
                        tipsText.Text = "Load file success";
                        search();
                    }
                        
                }
                catch (Exception exception)
                {
                    MessageBox.Show(exception.Message);
                }

            }
        }

        private void search()
        {
            loadExcelBtn.Enabled = false;
            listSheet = excelHelper.GetSheet("list");
            if (listSheet == null)
            {
                MessageBox.Show(" Sheet list not found");
                return;
            }
            listSheet.Cells.Clear();
            tipsText.Text = "parsing urls...";
            int totalRow = catagorySheet.UsedRange.Rows.Count;

            totalRow = 2; //test 1 rows ----------------------------------------test, delete this line after test
            int rowIndex = 2;
            for (int i = 2; i <= totalRow; i++)
            {
                string u = catagorySheet.Cells[i, 1].Text;
                for (int j = 1; j <= 5; j++)
                {
                    tipsText.Text = "load url:" + u + "?pg=" + j;
                    try
                    {
                        Thread.Sleep(500);
                        Console.WriteLine(u + "?pg=" + j);
                        //Byte[] responseData = wc.DownloadData(u + "?pg=" + j);
                        //string srcString = gz(responseData); 
                        loading = true;
                        webBrowser.Navigate(u + "?pg=" + j);
                        while (loading) {
                            System.Windows.Forms.Application.DoEvents();
                        }
                        string srcString = webBrowser.DocumentText;
                        HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                        doc.LoadHtml(srcString);
                        HtmlNodeCollection allItems = doc.DocumentNode.SelectNodes("//div[@class='zg_itemImmersion']/div[@class='zg_itemWrapper']/div[@class='a-section a-spacing-none p13n-asin']/a[@class='a-link-normal']");
                        HtmlNodeCollection allItems2 = doc.DocumentNode.SelectNodes("//li[@class='zg-item-immersion']/span[@class='a-list-item']/div[@class='a-section a-spacing-none aok-relative']/span[@class='aok-inline-block zg-item']/a[@class='a-link-normal']");
                        if (allItems == null&&allItems2 == null)
                        {
                            Console.WriteLine("allItems.Count==0" );
                            break;
                        }
                        if (allItems != null)
                        {
                            Console.WriteLine("allItems.Count==" + allItems.Count);
                            foreach (HtmlNode node in allItems)
                            {
                                string tempUrl = node.GetAttributeValue("href", "");
                                if (tempUrl != "" && !tempUrl.StartsWith("http"))
                                {
                                    listSheet.Cells[rowIndex, 1] = "https://www.amazon.com" + tempUrl;
                                    rowIndex++;
                                }
                            }
                        }
                        if (allItems2 != null) {
                            Console.WriteLine("allItems2.Count==" + allItems2.Count);
                            foreach (HtmlNode node in allItems2)
                            {
                                string tempUrl = node.GetAttributeValue("href", "");
                                if (tempUrl != "" && !tempUrl.StartsWith("http"))
                                {
                                    listSheet.Cells[rowIndex, 1] = "https://www.amazon.com" + tempUrl;
                                    rowIndex++;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        break;
                    }
                }
            }
            listSheet.Cells.EntireColumn.AutoFit();
            bool f=excelHelper.Save();
            Console.WriteLine(f);
            findUnavailable();
        }

        private void findUnavailable()
        {
            unavailableSheet = excelHelper.GetSheet("Unavailable");
            if (unavailableSheet == null)
            {
                MessageBox.Show(" Sheet Unavailable not found");
                return;
            }
            unavailableSheet.Cells.Clear();            
            int totalRow = listSheet.UsedRange.Rows.Count; 
            Console.WriteLine("total row=" + totalRow); 
            totalRow = 5;//----------------------------------------test, delete this line after test
            int rowIndex = 2;
            for (int i = 2; i <= totalRow; i++)
            {
                string url = listSheet.Cells[i, 1].Text;
                Console.WriteLine("get url=" + url);
                tipsText.Text = "get product info from url:" + url;
                try
                {
                    //Byte[] responseData = wc.DownloadData(url);
                    //string srcString = gz(responseData); 
                    loading = true;
                    webBrowser.Navigate(url);
                    while (loading)
                    {
                        System.Windows.Forms.Application.DoEvents();
                    }
                    string srcString = webBrowser.DocumentText;
                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                    doc.LoadHtml(srcString);
                    string title = doc.DocumentNode.SelectSingleNode("//div[@id='titleSection']/h1[@id='title']/span[@id='productTitle']").InnerText;
                    string price = "";
                    if (doc.DocumentNode.SelectSingleNode("//body").InnerText.IndexOf("With Deal:") >= 0)
                    {
                        price = doc.DocumentNode.SelectSingleNode("//span[@id='priceblock_dealprice']").InnerText;
                    }
                    else
                    {
                        price = doc.DocumentNode.SelectSingleNode("//span[@id='priceblock_ourprice']").InnerText;
                    }
                    string availability = "In Stock.";
                    if ( doc.DocumentNode.SelectSingleNode("//div[@id='availability']/span").InnerText!= "")
                        availability = doc.DocumentNode.SelectSingleNode("//div[@id='availability']/span").InnerText;

                    Console.WriteLine("title="+title.Trim());
                    Console.WriteLine(price.Trim());
                    Console.WriteLine(availability.Trim());

                    listSheet.Cells[i, 2] = title.Trim();
                    listSheet.Cells[i, 3] = price.Trim();
                    listSheet.Cells[i, 4] = availability.Trim();

                    if (availability.IndexOf("In Stock") < 0)
                    {
                        unavailableSheet.Cells[rowIndex, 1] = url;
                        unavailableSheet.Cells[rowIndex, 2] = title.Trim();
                    }
                    Thread.Sleep(500);
                }
                catch (Exception ex) {
                    Console.WriteLine(ex.Message);
                }
            }
            listSheet.Cells.EntireColumn.AutoFit();
            catagorySheet.Cells.EntireColumn.AutoFit();
            if (excelHelper.Save())
            {
                excelHelper.Close();
                tipsText.Text = "all task complet";
                MessageBox.Show("task is complete, file saved at:"+fileName);
            }
            else
            {
                MessageBox.Show("save error");
            }

            loadExcelBtn.Enabled = true;
        }

        private void webBrowser_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            string url = e.Url.ToString();
            Console.WriteLine("completed:" + url);
            if (!(url.StartsWith("http://") || url.StartsWith("https://")))
            {
                // in AJAX
            }

            if (e.Url.AbsolutePath == webBrowser.Url.AbsolutePath)
            {
                // REAL DOCUMENT COMPLETE
                Console.WriteLine("loading=false");
                loading = false;
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                listSheet = null;
                catagorySheet = null;
                unavailableSheet = null;
                excelHelper.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }


        public string gz(byte[] cbytes)
        {
            using (MemoryStream dms = new MemoryStream())
            {
                using (MemoryStream cms = new MemoryStream(cbytes))
                {
                    using (System.IO.Compression.GZipStream gzip = new System.IO.Compression.GZipStream(cms, System.IO.Compression.CompressionMode.Decompress))
                    {
                        byte[] bytes = new byte[10240];
                        int len = 0;
                        while ((len = gzip.Read(bytes, 0, bytes.Length)) > 0)
                        {
                            dms.Write(bytes, 0, len);
                        }
                    }
                }
                return (Encoding.UTF8.GetString(dms.ToArray()));
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            webBrowser.DocumentCompleted += new WebBrowserDocumentCompletedEventHandler(webBrowser_DocumentCompleted);
        }

    }
}
