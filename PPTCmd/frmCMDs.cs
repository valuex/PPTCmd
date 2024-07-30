using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Net;
using System.Windows.Forms;
using System.Xml.Linq;

namespace PPTCmd
{

    public partial class frmCMDs : Form
    {
        String XmlFilePath;
        public frmCMDs()
        {
            InitializeComponent();
        }

        private void frmCMDs_Load(object sender, EventArgs e)
        {
            LoadXMLIntoListView(sender,e);
            listView1.FullRowSelect = true;
            listView1.MultiSelect = false;
        }

        private void LoadXMLIntoListView(object sender, EventArgs e)
        {
            // load cmd list into list view

            listView1.View = View.Details;
            listView1.GridLines = true;
            //listView1.Sorting = SortOrder.Descending;
            listView1.Columns.Add("ID", 0);
            listView1.Columns.Add("Group", 60);
            listView1.Columns.Add("CName", 120);
            listView1.Columns.Add("Command", 250);
            listView1.Columns.Add("Times", 0);
            listView1.Columns.Add("Type", 0);
            listView1.Items.Clear();

            XmlFilePath = GetExternalXmlPath();

            var doc = XDocument.Load(XmlFilePath);
            var output = from x in doc.Descendants("cmd")
                         orderby (int)x.Attribute("RTimes") descending
                         select new ListViewItem(new[]
                         {
                             x.Attribute("Id").Value,
                             x.Attribute("GName").Value,
                             x.Attribute("CName").Value,
                             x.Attribute("Cmdlet").Value,
                             x.Attribute("RTimes").Value,
                             x.Attribute("CmdType").Value
                         });

            listView1.Items.AddRange(output.Take(10).ToArray());

        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string Kw=textBox1.Text;
            listView1.Items.Clear();

            var doc = XDocument.Load(XmlFilePath);
            bool HasSpace = Kw.Contains(" ");
            int FirstSpacePos = Kw.IndexOf(" ");
            int LastSpacePos = Kw.LastIndexOf(" ");
            bool HasOnlyOneSpace = HasSpace && (LastSpacePos == FirstSpacePos);
            string trimKw=Kw.Trim();
            IEnumerable<ListViewItem> output;
            if (HasOnlyOneSpace)
            {
                if (Kw.EndsWith(" "))
                {
                    // last char is the first Space
                    output = from x in doc.Descendants("cmd")
                                  let ItemGName = NPinyin.Pinyin.GetInitials(x.Attribute("GName").Value).ToLower()
                                  where ItemGName.Contains(trimKw)
                                  orderby (int)x.Attribute("RTimes") descending
                                  select new ListViewItem(new[]
                                  {
                             x.Attribute("Id").Value,
                             x.Attribute("GName").Value,
                             x.Attribute("CName").Value,
                             x.Attribute("Cmdlet").Value,
                             x.Attribute("RTimes").Value,
                             x.Attribute("CmdType").Value
                         });
                    listView1.Items.AddRange(output.Take(10).ToArray());

                }
                else
                {
                    string Kw_p1 = Kw.Substring(0, FirstSpacePos);
                    string Kw_p2 = Kw.Substring(FirstSpacePos + 1).Trim();
                    output = from x in doc.Descendants("cmd")
                                  let ItemGName = NPinyin.Pinyin.GetInitials(x.Attribute("GName").Value).ToLower()
                                  let ItemCName = NPinyin.Pinyin.GetInitials(x.Attribute("CName").Value).ToLower()
                                  where ItemGName.Contains(Kw_p1) && ItemCName.Contains(Kw_p2)
                                  orderby (int)x.Attribute("RTimes") descending
                                  select new ListViewItem(new[]
                                  {
                             x.Attribute("Id").Value,
                             x.Attribute("GName").Value,
                             x.Attribute("CName").Value,
                             x.Attribute("Cmdlet").Value,
                             x.Attribute("RTimes").Value,
                             x.Attribute("CmdType").Value
                         });
                    listView1.Items.AddRange(output.Take(10).ToArray());

                }
            }
            else if (trimKw.Length==0)
            {
                // input is blank
                output = from x in doc.Descendants("cmd")
                             orderby (int)x.Attribute("RTimes") descending
                             select new ListViewItem(new[]
                             {
                             x.Attribute("Id").Value,
                             x.Attribute("GName").Value,
                             x.Attribute("CName").Value,
                             x.Attribute("Cmdlet").Value,
                             x.Attribute("RTimes").Value,
                             x.Attribute("CmdType").Value
                         });

                listView1.Items.AddRange(output.Take(10).ToArray());
            }
            else
            {
                // no space or more than one space, index in CName
                output = from x in doc.Descendants("cmd")
                         let ItemValue = NPinyin.Pinyin.GetInitials(x.Attribute("CName").Value).ToLower()
                         where ItemValue.Contains(trimKw)
                         orderby (int)x.Attribute("RTimes") descending
                         select new ListViewItem(new[]
                         {
                             x.Attribute("Id").Value,
                             x.Attribute("GName").Value,
                             x.Attribute("CName").Value,
                             x.Attribute("Cmdlet").Value,
                             x.Attribute("RTimes").Value,
                             x.Attribute("CmdType").Value
                         });
                listView1.Items.AddRange(output.Take(10).ToArray());
            }
        }



        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            MessageBox.Show(listView1.SelectedItems[0].SubItems[2].Text);
        }

        private void listView1_MouseDown(object sender, MouseEventArgs e)
        {
            ListViewHitTestInfo info = listView1.HitTest(e.X, e.Y);
            ListViewItem item = info.Item;
            item.Selected = true;
            if (item != null)
            {
                string CmdId = item.SubItems[0].Text;
                string msoCMD = item.SubItems[3].Text;
                string CMDType = item.SubItems[5].Text;
                this.Close();
                RunCMD(CMDType, msoCMD);
                RunTimeIncreaseByOne(CmdId);
            }
        }
        private void RunTimeIncreaseByOne(string ThisId)
        {
            var doc = XDocument.Load(XmlFilePath);
            var nodesToUpdate = from x in doc.Descendants("cmd")
                                where (string)x.Attribute("Id") == ThisId
                                select x;
            foreach (XElement el in nodesToUpdate)
            {
                var CurTime = Int32.Parse(el.Attribute("RTimes").Value);
                el.Attribute("RTimes").Value = (CurTime + 1).ToString();
                doc.Save(XmlFilePath);
                break;
            }
        }
        private void RunCMD(string cmdType, string cmdName)
        {
            if (cmdType == "sys")
            {
                try
                {
                    Globals.ThisAddIn.Application.CommandBars.ExecuteMso(cmdName);
                }
                catch
                {
                    string eMsg = "Wrong Command Name in xml file Or \n No correct content is select for the command";
                    MessageBox.Show(eMsg);
                }
            }
            else if (cmdType == "usr")
            {
                try
                {
                    Globals.ThisAddIn.RunMacro(Globals.ThisAddIn.Application, new object[] { cmdName });
                }
                catch
                {
                    string eMsg = "Wrong Macro Name in xml file Or \n No correct content is select for the command";
                    MessageBox.Show(eMsg);
                }
            }
        }
        private void frmCMDs_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
            else if (e.KeyCode == Keys.Enter)
            {
                // Press Enter to excute the first matched one directly
                ListViewItem item = listView1.Items[0];
                string CmdId = item.SubItems[0].Text;
                string msoCMD = item.SubItems[3].Text;
                string CMDType = item.SubItems[5].Text;
                this.Close();
                RunCMD(CMDType, msoCMD);
                RunTimeIncreaseByOne(CmdId);
            }
            else if (e.KeyCode == Keys.Down)
            {
                // Press ArrowDown to focus on the listview box's first item
                if (textBox1.Focused)
                {
                    listView1.Focus();
                    listView1.Items[0].Selected = true;
                }
                else if(listView1.Focused && listView1.Items[listView1.Items.Count-1].Selected == true)
                { 
                    // when the last item in the listview is focused, press Down, go to textbox
                    textBox1.Focus();
                }
                else
                {
                    // focus on next item in the listView
                    for (int i = 0; i < 9; i++)
                    {
                        if (listView1.Items[i].Selected == true)
                        {
                            listView1.Items[i+1].Selected = true;
                            break;
                        }
                    }
                }
            }
            else if (e.KeyCode == Keys.Up)
            {
                // Press ArrowUP to focus on the texbox if the first item in the listview is focused
                if (listView1.Focused && listView1.Items[0].Selected == true)
                {
                    textBox1.Focus();
                }
            }
        }

        private string GetExternalXmlPath()
        {
            String AppDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            String XmlFilePath = AppDataDir + "\\PPTCmd\\CMDList.xml";
            return XmlFilePath;
        }
    }

}
