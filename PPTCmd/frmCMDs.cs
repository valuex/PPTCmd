using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Linq.Expressions;
using System.Net;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Runtime.CompilerServices.RuntimeHelpers;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

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
            try
            {
                ListViewItem item = listView1.Items[0];
                item.Selected = true;
            }
            catch { }

        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string Kw=textBox1.Text;
            bool IsOpenBracketPressed=Kw.Contains("[");
            bool IsCloseBracketPressed=Kw.Contains("]");
            Kw = Kw.Replace("[", "");
            Kw = Kw.Replace("]", "");
            textBox1.Text=Kw;
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
            try
            {
                ListViewItem item = listView1.Items[0];
                item.Selected = true;
            }
            catch { }
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
                // Press Enter to excute the selected item or first item directly
                ListViewItem item;
                try
                {
                    item = listView1.SelectedItems[0];
                }
                catch
                {
                    item = listView1.Items[0];
                }                
                string CmdId = item.SubItems[0].Text;
                string msoCMD = item.SubItems[3].Text;
                string CMDType = item.SubItems[5].Text;
                this.Close();
                RunCMD(CMDType, msoCMD);
                RunTimeIncreaseByOne(CmdId);

            }
            else if (e.KeyCode == Keys.Down || e.KeyCode == Keys.OemOpenBrackets )
            {
                SelectBelowRow();
            }
            else if (e.KeyCode == Keys.Up || e.KeyCode == Keys.OemCloseBrackets)
            {
                SelectAboveRow();
                // Send right arrow, otherwise the caret will move left 
                SendKeys.Send("{Right}"); 
            }

        }

        private void SelectAboveRow()
        {
            int PreRowIndex;
            int LastRowIndex = listView1.Items.Count - 1;
            try
            {
                ListViewItem item = listView1.SelectedItems[0];
                PreRowIndex = item.Index - 1;
                if (item.Index == 0) { PreRowIndex = LastRowIndex; }
            }
            catch
            { PreRowIndex = LastRowIndex; }
            // focus on next row
            try
            {
                ListViewItem item = listView1.Items[PreRowIndex];
                item.Selected = true;
            }
            catch
            { }
        }
        private void SelectBelowRow()
        {
            int NexRowIndex;
            int LastRowIndex = listView1.Items.Count - 1;

            try
            {
                ListViewItem item = listView1.SelectedItems[0];
                NexRowIndex = item.Index + 1;
                if (item.Index == LastRowIndex) { NexRowIndex = 0; }
            }
            catch
            { NexRowIndex = 0; }
            // focus on next row
            try
            {
                ListViewItem item = listView1.Items[NexRowIndex];
                item.Selected = true;
            }
            catch
            { }
        }
        private string GetExternalXmlPath()
        {
            String AppDataDir = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            String XmlFilePath = AppDataDir + "\\PPTCmd\\CMDList.xml";
            return XmlFilePath;
        }
        #region set listView selected item color
            bool lvEditMode = false;
            Color listViewSelectionColor = Color.DodgerBlue;
            private void listView1_DrawItem(object sender, DrawListViewItemEventArgs e)
            {
                var lView = sender as System.Windows.Forms.ListView;

                if (lvEditMode || lView.View == View.Details) return;
                TextFormatFlags flags = GetTextAlignment(lView, 0);
                Color itemColor = e.Item.ForeColor;

                if (e.Item.Selected)
                {
                    using (var bkBrush = new SolidBrush(listViewSelectionColor))
                    {
                        e.Graphics.FillRectangle(bkBrush, e.Bounds);
                    }
                    itemColor = e.Item.BackColor;
                }
                else
                {
                    e.DrawBackground();
                }

                TextRenderer.DrawText(e.Graphics, e.Item.Text, e.Item.Font, e.Bounds, itemColor, flags);

                if (lView.View == View.Tile && e.Item.SubItems.Count > 1)
                {
                    var subItem = e.Item.SubItems[1];
                    flags = GetTextAlignment(lView, 1);
                    TextRenderer.DrawText(e.Graphics, subItem.Text, subItem.Font, e.Bounds, SystemColors.GrayText, flags);
                }
            }

        
            private void listView1_DrawSubItem(object sender, DrawListViewSubItemEventArgs e)
            {
                var lView = sender as System.Windows.Forms.ListView;
                TextFormatFlags flags = GetTextAlignment(lView, e.ColumnIndex);
                Color itemColor = e.Item.ForeColor;

                if (e.Item.Selected && !lvEditMode)
                {
                    if (e.ColumnIndex == 0 || lView.FullRowSelect)
                    {
                        using (var bkgrBrush = new SolidBrush(listViewSelectionColor))
                        {
                            e.Graphics.FillRectangle(bkgrBrush, e.Bounds);
                        }
                        itemColor = e.Item.BackColor;
                    }
                }
                else
                {
                    e.DrawBackground();
                }
                TextRenderer.DrawText(e.Graphics, e.SubItem.Text, e.SubItem.Font, e.Bounds, itemColor, flags);
            }

            private void listView1_DrawColumnHeader(object sender, DrawListViewColumnHeaderEventArgs e)
            {
                e.DrawDefault = true;
            }

            private TextFormatFlags GetTextAlignment(System.Windows.Forms.ListView lstView, int colIndex)
            {
                TextFormatFlags flags = (lstView.View == View.Tile)
                    ? (colIndex == 0) ? TextFormatFlags.Default : TextFormatFlags.Bottom
                    : TextFormatFlags.VerticalCenter;

                if (lstView.View == View.Details) flags |= TextFormatFlags.LeftAndRightPadding;

                if (lstView.Columns[colIndex].TextAlign != HorizontalAlignment.Left)
                {
                    flags |= (TextFormatFlags)((int)lstView.Columns[colIndex].TextAlign ^ 3);
                }
                return flags;
            }
        #endregion

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            // surpress the input of '[' or ']'
            if (e.KeyChar == '[' || e.KeyChar == ']')
            {
                e.Handled = true;
            }
        }
    }

}
