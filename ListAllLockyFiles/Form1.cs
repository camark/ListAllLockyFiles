using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;


namespace ListAllLockyFiles
{
    public partial class Form1 : Form
    {
        private ArrayList lstNetComputers;
        public Form1()
        {
            InitializeComponent();
        }

        delegate void SetTextCallback(string text);

        private void SetText(string text)
        {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (this.lblComputer.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(SetText);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                this.lblComputer.Text = text;
            }
        }

        delegate void RefreshLabelCallback();

        private void RefreshLabel()
        {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (this.lblComputer.InvokeRequired)
            {
                RefreshLabelCallback d = new RefreshLabelCallback(RefreshLabel);
                this.Invoke(d, new object[] { });
            }
            else
            {
                this.lblComputer.Refresh();
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            _worker = new BackgroundWorker { WorkerSupportsCancellation = true };

            _worker.DoWork += new DoWorkEventHandler((state, args) =>
            {
                ListAllServers();

                foreach (string sComputer in lstNetComputers)
                {
                    SetText(sComputer);
                    RefreshLabel();
                    if (_worker.CancellationPending)
                        break;
                    ArrayList lstShares = ListShares((string)sComputer);

                    ListAllLocky(lstShares);
                }

                button1.Enabled = true;
                button2.Enabled = false;
                _worker.CancelAsync();
                lblComputer.Text = "Not started";

            });

            _worker.RunWorkerAsync();
            button1.Enabled = false;
            button2.Enabled = true;


        }
        private delegate void AddItemCallback(object o);

        private void AddItem(object oItem)
        {
            if (this.listView1.InvokeRequired)
            {
                AddItemCallback d = new AddItemCallback(AddItem);
                this.Invoke(d, new object[] { oItem });
            }
            else
            {
                if (oItem is string)
                    listView1.Items.Add((string)oItem);
            }
        }

        private delegate void EnsureVisibleCallback(object o);

        private void EnsureVisible(object o)
        {
            if (this.listView1.InvokeRequired)
            {
                EnsureVisibleCallback d = new EnsureVisibleCallback(AddItem);
                this.Invoke(d, new object[] { o });
            }
            else
            {
                listView1.Items[(int)o].EnsureVisible();
            }
        }

        delegate void RefreshListCallback();

        private void RefreshList()
        {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (this.lblComputer.InvokeRequired)
            {
                RefreshListCallback d = new RefreshListCallback(RefreshList);
                this.Invoke(d, new object[] { });
            }
            else
            {
                this.listView1.Refresh();
            }
        }
        private void ListAllLocky(ArrayList lstShares)
        {
            foreach (string oFolderPath in lstShares)
            {
                if (_worker.CancellationPending)
                    break;
                String[] allfiles;
                try
                {


                    allfiles = System.IO.Directory.GetFiles(oFolderPath, "*._locky",
                        System.IO.SearchOption.AllDirectories);
                }
                catch (Exception)
                {

                    return;
                }
                foreach (string sFile in allfiles)
                {
                    if (_worker.CancellationPending)
                        break;
                    AddItem(sFile);

                    EnsureVisible(listView1.Items.Count - 1);
                    RefreshList();
                }
            }
        }

        private void ListAllServers()
        {
            if (_worker.CancellationPending)
                return;
            NetworkBrowser oBrowser = new NetworkBrowser();
            lstNetComputers = oBrowser.getNetworkComputers();
        }

        private ArrayList ListShares(string server)
        {

            ArrayList sSharesList = new ArrayList();

            if (_worker.CancellationPending)
                return sSharesList;
            //// Enumerate shares on local computer
            //Console.WriteLine("\nShares on local computer:");
            ShareCollection shi = ShareCollection.LocalShares;

            // Enumerate shares on a remote computer

            if (server != null && server.Trim().Length > 0)
            {
                Console.WriteLine("\nShares on {0}:", server);
                shi = ShareCollection.GetShares(server);
                if (shi != null)
                {
                    foreach (Share si in shi)
                    {
                        if (_worker.CancellationPending)
                            break;
                        if ((si.Root == null) || si.Root.FullName.Contains("$"))
                            continue;

                        sSharesList.Add(si.Root.FullName);
                    }
                }
                else
                    Console.WriteLine("Unable to enumerate the shares on {0}.\n"
                        + "Make sure the machine exists, and that you have permission to access it.",
                        server);

                Console.WriteLine();
            }

            return sSharesList;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            button1.Enabled = true;
            button2.Enabled = false;
            _worker.CancelAsync();
            lblComputer.Text = "Not started";
        }

        public BackgroundWorker _worker { get; set; }

        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            String sPath = listView1.SelectedItems[0].Text;

            string sDir = Path.GetDirectoryName(sPath);

            if (sDir != null) Process.Start(sDir);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application xla = new Microsoft.Office.Interop.Excel.Application();

            xla.Visible = true;

            Microsoft.Office.Interop.Excel.Workbook wb = xla.Workbooks.Add(Microsoft.Office.Interop.Excel.XlSheetType.xlWorksheet);

            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)xla.ActiveSheet;

            int i = 1;

            int j = 1;

            foreach (ListViewItem comp in listView1.Items)
            {

                ws.Cells[i, j] = comp.Text.ToString();

                //MessageBox.Show(comp.Text.ToString());

                foreach (ListViewItem.ListViewSubItem drv in comp.SubItems)
                {

                    ws.Cells[i, j] = drv.Text.ToString();

                    j++;

                }

                j = 1;

                i++;

            }
        }
    }
}

