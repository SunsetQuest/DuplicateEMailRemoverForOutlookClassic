// Released under the MIT License. Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
// The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

// DISCLAIMER
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
// THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THIS CODE LIES SOLELY WITH THE USER.

// Created by Ryan Scott White on Nov. 2023, Updated 12/20/2024 (note: no AI used for coding as of 12/20/2024, may change with the future drafts)


using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace DuplicateEMailRemover
{
    public partial class MainForm : Form
    {
        // Enabling will allow emails to be acted on(like delete or moved).
        private readonly bool ALLOW_MOVES_DELETES = true;

        private int authorizedTabIndex = 0;
        private int totalFoldersToSearch = 0;
        private int totalItemsToSearch = 0;
        private readonly string logToPath = $"{Path.GetTempPath()}\\DuplicateEMailRemoverOutput.txt";
        private CancellationTokenSource? tokenSource;
        private static readonly Outlook.Application outlook = new();

        // If this is specified, then it will save each email to the specified folder.
        private string? filePathToSaveEmails;

        // Folders we should skip - these folders should not really be scanned for duplicate emails.
        private static readonly SortedSet<string> FoldersToSkip = [
            "Calendar",
            "Contacts",
            "Conflicts",
            "Conversation Action Settings",
            "Conversation History",
            "Drafts",
            "ExternalContacts",
            "Files",
            "Journal",
            "Local Failures",
            "Lync Contacts",
            "MeContact",
            "News Feed",
            "Notes",
            "Outbox",
            "PersonMetadata",
            "Quick Step Settings",
            "RSS Feeds",
            "Server Failures",
            "Suggested Contacts",
            "Sync Issues",
            "Tasks",
            "Yammer Root",
            "Your feeds",
            ];

        public MainForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //tabControl1.Appearance = TabAppearance.FlatButtons;
            //tabControl1.ItemSize = new Size(0, 1);
            //tabControl1.SizeMode = TabSizeMode.Fixed;
            radioButtonLogOnly.Text = $"Log only: {logToPath}";

            if (!ALLOW_MOVES_DELETES)
            {
                Text += " (DEBUG MODE)";
            }
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            _ = MessageBox.Show("Avoid opening or closing Outlook while running this tool.");

            SuspendLayout();
            Outlook.Folders mailboxes = outlook.GetNamespace("MAPI").Folders;
            int folderCount = 0;
            //treeView1.BeginUpdate();
            foreach (Outlook.Folder f in mailboxes)
            {
                try
                {
                    int mailboxNameLen = f.FolderPath.Length + 1;
                    FolderTreeNode node = new(f);
                    node.Expand();
                    visitTopLevelFolders(f, node, mailboxNameLen, ref folderCount);
                    _ = treeView1.Nodes.Add(node);
                }
                catch (Exception ee)
                {
                    _ = MessageBox.Show($"""
                            There was an error with the mailbox {f.Name}-{f.AddressBookName}-{f.FolderPath}.
                            The error message was: {ee.Message}");
                            {ee.InnerException}
                            """);
                }

            }
            //treeView1.EndUpdate();
            ResumeLayout();

            // Make sure some folders were found.
            if (folderCount == 0)
            {
                _ = MessageBox.Show("No folders found. Please make sure you have office installed.");
                System.Windows.Forms.Application.Exit();
            }
            //MessageBox.Show("Form1_END");
        }

        private void visitTopLevelFolders(Outlook.Folder folder, TreeNode node, int mailboxNameLen, ref int folderCount)
        {
            foreach (Outlook.Folder f in folder.Folders)
            {
                string folderPath = f.FolderPath[mailboxNameLen..];
                if (FoldersToSkip.Contains(folderPath))
                {
                    continue;
                }

                FolderTreeNode childFolderNode = new(f);
                folderCount++;
                _ = node.Nodes.Add(childFolderNode);

                try
                {
                    visitSubFolder(f, childFolderNode, ref folderCount);
                }
                catch (Exception error)
                {
                    _ = MessageBox.Show($"""
                            Errors can sometimes happen if Outlook is opened/closed while 
                            $"this app is open.
                            There was an error with the mailbox {f.Name}-{f.AddressBookName}-{f.FolderPath}
                            The error message was: {error.Message}");
                            {error.InnerException}
                            """);
                    //Application.Exit();
                }
            }
        }


        private void visitSubFolder(Outlook.Folder folder, TreeNode node, ref int folderCount)
        {
            foreach (Outlook.Folder f in folder.Folders)
            {
                FolderTreeNode childFolderNode = new(f);
                folderCount++;
                _ = node.Nodes.Add(childFolderNode);
                visitSubFolder(f, childFolderNode, ref folderCount);
            }
        }

        private void btnTreeViewSelectNextPanel_Click(object sender, EventArgs e)
        {
            totalItemsToSearch = 0;
            totalFoldersToSearch = 0;

            try
            {
                foreach (TreeNode nd in treeView1.Nodes)
                {
                    CallRecursiveOnTreeNodeToFillOrderByListBox(nd);
                }
            }
            catch (Exception error)
            {
                _ = MessageBox.Show($"Errors can sometimes happen if Outlook is opened/closed while " +
                    $"this app is open.\n\n Details: {error.Message})");
                Application.Exit();

            }

            if (totalFoldersToSearch == 0)
            {
                _ = MessageBox.Show("There should be at least one folder selected.");
                return;
            }

            // If there is only one item selected then we don't need to do any sorting.
            else if (totalFoldersToSearch == 1)
            {
                authorizedTabIndex = 2;
                tabControl1.SelectedIndex = 2;
            }

            // there was at last two folders selected so we should proceed to the folder-ordering tab
            else
            {

                // Generally we want anything under Deleted Items last
                for (int i = listBoxOfFoldersToProcess.Items.Count - 1; i >= 0; i--)
                {
                    FolderTreeNode node = (FolderTreeNode)listBoxOfFoldersToProcess.Items[i];
                    if (Regex.IsMatch(node.Text, @"\\\\[^\\]+?\\Deleted Items(\\|$)"))
                    {
                        _ = listBoxOfFoldersToProcess.Items.Add(node);
                        listBoxOfFoldersToProcess.Items.RemoveAt(i);
                    }
                }

                authorizedTabIndex = 1;
                tabControl1.SelectedIndex = 1; // = (tabControl1.SelectedIndex + 1 < tabControl1.TabCount) ?   tabControl1.SelectedIndex + 1 : tabControl1.SelectedIndex;
            }

            txtItemsTotalCount.Text = totalItemsToSearch.ToString();
            txtFoldersTotalCount.Text = totalFoldersToSearch.ToString();
        }

        private void CallRecursiveOnTreeNodeToFillOrderByListBox(TreeNode nd)
        {
            if (nd.Checked)
            {
                _ = listBoxOfFoldersToProcess.Items.Add(nd);
                totalItemsToSearch += (nd as FolderTreeNode)?.OutlookFolder.Items.Count ?? 0;
                totalFoldersToSearch++;
            }

            _ = cboDestinationFolder.Items.Add(nd);

            foreach (TreeNode n in nd.Nodes)
            {
                CallRecursiveOnTreeNodeToFillOrderByListBox(n);
            }
        }

        private void treeView1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                // Select the clicked node
                treeView1.SelectedNode = treeView1.GetNodeAt(e.X, e.Y);

                if (treeView1.SelectedNode != null)
                {
                    contextMenuStrip1.Show(treeView1, e.Location);
                }
            }
        }

        private void toolStripMenuItemCheckAllChildNodes_Click(object sender, EventArgs e)
        {
            CheckOrUncheckSubnodes(treeView1.SelectedNode, true);
        }

        private void toolStripMenuItemUnCheckAllChildNodes_Click(object sender, EventArgs e)
        {
            CheckOrUncheckSubnodes(treeView1.SelectedNode, false);
        }

        private void CheckOrUncheckSubnodes(TreeNode node, bool value)
        {
            node.Checked = value;
            foreach (TreeNode childNode in node.Nodes)
            {
                CheckOrUncheckSubnodes(childNode, value);
            }
        }

        private void listBox1_MouseDown(object sender, MouseEventArgs e)
        {
            if (listBoxOfFoldersToProcess.SelectedItem == null)
            {
                return;
            }

            _ = listBoxOfFoldersToProcess.DoDragDrop(listBoxOfFoldersToProcess.SelectedItem, DragDropEffects.Move);
        }

        private void listBox1_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void listBox1_DragDrop(object sender, DragEventArgs e)
        {
            Point point = listBoxOfFoldersToProcess.PointToClient(new Point(e.X, e.Y));
            int index = listBoxOfFoldersToProcess.IndexFromPoint(point);
            if (index < 0)
            {
                index = listBoxOfFoldersToProcess.Items.Count - 1;
            }

            object? data = e.Data?.GetData(typeof(FolderTreeNode)) ?? null;
            if (data != null)
            {
                listBoxOfFoldersToProcess.Items.Remove(data);
                listBoxOfFoldersToProcess.Items.Insert(index, data);
            }
        }

        private void btnNextToOptions_Click(object sender, EventArgs e)
        {
            authorizedTabIndex = 2;
            tabControl1.SelectedIndex = 2;
        }

        private void radioButtonActionUpdateGUI_CheckedChanged(object sender, EventArgs e)
        {
            cboDestinationFolder.Enabled = radioButtonMoveToFolder.Checked | radioButtonCopyToFolder.Checked;
        }

        private void btnMoveToBeginProcessing_Click(object sender, EventArgs e)
        {
            // Lets make sure the destination folder is valid.
            if ((radioButtonMoveToFolder.Checked | radioButtonCopyToFolder.Checked)
                && cboDestinationFolder.SelectedIndex == -1)
            {
                _ = MessageBox.Show("Please select a destination folder.");
                return;
            }

            if (!checkBoxMatchOnBody.Checked && !checkBoxMatchOnHTMLBody.Checked)
            {
                DialogResult result = MessageBox.Show("Neither the \"Email Body\" or \"Body(HTLM)\" were selected.\n" +
                    "This can cause unintended non-duplicate messages to be moved/deleted.\n\n" +
                    "Are you sure you would like to continue? ", "Confirm", MessageBoxButtons.YesNo);

                if (result != DialogResult.Yes)
                {
                    return;
                }
            }

            if (radioButtonOpenEachInOutlook.Checked)
            {
                _ = MessageBox.Show("Please open Outlook and leave it open. Closing it will cause errors.");
            }


            authorizedTabIndex = 3;
            tabControl1.SelectedIndex = 3;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // only needed if user somehow clicks back to a previous tab
            if (tabControl1.SelectedIndex == 0)
            {
                listBoxOfFoldersToProcess.Items.Clear();
            }
        }
        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
            if (ALLOW_MOVES_DELETES)
            {
                if (tabControl1.SelectedIndex != authorizedTabIndex)
                {
                    e.Cancel = true;
                }
            }
        }

        private void checkBoxMatchOn_CheckedChanged(object sender, EventArgs e)
        {

            int checkCount =
                  (checkBoxMatchOnSentOn.Checked ? 1 : 0)
                + (checkBoxMatchOnReceivedTime.Checked ? 1 : 0)
                + (checkBoxMatchOnLastModTime.Checked ? 1 : 0)
                + (checkBoxMatchOnSenderEmail.Checked ? 1 : 0)
                + (checkBoxMatchOnTo.Checked ? 1 : 0)
                + (checkBoxMatchOnCC.Checked ? 1 : 0)
                + (checkBoxMatchOnBCC.Checked ? 1 : 0)
                + (checkBoxMatchOnSubject.Checked ? 1 : 0)
                + (checkBoxMatchOnBody.Checked ? 1 : 0)
                + (checkBoxMatchOnHTMLBody.Checked ? 1 : 0)
                + (checkBoxMatchOnAttachment.Checked ? 1 : 0);

            btnMoveToBeginProcessing.Enabled = checkCount > 1;
        }

        // some help from here https://stackoverflow.com/a/18033198/2352507 for async and progress bars
        private async void btnStart_Click(object sender, EventArgs e)
        {
            btnStart.Enabled = false;
            btnStop.Enabled = true;
            btnStart.Text = "Running";
            try
            {
                Progress<int> folderProcessedCount = new(s => txtFoldersProcessed.Text = s.ToString());
                Progress<int> mailProcessedCount = new(s => txtEmailsProcessed.Text = s.ToString());
                Progress<int> duplicatesFound = new(s => txtDuplicatesFound.Text = s.ToString());
                Progress<int> nonMailItemsSkipped = new(s => txtNonMailItemsSkipped.Text = s.ToString());
                //outlook.GetNamespace("MAPI").Folders
                tokenSource = new CancellationTokenSource();
                CancellationToken token = tokenSource.Token;
                await Task.Run(() => ScanFolders(folderProcessedCount, mailProcessedCount, duplicatesFound, nonMailItemsSkipped, token));
            }
            catch (Exception exception)
            {
                _ = MessageBox.Show(exception.Message);
            }

            // Lets show the result
            _ = new Process
            {
                StartInfo = new ProcessStartInfo(logToPath)
                {
                    UseShellExecute = true
                }
            }.Start();

            btnStart.Text = "Start";
            btnStart.Enabled = true;
            btnStop.Enabled = false;
        }

        public void ScanFolders(IProgress<int> foldersProgress,
            IProgress<int> fileProgress,
            IProgress<int> duplicateCount,
            IProgress<int> nonMailItemCount, CancellationToken token)
        {
            // Set myItems = myContacts.Restrict("[LastModificationTime] > '01/1/2003'") 
            // As a safeguard, lets check to make sure we do not have any duplicate folders in listBoxOfFoldersToProcess
            int folderCount = listBoxOfFoldersToProcess.Items.Count;
            string[] folders = new string[folderCount];
            for (int i = 0; i < folderCount; i++)
            {
                folders[i] = ((FolderTreeNode)listBoxOfFoldersToProcess.Items[i]).Path;
                //need to test
                List<int> ints = [1];

                List<int> ints2 = ints.ToArray().Select(i => i + 1).ToList();

                //folders[i] = listBoxOfFoldersToProcess.Items[i].ddTostring();
            }

            for (int i = 0; i < folderCount; i++)
            {
                foldersProgress.Report(i);

                //string curFolder = listBoxOfFoldersToProcess.Items[i].ToString();
                for (int j = i + 1; j < folderCount; j++)
                {
                    if (folders[i] == folders[j])
                    {
                        _ = MessageBox.Show("Duplicate folders found in listBoxOfFoldersToProcess. Please remove duplicate folders.");
                        return;
                    }
                }
            }

            Outlook.MAPIFolder? moveTo = null;
            if (radioButtonCopyToFolder.Checked || radioButtonCopyToFolder.Checked)
            {
                FolderTreeNode? item = (FolderTreeNode?)cboDestinationFolder.SelectedItem;
                if (item != null)
                {
                    moveTo = item!.OutlookFolder;
                }
            }

            // Lets start removing the duplicates. We will iterate through the listBoxOfFoldersToProcess
            // and remove the duplicates as we find them.
            Dictionary<string, string> hs = [];

            bool warningNotificationEnabled = true;
            int totalItemCount = 0;
            int totalDuplicateCount = 0;
            int totalNonMailItemCount = 0;

            using StreamWriter sw = new(logToPath, false, Encoding.UTF8, 65536);
            StringBuilder sb = new(1024);
            int count = listBoxOfFoldersToProcess.Items.Count;
            for (int i = 0; i < count; i++)
            {
                FolderTreeNode node = (FolderTreeNode)listBoxOfFoldersToProcess.Items[i];
                Outlook.Items items = node.OutlookFolder.Items;
                foldersProgress.Report(i);

                foreach (object? item in items)
                {
                    fileProgress.Report(++totalItemCount);

                    if (item is not Outlook.MailItem)
                    {
                        nonMailItemCount.Report(++totalNonMailItemCount);
                        continue;
                    }

                    Outlook.MailItem mailItem = (Outlook.MailItem)item;

                    // Build the attributes to check for duplicates. (Subject/Received date/Sender/etc.)
                    _ = sb.Clear();
                    if (checkBoxMatchOnSentOn.Checked)
                    {
                        _ = sb.Append(mailItem.SentOn.ToString());
                    }

                    if (checkBoxMatchOnReceivedTime.Checked)
                    {
                        _ = sb.Append(mailItem.ReceivedTime.ToString());
                    }

                    if (checkBoxMatchOnLastModTime.Checked)
                    {
                        _ = sb.Append(mailItem.LastModificationTime.ToString());
                    }

                    if (checkBoxMatchOnSenderEmail.Checked)
                    {
                        _ = sb.Append(mailItem.SenderEmailAddress);
                    }

                    if (checkBoxMatchOnTo.Checked)
                    {
                        _ = sb.Append(mailItem.To);
                    }

                    if (checkBoxMatchOnCC.Checked)
                    {
                        _ = sb.Append(mailItem.CC);
                    }

                    if (checkBoxMatchOnBCC.Checked)
                    {
                        _ = sb.Append(mailItem.BCC);
                    }

                    if (checkBoxMatchOnSubject.Checked)
                    {
                        _ = sb.Append(mailItem.Subject);
                    }

                    if (checkBoxMatchOnBody.Checked)
                    {
                        _ = sb.Append(mailItem.Body);
                    }

                    if (checkBoxMatchOnHTMLBody.Checked)
                    {
                        _ = sb.Append(mailItem.HTMLBody);
                    }

                    if (checkBoxMatchOnAttachment.Checked)
                    {
                        foreach (Outlook.Attachment attachment in mailItem.Attachments)
                        {
                            _ = sb.Append(attachment.FileName);
                        }
                    }

                    byte[] buffer = Encoding.UTF8.GetBytes(sb.ToString());
                    byte[] digest = MD5.HashData(buffer);

                    //StringBuilder result = new StringBuilder(digest.Length * 2);
                    _ = sb.Clear();

                    for (int j = 0; j < digest.Length; j++)
                    {
                        _ = sb.Append(digest[j].ToString("X2"));
                    }
                    string hash = sb.ToString();

                    if (token.IsCancellationRequested)
                    {
                        sw.Write("  User canceled via Stop button.\r\n");
                        return;
                    }

                    if (hs.ContainsKey(hash))
                    {
                        duplicateCount.Report(++totalDuplicateCount);

                        // We have a duplicate, lets first log it.
                        _ = sb.Clear();
                        _ = sb.AppendLine($"===================================\nDuplicate Found:" +
                            $"\n  Sent On: {mailItem.SentOn} " +
                            $"\n  Sender:  {mailItem.SenderEmailAddress} " +
                            $"\n  Subject: {mailItem.Subject ?? ""}");
                        string OrigFolder = hs[hash];
                        if (OrigFolder != node.Path)
                        {
                            _ = sb.AppendLine($"\n  Original was here:\n    {OrigFolder}");
                        }

                        _ = sb.AppendLine($"\n  Duplicate found here:\n    {node.Path}");
                        _ = sb.AppendLine($"\n  HASH: {hash}");  //for debug
                        string msg = sb.ToString();
                        sw.WriteLine(msg);

                        // Lets check with the user if we should take an action on this duplicate email.
                        if (warningNotificationEnabled)
                        {
                            DialogResult dialog = MessageBox.Show($"Do you want to continue to be notified?\r\n Yes = Continue to notify me on each message\n" +
                                $" No = Stop notifying and process ALL remaining emails\n\n{msg}", "Duplicate found", MessageBoxButtons.YesNoCancel);
                            if (dialog == DialogResult.No)
                            {
                                warningNotificationEnabled = false;
                            }
                            else if (dialog == DialogResult.Cancel)
                            {
                                sw.WriteLine("  User canceled via Cancel button.\r\n");
                                return;
                            }
                        }

                        // If the user requested to save the email to a folder, do that first before moving or deleting. 
                        if (filePathToSaveEmails != null)
                        {
                            string sentOnDate = mailItem.SentOn.ToString();
                            string filenameFriendlyDate = string.Join("_", sentOnDate.Split(Path.GetInvalidFileNameChars()));
                            string path2 = $"{filePathToSaveEmails}{node.Path}\\{filenameFriendlyDate} {hash}.msg";
                            _ = Directory.CreateDirectory(filePathToSaveEmails + node.Path);
                            mailItem.SaveAs(path2);
                            sw.WriteLine($" Email saved to {path2}");
                        }

                        if (radioButtonCopyToFolder.Checked)
                        {
                            if (ALLOW_MOVES_DELETES)
                            {
                                try
                                {
                                    Outlook.MailItem copiedItem = mailItem.Copy();
                                    copiedItem.Move(moveTo);
                                }
                                catch (Exception error)
                                {
                                    LogErrorMessage(sw, error, "Email copied to mailbox folder");
                                }
                            }
                            else
                            {
                                sw.Write("  Simulated: Email copied to mailbox folder\r\n");
                            }
                        }
                        else if (radioButtonMoveToFolder.Checked)
                        {
                            if (ALLOW_MOVES_DELETES)
                            {
                                try
                                {
                                    mailItem.Move(moveTo);
                                }
                                catch (Exception error)
                                {
                                    LogErrorMessage(sw, error, "Email moved to mailbox folder");
                                }
                            }
                            else
                            {
                                sw.Write("  Simulated: Email moved to mailbox folder\r\n");
                            }
                        }
                        else if (radioButtonOpenEachInOutlook.Checked)
                        {
                            try
                            {
                                mailItem.Display(true);
                            }
                            catch (Exception error)
                            {
                                LogErrorMessage(sw, error, "Email opened in Outlook");
                            }
                        }
                        else if (radioButtonDeleteEmails.Checked)
                        {
                            sw.Write("  Email Deleted\r\n");

                            if (ALLOW_MOVES_DELETES)
                            {
                                try
                                {
                                    mailItem.Delete();
                                }
                                catch (Exception error)
                                {
                                    LogErrorMessage(sw, error, "Email Deleted");
                                }
                            }
                            else
                            {
                                sw.Write("  Simulated: Email Deleted\r\n");
                            }
                        }

                    }
                    else
                    {
                        hs.Add(sb.ToString(), node.Path);
                    }
                }
            }
        }

        private static void LogErrorMessage(StreamWriter sw, Exception error, string action)
        {
            string message = $"  Failed {action} " +
                $"Message: {error.Message}\r\n" +
                $"Source: {error.Source}\r\n" +
                $"InnerException: {error.InnerException}\r\n";
            sw.Write(message);
            _ = MessageBox.Show(message);
        }

        private void checkBoxSaveEachDupToFilePath_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBoxSaveEachDupToFilePath.Checked)
            {
                using FolderBrowserDialog fbd = new();
                DialogResult result = fbd.ShowDialog(this);
                if (result == DialogResult.OK && Directory.Exists(fbd.SelectedPath))
                {
                    filePathToSaveEmails = fbd.SelectedPath;
                    txtSaveEachDupToFilePath.Text = fbd.SelectedPath;
                }
                else
                {
                    filePathToSaveEmails = null;
                    txtSaveEachDupToFilePath.Text = "";
                }
            }
            else
            {
                filePathToSaveEmails = null;
                txtSaveEachDupToFilePath.Text = "";
            }

            checkBoxSaveEachDupToFilePath.Text = "Save each duplicate to " + (filePathToSaveEmails ?? "a file.");
            txtSaveEachDupToFilePath.Visible = checkBoxSaveEachDupToFilePath.Checked;
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            tokenSource?.Cancel();
        }

        private void btnI_Agree_Click(object sender, EventArgs e)
        {
            authorizedTabIndex = 4;
            tabControl1.SelectedIndex = 4;
        }
    }


    internal sealed class FolderTreeNode : TreeNode
    {
        public override string ToString()
        {
            return $"{Text.TrimStart('\\')} ({ItemCount})";
        }

        public Outlook.Folder OutlookFolder { get; }

        public Outlook.Items OutlookItems { get; }

        public int ItemCount { get; }

        public string Path
        {
            get => Text;
            set => Text = value;
        }

        public FolderTreeNode(Outlook.Folder outlookFolder)
        {
            OutlookFolder = outlookFolder;
            Text = outlookFolder.FolderPath;
            OutlookItems = OutlookFolder.Items;
            ItemCount = OutlookFolder.Items.Count;
        }
    }
}