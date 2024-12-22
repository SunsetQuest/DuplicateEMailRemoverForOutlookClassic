namespace DuplicateEMailRemover
{
    partial class MainForm
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            tabControl1 = new TabControl();
            tabPage1 = new TabPage();
            btnTreeViewSelectNextPanel = new Button();
            label1 = new Label();
            treeView1 = new TreeView();
            tabPage2 = new TabPage();
            label7 = new Label();
            label6 = new Label();
            button2 = new Button();
            listBoxOfFoldersToProcess = new ListBox();
            tabPage3 = new TabPage();
            groupBoxWhatToMatchOn = new GroupBox();
            checkBoxMatchOnBCC = new CheckBox();
            checkBoxMatchOnCC = new CheckBox();
            checkBoxMatchOnTo = new CheckBox();
            checkBoxMatchOnHTMLBody = new CheckBox();
            checkBoxMatchOnBody = new CheckBox();
            checkBoxMatchOnSubject = new CheckBox();
            checkBoxMatchOnSenderEmail = new CheckBox();
            checkBoxMatchOnAttachment = new CheckBox();
            checkBoxMatchOnLastModTime = new CheckBox();
            checkBoxMatchOnReceivedTime = new CheckBox();
            checkBoxMatchOnSentOn = new CheckBox();
            btnMoveToBeginProcessing = new Button();
            groupBoxWhatToDoWithEmails = new GroupBox();
            radioButtonLogOnly = new RadioButton();
            txtSaveEachDupToFilePath = new TextBox();
            checkBoxSaveEachDupToFilePath = new CheckBox();
            radioButtonOpenEachInOutlook = new RadioButton();
            radioButtonCopyToFolder = new RadioButton();
            cboDestinationFolder = new ComboBox();
            radioButtonDeleteEmails = new RadioButton();
            radioButtonMoveToFolder = new RadioButton();
            tabPage4 = new TabPage();
            button1 = new Button();
            richTextBox1 = new RichTextBox();
            tabPage6 = new TabPage();
            txtNonMailItemsSkipped = new TextBox();
            txtDuplicatesFound = new TextBox();
            label5 = new Label();
            btnStop = new Button();
            label3 = new Label();
            label2 = new Label();
            txtItemsTotalCount = new TextBox();
            txtFoldersTotalCount = new TextBox();
            txtEmailsProcessed = new TextBox();
            txtFoldersProcessed = new TextBox();
            label4 = new Label();
            lblFolders = new Label();
            lblEmails = new Label();
            btnStart = new Button();
            contextMenuStrip1 = new ContextMenuStrip(components);
            toolStripMenuItemCheckAllChildNodes = new ToolStripMenuItem();
            toolStripMenuItemUnCheckAllChildNodes = new ToolStripMenuItem();
            toolTip1 = new ToolTip(components);
            tabControl1.SuspendLayout();
            tabPage1.SuspendLayout();
            tabPage2.SuspendLayout();
            tabPage3.SuspendLayout();
            groupBoxWhatToMatchOn.SuspendLayout();
            groupBoxWhatToDoWithEmails.SuspendLayout();
            tabPage4.SuspendLayout();
            tabPage6.SuspendLayout();
            contextMenuStrip1.SuspendLayout();
            SuspendLayout();
            // 
            // tabControl1
            // 
            tabControl1.Controls.Add(tabPage1);
            tabControl1.Controls.Add(tabPage2);
            tabControl1.Controls.Add(tabPage3);
            tabControl1.Controls.Add(tabPage4);
            tabControl1.Controls.Add(tabPage6);
            tabControl1.Dock = DockStyle.Fill;
            tabControl1.Location = new Point(0, 0);
            tabControl1.Name = "tabControl1";
            tabControl1.SelectedIndex = 0;
            tabControl1.Size = new Size(648, 549);
            tabControl1.TabIndex = 0;
            tabControl1.SelectedIndexChanged += tabControl1_SelectedIndexChanged;
            tabControl1.Selecting += tabControl1_Selecting;
            // 
            // tabPage1
            // 
            tabPage1.Controls.Add(btnTreeViewSelectNextPanel);
            tabPage1.Controls.Add(label1);
            tabPage1.Controls.Add(treeView1);
            tabPage1.Location = new Point(4, 29);
            tabPage1.Name = "tabPage1";
            tabPage1.Padding = new Padding(3);
            tabPage1.Size = new Size(640, 516);
            tabPage1.TabIndex = 0;
            tabPage1.Text = "Select Folders";
            tabPage1.UseVisualStyleBackColor = true;
            // 
            // btnTreeViewSelectNextPanel
            // 
            btnTreeViewSelectNextPanel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnTreeViewSelectNextPanel.Location = new Point(482, 458);
            btnTreeViewSelectNextPanel.Name = "btnTreeViewSelectNextPanel";
            btnTreeViewSelectNextPanel.Size = new Size(150, 50);
            btnTreeViewSelectNextPanel.TabIndex = 2;
            btnTreeViewSelectNextPanel.Text = "Next";
            btnTreeViewSelectNextPanel.UseVisualStyleBackColor = true;
            btnTreeViewSelectNextPanel.Click += btnTreeViewSelectNextPanel_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            label1.Location = new Point(6, 9);
            label1.Name = "label1";
            label1.Size = new Size(507, 20);
            label1.TabIndex = 1;
            label1.Text = "Select folders to scan:    (Right-click to show menu to select subfolders)";
            // 
            // treeView1
            // 
            treeView1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            treeView1.CheckBoxes = true;
            treeView1.Location = new Point(6, 38);
            treeView1.Name = "treeView1";
            treeView1.Size = new Size(626, 414);
            treeView1.TabIndex = 0;
            treeView1.MouseClick += treeView1_MouseClick;
            // 
            // tabPage2
            // 
            tabPage2.Controls.Add(label7);
            tabPage2.Controls.Add(label6);
            tabPage2.Controls.Add(button2);
            tabPage2.Controls.Add(listBoxOfFoldersToProcess);
            tabPage2.Location = new Point(4, 29);
            tabPage2.Name = "tabPage2";
            tabPage2.Padding = new Padding(3);
            tabPage2.Size = new Size(640, 516);
            tabPage2.TabIndex = 1;
            tabPage2.Text = "Folder Order";
            tabPage2.UseVisualStyleBackColor = true;
            // 
            // label7
            // 
            label7.AutoSize = true;
            label7.Font = new Font("Segoe UI", 14F);
            label7.Location = new Point(6, 11);
            label7.Name = "label7";
            label7.Size = new Size(269, 32);
            label7.TabIndex = 5;
            label7.Text = "Folder Processing Order";
            // 
            // label6
            // 
            label6.AutoSize = true;
            label6.Font = new Font("Segoe UI", 9F);
            label6.Location = new Point(6, 43);
            label6.Name = "label6";
            label6.Size = new Size(434, 80);
            label6.TabIndex = 4;
            label6.Text = resources.GetString("label6.Text");
            // 
            // button2
            // 
            button2.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            button2.Location = new Point(482, 458);
            button2.Name = "button2";
            button2.Size = new Size(150, 50);
            button2.TabIndex = 3;
            button2.Text = "Next";
            button2.UseVisualStyleBackColor = true;
            button2.Click += btnNextToOptions_Click;
            // 
            // listBoxOfFoldersToProcess
            // 
            listBoxOfFoldersToProcess.AllowDrop = true;
            listBoxOfFoldersToProcess.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            listBoxOfFoldersToProcess.FormattingEnabled = true;
            listBoxOfFoldersToProcess.Location = new Point(6, 126);
            listBoxOfFoldersToProcess.Name = "listBoxOfFoldersToProcess";
            listBoxOfFoldersToProcess.Size = new Size(626, 324);
            listBoxOfFoldersToProcess.TabIndex = 0;
            listBoxOfFoldersToProcess.DragDrop += listBox1_DragDrop;
            listBoxOfFoldersToProcess.DragOver += listBox1_DragOver;
            listBoxOfFoldersToProcess.MouseDown += listBox1_MouseDown;
            // 
            // tabPage3
            // 
            tabPage3.Controls.Add(groupBoxWhatToMatchOn);
            tabPage3.Controls.Add(btnMoveToBeginProcessing);
            tabPage3.Controls.Add(groupBoxWhatToDoWithEmails);
            tabPage3.Location = new Point(4, 29);
            tabPage3.Name = "tabPage3";
            tabPage3.Size = new Size(640, 516);
            tabPage3.TabIndex = 2;
            tabPage3.Text = "Action to Take";
            tabPage3.UseVisualStyleBackColor = true;
            // 
            // groupBoxWhatToMatchOn
            // 
            groupBoxWhatToMatchOn.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            groupBoxWhatToMatchOn.Controls.Add(checkBoxMatchOnBCC);
            groupBoxWhatToMatchOn.Controls.Add(checkBoxMatchOnCC);
            groupBoxWhatToMatchOn.Controls.Add(checkBoxMatchOnTo);
            groupBoxWhatToMatchOn.Controls.Add(checkBoxMatchOnHTMLBody);
            groupBoxWhatToMatchOn.Controls.Add(checkBoxMatchOnBody);
            groupBoxWhatToMatchOn.Controls.Add(checkBoxMatchOnSubject);
            groupBoxWhatToMatchOn.Controls.Add(checkBoxMatchOnSenderEmail);
            groupBoxWhatToMatchOn.Controls.Add(checkBoxMatchOnAttachment);
            groupBoxWhatToMatchOn.Controls.Add(checkBoxMatchOnLastModTime);
            groupBoxWhatToMatchOn.Controls.Add(checkBoxMatchOnReceivedTime);
            groupBoxWhatToMatchOn.Controls.Add(checkBoxMatchOnSentOn);
            groupBoxWhatToMatchOn.Location = new Point(3, 298);
            groupBoxWhatToMatchOn.Name = "groupBoxWhatToMatchOn";
            groupBoxWhatToMatchOn.Size = new Size(615, 154);
            groupBoxWhatToMatchOn.TabIndex = 3;
            groupBoxWhatToMatchOn.TabStop = false;
            groupBoxWhatToMatchOn.Text = "Match on what fields?";
            // 
            // checkBoxMatchOnBCC
            // 
            checkBoxMatchOnBCC.AutoSize = true;
            checkBoxMatchOnBCC.Location = new Point(228, 116);
            checkBoxMatchOnBCC.Name = "checkBoxMatchOnBCC";
            checkBoxMatchOnBCC.Size = new Size(145, 24);
            checkBoxMatchOnBCC.TabIndex = 7;
            checkBoxMatchOnBCC.Text = "BCC Addresses(s)";
            checkBoxMatchOnBCC.UseVisualStyleBackColor = true;
            checkBoxMatchOnBCC.CheckedChanged += checkBoxMatchOn_CheckedChanged;
            // 
            // checkBoxMatchOnCC
            // 
            checkBoxMatchOnCC.AutoSize = true;
            checkBoxMatchOnCC.Location = new Point(228, 86);
            checkBoxMatchOnCC.Name = "checkBoxMatchOnCC";
            checkBoxMatchOnCC.Size = new Size(122, 24);
            checkBoxMatchOnCC.TabIndex = 6;
            checkBoxMatchOnCC.Text = "CC Address(s)";
            checkBoxMatchOnCC.UseVisualStyleBackColor = true;
            checkBoxMatchOnCC.CheckedChanged += checkBoxMatchOn_CheckedChanged;
            // 
            // checkBoxMatchOnTo
            // 
            checkBoxMatchOnTo.AutoSize = true;
            checkBoxMatchOnTo.Checked = true;
            checkBoxMatchOnTo.CheckState = CheckState.Checked;
            checkBoxMatchOnTo.Location = new Point(228, 56);
            checkBoxMatchOnTo.Name = "checkBoxMatchOnTo";
            checkBoxMatchOnTo.Size = new Size(120, 24);
            checkBoxMatchOnTo.TabIndex = 5;
            checkBoxMatchOnTo.Text = "To Address(s)";
            checkBoxMatchOnTo.UseVisualStyleBackColor = true;
            checkBoxMatchOnTo.CheckedChanged += checkBoxMatchOn_CheckedChanged;
            // 
            // checkBoxMatchOnHTMLBody
            // 
            checkBoxMatchOnHTMLBody.AutoSize = true;
            checkBoxMatchOnHTMLBody.Location = new Point(421, 86);
            checkBoxMatchOnHTMLBody.Name = "checkBoxMatchOnHTMLBody";
            checkBoxMatchOnHTMLBody.Size = new Size(114, 24);
            checkBoxMatchOnHTMLBody.TabIndex = 4;
            checkBoxMatchOnHTMLBody.Text = "Body(HTML)";
            checkBoxMatchOnHTMLBody.UseVisualStyleBackColor = true;
            checkBoxMatchOnHTMLBody.CheckedChanged += checkBoxMatchOn_CheckedChanged;
            // 
            // checkBoxMatchOnBody
            // 
            checkBoxMatchOnBody.AutoSize = true;
            checkBoxMatchOnBody.Checked = true;
            checkBoxMatchOnBody.CheckState = CheckState.Checked;
            checkBoxMatchOnBody.Location = new Point(421, 56);
            checkBoxMatchOnBody.Name = "checkBoxMatchOnBody";
            checkBoxMatchOnBody.Size = new Size(106, 24);
            checkBoxMatchOnBody.TabIndex = 4;
            checkBoxMatchOnBody.Text = "Email Body";
            checkBoxMatchOnBody.UseVisualStyleBackColor = true;
            checkBoxMatchOnBody.CheckedChanged += checkBoxMatchOn_CheckedChanged;
            // 
            // checkBoxMatchOnSubject
            // 
            checkBoxMatchOnSubject.AutoSize = true;
            checkBoxMatchOnSubject.Checked = true;
            checkBoxMatchOnSubject.CheckState = CheckState.Checked;
            checkBoxMatchOnSubject.Location = new Point(421, 26);
            checkBoxMatchOnSubject.Name = "checkBoxMatchOnSubject";
            checkBoxMatchOnSubject.Size = new Size(80, 24);
            checkBoxMatchOnSubject.TabIndex = 4;
            checkBoxMatchOnSubject.Text = "Subject";
            checkBoxMatchOnSubject.UseVisualStyleBackColor = true;
            checkBoxMatchOnSubject.CheckedChanged += checkBoxMatchOn_CheckedChanged;
            // 
            // checkBoxMatchOnSenderEmail
            // 
            checkBoxMatchOnSenderEmail.AutoSize = true;
            checkBoxMatchOnSenderEmail.Checked = true;
            checkBoxMatchOnSenderEmail.CheckState = CheckState.Checked;
            checkBoxMatchOnSenderEmail.Location = new Point(228, 26);
            checkBoxMatchOnSenderEmail.Name = "checkBoxMatchOnSenderEmail";
            checkBoxMatchOnSenderEmail.Size = new Size(143, 24);
            checkBoxMatchOnSenderEmail.TabIndex = 4;
            checkBoxMatchOnSenderEmail.Text = "Sender's Address";
            checkBoxMatchOnSenderEmail.UseVisualStyleBackColor = true;
            checkBoxMatchOnSenderEmail.CheckedChanged += checkBoxMatchOn_CheckedChanged;
            // 
            // checkBoxMatchOnAttachment
            // 
            checkBoxMatchOnAttachment.AutoSize = true;
            checkBoxMatchOnAttachment.Location = new Point(15, 116);
            checkBoxMatchOnAttachment.Name = "checkBoxMatchOnAttachment";
            checkBoxMatchOnAttachment.Size = new Size(185, 24);
            checkBoxMatchOnAttachment.TabIndex = 3;
            checkBoxMatchOnAttachment.Text = "Attachment File Names";
            checkBoxMatchOnAttachment.UseVisualStyleBackColor = true;
            checkBoxMatchOnAttachment.CheckedChanged += checkBoxMatchOn_CheckedChanged;
            // 
            // checkBoxMatchOnLastModTime
            // 
            checkBoxMatchOnLastModTime.AutoSize = true;
            checkBoxMatchOnLastModTime.Location = new Point(15, 86);
            checkBoxMatchOnLastModTime.Name = "checkBoxMatchOnLastModTime";
            checkBoxMatchOnLastModTime.Size = new Size(191, 24);
            checkBoxMatchOnLastModTime.TabIndex = 2;
            checkBoxMatchOnLastModTime.Text = "Date/Time Modification";
            checkBoxMatchOnLastModTime.UseVisualStyleBackColor = true;
            checkBoxMatchOnLastModTime.CheckedChanged += checkBoxMatchOn_CheckedChanged;
            // 
            // checkBoxMatchOnReceivedTime
            // 
            checkBoxMatchOnReceivedTime.AutoSize = true;
            checkBoxMatchOnReceivedTime.Location = new Point(15, 56);
            checkBoxMatchOnReceivedTime.Name = "checkBoxMatchOnReceivedTime";
            checkBoxMatchOnReceivedTime.Size = new Size(166, 24);
            checkBoxMatchOnReceivedTime.TabIndex = 1;
            checkBoxMatchOnReceivedTime.Text = "Date/Time Received";
            checkBoxMatchOnReceivedTime.UseVisualStyleBackColor = true;
            checkBoxMatchOnReceivedTime.CheckedChanged += checkBoxMatchOn_CheckedChanged;
            // 
            // checkBoxMatchOnSentOn
            // 
            checkBoxMatchOnSentOn.AutoSize = true;
            checkBoxMatchOnSentOn.Checked = true;
            checkBoxMatchOnSentOn.CheckState = CheckState.Checked;
            checkBoxMatchOnSentOn.Location = new Point(15, 26);
            checkBoxMatchOnSentOn.Name = "checkBoxMatchOnSentOn";
            checkBoxMatchOnSentOn.Size = new Size(135, 24);
            checkBoxMatchOnSentOn.TabIndex = 0;
            checkBoxMatchOnSentOn.Text = "Date/Time Sent";
            checkBoxMatchOnSentOn.UseVisualStyleBackColor = true;
            checkBoxMatchOnSentOn.CheckedChanged += checkBoxMatchOn_CheckedChanged;
            // 
            // btnMoveToBeginProcessing
            // 
            btnMoveToBeginProcessing.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnMoveToBeginProcessing.Location = new Point(482, 458);
            btnMoveToBeginProcessing.Name = "btnMoveToBeginProcessing";
            btnMoveToBeginProcessing.Size = new Size(150, 50);
            btnMoveToBeginProcessing.TabIndex = 2;
            btnMoveToBeginProcessing.Text = "Next";
            btnMoveToBeginProcessing.UseVisualStyleBackColor = true;
            btnMoveToBeginProcessing.Click += btnMoveToBeginProcessing_Click;
            // 
            // groupBoxWhatToDoWithEmails
            // 
            groupBoxWhatToDoWithEmails.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            groupBoxWhatToDoWithEmails.Controls.Add(radioButtonLogOnly);
            groupBoxWhatToDoWithEmails.Controls.Add(txtSaveEachDupToFilePath);
            groupBoxWhatToDoWithEmails.Controls.Add(checkBoxSaveEachDupToFilePath);
            groupBoxWhatToDoWithEmails.Controls.Add(radioButtonOpenEachInOutlook);
            groupBoxWhatToDoWithEmails.Controls.Add(radioButtonCopyToFolder);
            groupBoxWhatToDoWithEmails.Controls.Add(cboDestinationFolder);
            groupBoxWhatToDoWithEmails.Controls.Add(radioButtonDeleteEmails);
            groupBoxWhatToDoWithEmails.Controls.Add(radioButtonMoveToFolder);
            groupBoxWhatToDoWithEmails.Location = new Point(3, 3);
            groupBoxWhatToDoWithEmails.Name = "groupBoxWhatToDoWithEmails";
            groupBoxWhatToDoWithEmails.Size = new Size(615, 285);
            groupBoxWhatToDoWithEmails.TabIndex = 1;
            groupBoxWhatToDoWithEmails.TabStop = false;
            groupBoxWhatToDoWithEmails.Text = "Action to take";
            groupBoxWhatToDoWithEmails.UseCompatibleTextRendering = true;
            // 
            // radioButtonLogOnly
            // 
            radioButtonLogOnly.AutoSize = true;
            radioButtonLogOnly.Location = new Point(13, 155);
            radioButtonLogOnly.Name = "radioButtonLogOnly";
            radioButtonLogOnly.Size = new Size(87, 24);
            radioButtonLogOnly.TabIndex = 8;
            radioButtonLogOnly.Text = "Log only";
            radioButtonLogOnly.UseVisualStyleBackColor = true;
            // 
            // txtSaveEachDupToFilePath
            // 
            txtSaveEachDupToFilePath.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtSaveEachDupToFilePath.Location = new Point(36, 238);
            txtSaveEachDupToFilePath.Name = "txtSaveEachDupToFilePath";
            txtSaveEachDupToFilePath.ReadOnly = true;
            txtSaveEachDupToFilePath.Size = new Size(565, 27);
            txtSaveEachDupToFilePath.TabIndex = 7;
            // 
            // checkBoxSaveEachDupToFilePath
            // 
            checkBoxSaveEachDupToFilePath.AutoSize = true;
            checkBoxSaveEachDupToFilePath.Location = new Point(16, 207);
            checkBoxSaveEachDupToFilePath.Name = "checkBoxSaveEachDupToFilePath";
            checkBoxSaveEachDupToFilePath.Size = new Size(361, 24);
            checkBoxSaveEachDupToFilePath.TabIndex = 5;
            checkBoxSaveEachDupToFilePath.Text = "Also, save each duplicate to a windows file folder.";
            toolTip1.SetToolTip(checkBoxSaveEachDupToFilePath, "This will save the emails to a file folder.  The directory structure is preserved when saving to files.");
            checkBoxSaveEachDupToFilePath.UseVisualStyleBackColor = true;
            checkBoxSaveEachDupToFilePath.CheckedChanged += checkBoxSaveEachDupToFilePath_CheckedChanged;
            // 
            // radioButtonOpenEachInOutlook
            // 
            radioButtonOpenEachInOutlook.AutoSize = true;
            radioButtonOpenEachInOutlook.Location = new Point(13, 95);
            radioButtonOpenEachInOutlook.Name = "radioButtonOpenEachInOutlook";
            radioButtonOpenEachInOutlook.Size = new Size(528, 24);
            radioButtonOpenEachInOutlook.TabIndex = 4;
            radioButtonOpenEachInOutlook.Text = "Open each email in Outlook - action can be taken on each email in Outlook";
            radioButtonOpenEachInOutlook.UseVisualStyleBackColor = true;
            radioButtonOpenEachInOutlook.CheckedChanged += radioButtonActionUpdateGUI_CheckedChanged;
            // 
            // radioButtonCopyToFolder
            // 
            radioButtonCopyToFolder.AutoSize = true;
            radioButtonCopyToFolder.Location = new Point(164, 31);
            radioButtonCopyToFolder.Name = "radioButtonCopyToFolder";
            radioButtonCopyToFolder.Size = new Size(130, 24);
            radioButtonCopyToFolder.TabIndex = 3;
            radioButtonCopyToFolder.Text = "Copy To Folder";
            toolTip1.SetToolTip(radioButtonCopyToFolder, "This will create a copy each email into the folder below. (creating even more duplicates)");
            radioButtonCopyToFolder.UseVisualStyleBackColor = true;
            radioButtonCopyToFolder.CheckedChanged += radioButtonActionUpdateGUI_CheckedChanged;
            // 
            // cboDestinationFolder
            // 
            cboDestinationFolder.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            cboDestinationFolder.FormattingEnabled = true;
            cboDestinationFolder.Location = new Point(36, 61);
            cboDestinationFolder.Name = "cboDestinationFolder";
            cboDestinationFolder.Size = new Size(565, 28);
            cboDestinationFolder.TabIndex = 2;
            // 
            // radioButtonDeleteEmails
            // 
            radioButtonDeleteEmails.AutoSize = true;
            radioButtonDeleteEmails.Location = new Point(13, 125);
            radioButtonDeleteEmails.Name = "radioButtonDeleteEmails";
            radioButtonDeleteEmails.Size = new Size(74, 24);
            radioButtonDeleteEmails.TabIndex = 1;
            radioButtonDeleteEmails.Text = "Delete";
            radioButtonDeleteEmails.UseVisualStyleBackColor = true;
            radioButtonDeleteEmails.CheckedChanged += radioButtonActionUpdateGUI_CheckedChanged;
            // 
            // radioButtonMoveToFolder
            // 
            radioButtonMoveToFolder.AutoSize = true;
            radioButtonMoveToFolder.Checked = true;
            radioButtonMoveToFolder.Location = new Point(13, 31);
            radioButtonMoveToFolder.Name = "radioButtonMoveToFolder";
            radioButtonMoveToFolder.Size = new Size(133, 24);
            radioButtonMoveToFolder.TabIndex = 0;
            radioButtonMoveToFolder.TabStop = true;
            radioButtonMoveToFolder.Text = "Move To Folder";
            toolTip1.SetToolTip(radioButtonMoveToFolder, "This will move each duplicate email into the folder below.");
            radioButtonMoveToFolder.UseVisualStyleBackColor = true;
            radioButtonMoveToFolder.CheckedChanged += radioButtonActionUpdateGUI_CheckedChanged;
            // 
            // tabPage4
            // 
            tabPage4.Controls.Add(button1);
            tabPage4.Controls.Add(richTextBox1);
            tabPage4.Location = new Point(4, 29);
            tabPage4.Name = "tabPage4";
            tabPage4.Padding = new Padding(3);
            tabPage4.Size = new Size(640, 516);
            tabPage4.TabIndex = 4;
            tabPage4.Text = "Accept";
            tabPage4.UseVisualStyleBackColor = true;
            // 
            // button1
            // 
            button1.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            button1.Location = new Point(482, 458);
            button1.Name = "button1";
            button1.Size = new Size(150, 50);
            button1.TabIndex = 4;
            button1.Text = "I Agree";
            button1.UseVisualStyleBackColor = true;
            button1.Click += btnI_Agree_Click;
            // 
            // richTextBox1
            // 
            richTextBox1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            richTextBox1.Location = new Point(10, 0);
            richTextBox1.Name = "richTextBox1";
            richTextBox1.ReadOnly = true;
            richTextBox1.Size = new Size(634, 437);
            richTextBox1.TabIndex = 0;
            richTextBox1.Text = resources.GetString("richTextBox1.Text");
            // 
            // tabPage6
            // 
            tabPage6.Controls.Add(txtNonMailItemsSkipped);
            tabPage6.Controls.Add(txtDuplicatesFound);
            tabPage6.Controls.Add(label5);
            tabPage6.Controls.Add(btnStop);
            tabPage6.Controls.Add(label3);
            tabPage6.Controls.Add(label2);
            tabPage6.Controls.Add(txtItemsTotalCount);
            tabPage6.Controls.Add(txtFoldersTotalCount);
            tabPage6.Controls.Add(txtEmailsProcessed);
            tabPage6.Controls.Add(txtFoldersProcessed);
            tabPage6.Controls.Add(label4);
            tabPage6.Controls.Add(lblFolders);
            tabPage6.Controls.Add(lblEmails);
            tabPage6.Controls.Add(btnStart);
            tabPage6.Location = new Point(4, 29);
            tabPage6.Name = "tabPage6";
            tabPage6.Size = new Size(640, 516);
            tabPage6.TabIndex = 3;
            tabPage6.Text = "Go";
            tabPage6.UseVisualStyleBackColor = true;
            // 
            // txtNonMailItemsSkipped
            // 
            txtNonMailItemsSkipped.Location = new Point(252, 234);
            txtNonMailItemsSkipped.Name = "txtNonMailItemsSkipped";
            txtNonMailItemsSkipped.ReadOnly = true;
            txtNonMailItemsSkipped.Size = new Size(75, 27);
            txtNonMailItemsSkipped.TabIndex = 12;
            // 
            // txtDuplicatesFound
            // 
            txtDuplicatesFound.Location = new Point(252, 207);
            txtDuplicatesFound.Name = "txtDuplicatesFound";
            txtDuplicatesFound.ReadOnly = true;
            txtDuplicatesFound.Size = new Size(75, 27);
            txtDuplicatesFound.TabIndex = 11;
            // 
            // label5
            // 
            label5.AutoSize = true;
            label5.Location = new Point(65, 237);
            label5.Name = "label5";
            label5.Size = new Size(181, 20);
            label5.TabIndex = 10;
            label5.Text = "Non-Email Items Skipped:";
            // 
            // btnStop
            // 
            btnStop.Enabled = false;
            btnStop.Location = new Point(327, 15);
            btnStop.Name = "btnStop";
            btnStop.Size = new Size(263, 60);
            btnStop.TabIndex = 9;
            btnStop.Text = "Stop";
            btnStop.UseVisualStyleBackColor = true;
            btnStop.Click += btnStop_Click;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(252, 129);
            label3.Name = "label3";
            label3.Size = new Size(75, 20);
            label3.TabIndex = 8;
            label3.Text = "Processed";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(343, 129);
            label2.Name = "label2";
            label2.Size = new Size(42, 20);
            label2.TabIndex = 8;
            label2.Text = "Total";
            // 
            // txtItemsTotalCount
            // 
            txtItemsTotalCount.Location = new Point(327, 180);
            txtItemsTotalCount.Name = "txtItemsTotalCount";
            txtItemsTotalCount.ReadOnly = true;
            txtItemsTotalCount.Size = new Size(75, 27);
            txtItemsTotalCount.TabIndex = 7;
            // 
            // txtFoldersTotalCount
            // 
            txtFoldersTotalCount.Location = new Point(327, 153);
            txtFoldersTotalCount.Name = "txtFoldersTotalCount";
            txtFoldersTotalCount.ReadOnly = true;
            txtFoldersTotalCount.Size = new Size(75, 27);
            txtFoldersTotalCount.TabIndex = 6;
            // 
            // txtEmailsProcessed
            // 
            txtEmailsProcessed.Location = new Point(252, 180);
            txtEmailsProcessed.Name = "txtEmailsProcessed";
            txtEmailsProcessed.ReadOnly = true;
            txtEmailsProcessed.Size = new Size(75, 27);
            txtEmailsProcessed.TabIndex = 5;
            // 
            // txtFoldersProcessed
            // 
            txtFoldersProcessed.Location = new Point(252, 153);
            txtFoldersProcessed.Name = "txtFoldersProcessed";
            txtFoldersProcessed.ReadOnly = true;
            txtFoldersProcessed.Size = new Size(75, 27);
            txtFoldersProcessed.TabIndex = 4;
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Location = new Point(78, 210);
            label4.Name = "label4";
            label4.Size = new Size(168, 20);
            label4.TabIndex = 3;
            label4.Text = "Duplicate Emails Found:";
            // 
            // lblFolders
            // 
            lblFolders.AutoSize = true;
            lblFolders.Location = new Point(186, 156);
            lblFolders.Name = "lblFolders";
            lblFolders.Size = new Size(60, 20);
            lblFolders.TabIndex = 2;
            lblFolders.Text = "Folders:";
            // 
            // lblEmails
            // 
            lblEmails.AutoSize = true;
            lblEmails.Location = new Point(100, 183);
            lblEmails.Name = "lblEmails";
            lblEmails.Size = new Size(146, 20);
            lblEmails.TabIndex = 1;
            lblEmails.Text = "Current Folder Items:";
            // 
            // btnStart
            // 
            btnStart.Location = new Point(24, 15);
            btnStart.Name = "btnStart";
            btnStart.Size = new Size(263, 60);
            btnStart.TabIndex = 0;
            btnStart.Text = "Start";
            btnStart.UseVisualStyleBackColor = true;
            btnStart.Click += btnStart_Click;
            // 
            // contextMenuStrip1
            // 
            contextMenuStrip1.ImageScalingSize = new Size(20, 20);
            contextMenuStrip1.Items.AddRange(new ToolStripItem[] { toolStripMenuItemCheckAllChildNodes, toolStripMenuItemUnCheckAllChildNodes });
            contextMenuStrip1.Name = "contextMenuStrip1";
            contextMenuStrip1.Size = new Size(231, 52);
            // 
            // toolStripMenuItemCheckAllChildNodes
            // 
            toolStripMenuItemCheckAllChildNodes.Name = "toolStripMenuItemCheckAllChildNodes";
            toolStripMenuItemCheckAllChildNodes.Size = new Size(230, 24);
            toolStripMenuItemCheckAllChildNodes.Text = "Check All Subfolders";
            toolStripMenuItemCheckAllChildNodes.Click += toolStripMenuItemCheckAllChildNodes_Click;
            // 
            // toolStripMenuItemUnCheckAllChildNodes
            // 
            toolStripMenuItemUnCheckAllChildNodes.Name = "toolStripMenuItemUnCheckAllChildNodes";
            toolStripMenuItemUnCheckAllChildNodes.Size = new Size(230, 24);
            toolStripMenuItemUnCheckAllChildNodes.Text = "Uncheck All Subfolders";
            toolStripMenuItemUnCheckAllChildNodes.Click += toolStripMenuItemUnCheckAllChildNodes_Click;
            // 
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            AutoSize = true;
            ClientSize = new Size(648, 549);
            Controls.Add(tabControl1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            MaximizeBox = false;
            MinimumSize = new Size(500, 500);
            Name = "MainForm";
            Text = "Duplicate Email Remover For Outlook Classic";
            Load += Form1_Load;
            Shown += Form1_Shown;
            tabControl1.ResumeLayout(false);
            tabPage1.ResumeLayout(false);
            tabPage1.PerformLayout();
            tabPage2.ResumeLayout(false);
            tabPage2.PerformLayout();
            tabPage3.ResumeLayout(false);
            groupBoxWhatToMatchOn.ResumeLayout(false);
            groupBoxWhatToMatchOn.PerformLayout();
            groupBoxWhatToDoWithEmails.ResumeLayout(false);
            groupBoxWhatToDoWithEmails.PerformLayout();
            tabPage4.ResumeLayout(false);
            tabPage6.ResumeLayout(false);
            tabPage6.PerformLayout();
            contextMenuStrip1.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private TabControl tabControl1;
        private TabPage tabPage1;
        private TabPage tabPage2;
        private Button btnTreeViewSelectNextPanel;
        private Label label1;
        private TreeView treeView1;
        private ContextMenuStrip contextMenuStrip1;
        private ToolStripMenuItem toolStripMenuItemCheckAllChildNodes;
        private ToolStripMenuItem toolStripMenuItemUnCheckAllChildNodes;
        private ListBox listBoxOfFoldersToProcess;
        private Button button2;
        private TabPage tabPage3;
        private TabPage tabPage6;
        private Button btnMoveToBeginProcessing;
        private GroupBox groupBoxWhatToDoWithEmails;
        private RadioButton radioButtonDeleteEmails;
        private RadioButton radioButtonMoveToFolder;
        private Button btnStart;
        private ComboBox cboDestinationFolder;
        private GroupBox groupBoxWhatToMatchOn;
        private CheckBox checkBoxMatchOnBCC;
        private CheckBox checkBoxMatchOnCC;
        private CheckBox checkBoxMatchOnTo;
        private CheckBox checkBoxMatchOnSenderEmail;
        private CheckBox checkBoxMatchOnAttachment;
        private CheckBox checkBoxMatchOnLastModTime;
        private CheckBox checkBoxMatchOnReceivedTime;
        private CheckBox checkBoxMatchOnSentOn;
        private CheckBox checkBoxMatchOnHTMLBody;
        private CheckBox checkBoxMatchOnBody;
        private CheckBox checkBoxMatchOnSubject;
        private RadioButton radioButtonCopyToFolder;
        private RadioButton radioButtonOpenEachInOutlook;
        private ToolTip toolTip1;
        private CheckBox checkBoxSaveEachDupToFilePath;
        private TextBox txtSaveEachDupToFilePath;
        private RadioButton radioButtonLogOnly;
        private Label label3;
        private Label label2;
        private TextBox txtItemsTotalCount;
        private TextBox txtFoldersTotalCount;
        private TextBox txtEmailsProcessed;
        private TextBox txtFoldersProcessed;
        private Label label4;
        private Label lblFolders;
        private Label lblEmails;
        private Button btnStop;
        private Label label5;
        private Label label7;
        private Label label6;
        private TabPage tabPage4;
        private Button button1;
        private RichTextBox richTextBox1;
        private TextBox txtNonMailItemsSkipped;
        private TextBox txtDuplicatesFound;
    }
}