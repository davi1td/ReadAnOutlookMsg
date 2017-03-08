using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using iwantedue;
using iwantedue.Windows.Forms;

namespace ReadAnOutlookMsg
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult msgFileSelectResult = this.openFileDialog1.ShowDialog();
            if (msgFileSelectResult == DialogResult.OK)
            {
                foreach (string msgfile in this.openFileDialog1.FileNames)
                {
                    Stream messageStream = File.Open(msgfile, FileMode.Open, FileAccess.Read);
                    OutlookStorage.Message message = new OutlookStorage.Message(messageStream);
                    messageStream.Close();

                    this.LoadMsgToTree(message, this.treeView1.Nodes.Add("MSG"));
                    message.Dispose();
                }
            }
        }

        private void treeView1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void treeView1_DragDrop(object sender, DragEventArgs e)
        {
            //wrap standard IDataObject in OutlookDataObject
            OutlookDataObject dataObject = new OutlookDataObject(e.Data);

            //get the names and data streams of the files dropped
            string[] filenames = (string[])dataObject.GetData("FileGroupDescriptor");
            MemoryStream[] filestreams = (MemoryStream[])dataObject.GetData("FileContents");

            for (int fileIndex = 0; fileIndex < filenames.Length; fileIndex++)
            {
                //use the fileindex to get the name and data stream
                string filename = filenames[fileIndex];
                MemoryStream filestream = filestreams[fileIndex];

                OutlookStorage.Message message = new OutlookStorage.Message(filestream);
                this.LoadMsgToTree(message, this.treeView1.Nodes.Add("MSG"));
                message.Dispose();
            }
        }

        private void LoadMsgToTree(OutlookStorage.Message message, TreeNode messageNode)
        {
            messageNode.Text = message.Subject;
            messageNode.Nodes.Add("Subject: " + message.Subject);
            TreeNode bodyNode = messageNode.Nodes.Add("Body: (double click to view)");
            bodyNode.Tag = new string[] {message.BodyText, message.BodyRTF};

            TreeNode recipientNode = messageNode.Nodes.Add("Recipients: " + message.Recipients.Count);
            foreach (OutlookStorage.Recipient recipient in message.Recipients)
            {
                recipientNode.Nodes.Add(recipient.Type + ": " + recipient.Email);
            }

            TreeNode attachmentNode = messageNode.Nodes.Add("Attachments: " + message.Attachments.Count);
            foreach (OutlookStorage.Attachment attachment in message.Attachments)
            {
                attachmentNode.Nodes.Add(attachment.Filename + ": " + attachment.Data.Length + "b");
            }

            TreeNode subMessageNode = messageNode.Nodes.Add("Sub Messages: " + message.Messages.Count);
            foreach (OutlookStorage.Message subMessage in message.Messages)
            {
                this.LoadMsgToTree(subMessage, subMessageNode.Nodes.Add("MSG"));
            }
        }

        private void treeView1_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Tag is string[])
            {
                string[] body = e.Node.Tag as string[];
                Form2 bodyTextForm = new Form2(body[0], body[1]);
                bodyTextForm.ShowDialog(this);
                bodyTextForm.Dispose();
            }
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}