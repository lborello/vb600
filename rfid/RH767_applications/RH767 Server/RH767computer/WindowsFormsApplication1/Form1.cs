using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.IO;
using Microsoft.Win32;
using System.Diagnostics;
using System.Threading;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
    //const string PAPY_IP = "localhost";
        const UInt16 PAPY_PORT = 1999;

        public Form1()
        {
            InitializeComponent();
        }

        public int numeral = 0;

        public void writeList()
        {
            char[] EPC = new char[24];
            char[] times = new char[4];
            int veces = 0;

            for (int i = 0; data[i] != '%'; i++)
            {
                if (data[i] == '#')
                    numeral++;
            }

            ListViewItem linea = new ListViewItem();
            for (int j = 0; j < numeral; j++)
            {
                for (int h = 0; h < 24; h++)
                {
                    EPC[h] = data[(veces*30) + h];
                }
                for(int l = 0; l < 4 ; l++)
                {
                    times[l] = data[(veces * 30) + 25 + l];
                }
                string ID = new string(EPC);
                ID = ID.ToUpper();
                string timesNumber = new string(times);
                int numVeces = Convert.ToInt32(timesNumber);
                timesNumber = numVeces.ToString();
                listViewTags.Items.Add(ID).SubItems.Add(timesNumber);
                veces++;

                if (chkPap.Checked)
                {
                    try
                    {
                        //Socket sender = new Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp);
                        //sender.Connect(PAPY_IP, PAPY_PORT);
                        //sender.Send(Encoding.ASCII.GetBytes(ID), SocketFlags.None);
                        //sender.Close();
                        string data = "s" + ID + ";" + DateTime.Now;
                        Socket sender = new Socket(AddressFamily.InterNetwork, SocketType.Dgram, ProtocolType.Udp);
                        sender.Connect(txtPapIP.Text, PAPY_PORT);
                        sender.Send(Encoding.ASCII.GetBytes(data), SocketFlags.None);
                        sender.Close();
                        Thread.Sleep(100);
                    }
                    catch (Exception ex)
                    {
                        labelText.Text = ex.Message;
                    }
                }
            }
            numeral = 0;
        }

        public string data = null;

        private void btnConnect_Click(object sender, EventArgs e)
        {
            bwServer.RunWorkerAsync();
            
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            labelText.Text = "Saving";
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = "Tags RH767"; // Default file name
            dlg.DefaultExt = ".txt"; // Default file extension
            dlg.Filter = "Text files (.txt)|*.txt"; // Filter files by extension

            // Show save file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            string[] array = new string[(listViewTags.Items.Count) + 1];

            string tag = "EPC";
            string times = "Times";
            string space = "                         ";
            array[0] = String.Concat(tag, space, times);

            for (int i = 0; i < listViewTags.Items.Count; i++)
            {
                tag = listViewTags.Items[i].Text;
                times = listViewTags.Items[i].SubItems[1].Text;
                space = "    ";

                array[i + 1] = String.Concat(tag, space, times);
            }
            // Process save file dialog box results
            if (result == true)
            {
                // Save document
                string filename = dlg.FileName;

                File.WriteAllLines(filename, array, Encoding.UTF8); //array is your array of strings
                labelText.Text = "File saved";
            }
            else
            {
                labelText.Text = "";
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            listViewTags.Items.Clear();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            byte[] bytes = new Byte[4096];

            //IPHostEntry ipHostInfo = Dns.Resolve(Dns.GetHostName());
            //IPAddress ipAddress = ipHostInfo.AddressList[0];
            IPEndPoint localEndPoint = new IPEndPoint(IPAddress.Any, 35500);

            Socket listener = new Socket(AddressFamily.InterNetwork,
                SocketType.Stream, ProtocolType.Tcp);

            listener.Bind(localEndPoint);
            listener.Listen(5);

            labelText.Text = "Waiting for connection...";

            while (true)
            {

                // Start listening for connections.
                //Console.WriteLine("Waiting for a connection...");
                // Program is suspended while waiting for an incoming connection.
                Socket handler = listener.Accept();
                data = null;

                labelText.Text = "Connected";

                // An incoming connection needs to be processed.
                while (true)
                {
                    bytes = new byte[4096];
                    int bytesRec = handler.Receive(bytes);
                    data += Encoding.ASCII.GetString(bytes, 0, bytesRec);
                    if (data.IndexOf("%") > -1)
                    {
                        break;
                    }
                }
                writeList();
                handler.Shutdown(SocketShutdown.Both);
                handler.Close();

                labelText.Text = "Ready";
             }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CheckForIllegalCrossThreadCalls = false;
           
        }

        private void listViewTags_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
