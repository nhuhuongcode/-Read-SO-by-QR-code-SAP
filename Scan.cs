using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using AForge;
using AForge.Video;
using AForge.Video.DirectShow;
using ZXing;
using ZXing.Aztec;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;

namespace WindowsFormsApp1
{
	public partial class Scan : Form
	{
		private FilterInfoCollection CaptureDevice;
		private VideoCaptureDevice FinalFrame;
		public Scan()
		{
			InitializeComponent();
		}

		private void Scan_Load(object sender, EventArgs e)
		{
			CaptureDevice = new FilterInfoCollection(FilterCategory.VideoInputDevice);
			foreach(FilterInfo Device in CaptureDevice) 
			{
				comboBox1.Items.Add(Device.Name);
			}

			comboBox1.SelectedIndex = 0;
			FinalFrame= new VideoCaptureDevice();


		}

		private void button1_Click(object sender, EventArgs e)
		{
			FinalFrame = new VideoCaptureDevice(CaptureDevice[comboBox1.SelectedIndex].MonikerString);
			FinalFrame.NewFrame += new NewFrameEventHandler(FinalFrame_NewFrame);
			FinalFrame.Start();
		}

		private void FinalFrame_NewFrame(object sender, NewFrameEventArgs eventArgs)
		{
			pictureBox1.Image=(Bitmap)eventArgs.Frame.Clone();
		}

		private void Scan_FormClosing(object sender, FormClosingEventArgs e)
		{
			if(FinalFrame.IsRunning == true)
			{
				FinalFrame.Stop();
			}
		}

		private void timer1_Tick(object sender, EventArgs e)
		{
			BarcodeReader reader= new BarcodeReader();
			Result result = reader.Decode((Bitmap)pictureBox1.Image);
			try
			{
				string decode = result.ToString().Trim();
				if(decode != "")
				{
					textBox1.Text = decode.ToString();
					MessageBox.Show(decode);
					if (FinalFrame.IsRunning == true)
					{
						timer1.Stop();
					}
				}
				else
				{
					MessageBox.Show(decode);
				}
			}
			catch 
			{
				
			}
		}

		private void button2_Click(object sender, EventArgs e)
		{
			timer1.Start();
			FinalFrame.Stop();
		}

		private void textBox1_TextChanged(object sender, EventArgs e)
		{
			Form1 form1 = new Form1();
			form1.res = getdocnum(textBox1.Text.Trim());
			form1.ShowDialog();
			this.Close();
		}
		public int getdocnum(string dn)
		{
			try
			{
				string[] tuTrongChuoi = dn.Split(' '); // Tách chuỗi thành các từ dựa trên khoảng trắng

				if (tuTrongChuoi.Length > 0)
				{
					string tuDauTien = tuTrongChuoi[0]; // Lấy từ đầu tiên sau khi tách
					return int.Parse(tuDauTien);
				}
			}
			catch (Exception ex)
			{
				return 0; // Use double quotes here to return an empty string
			}

			// Return a default value here if neither the try nor the catch block is executed
			return 0;
		}

	}
}
