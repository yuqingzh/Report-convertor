/*
 * Created by SharpDevelop.
 * User: yuqingzh
 * Date: 2011-6-27
 * Time: 14:55
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Text;

namespace Report_Convertor
{
	/// <summary>
	/// Description of Form4.
	/// </summary>
	public partial class frmHelp : Form
	{
		public frmHelp()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();
			
			//
			// TODO: Add constructor code after the InitializeComponent() call.
			//
			
			this.btnOK.Location = new System.Drawing.Point(this.Width / 2 - this.btnOK.Width / 2, 15);
		}
		
		void BtnOKClick(object sender, EventArgs e)
		{
			this.Dispose();
		}
		
		void FrmHelpLoad(object sender, EventArgs e)
		{
			string fileName = System.Environment.CurrentDirectory + "\\readme.txt";
			Encoding fileEncoding = TxtFileEncoding.GetEncoding(fileName, Encoding.GetEncoding("GB2312"));
			StreamReader reader = new StreamReader(fileName, fileEncoding);//用该编码创建StreamReader
			
			try   
			{    
				textBox1.Text = "";
				if (File.Exists(fileName))
				{
					do
					{
						textBox1.Text += reader.ReadLine() + "\r\n";					
					}while(reader.Peek() != -1);
				}
			} 
			catch 
			{ 
//				textBox1.Text =  "Read file failure";
			}
			finally
			{
				reader.Close();
			} 
		}
		
		void FrmHelpSizeChanged(object sender, EventArgs e)
		{
			this.btnOK.Location = new System.Drawing.Point(this.Width / 2 - this.btnOK.Width / 2, 15);
		}
	}
	
	
	/// <summary>
	/// ////////////////////////////////////////////
	/// </summary>
	public class TxtFileEncoding
	{
		public TxtFileEncoding()
		{

		} 

		public static Encoding GetEncoding(string fileName)
		{
			return GetEncoding(fileName, Encoding.Default);
		}

		public static Encoding GetEncoding(FileStream stream)
		{
			return GetEncoding(stream, Encoding.Default);
		}

		public static Encoding GetEncoding(string fileName, Encoding defaultEncoding)
		{
			FileStream fs = new FileStream(fileName, FileMode.Open);

			Encoding targetEncoding = GetEncoding(fs, defaultEncoding);

			fs.Close();

			return targetEncoding;
		}

		public static Encoding GetEncoding(FileStream stream, Encoding defaultEncoding)
		{
			Encoding targetEncoding = defaultEncoding;
			
			if(stream != null && stream.Length >= 2)
			{
				//保存文件流的前4个字节
				byte byte1 = 0;
				byte byte2 = 0;
				byte byte3 = 0;
				byte byte4 = 0; 

				//保存当前Seek位置
				long origPos = stream.Seek(0, SeekOrigin.Begin);
				stream.Seek(0, SeekOrigin.Begin);
				int nByte = stream.ReadByte();
				byte1 = Convert.ToByte(nByte);
				byte2 = Convert.ToByte(stream.ReadByte()); 

				if(stream.Length >= 3)
				{
					byte3 = Convert.ToByte(stream.ReadByte());
				}

				if(stream.Length >= 4)
				{
					byte4 = Convert.ToByte(stream.ReadByte());
				} 

				//根据文件流的前4个字节判断Encoding
				//Unicode {0xFF, 0xFE};
				//BE-Unicode {0xFE, 0xFF};
				//UTF8 = {0xEF, 0xBB, 0xBF};
				if(byte1 == 0xFE && byte2 == 0xFF)//UnicodeBe
				{
					targetEncoding = Encoding.BigEndianUnicode;
				}

				if(byte1 == 0xFF && byte2 == 0xFE && byte3 != 0xFF)//Unicode
				{
					targetEncoding = Encoding.Unicode;
				}


				if(byte1 == 0xEF && byte2 == 0xBB && byte3 == 0xBF)//UTF8
				{
					targetEncoding = Encoding.UTF8;
				}


				//恢复Seek位置       
				stream.Seek(origPos, SeekOrigin.Begin);
			}
			return targetEncoding;
		}
	}
}
