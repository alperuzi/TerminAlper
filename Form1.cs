using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Threading;
using System.Linq;

namespace TerminAlper
{
    public partial class Form1 : Form
    {
		/// <summary>
		/// Class to keep track of string and color for lines in output window.
		/// </summary>
		private class Line
		{
			public string Str;
			public Color ForeColor;

			public Line(string str, Color color)
			{
				Str = str;
				ForeColor = color;
			}
		};

		ArrayList lines = new ArrayList();

        Font origFont;
        Font monoFont;

		public Form1()
        {
            InitializeComponent();

            this.Text += " " + System.Reflection.Assembly.GetEntryAssembly().GetName().Version;

            splitContainer1.FixedPanel = FixedPanel.Panel1;
            splitContainer2.FixedPanel = FixedPanel.Panel2;

            AcceptButton = button5; //Send
            CancelButton = button4; //Close

			outputList_Initialize();

			Settings.Read();
            TopMost = Settings.Option.StayOnTop;

			// let form use multiple fonts
            origFont = Font;
            FontFamily ff = new FontFamily("Courier New");
            monoFont = new Font(ff, 10, FontStyle.Regular);
            Font = Settings.Option.MonoFont ? monoFont : origFont;

            CommPort com = CommPort.Instance;
            com.StatusChanged += OnStatusChanged;
            com.DataReceived += OnDataReceived;
            com.StartComPortThread();

			var timer = new System.Windows.Forms.Timer();
			timer.Tick += new EventHandler(timer_Tick);
			timer.Interval = 200; //200us
			timer.Start();
		}

		private void timer_Tick(object sender, EventArgs e)
		{
			CommPort com = CommPort.Instance;
			OnStatusChanged(com.StatusText);
		}

		private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            CommPort com = CommPort.Instance;
            if (com.IsOpen)
            {
                com.Close();
            }

            Environment.Exit(1);
        }

        /// <summary>
        /// output string to log file
        /// </summary>
        /// <param name="stringOut">string to output</param>
        public void logFile_writeLine(string stringOut)
		{
			if (Settings.Option.LogFileName != "")
			{
				Stream myStream = File.Open(Settings.Option.LogFileName,
					FileMode.Append, FileAccess.Write, FileShare.Read);
				if (myStream != null)
				{
					StreamWriter myWriter = new StreamWriter(myStream, Encoding.UTF8);
					myWriter.WriteLine(stringOut);
					myWriter.Close();
				}
			}
		}

		#region Output window

		string filterString = "";
		bool scrolling = true;
		Color receivedColor = Color.Yellow;
		Color sentColor = Color.White;

		/// <summary>
		/// context menu for the output window
		/// </summary>
		ContextMenu popUpMenu;

		/// <summary>
		/// check to see if filter matches string
		/// </summary>
		/// <param name="s">string to check</param>
		/// <returns>true if matches filter</returns>
		bool outputList_ApplyFilter(String s)
		{
			List<string> filters = filterString.Split(Settings.Option.filterDelimiter, StringSplitOptions.None).ToList();
			bool ret = false;

			foreach (string f in filters)
			{
				if (f == "")
				{
					ret = true;
				}
				else if (s == "")
				{
					ret = false;
				}
				else if (Settings.Option.FilterUseCase)
				{
					ret = s.Contains(f);
				}
				else
				{
					string upperString = s.ToUpper();
					string upperFilter = f.ToUpper();
					ret = upperString.Contains(upperFilter);
				}
				if (ret)
					return ret;
			}
			return ret;
		}

		/// <summary>
		/// clear the output window
		/// </summary>
		void outputList_ClearAll()
		{
			lines.Clear();
			partialLine = null;

			outputList.Items.Clear();
		}

		/// <summary>
		/// refresh the output window
		/// </summary>
		void outputList_Refresh()
		{
			outputList.BeginUpdate();
			outputList.Items.Clear();
			foreach (Line line in lines)
			{
				if (outputList_ApplyFilter(line.Str))
				{
					outputList.Items.Add(line);
				}
			}
			outputList.EndUpdate();
			outputList_Scroll();
		}
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            bool keyProcessed = false;
            switch (keyData)
            {
                case Keys.Escape:
                    outputList_ClearAll();
                    keyProcessed = true;
                    break;
                case Keys.F5:
                    outputList_ClearAll();
                    keyProcessed = true;
                    break;
                case Keys.F1:
                    button4_Click(null, null);
                    keyProcessed = true;
                    break;
                    
            }
            return keyProcessed;
        }
        /// <summary>
        /// add a new line to output window
        /// </summary>
        Line outputList_Add(string str, Color color)
		{
            while (outputList.Items.Count > 0 && outputList.Items.Count > Settings.Option.MaximumNumberOfDisplayLines)
            {
                outputList.Items.RemoveAt(0);
            }

            while (lines.Count > 0 && lines.Count > Settings.Option.MaximumNumberOfDisplayLines)
            {
                lines.RemoveAt(0);
            }

            Line newLine = new Line(str, color);
			lines.Add(newLine);

			if (outputList_ApplyFilter(newLine.Str))
			{
				outputList.Items.Add(newLine);
				outputList_Scroll();
			}

			return newLine;
		}

		/// <summary>
		/// Update a line in the output window.
		/// </summary>
		/// <param name="line">line to update</param>
		void outputList_Update(Line line)
		{
            /*logFile_writeLine("CONTENTS OF OUTPUTLIST:"); //CR LF meselesi galiba sadece LF ler var ucurdugu yerlerde.

            for (int i = 0; i < outputList.Items.Count; i++)
            {
                logFile_writeLine(((Line)outputList.Items[i]).Str);
            }*/
            // should we add to output?
            if (outputList_ApplyFilter(line.Str))
			{
				// is the line already displayed?
				bool found = false;
				for (int i = 0; i < outputList.Items.Count; ++i)
				{
					int index = (outputList.Items.Count - 1) - i;
					if (line == outputList.Items[index])
					{
						// is item visible?
						int itemsPerPage = (int)(outputList.Height / outputList.ItemHeight);
						if (index >= outputList.TopIndex &&
							index < (outputList.TopIndex + itemsPerPage))
						{
							// is there a way to refresh just one line
							// without redrawing the entire listbox?
							// changing the item value has no effect
							outputList.Refresh();
						}
						found = true;
						break;
					}
				}
				if (!found)
				{
					// not found, so add it
					outputList.Items.Add(line);
				}
			}
		}

		/// <summary>
		/// Initialize the output window
		/// </summary>
		private void outputList_Initialize()
		{
			// owner draw for listbox so we can add color
			outputList.DrawMode = DrawMode.OwnerDrawFixed;
			outputList.DrawItem += new DrawItemEventHandler(outputList_DrawItem);
			outputList.ClearSelected();

			// build the outputList context menu
			popUpMenu = new ContextMenu();
			popUpMenu.MenuItems.Add("&Copy", new EventHandler(outputList_Copy));
			popUpMenu.MenuItems[0].Visible = true;
			popUpMenu.MenuItems[0].Enabled = false;
			popUpMenu.MenuItems[0].Shortcut = Shortcut.CtrlC;
			popUpMenu.MenuItems[0].ShowShortcut = true;
			popUpMenu.MenuItems.Add("Copy All", new EventHandler(outputList_CopyAll));
			popUpMenu.MenuItems[1].Visible = true;
			popUpMenu.MenuItems.Add("Select &All", new EventHandler(outputList_SelectAll));
			popUpMenu.MenuItems[2].Visible = true;
			popUpMenu.MenuItems[2].Shortcut = Shortcut.CtrlA;
			popUpMenu.MenuItems[2].ShowShortcut = true;
			popUpMenu.MenuItems.Add("Clear Selected", new EventHandler(outputList_ClearSelected));
			popUpMenu.MenuItems[3].Visible = true;
			outputList.ContextMenu = popUpMenu;
		}

		/// <summary>
		/// draw item with color in output window
		/// </summary>
		void outputList_DrawItem(object sender, DrawItemEventArgs e)
		{
			e.DrawBackground();
			if (e.Index >= 0 && e.Index < outputList.Items.Count)
			{
				Line line = (Line)outputList.Items[e.Index];

				// if selected, make the text color readable
				Color color = line.ForeColor;
				if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
				{
					color = Color.Black;	// make it readable
				}
                //e.Graphics.DrawString(line.Str, e.Font, new SolidBrush(color),e.Bounds, StringFormat.GenericDefault);
                e.Graphics.DrawString(line.Str.Length > 0 ? (line.Str.StartsWith("\n") ? line.Str.Substring(1, line.Str.Length - 1) : line.Str) : "", e.Font, new SolidBrush(Color.Yellow), e.Bounds, StringFormat.GenericDefault);
            }
			e.DrawFocusRectangle();
		}

		/// <summary>
		/// Scroll to bottom of output window
		/// </summary>
		void outputList_Scroll()
		{
			if (scrolling)
			{
				int itemsPerPage = (int)(outputList.Height / outputList.ItemHeight);
				outputList.TopIndex = outputList.Items.Count - itemsPerPage;
			}
		}

		/// <summary>
		/// Enable/Disable copy selection in output window
		/// </summary>
		private void outputList_SelectedIndexChanged(object sender, EventArgs e)
		{
			popUpMenu.MenuItems[0].Enabled = (outputList.SelectedItems.Count > 0);
		}

		/// <summary>
		/// copy selection in output window to clipboard
		/// </summary>
		private void outputList_Copy(object sender, EventArgs e)
		{
			int iCount = outputList.SelectedItems.Count;
			if (iCount > 0)
			{
				String[] source = new String[iCount];
				for (int i = 0; i < iCount; ++i)
				{
					source[i] = ((Line)outputList.SelectedItems[i]).Str;
				}

				String dest = String.Join("\r\n", source);
				Clipboard.SetText(dest);
			}
		}

		/// <summary>
		/// copy all lines in output window
		/// </summary>
		private void outputList_CopyAll(object sender, EventArgs e)
		{
			int iCount = outputList.Items.Count;
			if (iCount > 0)
			{
				String[] source = new String[iCount];
				for (int i = 0; i < iCount; ++i)
				{
					source[i] = ((Line)outputList.Items[i]).Str;
				}

				String dest = String.Join("\r\n", source);
				Clipboard.SetText(dest);
			}
		}

		/// <summary>
		/// select all lines in output window
		/// </summary>
		private void outputList_SelectAll(object sender, EventArgs e)
		{
			outputList.BeginUpdate();
			for (int i = 0; i < outputList.Items.Count; ++i)
			{
			    outputList.SetSelected(i, true);
			}
			outputList.EndUpdate();
		}

		/// <summary>
		/// clear selected in output window
		/// </summary>
		private void outputList_ClearSelected(object sender, EventArgs e)
		{
			outputList.ClearSelected();
			outputList.SelectedItem = -1;
		}

		#endregion

		#region Event handling - data received and status changed

		/// <summary>
		/// Prepare a string for output by converting non-printable characters.
		/// </summary>
		/// <param name="StringIn">input string to prepare.</param>
		/// <returns>output string.</returns>
		private String PrepareData(String StringIn)
		{
			// The names of the first 32 characters
			string[] charNames = { "NUL", "SOH", "STX", "ETX", "EOT",
				"ENQ", "ACK", "BEL", "BS", "TAB", "LF", "VT", "FF", "CR", "SO", "SI",
				"DLE", "DC1", "DC2", "DC3", "DC4", "NAK", "SYN", "ETB", "CAN", "EM", "SUB",
				"ESC", "FS", "GS", "RS", "US", "Space"};

			string StringOut = "";

			foreach (char c in StringIn)
            {
                if (Settings.Option.HexOutput)
                {
                    StringOut = StringOut + String.Format("{0:X2} ", (int)c);
                }
                else if (c < 32 && c != 9 && c != 10)
                {
                    StringOut = StringOut + "<" + charNames[c] + ">";

                    //Uglier "Termite" style
                    //StringOut = StringOut + String.Format("[{0:X2}]", (int)c);
                }
                else
                {
                    StringOut = StringOut + c;
                }
            }
			return StringOut;
		}

		/// <summary>
		/// Partial line for AddData().
		/// </summary>
		private Line partialLine = null;

		/// <summary>
		/// Add data to the output.
		/// </summary>
		/// <param name="StringIn"></param>
		/// <returns></returns>
		private Line AddData(String StringIn)
		{
			String StringOut = PrepareData(StringIn);

			// if we have a partial line, add to it.
			if (partialLine != null)
			{
				// tack it on
				partialLine.Str = partialLine.Str + StringOut;
				outputList_Update(partialLine);
				return partialLine;
			}
			return outputList_Add(StringOut, receivedColor);
		}

		// delegate used for Invoke
		internal delegate void StringDelegate(string data);

		/// <summary>
		/// Handle data received event from serial port.
		/// </summary>
		/// <param name="data">incoming data</param>
		public void OnDataReceived(string dataIn)
        {
            //Handle multi-threading
            if (InvokeRequired)
            {
				BeginInvoke(new StringDelegate(OnDataReceived), new object[] { dataIn });
                return;
            }

			// pause scrolling to speed up output of multiple lines
			bool saveScrolling = scrolling;
			scrolling = false;

            // if we detect a line terminator, add line to output
            int index;
			while (dataIn.Length > 0 &&
				((index = dataIn.IndexOf("\r")) != -1 ||
				(index = dataIn.IndexOf("\n")) != -1))
            {
				String StringIn = dataIn.Substring(0, index);
				dataIn = dataIn.Remove(0, index + 1);

				logFile_writeLine(AddData(StringIn).Str);
				partialLine = null;	// terminate partial line
            }

			// if we have data remaining, add a partial line
			if (dataIn.Length > 0)
			{
				partialLine = AddData(dataIn);
			}

			// restore scrolling
			scrolling = saveScrolling;
			outputList_Scroll();
		}

		/// <summary>
		/// Update the connection status
		/// </summary>
		public void OnStatusChanged(string status)
		{
			//Handle multi-threading
			//This is not needed any more, UI update is timer tick based.
			if (InvokeRequired)
			{
				BeginInvoke(new StringDelegate(OnStatusChanged), new object[] { status });
				//return;
			}
            else
            {
                textBox1.Text = status;
            }
		}

		#endregion

		#region User interaction

		/// <summary>
		/// toggle connection status
		/// </summary>
		private void textBox1_Click(object sender, MouseEventArgs e)
		{ 	
			outputList.Focus();
		}

		/// <summary>
		/// Change filter
		/// </summary>
		private void textBox2_TextChanged(object sender, EventArgs e)
        {
            filterString = textBox2.Text;
			outputList_Refresh();
        }

		/// <summary>
		/// Show settings dialog
		/// </summary>
		private void button1_Click(object sender, EventArgs e)
		{
			TopMost = false;

			Form2 form2 = new Form2();
			form2.ShowDialog();

			TopMost = Settings.Option.StayOnTop;
			Font = Settings.Option.MonoFont ? monoFont : origFont;
            outputList.Focus();
        }

		/// <summary>
		/// Clear the output window
		/// </summary>
		private void button2_Click(object sender, EventArgs e)
		{
			outputList_ClearAll();
            outputList.Focus();
        }

		/// <summary>
		/// Show about dialog
		/// </summary>
		private void button3_Click(object sender, EventArgs e)
		{
			TopMost = false;
			AboutBox about = new AboutBox();          
            about.ShowDialog();
            TopMost = Settings.Option.StayOnTop;
            outputList.Focus();
        }

        /// <summary>
        /// Alper: Don't close app, just disconnect!
        /// </summary>
		private void button4_Click(object sender, EventArgs e)
		{
            CommPort com = CommPort.Instance;

            /*if (!com.IsOpen)
                return;*/

            if (com.IsReading())
            {
                com.StopComPortThread();
                com.Close();
                button4.Text = "Connect";
            }
           else
            {
                com.Open();
                com.StartComPortThread();
                button4.Text = "Disconnect";
            }
            outputList.Focus();
        }

		/// <summary>
		/// Send command
		/// </summary>
		private void button5_Click(object sender, EventArgs e)
        {
            string command = textBox3.Text;
            //textBox3.Text = "";
            textBox3.Focus();

			if (command.Length > 0)
			{
				CommPort com = CommPort.Instance;

                
     
                if (Settings.Option.LocalEcho)
				{
					outputList_Add(command + "\n", sentColor);
				}

                com.Send(command);

            }
        }

		/// <summary>
		/// send file to serial port
		/// </summary>
		private void button6_Click(object sender, EventArgs e)
		{
			OpenFileDialog dialog = new OpenFileDialog();
			dialog.RestoreDirectory = false;
			dialog.Title = "Select a file";
			if (dialog.ShowDialog() == DialogResult.OK)
			{
				String text = System.IO.File.ReadAllText(dialog.FileName);

				CommPort com = CommPort.Instance;
				com.Send(text);

				if (Settings.Option.LocalEcho)
				{
					outputList_Add("SendFile " + dialog.FileName + "," +
						text.Length.ToString() + " byte(s)\n", sentColor);
				}
			}
		}

		/// <summary>
		/// toggle scrolling
		/// </summary>
		private void button7_Click(object sender, EventArgs e)
		{
			scrolling = !scrolling;
			outputList_Scroll();
            outputList.Focus();
        }

        #endregion
      
        private void outputList_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Middle)
            {
               outputList_ClearAll();
            }
        }
    }

    internal class FlickerFreeListBox : System.Windows.Forms.ListBox
    {
        public FlickerFreeListBox()
        {
            this.SetStyle(
                ControlStyles.OptimizedDoubleBuffer |
                ControlStyles.ResizeRedraw |
                ControlStyles.UserPaint,
                true);
            this.DrawMode = DrawMode.OwnerDrawFixed;
        }
        protected override void OnDrawItem(DrawItemEventArgs e)
        {
            if (this.Items.Count > 0)
            {
               e.DrawBackground();
               //e.Graphics.DrawString(this.Items[e.Index].ToString(), e.Font, new SolidBrush(this.ForeColor), new PointF(e.Bounds.X, e.Bounds.Y));
            }
            base.OnDrawItem(e);
        }
        protected override void OnPaint(PaintEventArgs e)
        {
            Region iRegion = new Region(e.ClipRectangle);
            e.Graphics.FillRegion(new SolidBrush(this.BackColor), iRegion);
            if (this.Items.Count > 0)
            {
                for (int i = 0; i < this.Items.Count; ++i)
                {
                    System.Drawing.Rectangle irect = this.GetItemRectangle(i);
                    if (e.ClipRectangle.IntersectsWith(irect))
                    {
                        if ((this.SelectionMode == SelectionMode.One && this.SelectedIndex == i)
                        || (this.SelectionMode == SelectionMode.MultiSimple && this.SelectedIndices.Contains(i))
                        || (this.SelectionMode == SelectionMode.MultiExtended && this.SelectedIndices.Contains(i)))
                        {
                            OnDrawItem(new DrawItemEventArgs(e.Graphics, this.Font,
                                irect, i,
                                DrawItemState.Selected, this.ForeColor,
                                this.BackColor));
                        }
                        else
                        {
                            OnDrawItem(new DrawItemEventArgs(e.Graphics, this.Font,
                                irect, i,
                                DrawItemState.Default, this.ForeColor,
                                this.BackColor));
                        }
                        iRegion.Complement(irect);
                    }
                }
            }
            base.OnPaint(e);
        }

    }
}
