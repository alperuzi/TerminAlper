using System;
using System.IO;
using System.IO.Ports;
using System.Collections;
using System.Threading;
using System.Management;
using System.Linq;
using System.Collections.Generic;

namespace TerminAlper
{
    /// <summary> CommPort class creates a singleton instance
    /// of SerialPort (System.IO.Ports) </summary>
    /// <remarks> When ready, you open the port.
    ///   <code>
    ///   CommPort com = CommPort.Instance;
    ///   com.StatusChanged += OnStatusChanged;
    ///   com.DataReceived += OnDataReceived;
    ///  com.Open();
    ///   </code>
    ///   Notice that delegates are used to handle status and data events.
    ///   When settings are changed, you close and reopen the port.
    ///   <code>
    ///   CommPort com = CommPort.Instance;
    ///   com.Close();
    ///   com.PortName = "COM4";
    ///   com.Open();
    ///   </code>
    /// </remarks>
	public class CommPort
    {
        SerialPort _serialPort;
		Thread _readThread;
		bool _keepReading;
        string busName;

        bool sentBusy = false;

        //begin Singleton pattern
        static readonly CommPort instance = new CommPort();

		// Explicit static constructor to tell C# compiler
        // not to mark type as beforefieldinit
        static CommPort()
        {
        }

        CommPort()
        {
			_serialPort = new SerialPort();
			_readThread = null;
			_keepReading = false;
		}

		public static CommPort Instance
        {
            get
            {
                return instance;
            }
        }
        //end Singleton pattern

		//begin Observer pattern
        public delegate void EventHandler(string param);
        public EventHandler StatusChanged;
        public EventHandler DataReceived;
        //end Observer pattern

        public String StatusText;

        public void StartComPortThread()
		{
			if (!_keepReading)
			{
				_keepReading = true;
				_readThread = new Thread(ReadPort);
				_readThread.Start();
			}
		}

        public void StopComPortThread()
		{
            if(_readThread != null&&_readThread.IsAlive)
            {
                if (_keepReading)
                {
                    _keepReading = false;
                    _readThread.Join(); //block until exits
                    _readThread.Abort();
                    _readThread = null;
                }
            }
			
		}

        public bool IsReading()
        {
            return _keepReading;
        }

        public void setBusName(string _bName)
        {
            busName = _bName;
        }

		/// <summary> Get the data and pass it on. </summary>
		private void ReadPort()
            
		{
            long lastEventGenerateTime = 0;
            long eventGeneratePeriodInMillis = 100;
            string dataAccumulator = "";

            while (_keepReading)
            {
                if (_serialPort.IsOpen)
                {
                    byte[] readBuffer = new byte[_serialPort.ReadBufferSize + 1];
                    try
                    {
                        // If there are bytes available on the serial port,
                        // Read returns up to "count" bytes, but will not block (wait)
                        // for the remaining bytes. If there are no bytes available
                        // on the serial port, Read will block until at least one byte
                        // is available on the port, up until the ReadTimeout milliseconds
                        // have elapsed, at which time a TimeoutException will be thrown.
                        int count = _serialPort.Read(readBuffer, 0, _serialPort.ReadBufferSize);
                        String SerialIn = System.Text.Encoding.ASCII.GetString(readBuffer, 0, count);

                        dataAccumulator += SerialIn;


                    }
                    catch (Exception ex)
                    {
                        if (ex is TimeoutException) { }
                        if (ex is IOException || ex is UnauthorizedAccessException)
                        {
                            try
                            {
                                _serialPort.Close();
                            }
                            catch (Exception)
                            {
                            }
                            _serialPort = new SerialPort();
                            Open();

                        }
                    }
                }
                else
                {
                    TimeSpan waitTime = new TimeSpan(0, 0, 0, 0, 50);
                    Thread.Sleep(waitTime);
                    //_serialPort = new SerialPort();
                    Open();
                }
                if (DateTimeOffset.Now.ToUnixTimeMilliseconds() - lastEventGenerateTime > eventGeneratePeriodInMillis)
                {
                    if (!dataAccumulator.Equals(""))
                    {
                        DataReceived(dataAccumulator);
                        dataAccumulator = "";
                    }
                   
                    lastEventGenerateTime = DateTimeOffset.Now.ToUnixTimeMilliseconds();
                }
            }

		}

		/// <summary> Open the serial port with current settings. </summary>
        public void Open()
        {
            if (_serialPort.IsOpen)
            {
                _serialPort.Close();
                _serialPort.Dispose();
                _serialPort = new SerialPort();
            }
            try
            {
                _serialPort.PortName = Settings.Port.PortName;
                _serialPort.BaudRate = Settings.Port.BaudRate;
                _serialPort.Parity = Settings.Port.Parity;
                _serialPort.DataBits = Settings.Port.DataBits;
                _serialPort.StopBits = Settings.Port.StopBits;
                _serialPort.Handshake = Settings.Port.Handshake;

				// Set the read/write timeouts
				_serialPort.ReadTimeout = 50;
				_serialPort.WriteTimeout = 50;
                
                do
                {
                    _serialPort.Open();
                } while (!_serialPort.IsOpen);
            }
            catch (IOException)
            {
                //StatusChanged(String.Format("{0} does not exist", Settings.Port.PortName));
                StatusText = String.Format("{0} does not exist", Settings.Port.PortName);
            }
            catch (UnauthorizedAccessException)
            {
                if (!sentBusy) {
                    //StatusChanged(String.Format("{0} already in use", Settings.Port.PortName));
                    StatusText = String.Format("{0} already in use", Settings.Port.PortName);
                    sentBusy = true;
                }
                
            }
            catch (InvalidOperationException)
            {
                //StatusChanged(String.Format("{0} wadaaaa", Settings.Port.PortName));
                StatusText = String.Format("{0} wadaaaa", Settings.Port.PortName);
            }

            // Update the status
            if (_serialPort.IsOpen)
            {
                string p = _serialPort.Parity.ToString().Substring(0, 1);   //First char
                string h = _serialPort.Handshake.ToString();
                if (_serialPort.Handshake == Handshake.None)
                    h = "no handshake"; // more descriptive than "None"

                /*StatusChanged(String.Format("{0}: {1} bps, {2}{3}{4}, {5}",
                    _serialPort.PortName, _serialPort.BaudRate,
                    _serialPort.DataBits, p, (int)_serialPort.StopBits, h));*/
                StatusText = String.Format("{0}: {1} bps, {2}",
                    _serialPort.PortName, _serialPort.BaudRate,
                    Settings.Port.busName);

                sentBusy = false;
            }
            else
            {
                //(String.Format("{0} already in use", Settings.Port.PortName));
            }
        }

        /// <summary> Close the serial port. </summary>
        public void Close()
        {
			StopComPortThread();

            if (_serialPort.IsOpen)
            {
                _serialPort.Close();
                _serialPort.Dispose();
            }

            //StatusChanged("Disconnected");
            StatusText = "Disconnected";
        }

        /// <summary> Get the status of the serial port. </summary>
        public bool IsOpen
        {
            get
            {
                return _serialPort.IsOpen;
            }  
        }
     
        /// <summary> Get a list of the available ports. Already opened ports
        /// are not returend. </summary>
        /// Alper: not used anymore!

        /// <summary>Send data to the serial port after appending line ending. </summary>
        /// <param name="data">An string containing the data to send. </param>
        public void Send(string data)
        {
            if (IsOpen)
            {
                string lineEnding = "";
                switch (Settings.Option.AppendToSend)
                {
                    case Settings.Option.AppendType.AppendCR:
                        lineEnding = "\r"; break;
                    case Settings.Option.AppendType.AppendLF:
                        lineEnding = "\n"; break;
                    case Settings.Option.AppendType.AppendCRLF:
                        lineEnding = "\r\n"; break;
                }

                _serialPort.Write(data + lineEnding);
            }
        }

    }
    internal class ProcessConnection
    {
        public static ConnectionOptions ProcessConnectionOptions()
        {
            ConnectionOptions options = new ConnectionOptions();
            options.Impersonation = ImpersonationLevel.Impersonate;
            options.Authentication = AuthenticationLevel.Default;
            options.EnablePrivileges = true;
            return options;
        }


        public static ManagementScope ConnectionScope(string machineName, ConnectionOptions options, string path)
        {
            ManagementScope connectScope = new ManagementScope();
            connectScope.Path = new ManagementPath(@"\\" + machineName + path);
            connectScope.Options = options;
            connectScope.Connect();
            return connectScope;
        }
    }
}
