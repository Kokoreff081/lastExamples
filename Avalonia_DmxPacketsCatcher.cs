using Acn.ArtNet.IO;
using Acn.ArtNet.Packets;
using Acn.ArtNet.Sockets;
using Avalonia.Threading;
using DMXCatcher.Models;
using ReactiveUI;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Net;
using System.Reactive;
using System;
using System.Linq;
using str = DMXCatcher.Resources;
using Avalonia.Controls;

namespace DMXCatcher.ViewModels
{
    public class MainWindowViewModel : ViewModelBase
    {
        #region Fields
        private Window hostWindow;
        private IPAddress _currentIp;
        private double _currentNet;
        private SubNetControl _currentSubNet;
        private UniverseControl _currentUniverse;
        public bool _sniffFlag = false;
        private Socket mainSocket;
        private ArtNetSocket artSocket;
        private System.Timers.Timer pollTimer;
        private List<ArtNetDevice> _artnetDevices;

        private byte[] byteData = new byte[8372];
        private Dictionary<short, byte[]> dmxDataPerUniverse;
        //private ObservableCollection<DmxDataToList> _dmxDataToList;
        private Dictionary<int, List<UniverseControl>> _universesList;
        private Dictionary<double, List<SubNetControl>> _subNetsList;
        private List<SubNetControl> _subnetsItems;
        private List<UniverseControl> _universesItems;
        private bool _dmxReceiveing;
        private List<string> artNetIps;
        private string _netActivity = str.Resources.LabelNetActivity;
        private string _buttonContent = str.Resources.ButtonStart;
        private int _newCurrentUniverseL;//0 - 32767
        private int _newCurrentUniverse;//0-15
        private int _newSubNet;
        private int _newNet;
        private bool universeLChanged;
        private bool _stopSniffButton = false;
        private bool _startSniffButton = true;
        #endregion

        #region Properties

        public bool StopSniffButton
        {
            get { return _stopSniffButton; }
            set { this.RaiseAndSetIfChanged(ref _stopSniffButton, value); }
        }
        public bool StartSniffButton
        {
            get { return _startSniffButton; }
            set { this.RaiseAndSetIfChanged(ref _startSniffButton, value); }
        }

        public int NewCurrentUniverseL
        {
            get { return _newCurrentUniverseL; }
            set
            {
                this.RaiseAndSetIfChanged(ref _newCurrentUniverseL, value);
                universeLChanged = true;
                RecalculateNetSubNet(value);
                universeLChanged = false;
            }
        }
        public int NewCurrentUniverse
        {
            get { return _newCurrentUniverse; }
            set
            {
                this.RaiseAndSetIfChanged(ref _newCurrentUniverse, value);
                if (!universeLChanged)
                    RecalculateUniverseL();
            }
        }
        public int NewSubNet
        {
            get { return _newSubNet; }
            set
            {
                this.RaiseAndSetIfChanged(ref _newSubNet, value);
                if (!universeLChanged)
                    RecalculateUniverseL();
            }
        }
        public int NewNet
        {
            get { return _newNet; }
            set
            {
                this.RaiseAndSetIfChanged(ref _newNet, value);
                if (!universeLChanged)
                    RecalculateUniverseL();
            }
        }
        public string NetActivity
        {
            get { return _netActivity; }
            set { this.RaiseAndSetIfChanged(ref _netActivity, value); }
        }
        public bool ActivityDetected
        {
            get { return _dmxReceiveing; }
            set { this.RaiseAndSetIfChanged(ref _dmxReceiveing, value); }
        }

        public string CurrentIp
        {
            get { return _currentIp.ToString(); }
            set { this.RaiseAndSetIfChanged(ref _currentIp, IPAddress.Parse(value)); }
        }

        public double CurrentNet
        {
            get { return _currentNet; }
            set
            {
                this.RaiseAndSetIfChanged(ref _currentNet, value);
                SubNetsItems = SubNetsList[value];
                CurrentSubNet = SubNetsList[value][0];
            }
        }

        public SubNetControl CurrentSubNet
        {
            get { return _currentSubNet; }
            set
            {
                this.RaiseAndSetIfChanged(ref _currentSubNet, value);
                UniversesItems = UniversesList[_currentSubNet.SubNetId];
                CurrentUniverse = UniversesList[_currentSubNet.SubNetId][0];
            }
        }

        public UniverseControl CurrentUniverse
        {
            get { return _currentUniverse; }
            set { this.RaiseAndSetIfChanged(ref _currentUniverse, value); }
        }

        public Dictionary<int, List<UniverseControl>> UniversesList
        {
            get { return _universesList; }
            set { this.RaiseAndSetIfChanged(ref _universesList, value); }
        }

        public Dictionary<double, List<SubNetControl>> SubNetsList
        {
            get { return _subNetsList; }
            set { this.RaiseAndSetIfChanged(ref _subNetsList, value); }
        }

        public List<SubNetControl> SubNetsItems
        {
            get { return _subnetsItems; }
            set { this.RaiseAndSetIfChanged(ref _subnetsItems, value); }
        }

        public List<UniverseControl> UniversesItems
        {
            get { return _universesItems; }
            set { this.RaiseAndSetIfChanged(ref _universesItems, value); }
        }

        public UniverseControl SelectedSubNet { get; set; }

        public List<ArtNetDevice> ArtNetDevices
        {
            get { return _artnetDevices; }
            set { this.RaiseAndSetIfChanged(ref _artnetDevices, value); }
        }

        /* public ObservableCollection<DmxDataToList> DmxDataToLabels
         {
             get { return _dmxDataToList; }
             set { this.RaiseAndSetIfChanged(ref _dmxDataToList, value); }
         }*/
        public string ButtonContent
        {
            get { return _buttonContent; }
            set { this.RaiseAndSetIfChanged(ref _buttonContent, value); }
        }

        public bool SniffFlag
        {
            get { return _sniffFlag; }
            set { this.RaiseAndSetIfChanged(ref _sniffFlag, value); }
        }
        public ReactiveCommand<Unit, Unit> StartSniff { get; }
        public ReactiveCommand<Unit, Unit> StopSniff { get; }
        #endregion

        public MainWindowViewModel(Window _hostWindow)
        {
            hostWindow = _hostWindow;
            StartSniff = ReactiveCommand.Create(StartSniffering);
            StopSniff = ReactiveCommand.Create(StopSniffering);
            GetSelfIpAddress();
            UniversesList = new Dictionary<int, List<UniverseControl>>();
            GetUniversesList();
            SubNetsList = new Dictionary<double, List<SubNetControl>>();
            GetSubNetsList();
            SubNetsItems = SubNetsList[CurrentNet];
            CurrentSubNet = SubNetsItems[0];

            dmxDataPerUniverse = new Dictionary<short, byte[]>();
            // DmxDataToLabels = new ObservableCollection<DmxDataToList>();
            artNetIps = new List<string>();
            artSocket = new ArtNetSocket();
            artSocket.EnableBroadcast = true;
            artSocket.NewPacket += ArtSocket_NewPacket;
            artSocket.Open(IPAddress.Parse(CurrentIp), IPAddress.Broadcast/*.Parse("255.255.255.0")*/);
            SendArtPoll();
            pollTimer = new System.Timers.Timer();
            pollTimer.Interval = 10000;
            pollTimer.Elapsed += PollTimer_Elapsed;
            pollTimer.Start();

        }

        private void RecalculateNetSubNet(int universeL)
        {
            var arr = BitConverter.GetBytes(universeL);

            byte SubnetUniverse = arr[0];
            byte net = arr[1];
            NewCurrentUniverse = (short)(SubnetUniverse & 0xF);
            NewSubNet = (SubnetUniverse >> 4) & 0xF;
            NewNet = Convert.ToInt32(net.ToString());//net & 0xF;
            CurrentUniverse = UniversesList[NewSubNet].Where(w => w.UniverseIndex == NewCurrentUniverse).First();
        }

        private void RecalculateUniverseL()
        {
            NewCurrentUniverseL = NewCurrentUniverse + (NewSubNet << 4) + (NewNet << 8);
            CurrentUniverse = UniversesList[NewSubNet].Where(w => w.UniverseIndex == NewCurrentUniverse).First();
        }
        private void StartStopSnifferring()
        {
            object curButtonContent = new object();
            if (!SniffFlag)
            {
                App.Current.Resources.MergedDictionaries.FirstOrDefault(f => f.TryGetResource("ButtonStop", out curButtonContent));
                StartSniffering();
            }
            else
            {
                App.Current.Resources.MergedDictionaries.FirstOrDefault(f => f.TryGetResource("ButtonStart", out curButtonContent)); //ButtonContent = str.Resources.ButtonStart;
                StopSniffering();
            }
            ButtonContent = curButtonContent.ToString();
            
        }
        private void StartSniffering()
        {
            SniffFlag = true;
            StopSniffButton = true;
            StartSniffButton = false;
            mainSocket = new Socket(AddressFamily.InterNetwork, SocketType.Raw, ProtocolType.IP);
            mainSocket.Bind(new IPEndPoint(IPAddress.Parse(CurrentIp), 0));

            mainSocket.SetSocketOption(SocketOptionLevel.IP,  //Принимать только IP пакеты
                                       SocketOptionName.HeaderIncluded, //Включать заголовок
                                       true);

            byte[] byTrue = new byte[4] { 1, 0, 0, 0 };
            byte[] byOut = new byte[4] { 1, 0, 0, 0 };

            //Socket.IOControl это аналог метода WSAIoctl в Winsock 2
            mainSocket.IOControl(IOControlCode.ReceiveAll,  //SIO_RCVALL of Winsock
                                 byTrue, byOut);
            IPHostEntry HosyEntry = Dns.GetHostEntry((Dns.GetHostName()));
            //Начинаем приём асинхронный приём пакетов
            mainSocket.BeginReceive(byteData, 0, byteData.Length, SocketFlags.None,
                                    new AsyncCallback(OnReceive), null);

        }

        private void StopSniffering()
        {
            SniffFlag = false;
            StopSniffButton = false;
            StartSniffButton = true;
            //setSubNetActivity(-1);
            //setUniverseActivity(-1);
            mainSocket.Close();
            mainSocket.Dispose();
        }
        void OnReceive(IAsyncResult ar)
        {
            try
            {
                int nReceived = mainSocket.EndReceive(ar);
                ParseData(byteData, nReceived);
                byteData = new byte[8372];
                mainSocket.BeginReceive(byteData, 0, byteData.Length, SocketFlags.None,
                                new AsyncCallback(OnReceive), null);

            }
            catch (ObjectDisposedException)
            {
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message + " in " + ex.StackTrace);
                byteData = new byte[8372];
                mainSocket.BeginReceive(byteData, 0, byteData.Length, SocketFlags.None,
                               new AsyncCallback(OnReceive), null);
            }
        }

        void ParseData(byte[] byteData, int nReceived)
        {
            IPHeader ipHeader = new IPHeader(byteData, nReceived);
            if (ipHeader.ProtocolType == Protocol.UDP)
            {
                UDPHeader udpHeader = new UDPHeader(ipHeader.Data, (int)ipHeader.MessageLength);  //IPHeader.Data stores the data being carried by the IP datagram
                                                                                                  //Length of the data field   

                if (artNetIps.Contains(ipHeader.DestinationAddress.ToString()))
                {
                    ArtNetReceiveData artNetData = new ArtNetReceiveData() { DataLength = udpHeader.Data.Length, buffer = udpHeader.Data };
                    ArtNetDmxPacket artPack = (ArtNetDmxPacket)ArtNetPacket.Create(artNetData);
                    byte SubnetUniverse = artNetData.buffer[14];
                    byte net = artNetData.buffer[15];
                    var Universe = (short)(SubnetUniverse & 0xF);
                    var SubNet = (SubnetUniverse >> 4) & 0xF;
                    var Net = net & 0xF;
                    NewDmxPacket(artPack);
                }
                /*else
                {
                    ActivityDetected = false;
                    setSubNetActivity(-1);
                    setUniverseActivity(-1);
                }*/
            }
        }

        private void NewDmxPacket(ArtNetPacket e)
        {
            if (e.OpCode == Acn.ArtNet.ArtNetOpCodes.Dmx)
            {
                ActivityDetected = true;
                var packet = e as ArtNetDmxPacket;
                //SetActivity(packet.Net, packet.SubNet, packet.Universe);
                if (NewCurrentUniverseL > 0)
                {
                    if (packet.Universe != NewCurrentUniverseL)
                        return;
                    else
                        ProcessNewDmxPacket(packet);
                }
                else
                {
                    ProcessNewDmxPacket(packet);
                }
            }
        }

        private void ProcessNewDmxPacket(ArtNetDmxPacket packet)
        {
            byte[] dmxData;
            if (!dmxDataPerUniverse.TryGetValue(packet.Universe, out dmxData))
            {
                dmxData = new byte[512];
                dmxDataPerUniverse.Add(packet.Universe, dmxData);
            }

            if (!packet.DmxData.SequenceEqual(dmxData))
            {
                Debug.WriteLine("New DMX data for universe {0}", packet.Universe);

                var uc = UniversesList[NewSubNet].Where(w => w.UniverseIndex == NewCurrentUniverse).First();
                CurrentUniverse.DmxData = packet.DmxData;
                


                Debug.WriteLine("---=====---");

                if (dmxData.Length < packet.DmxData.Length)
                {
                    dmxData = new byte[packet.DmxData.Length];
                    dmxDataPerUniverse[packet.Universe] = dmxData;
                }

                Array.Copy(packet.DmxData, dmxData, dmxData.Length);

                //UniverseActivityDetected.Invoke(this, new UniverseActivityDetectedArgs(packet.Universe));
            }
        }

        private void SetActivity(int net, int subnet, int universe)
        {
            if ((int)CurrentNet == net)
            {
                setSubNetActivity(subnet);
                if (CurrentSubNet.SubNetId == subnet)
                    setUniverseActivity(universe);
            }
            else
            {
                NetActivity = str.Resources.LabelNetActivity + " " + net;
            }
        }
        private void setSubNetActivity(int subNet)
        {
            if (subNet >= 0)
            {
                var sn = SubNetsList[CurrentNet].Where(w => w.SubNetId == subNet).First();
                if (ActivityDetected)
                {
                    Dispatcher.UIThread.InvokeAsync((Action)delegate ()
                    {
                        sn.ActivityDetected = true;
                    }, DispatcherPriority.ContextIdle);
                }
                else
                {
                    Dispatcher.UIThread.InvokeAsync((Action)delegate ()
                    {
                        sn.ActivityDetected = false;
                    }, DispatcherPriority.ContextIdle);
                }
            }
            else
            {
                foreach (var sn in SubNetsList[CurrentNet])
                {
                    Dispatcher.UIThread.InvokeAsync((Action)delegate ()
                    {
                        sn.ActivityDetected = false;
                    }, DispatcherPriority.ContextIdle);
                }
            }
        }

        private void setUniverseActivity(int universe)
        {
            if (universe >= 0)
            {
                var uc = UniversesItems.Where(w => w.UniverseIndex == universe).First();
                if (ActivityDetected)
                {
                    Dispatcher.UIThread.InvokeAsync((Action)delegate ()
                    {
                        uc.ActivityDetected = true;
                    }, DispatcherPriority.ContextIdle);
                }
                else
                {
                    Dispatcher.UIThread.InvokeAsync((Action)delegate ()
                    {
                        uc.ActivityDetected = false;
                    }, DispatcherPriority.ContextIdle);
                }
            }
            else
            {
                foreach (var uc in UniversesItems)
                {
                    Dispatcher.UIThread.InvokeAsync((Action)delegate ()
                    {
                        uc.ActivityDetected = false;
                    }, DispatcherPriority.ContextIdle);
                }
            }
        }
        private void GetSelfIpAddress()
        {
            var nics = NetworkInterface.GetAllNetworkInterfaces();
            foreach (var nic in nics)
            {
                if (nic.NetworkInterfaceType != NetworkInterfaceType.Loopback)
                {
                    IPInterfaceProperties ipProperties = nic.GetIPProperties();

                    for (int n = 0; n < ipProperties.UnicastAddresses.Count; n++)
                    {
                        if (ipProperties.UnicastAddresses[n].Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                        {
                            _currentIp = nic.GetIPProperties().UnicastAddresses[n].Address;
                        }
                    }
                    break;
                }
            }
        }

        private void GetUniversesList()
        {
            var ucc = new Dictionary<int, List<UniverseControl>>();
            for (int x = 0; x < 16; x++)
            {
                var list = new List<UniverseControl>();
                for (int i = 0; i < 16; i++)
                {
                    var uc = new UniverseControl(i);
                    list.Add(uc);

                }
                ucc.Add(x, list);
            }

            UniversesList = ucc;

        }
        private void GetSubNetsList()
        {
            for (int s = 0; s < 128; s++)
            {
                var list = new List<SubNetControl>();
                for (int i = 0; i < 16; i++)
                {
                    var snc = new SubNetControl() { SubNetId = i };
                    list.Add(snc);

                }
                SubNetsList.Add(s, list);
            }
        }

        private void ArtSocket_NewPacket(object? sender, Acn.Sockets.NewPacketEventArgs<ArtNetPacket> e)
        {
            if (e.Packet.OpCode == Acn.ArtNet.ArtNetOpCodes.PollReply)
            {
                var packet = e.Packet as ArtPollReplyPacket;
                var artnetDevices = new List<ArtNetDevice>();
                ArtNetDevice artNetDevice = new ArtNetDevice(e.Source.Address.GetAddressBytes(), packet.Port)
                {
                    IpAddressStr = e.Source.Address.ToString(),
                    EstaCode = packet.EstaCode,
                    FirmwareVersion = packet.FirmwareVersion,
                    GoodInput = packet.GoodInput,
                    GoodOutput = packet.GoodOutput,
                    LongName = packet.LongName,
                    MacAddress = packet.MacAddress,
                    NodeReport = packet.NodeReport,
                    Oem = packet.Oem,
                    Status = packet.Status,
                    Status2 = packet.Status2,
                    Style = packet.Style,
                    SwMacro = packet.SwMacro,
                    SwRemote = packet.SwRemote,
                    SwVideo = packet.SwVideo,
                    UbeaVersion = packet.UbeaVersion,
                };
                artnetDevices.Add(artNetDevice);
                artNetIps.Add(e.Source.Address.ToString());
                ArtNetDevices = artnetDevices;
            }
        }

        private void PollTimer_Elapsed(object? sender, System.Timers.ElapsedEventArgs e)
        {
            SendArtPoll();
        }
        private void SendArtPoll()
        {
            var pollPacket = new ArtPollPacket();
            pollPacket.TalkToMe = 6;
            artSocket.Send(pollPacket);
        }
        public void DmxDataFromHexToDec()
        {
            var data = CurrentUniverse.DmxData;
            CurrentUniverse.HexOrDecSelected = true;
            CurrentUniverse.DmxData = data;
        }

        public void DmxDataFromDecToHex()
        {
            var data = CurrentUniverse.DmxData;
            CurrentUniverse.HexOrDecSelected = false;
            CurrentUniverse.DmxData = data;
        }
    }
}