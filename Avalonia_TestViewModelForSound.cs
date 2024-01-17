using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reactive;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Threading;
using CSCore;
using CSCore.SoundIn;
using CSCore.Streams;
using CSCore.Streams.Effects;
using NWaves.Transforms;
using ReactiveUI;
using SoundIOSharp;
using SoundTest.Models;
using SoundTest.Views;
using Timer = System.Timers;

namespace SoundTest.ViewModels;

public class MainWindowViewModel : ViewModelBase
{
    private SoundIO soundApi;
    private List<SoundIODevice> _soundInputs;
    private List<SoundIODevice> _soundOutputs;
    private SoundIODevice _selectedSoundInput;
    private SoundIODevice _selectedSoundOutput;
    private double microphone_latency = 0.2; // seconds
    private string outfile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "SoundTest.wav");
    private bool _isCapturing;
    private bool _isFFTNeeded;
    private bool _fftChartVisibility = false;
    private bool _microChartVisibility = true;
    private bool _hpFftChartVisibility = false;
    private bool _hpChartVisibility = false;
    private bool _isInputCapturing = true;
    private bool _isOutputCapturing = false;
    private bool _isInputOrOutput = false;
    private bool _isNotWin = false;
    private double[] _readBytes;
    private double[] _microBytes;
    private double[] _fftBytes;
    private double[] _hpBytes;
    private double[] _hpFftBytes;
    private PitchShifter _pitchShifter;
    private SoundIOInStream instream;
    private SoundIOOutStream outstream;
    private Timer.Timer timer_micro_chart;
    private Timer.Timer timer_fft_chart;
    private Timer.Timer timer_hp_chart;
    private Timer.Timer timer_hpFft_chart;
    private IWaveSource _source;
    private WasapiLoopbackCapture loopback;
    public MicroChartViewModel chartVM { get; set; }
    static SoundIORingBuffer ring_buffer = null;
    private MainWindow _hostWindow;
    static SoundIOFormat [] prioritized_formats = {
        SoundIODevice.Float32NE,
        SoundIODevice.Float32FE,
        SoundIODevice.S32NE,
        SoundIODevice.S32FE,
        SoundIODevice.S24NE,
        SoundIODevice.S24FE,
        SoundIODevice.S16NE,
        SoundIODevice.S16FE,
        SoundIODevice.Float64NE,
        SoundIODevice.Float64FE,
        SoundIODevice.U32NE,
        SoundIODevice.U32FE,
        SoundIODevice.U24NE,
        SoundIODevice.U24FE,
        SoundIODevice.U16NE,
        SoundIODevice.U16FE,
        SoundIOFormat.S8,
        //SoundIOFormat.U8,
        SoundIOFormat.Invalid,
    };
    static readonly int [] prioritized_sample_rates = {
        48000,
        44100,
        96000,
        24000,
        0,
    };

    public bool IsNotWin
    {
        get { return _isNotWin; }
        set { this.RaiseAndSetIfChanged(ref _isNotWin, value); }
    }
    public double[] ReadBytes
    {
        get { return _readBytes;}
        set { this.RaiseAndSetIfChanged(ref _readBytes, value); }
    }
    public double[] MicroBytes
    {
        get { return _microBytes;}
        set { this.RaiseAndSetIfChanged(ref _microBytes, value); }
    }
    public double[] FftBytes
    {
        get { return _fftBytes;}
        set { this.RaiseAndSetIfChanged(ref _fftBytes, value); }
    }
    public double[] HpBytes
    {
        get { return _hpBytes;}
        set { this.RaiseAndSetIfChanged(ref _hpBytes, value); }
    }
    public double[] HpFftBytes
    {
        get { return _hpFftBytes;}
        set { this.RaiseAndSetIfChanged(ref _hpFftBytes, value); }
    }
    public double MicrophoneLatency
    {
        get { return microphone_latency; }
        set { this.RaiseAndSetIfChanged(ref microphone_latency, value); }
    }
    public List<SoundIODevice> SoundInputs
    {
        get { return _soundInputs; }
        set { this.RaiseAndSetIfChanged(ref _soundInputs, value); }
    }
    public SoundIODevice SelectedSoundOutput
    {
        get{ return _selectedSoundOutput; }
        set
        {
            this.RaiseAndSetIfChanged(ref _selectedSoundOutput, value);
            OnOutDeviceChanged();
        }
    }
    public List<SoundIODevice> SoundOutputs
    {
        get { return _soundOutputs; }
        set { this.RaiseAndSetIfChanged(ref _soundOutputs, value); }
    }
    public SoundIODevice SelectedSoundInput
    {
        get{ return _selectedSoundInput; }
        set
        {
            this.RaiseAndSetIfChanged(ref _selectedSoundInput, value);
            OnDeviceChanged();
        }
    }

    public bool IsCapturing
    {
        get { return _isCapturing; }
        set { this.RaiseAndSetIfChanged(ref _isCapturing, value); }
    }
    public bool IsFFTNeeded
    {
        get { return _isFFTNeeded; }
        set 
        { 
            this.RaiseAndSetIfChanged(ref _isFFTNeeded, value);
            if(!value)
                ClearCharts();
            var flag = false;
            if (IsCapturing)
            {
                IsCapturing = false;
                flag = true;
            }

            if (!IsInputOrOutput)
            {
                FftChartVisibility = !FftChartVisibility;
                if (FftChartVisibility)
                {
                    timer_micro_chart.Enabled = false;
                    timer_fft_chart.Enabled = true;
                }
                else
                {
                    timer_micro_chart.Enabled = true;
                    timer_fft_chart.Enabled = false;
                }

                MicroChartVisibility = !MicroChartVisibility;
                if (HPChartVisibility)
                {
                    timer_fft_chart.Enabled = false;
                    timer_micro_chart.Enabled = true;
                }
                else
                {
                    timer_fft_chart.Enabled = true;
                    timer_micro_chart.Enabled = false;
                }
            }
            else
            {
                HpFftChartVisibility = !HpFftChartVisibility;
                if (HpFftChartVisibility)
                {
                    timer_hp_chart.Enabled = false;
                    timer_hpFft_chart.Enabled = true;
                }
                else
                {
                    timer_hp_chart.Enabled = true;
                    timer_hpFft_chart.Enabled = false;
                }

                HPChartVisibility = !HPChartVisibility;
                if (HPChartVisibility)
                {
                    timer_hpFft_chart.Enabled = false;
                    timer_hp_chart.Enabled = true;
                }
                else
                {
                    timer_hpFft_chart.Enabled = true;
                    timer_hp_chart.Enabled = false;
                }
            }

            if (flag)
            {
                IsCapturing = true;
                flag = false;
            }
        }
    }
    public bool FftChartVisibility
    {
        get { return _fftChartVisibility; }
        set { this.RaiseAndSetIfChanged(ref _fftChartVisibility, value); }
    }

    public bool MicroChartVisibility
    {
        get { return _microChartVisibility; }
        set { this.RaiseAndSetIfChanged(ref _microChartVisibility, value); }
    }
    public bool HpFftChartVisibility
    {
        get { return _hpFftChartVisibility; }
        set { this.RaiseAndSetIfChanged(ref _hpFftChartVisibility, value); }
    }

    public bool HPChartVisibility
    {
        get { return _hpChartVisibility; }
        set { this.RaiseAndSetIfChanged(ref _hpChartVisibility, value); }
    }
    public bool IsInputCapturing
    {
        get { return _isInputCapturing; }
        set { this.RaiseAndSetIfChanged(ref _isInputCapturing, value); }
    }

    public bool IsOutputCapturing
    {
        get { return _isOutputCapturing; }
        set { this.RaiseAndSetIfChanged(ref _isOutputCapturing, value); }
    }

    public bool IsInputOrOutput
    {
        get { return _isInputOrOutput; }
        set
        {
            this.RaiseAndSetIfChanged(ref _isInputOrOutput, value);
            if (value)
            {
                if (MicroChartVisibility)
                    MicroChartVisibility = false;
                if (FftChartVisibility)
                    FftChartVisibility = false;
                if (IsFFTNeeded)
                {
                    HPChartVisibility = false;
                    HpFftChartVisibility = true;
                }
                else
                {
                    HPChartVisibility = true;
                    HpFftChartVisibility = false;
                }

                OnOutDeviceChanged();
            }
            else 
            {
                if (HPChartVisibility)
                    HPChartVisibility = false;
                if (HpFftChartVisibility)
                    HpFftChartVisibility = false;
                if (IsFFTNeeded)
                {
                    MicroChartVisibility = false;
                    FftChartVisibility = true;
                }
                else
                {
                    MicroChartVisibility = true;
                    FftChartVisibility = false;
                }
            }
            /*IsInputCapturing = !IsInputCapturing;
            IsOutputCapturing = !IsOutputCapturing;*/
        }
    }
    public ReactiveCommand<Unit,Unit> StartCapturingCommand { get; }
    public MainWindowViewModel() { }
    public MainWindowViewModel(Window hostWindow)
    {
        _hostWindow = hostWindow as MainWindow;
        chartVM = new MicroChartViewModel(2048);
        soundApi = new SoundIO();
        soundApi.Connect();
        soundApi.FlushEvents();
        
        timer_micro_chart = new Timer.Timer();
        timer_micro_chart.Interval = 200;
        timer_micro_chart.Elapsed += Timer_micro_chartOnElapsed;

        timer_fft_chart = new Timer.Timer();
        timer_fft_chart.Interval = 200;
        timer_fft_chart.Elapsed += Timer_fft_chartOnElapsed;
        
        timer_hp_chart = new Timer.Timer();
        timer_hp_chart.Interval = 200;
        timer_hp_chart.Elapsed += Timer_hp_chartOnElapsed;

        timer_hpFft_chart = new Timer.Timer();
        timer_hpFft_chart.Interval = 200;
        timer_hpFft_chart.Elapsed += Timer_hpFft_chartOnElapsed;
        
        SoundInputs = new List<SoundIODevice>();
        SoundOutputs = new List<SoundIODevice>();
        GetSoundInputs();
        GetSoundOutputs();
        StartCapturingCommand = ReactiveCommand.Create(OnSelectingDevice);
    }
    private bool aaa;
    private bool bbb;
    private void Timer_hpFft_chartOnElapsed(object? sender, Timer.ElapsedEventArgs e)
    {
        if (!bbb)
        {
            bbb = true;
            double[] paddedAudio = FftSharp.Pad.ZeroPad(ReadBytes);
            double[] fftMag = FftSharp.Transform.FFTmagnitude(paddedAudio);
            //Array.Copy(fftMag, HpFftBytes, fftMag.Length);
            for (int i = 0, j = 0; i < fftMag.Length; i += 100, j++)
            {
                var arr = fftMag.Skip(i).Take(100);
                HpFftBytes[j] = arr.Max();
            }
            _hostWindow._hpFftChart.Refresh();
            bbb = false;
        }
    }

    private void Timer_hp_chartOnElapsed(object? sender, Timer.ElapsedEventArgs e)
    {
        Array.Copy(ReadBytes, HpBytes, ReadBytes.Length);
        _hostWindow._hpChart.Refresh();
    }

    private void Timer_micro_chartOnElapsed(object? sender, Timer.ElapsedEventArgs e)
    {
        Array.Copy(ReadBytes, MicroBytes, ReadBytes.Length);
        _hostWindow._microChart.Refresh();
    }
    private void Timer_fft_chartOnElapsed(object? sender, Timer.ElapsedEventArgs e)
    {
        if (!aaa)
        {
            aaa = true;
            double[] paddedAudio = FftSharp.Pad.ZeroPad(ReadBytes);
            double[] fftMag = FftSharp.Transform.FFTmagnitude(paddedAudio);
            Array.Copy(fftMag, FftBytes, fftMag.Length);
            
            _hostWindow._fftChart.Refresh();
            aaa = false;
        }
    }

    private void GetSoundInputs()
    {
        for (int i = 0; i < soundApi.InputDeviceCount; i++)
        {
            var device = soundApi.GetInputDevice(i);
            if (device.IsRaw & !SoundInputs.Any(a=>a.Name == device.Name))
            {
                SoundInputs.Add(device);
            }
            else if(!device.IsRaw & device.Name.Contains("CABLE"))
                SoundInputs.Add(device);
        }
        
        
    }

    private void GetSoundOutputs()
    {
        for (int i = 0; i < soundApi.OutputDeviceCount; i++)
        {
            var device = soundApi.GetOutputDevice(i);
            if(device.IsRaw)
                SoundOutputs.Add(device);
        }
    }
   
    private void OnDeviceChanged()
    {
        if (SelectedSoundInput != null)
        {
            var sample_rate = prioritized_sample_rates.First(sr => SelectedSoundInput.SupportsSampleRate(sr));
            chartVM.Rate = sample_rate / 1000D;
            var fmt = SelectedSoundInput.Formats.First(); 
                // SoundIOFormat.S16LE;prioritized_formats.First(f => SelectedSoundInput.SupportsFormat(f));
            if(SelectedSoundInput.Name.Contains("CABLE"))
                fmt = SelectedSoundInput.Formats.First(f=>f == SoundIOFormat.S16LE);
            instream = SelectedSoundInput.CreateInStream();
            instream.Format = fmt;
            instream.SampleRate = sample_rate;
            instream.ReadCallback = (fmin, fmax) => read_callback(instream, fmin, fmax);

            instream.Open();
            chartVM.BytesPerSample = instream.BytesPerSample;
            const int ring_buffer_duration_seconds = 1;
            int capacity = (int)(ring_buffer_duration_seconds * instream.SampleRate * instream.BytesPerFrame);
            int finalCapacity = 0;
            if (IsPowerOfTwo(capacity))
                finalCapacity = capacity;
            else
            {
                finalCapacity = (int)Math.Pow(2, GetMinPowerOfTwoLargerThan(capacity));
            }
            ReadBytes = new double[finalCapacity];
            MicroBytes = new double[finalCapacity];
            double[] paddedAudio = FftSharp.Pad.ZeroPad(ReadBytes);
            double[] fftMag = FftSharp.Transform.FFTpower(paddedAudio);
            FftBytes = new double[fftMag.Length];
            double fftPeriod = FftSharp.Transform.FFTfreqPeriod(sample_rate, fftMag.Length);
            var chartRate = sample_rate / 10000D;
            _hostWindow._microChart.Plot.AddSignal(MicroBytes, chartRate);
            _hostWindow._fftChart.Plot.AddSignal(FftBytes, 1.0 / fftPeriod);
            _hostWindow._microChart.RefreshRequest();
            ring_buffer = soundApi.CreateRingBuffer (capacity);
            var buf = ring_buffer.WritePointer;
        }
    }

    private void OnOutDeviceChanged()
    {
        var sample_rate = 44100;//prioritized_sample_rates.First(sr => SelectedSoundOutput.SupportsSampleRate(sr));
        var fmt = new WaveFormat();
        loopback = new WasapiLoopbackCapture(500, fmt);
        loopback.Initialize();
        var soundInSource = new SoundInSource(loopback);
        ISampleSource source = soundInSource.ToSampleSource().AppendSource(x => new PitchShifter(x), out _pitchShifter);

        //SetupSampleSource(source);
        var notificationSource = new SingleBlockNotificationStream(source);
        //pass the intercepted samples as input data to the spectrumprovider (which will calculate a fft based on them)
        //notificationSource.SingleBlockRead += (s, a) => spectrumProvider.Add(a.Left, a.Right);

        _source = notificationSource.ToWaveSource(16);
        // We need to read from our source otherwise SingleBlockRead is never called and our spectrum provider is not populated
        int capacity = _source.WaveFormat.BytesPerSecond / 2;
        int finalCapacity = 0;
        if (IsPowerOfTwo(capacity))
            finalCapacity = capacity;
        else
        {
            finalCapacity = (int)Math.Pow(2, GetMinPowerOfTwoLargerThan(capacity));
        }
        var chartRate = sample_rate / 10000D;
        byte[] buffer = new byte[finalCapacity];
        ReadBytes = new double[finalCapacity];
        HpBytes = new double[finalCapacity];

        _hostWindow._hpChart.Plot.AddSignal(HpBytes, chartRate);
        double[] paddedAudio = FftSharp.Pad.ZeroPad(ReadBytes);
        double[] fftMag = FftSharp.Transform.FFTpower(paddedAudio);
        HpFftBytes = new double[fftMag.Length];
        double fftPeriod = FftSharp.Transform.FFTfreqPeriod(sample_rate, fftMag.Length);
        //_hostWindow._hpFftChart.Plot.AddSignal(HpFftBytes, fftPeriod);
        var positions = new double[fftMag.Length];
        for (int i = 0, j = 0; i < fftMag.Length; i += 10, j++)
        {
            positions[j] = i;
        }
        var bar = _hostWindow._hpFftChart.Plot.AddBar(HpFftBytes, positions);
        bar.BarWidth = 10D;
        soundInSource.DataAvailable += (s, aEvent) =>
        {
            Thread.Sleep(200);
            int read;
            var AudioDevice = s as SoundInSource;
            var arr = new byte [finalCapacity];
            while ((read = _source.Read(buffer, 0, buffer.Length)) > 0)
            {
                Array.Copy(buffer, arr, buffer.Length);
                int bytesPerSamplePerChannel = AudioDevice.WaveFormat.BitsPerSample / 8;
                int bytesPerSample = bytesPerSamplePerChannel * AudioDevice.WaveFormat.Channels;
                int bufferSampleCount = aEvent.Data.Length / bytesPerSample;
                for (int i = 0; i < bufferSampleCount; i++)
                {
                    ReadBytes[i] = BitConverter.ToInt16(arr, i * bytesPerSample); // >> 8;
                }

                if (!IsCapturing)
                {
                    loopback.Stop();
                    break;
                }
            };
        };
    }

    private void ClearCharts()
    {
        _hostWindow._hpFftChart.Plot.Clear();
        _hostWindow._microChart.Plot.Clear();
        _hostWindow._hpChart.Plot.Clear();
        _hostWindow._fftChart.Plot.Clear();
    }
    private bool IsPowerOfTwo(int x)
    {
        return ((x & (x - 1)) == 0) && (x > 0);
    }
    private double GetMinPowerOfTwoLargerThan(int number)
    {            
        if (number < 0)
            return 1;
 
        int power = (int)Math.Ceiling(Math.Log(number) / Math.Log(2));
        int tempResult = 1 << power;

        return power; //(tempResult == number) ? tempResult << 1 : tempResult;
    }
    private async void OnSelectingDevice()
    {
        if (!IsInputOrOutput)
        {
            await Task.Run(() =>
            {
                while (IsCapturing)
                {
                    const int ring_buffer_duration_seconds = 1;
                    int finalCapacity = 0;
                    int capacity = (int)(ring_buffer_duration_seconds * instream.SampleRate * instream.BytesPerFrame);
                    if (IsPowerOfTwo(capacity))
                        finalCapacity = capacity;
                    else
                    {
                        finalCapacity = (int)Math.Pow(2, GetMinPowerOfTwoLargerThan(capacity));
                    }

                    instream.Start();
                    if (!FftChartVisibility)
                        timer_micro_chart.Enabled = true;
                    else
                        timer_fft_chart.Enabled = true;
                    var arr = new byte [finalCapacity];
                    unsafe
                    {
                        fixed (void* arrptr = arr)
                        {

                            soundApi.FlushEvents();
                            Thread.Sleep(200);
                            int fill_bytes = ring_buffer.FillCount;
                            var read_buf = ring_buffer.ReadPointer;

                            Buffer.MemoryCopy((void*)read_buf, arrptr, fill_bytes, fill_bytes);
                            //chartVM.Arr = arr;
                            if (arr.Max() == 0)
                                continue;
                            //Thread.Sleep(200);
                            double aa = 0;
                            switch (instream.Format)
                            {
                                case SoundIOFormat.S16LE:
                                    int bufferSampleCount = arr.Length / instream.BytesPerSample;
                                    for (int i = 0; i < bufferSampleCount; i++)
                                    {
                                
                                        ReadBytes[i] = BitConverter.ToInt16(arr, i * instream.BytesPerSample); // >> 8;
                                        aa = Math.Max(aa, ReadBytes[i]);
                                    }
                                    break;
                                    case SoundIOFormat.U8:
                                        int bytesPerSamplePerChannel = instream.BytesPerSample;
                                        int bytesPerSample = bytesPerSamplePerChannel * SelectedSoundInput.Layouts.First(w=>w.ChannelCount == 2).ChannelCount;
                                        int bufferSampleCountU8 = arr.Length / bytesPerSample;
                                        for (int i = 0; i < bufferSampleCountU8; i++)
                                        {
                                            ReadBytes[i] = Convert.ToInt16(arr[i]);
                                            int point = 0;
                                        }

                                        break;
                                
                                
                            }
                            ring_buffer.AdvanceReadPointer(fill_bytes);
                            if (ring_buffer.ReadPointer == IntPtr.Zero)
                            {
                                IsCapturing = false;
                                break;
                            }
                        }
                    }
                }

                Thread.Sleep(1000);
                chartVM.Rate = 1D;
                chartVM.Arr = new byte[2048];
                if (!FftChartVisibility)
                    timer_micro_chart.Enabled = false;
                else
                    timer_fft_chart.Enabled = false;
                instream.Dispose();
            });
        }
        else
        {
            await Task.Run(() =>
            {
                loopback.Start();
                if (IsFFTNeeded)
                    timer_hpFft_chart.Enabled = true;
                else
                {
                    timer_hp_chart.Enabled = true;
                }
            });
        }
    }

    
    private void read_callback (SoundIOInStream instream, int frame_count_min, int frame_count_max)
    {
        
        var write_ptr = ring_buffer.WritePointer;
        if (write_ptr == IntPtr.Zero)
            return;
        int free_bytes = ring_buffer.FreeCount;
        int free_count = free_bytes / instream.BytesPerFrame;

        if (frame_count_min > free_count)
            throw new InvalidOperationException ("ring buffer overflow"); // panic()

        int write_frames = Math.Min (free_count, frame_count_max);
        int frames_left = write_frames;

        for (; ; ) {
            int frame_count = frames_left;

            var areas = instream.BeginRead (ref frame_count);

            if (frame_count == 0)
                break;

            if (areas.IsEmpty) {
                // Due to an overflow there is a hole. Fill the ring buffer with
                // silence for the size of the hole.
                for (int i = 0; i < frame_count * instream.BytesPerFrame; i++)
                    Marshal.WriteByte (write_ptr + i, 0);
                Debug.WriteLine ("Dropped {0} frames due to internal overflow", frame_count);
            } else {
                for (int frame = 0; frame < frame_count; frame += 1) {
                    int chCount = instream.Layout.ChannelCount;
                    int copySize = instream.BytesPerSample;
                    unsafe {
                        for (int ch = 0; ch < chCount; ch += 1) {
                            var area = areas.GetArea (ch);
                            Buffer.MemoryCopy ((void*)area.Pointer, (void*)write_ptr, copySize, copySize);
                            area.Pointer += area.Step;
                            write_ptr += copySize;
                        }
                    }
                }
            }

            instream.EndRead ();

            frames_left -= frame_count;
            if (frames_left <= 0)
                break;
        }

        int advance_bytes = write_frames * instream.BytesPerFrame;
        ring_buffer.AdvanceWritePointer (advance_bytes);

    }

    
    private void CheckFFTOutput(byte[] readBytes)
    {
        if (IsFFTNeeded)
        {
            var fft = new Fft(8);
            var newArray = new float[readBytes.Length];
            float[] arr = new float[newArray.Length];
            for (int i=0;i<readBytes.Length;i++)
            {
                var item = readBytes[i];
                newArray.SetValue((float)item, i);
            }

            
            //fft.Direct(newArray, arr);
            
            Debug.WriteLine($"Volume: {arr[0],5} {arr[1],5} {arr[2],5} {arr[3],5}");
        }
    }
}