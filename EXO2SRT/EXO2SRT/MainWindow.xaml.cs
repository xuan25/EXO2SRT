using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace EXO2SRT
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        Thread convertingThread;
        private delegate void DelFinished();
        private event DelFinished Finished;

        private void StartProcess(string path)
        {
            InfoBox.Text = "转换中...";
            this.AllowDrop = false;
            Finished -= MainWindow_Finished;
            Finished += MainWindow_Finished;

            convertingThread = new Thread(delegate ()
            {
                StreamReader streamReader = new StreamReader(path, Encoding.GetEncoding("GBK"));
                string content = streamReader.ReadToEnd();
                streamReader.Close();

                Match exoInfoMatch = Regex.Match(content, @"\[exedit\]\r\nwidth=[0-9]+\r\nheight=[0-9]+\r\nrate=(?<FrameRate>[0-9]+)\r\nscale=[0-9]+\r\nlength=[0-9]+\r\naudio_rate=[0-9]+\r\naudio_ch=[0-9]+\r\n");
                TextItem.FrameRate = int.Parse(exoInfoMatch.Groups["FrameRate"].Value);

                MatchCollection textItemMatchCollection = Regex.Matches(content, @"\[[0-9]+\]\r\nstart=(?<StartFrame>[0-9]+)\r\nend=(?<EndFrame>[0-9]+)\r\nlayer=[0-9]+\r\noverlay=[0-9]+\r\ncamera=[0-9]+\r\n\[[0-9]+\.0\]\r\n_name=文本[\s\S]+?text=(?<Code>[0-9a-z]+)\r\n\[[0-9]+.1\]\r\n_name=标准变换[\s\S]+?blend=[0-9]+");
                List<TextItem> textItemlist = new List<TextItem>();
                foreach (Match m in textItemMatchCollection)
                {
                    long startFrame = long.Parse(m.Groups["StartFrame"].Value);
                    long endFrame = long.Parse(m.Groups["EndFrame"].Value);
                    string code = m.Groups["Code"].Value;
                    textItemlist.Add(new TextItem(startFrame, endFrame, code));
                }

                textItemlist = textItemlist.OrderBy(TextItem => TextItem.StartFrame).ToList();

                string srt = "";
                for (int i = 0; i < textItemlist.Count; i++)
                {
                    srt += string.Format("{0}\r\n{1} --> {2}\r\n{3}\r\n\r\n", i + 1, textItemlist[i].GetStartTime(), textItemlist[i].GetEndTime(), textItemlist[i].GetText());
                }

                StreamWriter streamWriter = new StreamWriter(path + ".srt");
                streamWriter.Write(srt);
                streamWriter.Close();
                Finished();
            });
            convertingThread.Start();
            
        }

        private void MainWindow_Finished()
        {
            Dispatcher.Invoke(new Action(() =>
            {
                InfoBox.Text = "转换完成!";
                this.AllowDrop = true;
            }));
        }

        private void Window_PreviewDragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Move;
                e.Handled = true;
            }
            else
            {
                e.Effects = DragDropEffects.None;
                e.Handled = false;
            }
        }

        private void Window_PreviewDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetData(DataFormats.FileDrop) == null)
                return;
            e.Handled = true;
            string filepath = ((Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();

            StartProcess(filepath);
        }
    }

    public class TextItem
    {
        public long StartFrame { get; internal set; }
        public long EndFrame{ get; internal set; }
        public string Code { get; internal set; }
        public static int FrameRate { get; set; }

        public TextItem(long startFrame, long endFrame, string code)
        {
            StartFrame = startFrame;
            EndFrame = endFrame;
            Code = code;
        }

        public string GetText()
        {
            byte[] byteArray = new byte[Code.Length/2];
            for(int i=0; i<Code.Length/2; i++)
            {
                string byteStr = "0x" + Code[2 * i] + Code[2 * i + 1];
                byteArray[i] = Convert.ToByte(byteStr, 16);
            }
            string str = Encoding.GetEncoding("UTF-16").GetString(byteArray);
            str = str.Substring(0, str.IndexOf('\0'));
            return str;
        }

        public string GetStartTime()
        {
            return FrameToTimeStamp(StartFrame);
        }

        public string GetEndTime()
        {
            return FrameToTimeStamp(EndFrame);
        }

        private string FrameToTimeStamp(long frame)
        {
            double sec = (double)frame / FrameRate;
            int h = (int)Math.Floor(sec / 3600);
            int m = (int)Math.Floor(sec % 3600 / 60);
            int s = (int)Math.Floor(sec % 3600 % 60);
            double ms = (sec % 3600 % 60) - s;

            string timeStamp = string.Format("{0}:{1}:{2},{3}", h.ToString("D2"), m.ToString("D2"), s.ToString("D2"), (ms * 1000).ToString("F0"));
            return timeStamp;

        }
    }
}
