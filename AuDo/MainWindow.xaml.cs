using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using System.Windows.Controls;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using Microsoft.Win32;
using System.Net;
using System.Management;

namespace AuDo
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            File.Delete(@"kbmlog.ads");
        }

        private void teachMe_Click(object sender, RoutedEventArgs e)
        {
            stop.IsEnabled = true;
            teachMe.IsEnabled = false;
            this.WindowState = WindowState.Minimized;
            MouseHook.Start();
            KBHook.Start();
        }

        private void stop_Click(object sender, RoutedEventArgs e)
        {
            stop.IsEnabled = false;
            teachMe.IsEnabled = true;
            MouseHook.Stop();
            KBHook.Stop();
        }

        private void _do_Click(object sender, RoutedEventArgs e)
        {
            if (!File.Exists(@"kbmlog.ads"))
            {
                MessageBox.Show("Sorry I haven't learnt anything");
                return;
            }
            MessageBox.Show("The process will start exactly after 3 seconds once you click OK button\r\nIf you need to full screen anything, do it in 3 seconds after clicking OK\r\nOnce started DO NOT use mouse or keyborad unless absolutely necessary\r\nThis window will minimize while the task is running and will come back normal once task completes\r\nRestore the software by clicking on the minimized window on taskbar for emergency stop", "Note", MessageBoxButton.OK, MessageBoxImage.Information);
            Thread.Sleep(3000);
            teachMe.IsEnabled = _do.IsEnabled = stop.IsEnabled = false;
            this.WindowState = WindowState.Minimized;
            int reptscr = 1;//For changing scroll value in each iteration, to invalidate just don't increase the value
            int reptdr = 1;//For increasing drag at each iteration
            for (int i = 0; i < Int32.Parse(times.Text); i++)
            {
                if (this.WindowState == WindowState.Normal)
                    break;
                string previousLine = "";
                string scriptLine;
                StreamReader scriptRead = new StreamReader(@"kbmlog.ads");
                while ((scriptLine = scriptRead.ReadLine()) != null)
                {
                    if (this.WindowState == WindowState.Normal)
                        break;
                    if (scriptLine.StartsWith("LD"))//Left button down
                    {
                        MatchCollection matches = Regex.Matches(scriptLine, "[0-9]+");
                        List<string> coordinates = new List<string>();
                        coordinates.Clear();
                        foreach (Match march in matches)
                        {
                            coordinates.Add(march.Value);
                        }
                        System.Windows.Forms.Cursor.Position = new System.Drawing.Point(Int32.Parse(coordinates[0]), Int32.Parse(coordinates[1]));
                        if (previousLine != scriptLine)
                            Thread.Sleep(100);
                        MouseHandler.LeftDown();
                        Thread.Sleep(100);
                    }
                    else if (scriptLine.StartsWith("LU"))//Left button up
                    {
                        MatchCollection matches = Regex.Matches(scriptLine, "[0-9]+");
                        List<string> coordinates = new List<string>();
                        coordinates.Clear();
                        foreach (Match march in matches)
                        {
                            coordinates.Add(march.Value);
                        }
                        System.Windows.Forms.Cursor.Position = new System.Drawing.Point(Int32.Parse(coordinates[0]), Int32.Parse(coordinates[1]));
                        //if (previousLine != scriptLine)
                            Thread.Sleep(100);
                        MouseHandler.LeftUp();
                        Thread.Sleep(100);
                    }
                    else if (scriptLine.StartsWith("RD"))//Right button down
                    {
                        MatchCollection matches = Regex.Matches(scriptLine, "[0-9]+");
                        List<string> coordinates = new List<string>();
                        coordinates.Clear();
                        foreach (Match march in matches)
                        {
                            coordinates.Add(march.Value);
                        }
                        System.Windows.Forms.Cursor.Position = new System.Drawing.Point(Int32.Parse(coordinates[0]), Int32.Parse(coordinates[1]));
                        //if (previousLine != scriptLine)
                            Thread.Sleep(100);
                        MouseHandler.RightDown();
                        Thread.Sleep(100);
                    }
                    else if (scriptLine.StartsWith("RU"))//Right button up
                    {
                        MatchCollection matches = Regex.Matches(scriptLine, "[0-9]+");
                        List<string> coordinates = new List<string>();
                        coordinates.Clear();
                        foreach (Match march in matches)
                        {
                            coordinates.Add(march.Value);
                        }
                        System.Windows.Forms.Cursor.Position = new System.Drawing.Point(Int32.Parse(coordinates[0]), Int32.Parse(coordinates[1]));
                        //if (previousLine != scriptLine)
                            Thread.Sleep(100);
                        MouseHandler.RightUp();
                        Thread.Sleep(100);
                    }
                    else if (scriptLine.StartsWith("LC"))//Left click
                    {
                        MatchCollection matches = Regex.Matches(scriptLine, "[0-9]+");
                        List<string> coordinates = new List<string>();
                        coordinates.Clear();
                        foreach (Match march in matches)
                        {
                            coordinates.Add(march.Value);
                        }
                        System.Windows.Forms.Cursor.Position = new System.Drawing.Point(Int32.Parse(coordinates[0]), Int32.Parse(coordinates[1]));
                        //if (previousLine != scriptLine)
                        Thread.Sleep(100);
                        MouseHandler.LeftClick();
                        Thread.Sleep(100);
                    }
                    else if (scriptLine.StartsWith("RC"))//Right click
                    {
                        MatchCollection matches = Regex.Matches(scriptLine, "[0-9]+");
                        List<string> coordinates = new List<string>();
                        coordinates.Clear();
                        foreach (Match march in matches)
                        {
                            coordinates.Add(march.Value);
                        }
                        System.Windows.Forms.Cursor.Position = new System.Drawing.Point(Int32.Parse(coordinates[0]), Int32.Parse(coordinates[1]));
                        //if (previousLine != scriptLine)
                        Thread.Sleep(100);
                        MouseHandler.RightClick();
                        Thread.Sleep(100);
                    }
                    else if (scriptLine.StartsWith("DC"))//Double click
                    {
                        MatchCollection matches = Regex.Matches(scriptLine, "[0-9]+");
                        List<string> coordinates = new List<string>();
                        coordinates.Clear();
                        foreach (Match march in matches)
                        {
                            coordinates.Add(march.Value);
                        }
                        System.Windows.Forms.Cursor.Position = new System.Drawing.Point(Int32.Parse(coordinates[0]), Int32.Parse(coordinates[1]));
                        //if (previousLine != scriptLine)
                        Thread.Sleep(100);
                        MouseHandler.LeftClick();
                        MouseHandler.LeftClick();
                        Thread.Sleep(100);
                    }
                    else if (scriptLine.StartsWith("SC"))//Scroll
                    {
                        if (!previousLine.StartsWith("SC"))
                            Thread.Sleep(2000);
                        string[] vals = scriptLine.Split(' ');
                        System.Windows.Forms.Cursor.Position = new System.Drawing.Point(Int32.Parse(Regex.Match(vals[1], @"\d+").Value), Int32.Parse(Regex.Match(vals[2], @"\d+").Value));
                        MouseHandler.Scroll(Int32.Parse(vals[3]) * reptscr);
                    }
                    else if (scriptLine.StartsWith("K"))//Keypress
                    {
                        string[] letter = scriptLine.Split(' ');
                        if (letter[1].Length > 1 && letter[1][0] == 'D')
                            letter[1] = letter[1][1].ToString();
                        KBHandler.Press(letter[1]);
                        Thread.Sleep(50);
                    }
                    else if(scriptLine.StartsWith("DR"))
                    {
                        Thread.Sleep(100);
                        MatchCollection matches = Regex.Matches(scriptLine, "[0-9]+");
                        List<string> coordinates = new List<string>();
                        coordinates.Clear();
                        foreach (Match march in matches)
                        {
                            coordinates.Add(march.Value);
                        }
                        int fromX = Int32.Parse(coordinates[0]);
                        int fromY = Int32.Parse(coordinates[1]);
                        int toX = Int32.Parse(coordinates[2]);
                        int toY = Int32.Parse(coordinates[3]);
                        for (int nod = 1; nod <= reptdr; nod++)
                        {
                            System.Windows.Forms.Cursor.Position = new System.Drawing.Point(fromX, fromY);
                            MouseHandler.LeftDown();
                            //Move y first
                            if (fromY > toY)
                            {
                                for (int shiftY = fromY; shiftY >= toY; shiftY--)
                                {
                                    Thread.Sleep(5);
                                    System.Windows.Forms.Cursor.Position = new System.Drawing.Point(fromX, shiftY);
                                }
                            }
                            else
                            {
                                for (int shiftY = fromY; shiftY <= toY; shiftY++)
                                {
                                    Thread.Sleep(5);
                                    System.Windows.Forms.Cursor.Position = new System.Drawing.Point(fromX, shiftY);
                                }
                            }
                            
                            //Move x then
                            if (fromX > toX)
                            {
                                for (int shiftX = fromX; shiftX >= toX; shiftX--)
                                {
                                    Thread.Sleep(5);
                                    System.Windows.Forms.Cursor.Position = new System.Drawing.Point(shiftX, toY);
                                }
                            }
                            else
                            {
                                for (int shiftX = fromX; shiftX <= toX; shiftX++)
                                {
                                    Thread.Sleep(5);
                                    System.Windows.Forms.Cursor.Position = new System.Drawing.Point(shiftX, toY);
                                }
                            }
                            MouseHandler.LeftUp();
                            Thread.Sleep(200);
                        }
                    }
                    if (previousLine.StartsWith("K") && scriptLine.StartsWith("K"))
                        Thread.Sleep(50);
                    else
                        Thread.Sleep(Int32.Parse(gap.SelectedIndex.ToString()) * 1000);
                    previousLine = scriptLine;
                    
                }
                if (inscr.IsChecked == true)
                    reptscr++;
                if (indr.IsChecked == true)
                    reptdr++;
            }
            teachMe.IsEnabled = _do.IsEnabled = true;
            this.WindowState = WindowState.Normal;
        }

        private void Window_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                this.DragMove();
            }   
        }

        private void clear_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Application will restart and your current training will be deleted, this may take a few seconds");
            System.Windows.Forms.Application.Restart();
            Application.Current.Shutdown();
        }

        private void load_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if (ofd.ShowDialog() == true)
            {
                File.Copy(ofd.FileName, @"kbmlog.ads", true);
            }
            MessageBox.Show("Script loaded successfuly");
        }

        private void save_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveLog = new SaveFileDialog();
            saveLog.RestoreDirectory = true;
            saveLog.FileName = "AuDoScript_" + DateTime.Now.ToShortDateString() + ".ads";
            if (saveLog.ShowDialog() == true)
            {
                try
                {
                    string log = File.ReadAllText(@"kbmlog.ads");
                    File.WriteAllText(saveLog.FileName, log);
                    MessageBox.Show("Script saved successfully");
                }
                catch (FileNotFoundException)
                {
                    MessageBox.Show("No script created yet");
                }
            }
        }

        private void options_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Under construction");
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void whatsFullScreen_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("To work with the software when operand software runs on full screen, you need to follow these steps in order\r\n\r\n1. Click on fullscreen button after clicking on Teach Me button\r\n2. Teach the software\r\n3. Press esc to exit full screen\r\n4. Stop teaching\r\n5. Click on Do button and click on the messagebox that appears\r\n6. Make full screen of the operand software\r\n7. Sit back and watch");
        }
    }
}
