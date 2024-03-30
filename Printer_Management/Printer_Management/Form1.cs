using System;
using System.Diagnostics;
using System.Management;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.ServiceProcess;
using System.Net;
using System.Threading;
using System.Net.NetworkInformation;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Drawing;

namespace Printer_Management
{
    public partial class Form1 : Form
    {

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int SendMessage(IntPtr hWnd, int wMsg, IntPtr wParam, IntPtr lParam);

        private const int WM_VSCROLL = 277; // Vertical scroll
        private const int SB_THUMBPOSITION = 4; // Thumb track position

        public Form1()
        {
            StartPrintService(); //启动打印服务
            InitializeComponent(); //初始化窗口
            GetPrinterInfo(); //获取本地打印机
            A05.Text = "A05";
            A06.Text = "A06";
            A07.Text = "A07";
            IPadd();
            textBox1.KeyDown += TextBox1_KeyDown;
        }

        // 列表窗口位置同步
        private void SyncListBoxScroll(ListBox sourceListBox, ListBox targetListBox)
        {
            int index = sourceListBox.SelectedIndex;
            if (index != -1)
            {
                int topIndex = sourceListBox.TopIndex;
                IntPtr scrollValue = (IntPtr)(SB_THUMBPOSITION + (0x10000 * topIndex));
                SendMessage(targetListBox.Handle, WM_VSCROLL, scrollValue, IntPtr.Zero);
            }
        }

        // 刷新 ListBox1 & ListBox2 列表内容
        private void GetPrinterInfo()
        {
            // 清空列表框
            listBox1.Items.Clear();
            listBox2.Items.Clear();

            // 创建一个 ManagementObjectSearcher 对象，用于执行 WMI 查询
            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Printer");

            // 遍历查询结果
            foreach (ManagementObject printer in searcher.Get())
            {
                // 将打印机名称添加到 listbox1
                listBox1.Items.Add(printer["Name"]);

                // 将打印机端口添加到 listbox2
                listBox2.Items.Add(printer["PortName"]);
            }
        }

        // ListBox1 窗口同步选项
        private void ListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 当用户选择 listBox1 中的项时，同步更新 listBox2 的选中项
            int selectedIndex = listBox1.SelectedIndex;
            if (selectedIndex >= 0 && selectedIndex < listBox2.Items.Count)
            {
                listBox2.SelectedIndex = selectedIndex;
                Prot_out.Text = listBox2.Items[selectedIndex].ToString();
                Print_name.Text = listBox1.Items[selectedIndex].ToString();

                // 提取IP地址并输出到 IPout.Text
                string ipAddress = ExtractIpAddress(Prot_out.Text);
                IPout.Text = ipAddress;
            }
            if (IPout.Text == "未找到端   口")
            {
                IPAddress_OLD.Text = "";
            }
            else if (IPout.Text != "未找到端   口")
            {
                IPAddress_OLD.Text = IPout.Text;
            }
            SyncListBoxScroll(listBox1, listBox2);
        }

        // ListBox2 窗口同步选项
        private void ListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 当用户选择 listBox2 中的项时，同步更新 listBox1 的选中项
            int selectedIndex = listBox2.SelectedIndex;
            if (selectedIndex >= 0 && selectedIndex < listBox1.Items.Count)
            {
                listBox1.SelectedIndex = selectedIndex;
                Prot_out.Text = listBox2.Items[selectedIndex].ToString();
                Print_name.Text = listBox1.Items[selectedIndex].ToString();
            }
            if (IPout.Text == "未找到端   口")
            {
                IPAddress_OLD.Text = "";
            }
            else if (IPout.Text != "未找到端   口")
            {
                IPAddress_OLD.Text = IPout.Text;
            }
        }

        // ListBox2 IP端口 筛选
        private string ExtractIpAddress(string input)
        {
            string pattern = @"\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}";
            Match match = Regex.Match(input, pattern);
            if (match.Success)
            {
                return match.Value;
            }
            else
            {
                return "无效的IPv4地址。";
            }

        }
        
        // 模块功能 - Ping选中的打印机
        private void TestPing_Click(object sender, EventArgs e)
        {
            string ipAddress = IPout.Text;

            // 检查是否是有效的IP地址
            if (!IPAddress.TryParse(ipAddress, out IPAddress ip))
            {
                MessageBox.Show(Print_name.Text + ipAddress + " 不是一个有效的IP地址", " - 测试 SelectPrint IPAddress -", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Ping pingSender = new Ping();
            PingReply reply = pingSender.Send(ipAddress);

            if (reply.Status == IPStatus.Success)
            {
                MessageBox.Show(Print_name.Text + "  " + ipAddress + " 可以连接", " - 测试 SelectPrint IPAddress -", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show(Print_name.Text + "  " + ipAddress + " 无法连接", " - 测试 SelectPrint IPAddress -", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 菜单选项 - 打开设备的打印机
        private void 新建ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("control", "printers");
        }

        // 菜单选项 - 重启打印机服务
        private void 打开ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RestartPrintService();
        }

        // 模块功能 - 重启打印机服务
        private void RestartPrintService()
        {
            string serviceName = "Spooler"; // 打印机服务的服务名称

            try
            {
                ServiceController sc = new ServiceController(serviceName);

                if (sc.Status == ServiceControllerStatus.Running)
                {
                    var bw = new BackgroundWorker();
                    bw.DoWork += (sender, e) =>
                    {
                        // 停止打印机服务
                        sc.Stop();
                        sc.WaitForStatus(ServiceControllerStatus.Stopped);
                        System.Threading.Thread.Sleep(1000); // 延时1秒

                        // 启动打印机服务
                        sc.Start();
                        sc.WaitForStatus(ServiceControllerStatus.Running);
                    };

                    bw.RunWorkerAsync();
                }
                else
                {
                    MessageBox.Show("打印机服务当前未运行，无需重启.");

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("发生异常: " + ex.Message);
            }

        }

        // 模块功能 - 启动打印机服务
        private void StartPrintService()
        {
            ServiceController service = new ServiceController("Spooler");

            try
            {
                TimeSpan timeout = TimeSpan.FromMilliseconds(5000);

                if (service.Status == ServiceControllerStatus.Stopped || service.Status == ServiceControllerStatus.StopPending)
                {
                    service.Start();
                    service.WaitForStatus(ServiceControllerStatus.Running, timeout);
                }
            }
            catch (Exception e)
            {
                // Handle exceptions here.
                Console.WriteLine(e.Message);
            }
        }

        // ListBox1 鼠标选中 事件 - 内容的操作
        private void ListBox1_MouseUp(object sender, MouseEventArgs e)
        {
            // 如果是鼠标右键单击
            if (e.Button == MouseButtons.Right)
            {
                int index = listBox1.IndexFromPoint(e.Location);
                if (index != ListBox.NoMatches)
                {
                    listBox1.SelectedIndex = index;
                    ContextMenu contextMenu = new ContextMenu();

                    MenuItem setDefaultMenuItem = new MenuItem("设置为默认打印机");
                    setDefaultMenuItem.Click += SetDefaultMenuItem_Click;
                    contextMenu.MenuItems.Add(setDefaultMenuItem);

                    MenuItem printTestPageMenuItem = new MenuItem("打印选中打印机的测试页");
                    printTestPageMenuItem.Click += PrintTestPageMenuItem_Click;
                    contextMenu.MenuItems.Add(printTestPageMenuItem);

                    MenuItem OpenPrintQueueMenuItem = new MenuItem("查看正在打印什么");
                    OpenPrintQueueMenuItem.Click += OpenPrintQueueMenuItem_Click;
                    contextMenu.MenuItems.Add(OpenPrintQueueMenuItem);

                    MenuItem TestPingItem = new MenuItem("测试选中打印机的IP");
                    TestPingItem.Click += TestPing_Click; // 注意这里应该是TestPingItem.Click，而不是OpenPrintQueueMenuItem.Click
                    contextMenu.MenuItems.Add(TestPingItem);

                    contextMenu.Show(listBox1, e.Location);
                }
            }
        }

        // ListBox1 鼠标右键选项 - 打印测试页
        private void PrintTestPageMenuItem_Click(object sender, EventArgs e)
        {
            string printerName = listBox1.SelectedItem.ToString().Replace("->", "").Trim();
            ThreadPool.QueueUserWorkItem(_ =>
            {
                using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Printer"))
                {
                    foreach (ManagementObject printer in searcher.Get())
                    {
                        if (printer["Name"].ToString() == printerName)
                        {
                            printer.InvokeMethod("PrintTestPage", null);
                            break;
                        }
                    }
                }
                this.Invoke(new Action(() =>
                {
                    MessageBox.Show("测试页已发送到 [ " + printerName + " ] 打印机。", " - 发送 Print Test Page -", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }));
            });
        }

        // ListBox1 鼠标右键选项 - 设置为默认打印机
        private void SetDefaultMenuItem_Click(object sender, EventArgs e)
        {
            string printerName = listBox1.SelectedItem.ToString().Replace("->", "").Trim();
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Printer"))
            {
                foreach (ManagementObject printer in searcher.Get())
                {
                    if (printer["Name"].ToString() == printerName)
                    {
                        printer.InvokeMethod("SetDefaultPrinter", null);
                        MessageBox.Show("已将 [ " + printerName + " ] 设置为默认打印机", " - 设置 DefaultPrint -", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        break;
                    }
                }
            }
        }

        // ListBox1 鼠标右键选项 - 打开选中的打印队列
        private void OpenPrintQueueMenuItem_Click(object sender, EventArgs e)
        {
            string printerName = listBox1.SelectedItem.ToString().Replace("->", "").Trim();
            using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Printer"))
            {
                foreach (ManagementObject printer in searcher.Get())
                {
                    if (printer["Name"].ToString() == printerName)
                    {
                        ProcessStartInfo psi = new ProcessStartInfo
                        {
                            FileName = "cmd.exe",
                            Arguments = string.Format("/c rundll32 printui.dll,PrintUIEntry /o /n \"{0}\"", printerName),
                            UseShellExecute = false,
                            CreateNoWindow = true
                        };
                        Process p = new Process
                        {
                            StartInfo = psi
                        };
                        p.Start();
                        Thread t = new Thread(() => p.WaitForExit())
                        {
                            IsBackground = true
                        };
                        t.Start();
                        break;
                    }
                }
            }
        }

        // 获取网卡IP地址信息
        private void IPadd()
        {
            listBox3.Items.Clear();
            listBox4.Items.Clear();
            NetworkInterface[] adapters = NetworkInterface.GetAllNetworkInterfaces();

            foreach (NetworkInterface adapter in adapters)
            {
                if (adapter.NetworkInterfaceType == NetworkInterfaceType.Ethernet ||
                    adapter.NetworkInterfaceType == NetworkInterfaceType.Wireless80211)
                {
                    bool isVirtualAdapter = adapter.Description.ToLower().Contains(A01.Text) ||
                                            adapter.Description.ToLower().Contains(A02.Text) ||
                                            adapter.Description.ToLower().Contains(A03.Text) ||
                                            adapter.Description.ToLower().Contains(A04.Text) ||
                                            adapter.Description.ToLower().Contains(A05.Text) ||
                                            adapter.Description.ToLower().Contains(A06.Text) ||
                                            adapter.Description.ToLower().Contains(A07.Text);

                    if (!isVirtualAdapter)
                    {
                        IPInterfaceProperties adapterProperties = adapter.GetIPProperties();
                        foreach (UnicastIPAddressInformation ip in adapterProperties.UnicastAddresses)
                        {
                            if (ip.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                            {
                                listBox3.Items.Add(adapter.Description);
                                listBox4.Items.Add(ip.Address.ToString());
                            }
                        }
                    }
                }
            }
        }
        
        // ListBox3 窗口同步选项
        private void ListBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 当用户选择 listBox3 中的项时，同步更新 listBox4 的选中项
            int selectedIndex = listBox3.SelectedIndex;
            if (selectedIndex >= 0 && selectedIndex < listBox4.Items.Count)
            {
                listBox4.SelectedIndex = selectedIndex;
                Prot_out.Text = listBox4.Items[selectedIndex].ToString();
                Print_name.Text = listBox3.Items[selectedIndex].ToString();

                // 提取IP地址并输出到 IPout.Text
                string ipAddress = ExtractIpAddress1(Prot_out.Text);
                IPout.Text = ipAddress;
            }
            if (IPout.Text == "未找到端   口")
            {
                textBox2.Text = "";
            }
            else if (IPout.Text != "未找到端   口")
            {
                textBox2.Text = IPout.Text;
            }
            SyncListBoxScroll(listBox3, listBox4);
        }
        
        // ListBox4 窗口同步选项
        private void ListBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            // 当用户选择 listBox4 中的项时，同步更新 listBox3 的选中项
            int selectedIndex = listBox4.SelectedIndex;
            if (selectedIndex >= 0 && selectedIndex < listBox3.Items.Count)
            {
                listBox3.SelectedIndex = selectedIndex;
                Prot_out.Text = listBox4.Items[selectedIndex].ToString();
                Print_name.Text = listBox3.Items[selectedIndex].ToString();
            }
            if (IPout.Text == "未找到端   口")
            {
                textBox2.Text = "";
            }
            else if (IPout.Text != "未找到端   口")
            {
                textBox2.Text = IPout.Text;
            }
        }

        // ListBox2 IP端口 筛选
        private string ExtractIpAddress1(string input)
        {
            string pattern = @"\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}";
            Match match = Regex.Match(input, pattern);
            if (match.Success)
            {
                return match.Value;
            }
            else
            {
                return "无效的IPv4地址。";
            }

        }

        // 点击按钮加载网页信息
        private void Button1_Click(object sender, EventArgs e)
        {
            webBrowser1.Navigate(IPAddress_OLD.Text);
        }

        // 输入密码后点击确认
        private void Button2_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "1234")
            {
                PasswordSuccess();
            }
            else
            {
                MessageBox.Show("密码错误！", "ERROR Message !");
                textBox1.Clear();
            }
        }

        // 密码正确
        private void PasswordSuccess()
        {
            pictureBox1.Visible = false;
            label1.Visible = false;
            button2.Visible = false;
            textBox1.Visible = false;
            //this.Size = new Size(1400, 900);
        }

        // 输入密码后按下回车
        private void TextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            
            if (e.KeyCode == Keys.Enter)
            {
                if (textBox1.Text == "1234")
                {
                    PasswordSuccess();
                }
                else
                {
                    MessageBox.Show("密码错误！","ERROR Message !", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    textBox1.Clear();
                }
            }
        }

        // 复制 IPv4 地址提示
        private void Button3_Click(object sender, EventArgs e)
        {
            if (textBox2.Text == "")
            {
                MessageBox.Show("地址为空 ！               ", "Tips Message !", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            else if (textBox2.Text != "")
            {
                Clipboard.SetText(textBox2.Text);
                MessageBox.Show("已复制 ！               ", "Tips Message !", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        // 关于软件
        private void 软件ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(" - 用来设置并操作打印机后台信息，前提需要打印机支持Web访问！！！ \n - 鼠标右击选项有功能可用！ \n - Ps:看似没用，实则好像就是没什么用...", "Software Message !", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // 关于作者
        private void 作者ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(" - Ae （2023/03/26） \n - FreyaGrace", "Author Message !", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // 广告位招租
        private void 广告位招租ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(" - 再这样下去，就要穷的吃土了！！！！", "Advertisement Message !", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // 假的刷新
        private void 刷新界面ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(" - 我真是太懒了，不想搞这个刷新咧！哼！", "Refresh UI Message !", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // A01
        private void CheckBox1_CheckedChanged_1(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                A01.Text = "bluetooth";
                IPadd();
            }
            else if (checkBox1.Checked == false)
            {
                A01.Text = "A01";
                IPadd();
            }
        }

        // A02
        private void CheckBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                A02.Text = "virtual";
                IPadd();
            }
            else if (checkBox2.Checked == false)
            {
                A02.Text = "A02";
                IPadd();
            }
        }

        // A03
        private void CheckBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                A03.Text = "vmware";
                IPadd();
            }
            else if (checkBox3.Checked == false)
            {
                A03.Text = "A03";
                IPadd();
            }
        }

        // A04
        private void CheckBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked == true)
            {
                A04.Text = "virtualbox";
                IPadd();
            }
            else if (checkBox4.Checked == false)
            {
                A04.Text = "A04";
                IPadd();
            }
        }

        // A05
        private void CheckBox5_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked == true)
            {
                A05.Text = "intel(r)";
                IPadd();
            }
            else if (checkBox5.Checked == false)
            {
                A05.Text = "A05";
                IPadd();
            }
        }

        // A06
        private void CheckBox6_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox6.Checked == true)
            {
                A06.Text = "realtek";
                IPadd();
            }
            else if (checkBox6.Checked == false)
            {
                A06.Text = "A06";
                IPadd();
            }
        }

        // 刷新网卡信息
        private void Button4_Click(object sender, EventArgs e)
        {
            IPadd();
        }

        // A07
        private void CheckBox7_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox7.Checked == true)
            {
                A07.Text = "tp-link";
                IPadd();
            }
            else if (checkBox6.Checked == false)
            {
                A07.Text = "A07";
                IPadd();
            }
        }
    }
}
