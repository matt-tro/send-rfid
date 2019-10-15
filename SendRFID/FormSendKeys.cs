using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO.Ports;

namespace SendRFID
{
    public partial class FormSendKeys : Form
    {
        [DllImport("kernel32.dll")]
        static extern void Sleep(int dwMilliseconds);

        [DllImport("MasterRD.dll")]
        static extern int lib_ver(ref uint pVer);

        [DllImport("MasterRD.dll")]
        static extern int rf_init_com(int port, int baud);

        [DllImport("MasterRD.dll")]
        static extern int rf_ClosePort();

        [DllImport("MasterRD.dll")]
        static extern int rf_antenna_sta(short icdev, byte mode);

        [DllImport("MasterRD.dll")]
        static extern int rf_init_type(short icdev, byte type);

        [DllImport("MasterRD.dll")]
        static extern int rf_request(short icdev, byte mode, ref ushort pTagType);

        [DllImport("MasterRD.dll")]
        static extern int rf_anticoll(short icdev, byte bcnt, IntPtr pSnr, ref byte pRLength);

        [DllImport("MasterRD.dll")]
        static extern int rf_select(short icdev, IntPtr pSnr, byte srcLen, ref sbyte Size);

        [DllImport("MasterRD.dll")]
        static extern int rf_halt(short icdev);

        [DllImport("MasterRD.dll")]
        static extern int rf_M1_authentication2(short icdev, byte mode, byte secnr, IntPtr key);

        [DllImport("MasterRD.dll")]
        static extern int rf_M1_initval(short icdev, byte adr, Int32 value);

        [DllImport("MasterRD.dll")]
        static extern int rf_M1_increment(short icdev, byte adr, Int32 value);

        [DllImport("MasterRD.dll")]
        static extern int rf_M1_decrement(short icdev, byte adr, Int32 value);

        [DllImport("MasterRD.dll")]
        static extern int rf_M1_readval(short icdev, byte adr, ref Int32 pValue);

        [DllImport("MasterRD.dll")]
        static extern int rf_M1_read(short icdev, byte adr, IntPtr pData, ref byte pLen);

        [DllImport("MasterRD.dll")]
        static extern int rf_M1_write(short icdev, byte adr, IntPtr pData);

        [DllImport("MasterRD.dll")]
        static extern int rf_beep(short icdev, int msec);

        [DllImport("MasterRD.dll")]
        static extern int rf_light(short icdev, int color);

        [DllImport("User32.dll")]
        static extern int SetForegroundWindow(IntPtr point);

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        bool bConnectedDevice;
        private bool allowshowdisplay = false;

        public FormSendKeys()
        {
            if (CheckInstance())
            {
                MessageBox.Show(this, "Un'altra istanza è già in esecuzione.","Impossibile avviare", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                Environment.Exit(1);
                Application.Exit();
            }

            InitializeComponent();

            backgroundWorker.RunWorkerAsync();

        }

        protected override void SetVisibleCore(bool value)
        {
            base.SetVisibleCore(allowshowdisplay ? value : allowshowdisplay);
        }

        private bool CheckInstance()
        {
            int np = Process.GetProcessesByName("SendRFID").Count<Process>();
            return np > 1;
        }

        private void SendToActiveWindow(String cardNumber)
        {
            IntPtr h = GetForegroundWindow();
            SendKeys.SendWait(cardNumber);
        }

        private void Connect()
        {
            List<int> ports = new List<int>();
            foreach(String nomePorta in SerialPort.GetPortNames())
            {
                int numeroPorta = Int32.Parse(nomePorta.Remove(0, 3));
                ports.Add(numeroPorta);
            }

            bConnectedDevice = false;
            if (!bConnectedDevice)
            {
                int baud = 0;
                int status;

                foreach(int port in ports)
                {
                    baud = Convert.ToInt32("14400");
                    status = rf_init_com(port, baud);
                    if (0 == status)
                    {
                        notifyIcon.Icon = Resources.logo_ok;
                        bConnectedDevice = true;
                        break;
                    }
                }
            }
        }

        private void Disconnect()
        {
            if (bConnectedDevice)
            {
                notifyIcon.Icon = Resources.logo_ko;
                rf_ClosePort();
                bConnectedDevice = false;
            }
        }

        public String GetCardNumber()
        {
            short icdev = 0x0000;
            int status;
            byte type = (byte)'A';//mifare one 
            byte mode = 0x52;
            ushort TagType = 0;
            byte bcnt = 0x04;//mifare 
            IntPtr pSnr;
            byte len = 255;
            sbyte size = 0;
            String m_cardNo = String.Empty;

            try
            {
                pSnr = Marshal.AllocHGlobal(1024);
                for (int i = 0; i < 2; i++)
                {
                    status = rf_antenna_sta(icdev, 0);//
                    if (status != 0)
                        throw new Exception();

                    Sleep(20);
                    status = rf_init_type(icdev, type);
                    if (status != 0)
                        continue;

                    Sleep(20);
                    status = rf_antenna_sta(icdev, 1);//
                    if (status != 0)
                        continue;

                    Sleep(50);
                    status = rf_request(icdev, mode, ref TagType);//
                    if (status != 0)
                        continue;

                    status = rf_anticoll(icdev, bcnt, pSnr, ref len);//
                    if (status != 0)
                        continue;

                    status = rf_select(icdev, pSnr, len, ref size);//ISO14443-3 TYPE_A 
                    if (status != 0)
                        continue;

                    byte[] szBytes = new byte[len];

                    for (int j = 0; j < len; j++)
                    {
                        szBytes[j] = Marshal.ReadByte(pSnr, j);
                    }

                    for (int q = 0; q < len; q++)
                    {
                        m_cardNo += byteHEX(szBytes[q]);
                    }

                    rf_beep(icdev, 10);
                    rf_light(icdev, 2);
                    Sleep(200);
                    rf_light(icdev, 1);

                    break;
                }
                Marshal.FreeHGlobal(pSnr);

                return m_cardNo;

            }
            catch (Exception)
            {
                Disconnect();
                return "";
            }

        }

        private static String byteHEX(Byte ib)
        {
            String _str = String.Empty;
            try
            {
                char[] Digit = { '0', '1', '2', '3', '4', '5', '6', '7', '8', '9', 'A',
                'B', 'C', 'D', 'E', 'F' };
                char[] ob = new char[2];
                ob[0] = Digit[(ib >> 4) & 0X0F];
                ob[1] = Digit[ib & 0X0F];
                _str = new String(ob);
            }
            catch (Exception)
            {
                new Exception("");
            }
            return _str;

        }

        private void toolStripEsci_Click(object sender, EventArgs e)
        {
            Environment.Exit(1);
        }

        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            notifyIcon.Icon = Resources.logo_ko;
            notifyIcon.Visible = true;
            notifyIcon.ShowBalloonTip(500, "SendRFID", "La app sta funzionando in background", ToolTipIcon.Info);

            do
            {
                while (!bConnectedDevice)
                    Connect();

                while (bConnectedDevice)
                {
                    String cardNumber = GetCardNumber();
                    if (cardNumber != String.Empty)
                    {
                        SendToActiveWindow(cardNumber);
                        Sleep(1000);
                    }
                }

            } while (true);
        }

        private void FormSendKeys_Shown(object sender, EventArgs e)
        {
            this.Hide();
            this.Visible = false;
        }
    }
}
