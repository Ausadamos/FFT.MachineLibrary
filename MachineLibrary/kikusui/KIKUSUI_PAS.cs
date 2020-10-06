using Ivi.Visa.Interop;
using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace MachineLibrary
{
    public class KIKUSUI_PAS
    {

        ResourceManager rm;
        IMessage msg;
        long _waittime = 50;

        public enum MODE
        {
            VOLTAGE = 0,
            CURRENT = 1
        }

        public KIKUSUI_PAS(String USB_NAME, int ADDRESS, MODE _MODE, Double MAX_VI, long WAITTIME = 50)
        {

            rm = new ResourceManager();
            msg = (IMessage)rm.Open(USB_NAME, AccessMode.NO_LOCK, 5000, "");
            _waittime = WAITTIME;

            msg.WriteString("NODE " + ADDRESS);
            msg.WriteString("OUT 0");

            if (_MODE == MODE.VOLTAGE)
            {

                msg.WriteString("ISET " + MAX_VI);
                msg.WriteString("VSET 0");

            }
            else
            {
                msg.WriteString("VSET " + MAX_VI);
                msg.WriteString("ISET 0");
            }


        }

        public bool close_kikusui()
        {
            msg.WriteString("ISET 0");
            msg.WriteString("VSET 0");
            msg.WriteString("OUT 0");
            return true;
        }

        public bool ps_off()
        {

            msg.WriteString("OUT 0");
            DELAY(_waittime);
            return true;
        }

        public bool ps_on()
        {
            msg.WriteString("OUT 1");
            DELAY(_waittime);

            return true;
        }

        public double set_voltage(Double voltage)
        {

            msg.WriteString("VSET " + voltage.ToString("0.00"));
            DELAY(_waittime);
            msg.WriteString("VOUT?");

            DELAY(_waittime);
            return Convert.ToDouble(String.Format("{0:0.00}", msg.ReadString(256)));

        }

        public double get_voltage()
        {

            msg.WriteString("VOUT?");
            DELAY(_waittime);
            if (Convert.ToDouble(String.Format("{0:0.00}", msg.ReadString(256))) < 0)
            {

                return 0;

            }
            else
            {

                return Convert.ToDouble(String.Format("{0:0.00}", msg.ReadString(256)));

            }


        }

        public void Dispose()
        {
            msg.WriteString("ISET 0");
            msg.WriteString("VSET 0");
            msg.WriteString("OUT 0");

            Dispose();
            GC.SuppressFinalize(this);
        }


        private void DELAY(long ms)
        {
            Stopwatch stw = new Stopwatch();
            stw.Reset();
            stw.Start();

            while (stw.ElapsedMilliseconds < ms)
            {
                Application.DoEvents();
            }

            stw.Stop();
            stw = null;
        }
    }
}
