using System;

namespace MachineLibrary
{
    public class Ophir
    {

        public int hDevice = 0;
        private OphirLMMeasurementLib.CoLMMeasurement lm = null;
        private string result = "";
        private object arrayValue, arrayTimestamp, arrayStatus;

        public string OpenUSBDevice(string serialNumbers)
        {
            result = "success";

            try
            {
                lm = new OphirLMMeasurementLib.CoLMMeasurement();

                object serials;
                lm.ScanUSB(out serials);
                lm.OpenUSBDevice(serialNumbers, out hDevice);
            }
            catch (Exception ex)
            {
                result = ex.Message;
            }

            return result;

        }

        public string StartStream()
        {
            result = "success";

            try
            {
                lm.StartStream(hDevice, 0);
            }
            catch (Exception ex)
            {
                result = ex.Message;
            }

            return result;

        }

        public string StopStreamAndClose()
        {
            result = "success";

            try
            {
                lm.StopStream(hDevice, 0);
                lm.Close(hDevice);
            }
            catch (Exception ex)
            {
                result = ex.Message;
            }

            return result;

        }

        public double GetData()
        {
            double result = -99;

            try
            {

                System.Threading.Thread.Sleep(1000);
                lm.GetData(hDevice, 0, out arrayValue, out arrayTimestamp, out arrayStatus);

                double[] values = (double[])arrayValue;
                double[] timestamps = (double[])arrayTimestamp;
                int[] statuses = (int[])arrayStatus;

                if (values.Length > 0)
                {
                    result = values[0];
                }
                else
                {
                    result = -99;
                }
            }
            catch (Exception ex)
            {
                result = -99;
            }

            return result;

        }
    }
}
