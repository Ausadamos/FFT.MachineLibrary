using NationalInstruments.NI4882;
using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace MachineLibrary
{
    public class GPIB_OSA
    {

        public Device gpib_dev;

        public bool fGpibInit(string SYSTEM_GPIB_OSA_ADDRESS)
        {
            try
            {
                Address address = new Address();
                address.PrimaryAddress = byte.Parse(SYSTEM_GPIB_OSA_ADDRESS);

                gpib_dev = new Device(0, address);
                gpib_dev.DefaultBufferSize = 1200000;
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
        }

        public bool fGpibRead(ref string ret)
        {
            try
            {
                string strBuff = "";
                strBuff = gpib_dev.ReadString();
                ret = strBuff;
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
        }

        public bool fGpibReadToFile(string file_name)
        {
            try
            {

                gpib_dev.ReadToFile(file_name);
                return true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
        }

        public bool fGpibWrite(string str)
        {
            try
            {

                gpib_dev.Write(str);
                Delay(100);
                return true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
        }

        public bool fGpibQuery(string str, ref string ret)
        {
            string strBuff = "";

            try
            {
                fGpibWrite(str);
                Delay(200);
                strBuff = gpib_dev.ReadString();
                ret = strBuff;
                return true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
        }

        public bool fOsaInit()
        {
            try
            {
                gpib_dev.Reset();
                fGpibWrite("*RST");
                fGpibWrite("CFORM1");
                return true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
        }

        public void fOsaWait()
        {
            string str = "";

            gpib_dev.IOTimeout = TimeoutValue.T100ms;
            fGpibWrite("*WAI");
            do
            {
                try
                {
                    gpib_dev.Write(":STATUS:OPERATION?");
                    Delay(50);
                    str = gpib_dev.ReadString();
                }
                catch (GpibException ex)
                {
                    str = ex.ErrorCode.ToString();
                }
            } while (double.Parse(str) > 0);

            gpib_dev.IOTimeout = TimeoutValue.T10s;
        }


        public bool fOsaWLShift(double offset)
        {
            try
            {
                fGpibWrite(string.Format(":SENS:CORR:WAV:SHifT {0}NM", offset.ToString("0.0000")));
                return true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
        }

        public bool fOsaSweepStop(int dev)
        {
            try
            {
                fGpibWrite(string.Format(":ABORT"));
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
        }

        public bool fOsaSetup(double dblRes, int intSen, int intAvg, int intSmpl, int intMesWL)
        {
            try
            {
                //MeasWL
                fGpibWrite(string.Format(":SENS:CORR:RVEL:MEDIUM {0}", intMesWL));

                //Resolution
                fGpibWrite(string.Format(":SENS:BAND:RES {0}NM", dblRes.ToString("F4")));

                //Average
                fGpibWrite(string.Format(":SENS:AVER:COUNT {0}", intAvg));

                //Sample
                fGpibWrite(string.Format(":SENS:SWEEP:POINTS {0}", intSmpl));

                //Sensitivity
                fGpibWrite(string.Format(":SENS:SENSE {0}", intSen));

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
        }

        public bool fOsaCenter(double dblCenter)
        {
            try
            {
                fGpibWrite(string.Format(":SENS:WAV:CENTER {0}NM", dblCenter.ToString("F4")));
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
        }

        public bool fOsaSpan(double dblSpan)
        {
            try
            {
                fGpibWrite(string.Format(":SENS:WAV:SPAN {0}NM", dblSpan.ToString("F4")));
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
        }

        //--------------------------------------------------------------------------------------
        // OSA Level
        //--------------------------------------------------------------------------------------


        public bool fOsaLevelLogLog(double RefLevel, double LogScale, double SubLogScale, double LevelOffset)
        {
            try
            {

                fGpibWrite(string.Format(":DISP:WINDOW:TRACE:Y1:SCALE:SPAC LOG;:DISP:WINDOW:TRACE:Y1:SCALE:PDIV {0}DB", LogScale.ToString("F2")));
                fGpibWrite(string.Format(":DISP:WINDOW:TRACE:Y1:SCALE:RLEVEL {0}DBM", RefLevel.ToString("F2")));
                fGpibWrite(string.Format(":DISP:WINDOW:TRACE:Y2:SCALE:SPAC LOG;:DISP:WINDOW:TRACE:Y2:SCALE:PDIV {0}DB", SubLogScale.ToString("F2")));
                fGpibWrite(string.Format(":DISP:WINDOW:TRACE:Y2:SCALE:OLEVEL {0}DB", LevelOffset.ToString("0.00")));

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
        }

        public bool fOsaLineMarkerClear()
        {
            try
            {

                fGpibWrite(string.Format(":CALC:LMARKER:AOFF; :CALC:LMARKER:SRANGE 0"));
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
        }

        public bool fOsaLineMarkerSet(double dblMark1 = 0, double dblMark2 = 0, double dblMark3 = 0, double dblMark4 = 0)
        {
            try
            {

                //Set Line Marker 1
                if (dblMark1 != 0) fGpibWrite(string.Format(":CALC:LMARKER:X 1,{0}NM", dblMark1.ToString("F4")));

                //Set Line Marker 2
                if (dblMark2 != 0) fGpibWrite(string.Format(":CALC:LMARKER:X 2,{0}NM", dblMark2.ToString("F4")));

                //Set Line Marker 3
                if (dblMark3 != 0) fGpibWrite(string.Format(":CALC:LMARKER:Y 3,{0}NM", dblMark3.ToString("F4")));

                //Set Line Marker 4
                if (dblMark4 != 0) fGpibWrite(string.Format(":CALC:LMARKER:Y 4,{0}NM", dblMark4.ToString("F4")));
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
        }


        public bool fOsaSweepSingle(bool L1L2Sweep = false)
        {
            try
            {

                fGpibWrite(string.Format(":ABORT"));
                if (L1L2Sweep)
                {
                    fGpibWrite(string.Format(":SENS:WAV:SRANGE 1"));
                }
                else
                {
                    fGpibWrite(string.Format(":SENS:WAV:SRANGE 0"));
                }

                fGpibWrite(string.Format(":INIT:SMODE 1"));
                fGpibWrite(string.Format("*CLS"));
                fGpibWrite(string.Format(":INIT"));

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
        }

        public bool fOsaSweepRepeat(bool L1L2Sweep = false)
        {
            try
            {

                fGpibWrite(string.Format(":ABORT"));

                if (L1L2Sweep)
                {
                    fGpibWrite(string.Format(":SENS:WAV:SRANGE 1"));
                }
                else
                {
                    fGpibWrite(string.Format(":SENS:WAV:SRANGE 0"));
                }

                fGpibWrite(string.Format(":INIT:SMODE 2"));

                fGpibWrite(string.Format("*CLS"));

                fGpibWrite(string.Format(":INIT"));
                return true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
        }

        //--------------------------------------------------------------------------------------
        //?@OSA Peak Search
        //--------------------------------------------------------------------------------------

        public bool fOsaPeakSearch(bool L1L2Sweep = false)
        {
            try
            {
                fGpibWrite(string.Format(":ABORT"));

                if (L1L2Sweep)
                {
                    fGpibWrite(string.Format(":CALC:LMAR:SRAN 1"));
                }
                else
                {
                    fGpibWrite(string.Format(":CALC:LMAR:SRAN 0"));
                }


                fGpibWrite(string.Format(":CALC:MARK:MAX"));
                return true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
        }

        public bool fOsaPeakSearch(ref string ret, bool L1L2Sweep = false)
        {

            string wl = "";
            string lv = "";

            try
            {
                if (L1L2Sweep)
                {

                    if (fGpibQuery(string.Format(":ABORT;:CALC:LMAR:SRAN 1;:CALC:MARK:MAX;:CALC:MARK:X? 0"), ref wl) == false)
                    {
                        ret = "";
                        return false;
                    }

                    if (fGpibQuery(string.Format(":ABORT;:CALC:LMAR:SRAN 1;:CALC:MARK:MAX;:CALC:MARK:Y? 0"), ref lv) == false)
                    {
                        ret = "";
                        return false;
                    }
                }
                else
                {
                    if (fGpibQuery(string.Format(":ABORT;:CALC:LMAR:SRAN 0;:CALC:MARK:MAX;:CALC:MARK:X? 0"), ref wl) == false)
                    {
                        ret = "";
                        return false;
                    }

                    if (fGpibQuery(string.Format(":ABORT;:CALC:LMAR:SRAN 0;:CALC:MARK:MAX;:CALC:MARK:Y? 0"), ref lv) == false)
                    {
                        ret = "";
                        return false;
                    }
                }

                ret = string.Format("{0},{1}", wl, lv);
                return true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
        }


        //--------------------------------------------------------------------------------------
        //?@OSA Bottom Search
        //--------------------------------------------------------------------------------------

        public bool fOsaTraceDelete(string trace_string)
        {
            string readstring = "0";
            try
            {

                fGpibWrite(string.Format(":TRACE:DELETE TR{0}", trace_string));
                return true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
        }

        public bool fOsaTraceGetWL(string trace_string)
        {

            try
            {

                fGpibWrite(string.Format(":TRACE:DATA:X? TR{0}", trace_string));
                Delay(200);
                fGpibReadToFile("tmp\\wl.txt");
                return true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
        }

        public bool fOsaTraceGetWL(string trace_string, ref string[] ret)
        {
            string readstring = "0";

            try
            {

                if (fGpibQuery(string.Format(":TRACE:DATA:X? TR{0}", trace_string), ref readstring) == false)
                {
                    ret = new string[] { };
                    return false;
                }



                readstring = readstring.Replace(Environment.NewLine, "");
                readstring = readstring.Replace("\r", "");
                readstring = readstring.Replace("\n", "");
                ret = readstring.Split(',');
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
        }

        public bool fOsaTraceGetLevel(string trace_string)
        {
            string readstring = "0";

            try
            {

                fGpibWrite(string.Format(":TRACE:DATA:Y? TR{0}", trace_string));
                Delay(200);
                fGpibReadToFile("tmp\\lv.txt");
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
        }

        public bool fOsaTraceGetLevel(string trace_string, ref string[] ret)
        {
            string readstring = "0";

            try
            {
                if (fGpibQuery(string.Format(":TRACE:DATA:Y? TR{0}", trace_string), ref readstring) == false)
                {
                    ret = new string[] { };
                    return false;
                }

                readstring = readstring.Replace(Environment.NewLine, "");
                readstring = readstring.Replace("\r", "");
                readstring = readstring.Replace("\n", "");
                ret = readstring.Split(',');
                return true;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return false;
            }
        }

        public Bitmap fOsaScreenCapture()
        {
            try
            {
                byte[] intDummy;

                gpib_dev.Clear();

                //Save bmp file to internal memory
                fGpibWrite(":mmem:stor:grap color,bmp,\"test\",int");

                //Read bmp file from OSA
                fGpibWrite(":mmem:data? \"test.bmp\",int");

                //gpib_dev.ReadToFile("tmp\test.bmp")
                intDummy = gpib_dev.ReadByteArray();

                //----- save data to file
                long startPos = 0;

                while ((char)(intDummy[startPos]) != 'B')
                {
                    startPos += 1;
                }

                return new Bitmap(new MemoryStream(intDummy, int.Parse(startPos.ToString()), (int.Parse(intDummy.Length.ToString()) - int.Parse(startPos.ToString()))));

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return null;
            }
        }


        public int fOsaScreenCapture(PictureBox picturebox)
        {
            try
            {
                byte[] intDummy;

                //Save bmp file to internal memory
                fGpibWrite(":mmem:stor:grap color,bmp,\"test\",int");

                //Read bmp file from OSA
                fGpibWrite(":mmem:data? \"test.bmp\",int");

                //gpib_dev.ReadToFile("tmp\test.bmp")
                intDummy = gpib_dev.ReadByteArray();

                //----- save data to file
                long startPos = 0;

                while ((char)(intDummy[startPos]) != 'B')
                {
                    startPos += 1;
                }

                picturebox.Image = new Bitmap(new MemoryStream(intDummy, int.Parse(startPos.ToString()), (int.Parse(intDummy.Length.ToString()) - int.Parse(startPos.ToString()))));

                return intDummy.Length;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return -1;
            }
        }


        public int fOsaScreenCapture(string fileName)
        {
            try
            {
                byte[] intDummy;

                //Save bmp file to internal memory
                fGpibWrite(":mmem:stor:grap color,bmp,\"test\",int");

                //Read bmp file from OSA
                fGpibWrite(":mmem:data? \"test.bmp\",int");

                //gpib_dev.ReadToFile("tmp\test.bmp")
                intDummy = gpib_dev.ReadByteArray();

                //----- save data to file
                long startPos = 0;

                while ((char)(intDummy[startPos]) != 'B')
                {
                    startPos += 1;
                }

                Bitmap im = new Bitmap(new MemoryStream(intDummy, int.Parse(startPos.ToString()), (int.Parse(intDummy.Length.ToString()) - int.Parse(startPos.ToString()))));
                im.Save(fileName);

                return intDummy.Length;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "เกิดข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return -1;
            }
        }

        private void Delay(long ms)
        {
            Stopwatch stw = new Stopwatch();
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
