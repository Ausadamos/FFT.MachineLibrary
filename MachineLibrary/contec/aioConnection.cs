using CaioCs;

namespace MachineLibrary
{
    public class aioConnection
    {

        Caio _aio;
        public short m_Id;
        string ErrorString;
        string machineName;

        public string Init(string aioName, string aioSerial)
        {
            _aio = new Caio();

            string result = "success";
            int Ret;
            Ret = _aio.Init(aioSerial, out m_Id);
            if (Ret != 0)
            {
                _aio.GetErrorString(Ret, out ErrorString);
                return "ไม่สามารถติดต่อกับเครื่อง " + aioName + " : " + System.Convert.ToString(Ret) + " : " + ErrorString;
            }

            setRangeContec();

            machineName = aioName;

            return result;
        }

        public float[] getMultiChannel()
        {

            int Ret;
            short AiChannels = 13;
            float[] AiData = new float[AiChannels];
            Ret = _aio.MultiAiEx(m_Id, AiChannels, AiData);
            if (Ret != 0)
            {
                AiData = null;
            }

            return AiData;
            //ถ้าส่ง null คือ Error
        }


        private string setRangeContec()
        {
            string result = "success";
            int Index;
            short AiRange = 0;
            Index = 0;
            switch (Index)
            {
                case 0: AiRange = (short)CaioConst.PM10; break;
                case 1: AiRange = (short)CaioConst.PM5; break;
                case 2: AiRange = (short)CaioConst.PM25; break;
                case 3: AiRange = (short)CaioConst.PM125; break;
                case 4: AiRange = (short)CaioConst.PM1; break;
                case 5: AiRange = (short)CaioConst.PM0625; break;
                case 6: AiRange = (short)CaioConst.PM05; break;
                case 7: AiRange = (short)CaioConst.PM03125; break;
                case 8: AiRange = (short)CaioConst.PM025; break;
                case 9: AiRange = (short)CaioConst.PM0125; break;
                case 10: AiRange = (short)CaioConst.PM01; break;
                case 11: AiRange = (short)CaioConst.PM005; break;
                case 12: AiRange = (short)CaioConst.PM0025; break;
                case 13: AiRange = (short)CaioConst.PM00125; break;
                case 14: AiRange = (short)CaioConst.PM001; break;
                case 15: AiRange = (short)CaioConst.P10; break;
                case 16: AiRange = (short)CaioConst.P5; break;
                case 17: AiRange = (short)CaioConst.P4095; break;
                case 18: AiRange = (short)CaioConst.P25; break;
                case 19: AiRange = (short)CaioConst.P125; break;
                case 20: AiRange = (short)CaioConst.P1; break;
                case 21: AiRange = (short)CaioConst.P05; break;
                case 22: AiRange = (short)CaioConst.P025; break;
                case 23: AiRange = (short)CaioConst.P01; break;
                case 24: AiRange = (short)CaioConst.P005; break;
                case 25: AiRange = (short)CaioConst.P0025; break;
                case 26: AiRange = (short)CaioConst.P00125; break;
                case 27: AiRange = (short)CaioConst.P001; break;
                case 28: AiRange = (short)CaioConst.P20MA; break;
                case 29: AiRange = (short)CaioConst.P4TO20MA; break;
                case 30: AiRange = (short)CaioConst.P1TO5; break;
            }

            // Set the input range
            int Ret;
            Ret = _aio.SetAiRangeAll(m_Id, AiRange);
            if (Ret != 0)
            {
                string ErrorString;
                _aio.GetErrorString(Ret, out ErrorString);
                return "ไม่สามารถ Set input range ของเครื่อง " + machineName + " : " + System.Convert.ToString(Ret) + " : " + ErrorString;
            }

            return result;
        }

        public int exit()
        {
            int Ret;
            Ret = _aio.Exit(m_Id);

            if (Ret != 0)
            {
                return -1; //-1 ไม่มีอยู่จริง
            }
            else
            {
                return 0;
            }

        }

    }
}
