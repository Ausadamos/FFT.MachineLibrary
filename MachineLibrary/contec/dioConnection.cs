using CdioCs;

namespace MachineLibrary
{
    public class dioConnection
    {
        Cdio _dio;
        public short m_Id;
        short m_InPortNum;
        short m_OutPortNum;
        string ErrorString;
        string machineName;



        public string Init(string dioName, string dioSerial)
        {
            _dio = new Cdio();
            string result = "success";
            int Ret;

            Ret = _dio.Init(dioSerial, out m_Id);
            _dio.GetErrorString(Ret, out ErrorString);
            if (Ret != (int)CdioConst.DIO_ERR_SUCCESS)
            {
                return "ไม่สามารถติดต่อกับเครื่อง " + dioName + " : " + System.Convert.ToString(Ret) + " : " + ErrorString;
            }

            //-----------------------------
            // Get all ports number
            //-----------------------------
            Ret = _dio.GetMaxPorts(m_Id, out m_InPortNum, out m_OutPortNum);
            _dio.GetErrorString(Ret, out ErrorString);
            if (Ret != (int)CdioConst.DIO_ERR_SUCCESS)
            {
                return "ไม่สามารถ Get all ports number ของเครื่อง " + dioName + " : " + System.Convert.ToString(Ret) + " : " + ErrorString;
            }

            if (m_InPortNum > 8)
            {
                m_InPortNum = 8;
            }
            if (m_OutPortNum > 8)
            {
                m_OutPortNum = 8;
            }

            machineName = dioName;

            return result;
        }


        public string setSigleBitChannel(int channel, string value)
        {

            string result = "success";
            int Ret;
            short OutBitNo;
            byte OutBitData;
            //-----------------------------
            // Get bit No. and output data from screen
            //-----------------------------
            OutBitNo = System.Convert.ToInt16(channel);
            OutBitData = System.Convert.ToByte(value, 16);

            //-----------------------------
            // Port input
            //-----------------------------
            Ret = _dio.OutBit(m_Id, OutBitNo, OutBitData);
            //-----------------------------
            // Error process
            //-----------------------------
            string ErrorString;
            _dio.GetErrorString(Ret, out ErrorString);
            if (Ret != (int)CdioConst.DIO_ERR_SUCCESS)
            {
                return "ไม่สามารถ set SD_IN ได้ : Ret = " + System.Convert.ToString(Ret) + " : " + ErrorString;
            }

            return result;

        }

        public int getSigleBitChannel(int channel)
        {
            int Ret;
            short OutBitNo;
            byte OutBitData;
            //-----------------------------
            // Get bit No. from screen
            //-----------------------------
            OutBitNo = System.Convert.ToInt16(channel);

            //-----------------------------
            // Port input
            //-----------------------------
            Ret = _dio.EchoBackBit(m_Id, OutBitNo, out OutBitData);
            //-----------------------------
            // Error process
            //-----------------------------
            string ErrorString;
            _dio.GetErrorString(Ret, out ErrorString);
            if (Ret != (int)CdioConst.DIO_ERR_SUCCESS)
            {
                return -1; //-1 ไม่มีอยู่จริง
            }

            //-----------------------------
            // Data express
            //-----------------------------

            return OutBitData;

        }

        public byte[] getMultiChannel()
        {

            int Ret;
            short[] BitNo = new short[8];
            byte[] Data = new byte[8];

            for (short InBit = 0; InBit < 8; InBit++)
            {
                BitNo[InBit] = InBit;
            }

            Ret = _dio.InpMultiBit(m_Id, BitNo, 8, Data);

            _dio.GetErrorString(Ret, out ErrorString);

            if (Ret != (int)CdioConst.DIO_ERR_SUCCESS)
            {
                return null;
            }

            return Data;

        }

        public int exit()
        {
            int Ret;
            Ret = _dio.Exit(m_Id);

            if (Ret != (int)CdioConst.DIO_ERR_SUCCESS)
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
