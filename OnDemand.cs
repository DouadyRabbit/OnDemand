using System;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;

//- Ref https://www.ibm.com/support/knowledgecenter/SSB2EG_7.4.0/com.ibm.ondemand.winclient.doc/dodwc249.htm

namespace OnDemand
{
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ProgId("OnDemand.OnDemand")]

    public class OnDemand
    {
        //-------------------- Parameters ----------------------------
        public string server = "SRPT-PROD";
        public string userName = "USER-NAME-HERE";
        public string password = "USER-PASSWORD-HERE";



        //--------------Error feedback---------------------------------
        public int errorCode = 0;
        public string errorMessage = "";


        //-----------------Search Operatos constants------------------
        const short ARS_OLE_OPR_EQUAL = 1;
        const short ARS_OLE_OPR_NOT_EQUAL = 2;
        const short ARS_OLE_OPR_LESS_THAN = 3;

        const short ARS_OLE_OPR_LESS_THAN_OR_EQUAL = 4;
        const short ARS_OLE_OPR_GREATER_THAN = 5;
        const short ARS_OLE_OPR_GREATER_THAN_OR_EQUAL = 6;

        const short ARS_OLE_OPR_BETWEEN = 7;
        const short ARS_OLE_OPR_NOT_BETWEEN = 8;
        const short ARS_OLE_OPR_IN = 9;

        const short ARS_OLE_OPR_NOT_IN = 10;
        const short ARS_OLE_OPR_LIKE = 11;
        const short ARS_OLE_OPR_NOT_LIKE = 12;

        //-----------------ARSOLE Find Types constants ------------------------
        const short ARS_OLE_FIND_FIRST = 1;
        const short ARS_OLE_FIND_PREV = 2;
        const short ARS_OLE_FIND_NEXT = 3;

        //--------------------------------------------------
        ArsOLEWrapper oOnD = null;

        //------------------------------------------------
        public OnDemand()
        {
            try
            {
                this.oOnD = new ArsOLEWrapper();
            }
            catch (Exception ex)
            {
                this.errorMessage = "Failed initializing OnDemand object" + ex.Message;
                this.errorCode = -1;
            }
        }

        //-------------------------------------
        public int execOD()
        {
            int error_Success = 0;

            //----------- Login -------------

            try
            {
                error_Success = oOnD.Logon(server, userName, password); ;
                if (error_Success != 0)
                {
                    errorCode = -1;
                    errorMessage = "Failed login to OnDemand";
                }

            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                errorCode = -2;
                throw;
            }

            return error_Success;
        }
    }

    //---------------------------------------------//
    public class ArsOLEWrapper : AxARSOLELib.AxArsOle
    {
        public ArsOLEWrapper()
        {
            this.CreateControl();
        }
    }
}



