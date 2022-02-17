using SAPFEWSELib;
using System;
using System.IO;
using System.Threading;

namespace DigitalVPM_DataCollector
{
    public class Program
    {
        public static AUTOSAP Session = new AUTOSAP();
        public static DateTime Connect, Check, Download;
        public static Thread DataCollection;
        public static Thread Inputs;
        public static CancellationTokenSource CTS = new CancellationTokenSource();

        public static void ConnectSAP()
        {
            if (!Session.CheckConnected())
            {
                Console.WriteLine("Attaching to SAP Scripting API");

                if (Session.GetSession())
                {
                    Console.WriteLine($"Session found on server {Session.GetSessionInfo().ApplicationServer} with client id {Session.GetSessionInfo().Client}");
                }
                else
                {
                    Console.WriteLine("Cannot find a SAP session, please verify SAP is running..");
                }

                Connect = DateTime.Now.AddMinutes(1);
            }
        }

        // Get data method
        public static void GetData(string workcenter, string startdate, string enddate, Settings setting)
        {

            Console.WriteLine($"Getting data from Work Center {workcenter}");

            Session.StartTransaction("ZIW37N");
            Session.SetVariant("/SMT-006");

            // De-select completed order
            Session.GetCheckBox("SP_MAB").Selected = false;
            Session.GetCheckBox("SP_HIS").Selected = false;

            // Set work centre
            Session.GetButton("%_S_GEWRK_%_APP_%-VALU_PUSH").Press();
            ((GuiButton)Session.GetFormById("wnd[1]/tbar[0]/btn[16]")).Press();
            ((GuiButton)Session.GetFormById("wnd[1]/tbar[0]/btn[8]")).Press();
            Session.GetCTextField("S_GEWRK-LOW").Text   = workcenter;

            // Set date
            Session.GetCTextField("S_DATUM-LOW").Text   = startdate;
            Session.GetCTextField("S_DATUM-HIGH").Text  = enddate;
            Session.GetTab("S_TAB2").Select();
            Session.GetCTextField("S_GLTRP-LOW").Text   = startdate;
            Session.GetCTextField("S_GLTRP-HIGH").Text  = enddate;

            // Set for lead
            Session.GetButton("%_S_PLNNR_%_APP_%-VALU_PUSH").Press();
            ((GuiCTextField)Session.GetFormById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]")).Text = "*001";
            ((GuiCTextField)Session.GetFormById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]")).Text = "WHE003";
            ((GuiButton)Session.GetFormById("wnd[1]/tbar[0]/btn[8]")).Press();

            // Set variant
            Session.GetTab("S_TAB9").Select();
            Session.GetCTextField("SP_VARI").Text = "SMT-006";
            Session.SendVKey(8);

            // Export to text
            GuiMenu Menu = (GuiMenu)Session.GetFormById("wnd[0]/mbar/menu[0]/menu[10]/menu[2]");

            Menu.Select();

            // Text file format popup
            GuiModalWindow Popup = (GuiModalWindow)Session.GetFormById("wnd[1]");
            ((GuiRadioButton)Popup.FindById("usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]")).Select();
            ((GuiButton)Popup.FindById("tbar[0]/btn[0]")).Press();

            // Filename popup
            Popup = (GuiModalWindow)Session.GetFormById("wnd[1]");
            ((GuiCTextField)Popup.FindByName("DY_PATH", "GuiCTextField")).Text = setting.GetSettings().DBPath;
            ((GuiCTextField)Popup.FindByName("DY_FILENAME", "GuiCTextField")).Text = $"ZIW37N_{workcenter}";
            ((GuiButton)Popup.FindById("tbar[0]/btn[11]")).Press();

            Session.EndTransaction();

            // Get planned/actual hours and costs
            Session.StartTransaction("ZCS_SERVICE_REPORT");
            Session.SetVariant("PERFORMANCE");
            Session.GetCTextField("S_GEWRK-LOW").Text = "SMB";
            Session.GetTextField("S_VAWRK-LOW").Text = "1002";
            Session.SendVKey(8);

            // Export to text
            Menu = (GuiMenu)Session.GetFormById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]");

            Menu.Select();

            // Text file format popup
            Popup = (GuiModalWindow)Session.GetFormById("wnd[1]");
            ((GuiRadioButton)Popup.FindById("usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]")).Select();
            ((GuiButton)Popup.FindById("tbar[0]/btn[0]")).Press();

            // Filename popup
            if (Directory.Exists(setting.GetSettings().DBPath))
            {
                Popup = (GuiModalWindow)Session.GetFormById("wnd[1]");
                ((GuiCTextField)Popup.FindByName("DY_PATH", "GuiCTextField")).Text = setting.GetSettings().DBPath;
                ((GuiCTextField)Popup.FindByName("DY_FILENAME", "GuiCTextField")).Text = $"ZCS_SERVICE_REPORT_{workcenter}";
                ((GuiButton)Popup.FindById("tbar[0]/btn[11]")).Press();
            } else
            {
                Console.WriteLine("Connection to PROD.LOCAL is not found..");
            }

            Session.EndTransaction();           
        }

        public static void ThreadOps(Settings DSet, CancellationToken Token)
        {
            // Set timers
            Connect = DateTime.Now;
            Download = DateTime.Now;

            while (!Token.IsCancellationRequested)
            {
                Check = DateTime.Now;

                // SAP Online check
                if (Check > Connect) ConnectSAP();

                // Download data check
                if (Check > Download && Session.CheckConnected())
                {
                    GetData("SMB",      "01.01.2018", "01.01.2222", DSet);
                    //GetData("MGLV",     "01.01.2018", "01.01.2222", DSet);
                    //GetData("ALT",      "01.01.2018", "01.01.2222", DSet);
                    //GetData("OHV",      "01.01.2018", "01.01.2222", DSet);
                    //GetData("WHEELSET", "01.01.2018", "01.01.2222", DSet);

                    Download = DateTime.Now.AddMinutes(DSet.GetSettings().DownloadTimeout);
                    Console.WriteLine($"Data collection completed, re-fresh @ {Download.ToShortTimeString()}", DataCollection);
                }

            }
        }

        public static void ReadInput(CancellationToken Token)
        {
            while(!Token.IsCancellationRequested)
            {
                Thread.Sleep(10);

                switch(Console.ReadLine())
                {
                    case "/exit":

                            CTS.Cancel();
                            

                        break;
                }
            }
        }


        // Main entry point
        private static void Main(string[] args)
        {
            
            // Variables
            Settings DSettings  = new Settings();

            DataCollection  = new Thread(new ThreadStart(delegate { ThreadOps(DSettings, CTS.Token); }));
            Inputs          = new Thread(new ThreadStart(delegate { ReadInput(CTS.Token); }));

            // Load settings
            DSettings.LoadSettings();

            Console.WriteLine("DIGITAL VPM - DATA COLLECTOR v 1");

            /* Start thread for data collection */
            DataCollection.Start();
            Inputs.Start();


            
        }
    }
}
