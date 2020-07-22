using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using LockheedMartin.Prepar3D.SimConnect;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using DiscordRPC;
using System.Threading;


namespace SimConnectForms {
    public partial class Form1 : Form {

        public Form1() {
            InitializeComponent();
            connectSim();
        }

        SimConnect simconnect = null;
        const int WM_USER_SIMCONNECT = 0x0402;
        enum DEFINITIONS {
            Struct1,
        }
        enum DATA_REQUESTS {
            REQUEST_1,
        };
        // this is how you declare a data structure so that
        // simconnect knows how to fill it/read it.
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi, Pack = 1)]
        struct Struct1 {
            // this is how you declare a fixed size string
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
            public String title;
            public double latitude;
            public double longitude;
            public double altitude;
            public double speed;
        };


        public DiscordRpcClient Client { get; private set; }

        void setup() {
            Client = new DiscordRpcClient("735335622323077204");
            Client.Initialize();
        }

        void updatePresence(string title, double altitude, double speed) {
            int alt = Convert.ToInt32(altitude);
            speed = Math.Round(speed);
            Client.SetPresence(new RichPresence() {
                Details = "Flying the " + title,
                State = "Going " + speed.ToString() + " knots at " + alt.ToString() + " feet",
                Assets = new Assets() {
                    LargeImageKey = "maxresdefault"
                }
            });
        }

        void cleanup() {
            Client.Dispose();
        }


        protected override void DefWndProc(ref Message m) {
            if (m.Msg == WM_USER_SIMCONNECT) {
                if (simconnect != null) {
                    simconnect.ReceiveMessage();
                }
            } else {
                base.DefWndProc(ref m);
            }
        }

        void connectSim() {
            if (simconnect == null) {
                try {
                    setup();
                    simconnect = new SimConnect("Managed Data Request", this.Handle, WM_USER_SIMCONNECT, null, 0);
                    initDataRequest();
                } catch (COMException ex) { }
            } else {
                closeConnection();
            }
        }


        private void initDataRequest() {
            try {
                // listen to connect and quit msgs
                simconnect.OnRecvOpen += new SimConnect.RecvOpenEventHandler(simconnect_OnRecvOpen);
                simconnect.OnRecvQuit += new SimConnect.RecvQuitEventHandler(simconnect_OnRecvQuit);

                // listen to exceptions
                simconnect.OnRecvException += new SimConnect.RecvExceptionEventHandler(simconnect_OnRecvException);

                // define a data structure
                simconnect.AddToDataDefinition(DEFINITIONS.Struct1, "title", null, SIMCONNECT_DATATYPE.STRING256, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                simconnect.AddToDataDefinition(DEFINITIONS.Struct1, "Plane Latitude", "degrees", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                simconnect.AddToDataDefinition(DEFINITIONS.Struct1, "Plane Longitude", "degrees", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                simconnect.AddToDataDefinition(DEFINITIONS.Struct1, "Plane Altitude", "feet", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);
                simconnect.AddToDataDefinition(DEFINITIONS.Struct1, "Airspeed Indicated", "knots", SIMCONNECT_DATATYPE.FLOAT64, 0.0f, SimConnect.SIMCONNECT_UNUSED);

                // IMPORTANT: register it with the simconnect managed wrapper marshaller
                // if you skip this step, you will only receive a uint in the .dwData field.
                simconnect.RegisterDataDefineStruct<Struct1>(DEFINITIONS.Struct1);
                // NearestVorCurrentICAO
                // catch a simobject data request
                simconnect.OnRecvSimobjectDataBytype += new SimConnect.RecvSimobjectDataBytypeEventHandler(simconnect_OnRecvSimobjectDataBytype);
            } catch (COMException ex) {
            }
        }

        void simconnect_OnRecvOpen(SimConnect sender, SIMCONNECT_RECV_OPEN data) {
        }

        void simconnect_OnRecvQuit(SimConnect sender, SIMCONNECT_RECV data) {
            closeConnection();
        }

        void simconnect_OnRecvException(SimConnect sender, SIMCONNECT_RECV_EXCEPTION data) {
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e) {
            closeConnection();
        }

        void simconnect_OnRecvSimobjectDataBytype(SimConnect sender, SIMCONNECT_RECV_SIMOBJECT_DATA_BYTYPE data) {

            switch ((DATA_REQUESTS)data.dwRequestID) {
                case DATA_REQUESTS.REQUEST_1:
                    Struct1 s1 = (Struct1)data.dwData[0];
                    updatePresence(s1.title, s1.altitude, s1.speed);
                    break;

                default:
                    break;
            }
        }

        private void getVars() {
            // The following call returns identical information to:
            // simconnect.RequestDataOnSimObject(DATA_REQUESTS.REQUEST_1, DEFINITIONS.Struct1, SimConnect.SIMCONNECT_OBJECT_ID_USER, SIMCONNECT_PERIOD.ONCE);

            simconnect.RequestDataOnSimObjectType(DATA_REQUESTS.REQUEST_1, DEFINITIONS.Struct1, 0, SIMCONNECT_SIMOBJECT_TYPE.USER);
        }
        void closeConnection() {
            if (simconnect != null) {
                // Dispose serves the same purpose as SimConnect_Close()
                simconnect.Dispose();
                simconnect = null;
            }
        }

        private void button1_Click(object sender, EventArgs e) {
            //getVars();
            while (true) {
                getVars();
            }

        }
    }
}
