using System;
using System.Collections.Generic;
using System.Linq;
using CoreScanner; //Add reference to dll, change Reference > Interop.CoreScanner.dll > Properties > "Embedded interop types" to false.
using System.Xml;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Globalization;

namespace Scanner
{
    public class Scanner
    {
        //All Hexidecimal values for corresponding ASCII characters. A-Z, a-z, 0-9, !@#$%^&*()-=_+/\.,[]{}
        //Any values not stored in this string[] will be ignored upon decode.
        public static readonly string[] strALLOWED_HEXIDECIMAL_VALUES = new string[] { "0x20", "0x61", "0x62", "0x63", "0x64", "0x65", "0x66", "0x67", "0x68", "0x69", "0x6A", "0x6B", "0x6C", "0x6D", "0x6E", "0x6F", "0x70", "0x71", "0x72", "0x73", "0x74", "0x75", "0x76", "0x77", "0x78", "0x79", "0x7A", "0x41", "0x42", "0x43", "0x44", "0x45", "0x46", "0x47", "0x48", "0x49", "0x4A", "0x4B", "0x4C", "0x4D", "0x4E", "0x4F", "0x50", "0x51", "0x52", "0x53", "0x54", "0x55", "0x56", "0x57", "0x58", "0x59", "0x5A", "0x31", "0x32", "0x33", "0x34", "0x35", "0x36", "0x37", "0x38", "0x39", "0x30", "0x3A", "0x2F", "0x25", "0x2D", "0x5C", "0x21", "0x40", "0x23", "0x24", "0x25", "0x5E", "0x26", "0x2A", "0x28", "0x29", "0x2B", "0x7B", "0x7D", "0x5B", "0x5D", "0x3C", "0x3E", "0x2C", "0x2E" };
        public CCoreScannerClass cCScanner;

        /// <summary>
        /// Create an instance of the CCoreScannerClass.
        /// </summary>
        public Scanner()
        {
            cCScanner = new CCoreScannerClass();
        }

        /// <summary>
        /// Will open a connection to the connected scanner, parameter "0" is reserved for the method Open(). 
        /// </summary>
        /// <param name="ScannerTypes">1 = SCANNER_TYPES_ALL, 2=SCANNER_TYPES_SNAPI, 3=SCANNER_TYPES_SSI, 6=SCANNER_TYPES_IBMHID, 7=SCANNER_TYPES_NIXMODB, 8=SCANNER_TYPES_HIDKB, 9=SCANNER_TYPES_IBMTT</param>
        /// <param name="ScannerQty">Amount of scanners to connect to a given type. Max 255</param>
        /// <returns>0 = success, 1 = fail</returns>
        public bool ConnectScanner(short[] ScannerTypes, short ScannerQty)
        {
            int status = 1; //1 fail 0 success

            
            cCScanner.Open(0, ScannerTypes, ScannerQty, out status);

            return (0 == status);
        }

        /// <summary>
        ///  ---TABLE OF ALL POSSIBLE SCANNER COMMANDS---
        ///  CommandsGET_VERSION: 1000
        ///  REGISTER_FOR_EVENTS: 1001
        ///  UNREGISTER_FOR_EVENTS: 1002
        ///  CLAIM_DEVICE: 1500
        ///  RELEASE_DEVICE: 1501
        ///  ABORT_MACROPDF: 2000
        ///  ABORT_UPDATE_FIRMWARE: 2001
        ///  AIM_OFF: 2002
        ///  AIM_ON: 2003
        ///  FLUSH_MACROPDF: 2005
        ///  DEVICE_PULL_TRIGGER: 2011
        ///  DEVICE_RELEASE_TRIGGER: 2012
        ///  SCAN_DISABLE: 2013
        ///  SCAN_ENABLE: 2014
        ///  SET_PARAMETER_DEFAULTS: 2015
        ///  DEVICE_SET_PARAMETERS: 2016
        ///  SET_PARAMETER_PERSISTANCE: 2017
        ///  REBOOT_SCANNER: 2019
        ///  DEVICE_CAPTURE_IMAGE: 3000
        ///  DEVICE_CAPTURE_BARCODE: 3500
        ///  DEVICE_CAPTURE_VIDEO: 4000
        /// </summary>
        /// <param name="NumOfEvents">Number of events to be executed</param>
        /// <param name="EventID">Array of event ID's corresponding to an event. amount must match number of events.</param>
        /// <returns></returns>
        public bool ExecuteScanCommand(short NumOfEvents, short[] EventID)
        {
            string strInput = ""; //Act as the input for the command
            string strEvents = ""; //Build for each Number of Events
            string strOutXML = ""; //Output result in XML, typically null unless requesting information
            int status = 1; // 1 = fail, 0 = Success

            for (int CurEvent = 0; NumOfEvents > CurEvent; CurEvent++) //Itterates Length of NumOfEvents
            {
                strEvents += EventID[CurEvent].ToString() + ","; //Formatting string for output.
            }
            
            //Input for the XML Command sent to the scanner.
            strInput = $"<inArgs><cmdArgs><arg-int>{NumOfEvents}</arg-int><arg-int>{strEvents.TrimEnd(',')}</arg-int></cmdArgs></inArgs>";
            //Executing scanner command
            cCScanner.ExecCommand(1001, strInput, out strOutXML, out status);

            return (0 == status); //Success or Failure
        }

        /// <summary>
        /// This method will decode the Hexidecimal between the <rawdata></rawdata> tags and convert it to plaintext (char) format.
        /// </summary>
        /// <param name="rawData">Data from the barcode scanner</param>
        /// <returns></returns>
        public static string DecodeData(string rawData)
        {
            XDocument XMLBarcode = XDocument.Parse(rawData); //Loading scanner output into XML Format
            string output = "";

            //Foreach Hexidecimal split from the string by Spaces.
            foreach (string Hexidecimal in XMLBarcode.Root.Descendants("rawdata").First().Value.Split(' '))
            {
                //If the current Hexidecimal value is allowed based on the readonly array class field.
                if (strALLOWED_HEXIDECIMAL_VALUES.Contains(Hexidecimal))
                {
                    //Converting the Hexidecimal value to an actual ASCII character
                    output += Convert.ToChar(Int32.Parse(Hexidecimal.Substring(2), NumberStyles.AllowHexSpecifier));
                }
            }
            //Final string output
            return output;
        }
    }

    public class Scanner_Example
    {
        //Base 32 Hexidecimal recognizes the these characters
        private string strBase32Hex = "0123456789ABCDEFGHIJKLMNOPQRSTUV";
        //Dictionary of military branches to corresponding single char values
        private static readonly Dictionary<string, string> MilitaryBranches = new Dictionary<string, string>()
        {
            { "F", "Air Force" }, {"A", "Army" }, {"M", "Marines" }, {"N", "Navy" },
            { "C", "Coast Guard" }, {"D", "Department of Defense" }, {"H", "Public Health Service" }, {"O", "National Oceanic and Atmospheric Administration" },
            { "1", "Foreign Army" }, {"2", "Foreign Navy" }, {"3", "Foreign Marine Corps" }, {"4", "Foreign Air Force" },  {"X", "Other"}
        };

        //USE CASE OF THIS CLASS
        private static void Main()
        {
            scanner = new Scanner(); //New instance of this class
            scanner.ConnectScanner(new short[] { 1 }, 1); //Connecting to the scanner with the given type
            scanner.ExecuteScanCommand(1, new short[] { 1 }); //Subscribes to scan events

            //Creating an event handler for the OnBarcodeScan event
            scanner.cCScanner.BarcodeEvent += new _ICoreScannerEvents_BarcodeEventEventHandler(OnBarcodeEvent);
        }

        //// <summary>
        //// METHOD: Base32 Hex Decoder
        ////
        //// I did some research on the DoD CAC and how they store member information, I found some
        //// very useful information on their API, but mainly on how their encoding / HASHing works on the following DoD Website:
        //// https://www.dmdc.osd.mil/smartcard/images/CACapplicationProgrammingInterface.pdf
        //// Discovered that they use Base32 encoding with Extended Hex (not standard Base32). I was previously trying
        //// to decode with standard Base32 (and was obviously not working). So I went to the following technical documentation on
        //// the IETF's website, and found that they have GREAT charts explaining the Base16-64 encoding methodologies.
        //// https://tools.ietf.org/html/rfc4648#section-7
        //// https://en.wikipedia.org/wiki/Base32
        //// 
        //// The following is a direct snippet from the IETF website on RFC4648 (Most recent HASHing used for Base16,32,64 with Hex
        //// 
        //// <QUOTE Source=https://tools.ietf.org/html/rfc4648#section-7>
        //// 
        //// *Base 32 Encoding with Extended Hex Alphabet*
        //// The following description of base 32 is derived from[7].  This
        //// encoding may be referred to as "base32hex".  This encoding should not
        //// be regarded as the same as the "base32" encoding and should not be
        //// referred to as only "base32".  This encoding is used by, e.g.,
        //// NextSECure3(NSEC3) [10].
        ////
        //// One property with this alphabet, which the base64 and base32
        //// alphabets lack, is that encoded data maintains its sort order when
        //// the encoded data is compared bit-wise.
        ////
        //// This encoding is identical to the previous one, except for the
        //// alphabet.  The new alphabet is found in Table 4.
        ////
        ////              Table 4: The "Extended Hex" Base 32 Alphabet
        ////
        ////******                       IMPORTANT NOTE!                          *******
        ////****** The below table was represented as the constant 'strBase32Hex' *******
        ////
        ////      Value Encoding Value Encoding Value Encoding Value Encoding
        ////          0 0             9 9            18 I            27 R
        ////          1 1            10 A            19 J            28 S
        ////          2 2            11 B            20 K            29 T
        ////          3 3            12 C            21 L            30 U
        ////          4 4            13 D            22 M            31 V
        ////          5 5            14 E            23 N
        ////          6 6            15 F            24 O(pad) =
        ////          7 7            16 G            25 P
        ////          8 8            17 H            26 Q
        ////</QUOTE>
        //// </summary>
        //// <param name="strBase32">Raw string in Base32Hex format</param>
        //// <returns>Returns the decoded string from a given Base32Hex Value</returns>
        private string Base32HexDecode(string strRawBase32Hex)
        {
            decimal decResult = 0m;

            for (int intReverseIndex = strRawBase32Hex.Length; intReverseIndex >= 1; intReverseIndex--)
            {
                decResult += (decimal)(strBase32Hex.IndexOf(strRawBase32Hex.Substring(intReverseIndex - 1, 1))
                    * Math.Pow(32, (strRawBase32Hex.Length - intReverseIndex)));
            }

            return decResult.ToString();
        }

        /// <summary>
        /// Method called when the scanner scans a barcode.
        /// </summary>
        /// <param name="eventType"></param>
        /// <param name="pscanData"></param>
        public void OnBarcodeEvent(short eventType, ref string pscanData)
        {
            //pscanData is the raw XML & Hexidecimal formatted barcode scan. We use the method DecodeData(String) to 
            //Decode to plain text.
            string DecodedBarcode = Scanner.DecodeData(pscanData);

            //Accessing another call stack or thread from this current thread is done with the Invoke(MethodInvoker)delegate{}
            //When you set an object that exists on this form from the thread the scanner runs on, you must use this.
            Invoke((MethodInvoker)delegate
            {
                //Do things with the DecodedBarcode
                //Ex. Textbox1.Text = DecodedBarcode;
            }
         }
    }

}
