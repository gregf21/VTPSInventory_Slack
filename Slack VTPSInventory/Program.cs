using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Threading;
using SlackAPI;


namespace VTPSInventory
{
    public partial class Form1
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "InventoryLocations";
        static readonly string SpreadsheetId = "14LTmOA1xTyFOZ0eHHOENEZzabqh5bnTXVgC6Zs_QZj4";
        static readonly string sheet = "InventoryLocations";
        static SheetsService service;

        static readonly String encryptedToken = "xoxb - 428502883536 - 873105840420 - OZDHxvmy5psq1cQzWHmRvBpy";


        static SlackSocketClient client;
        static ManualResetEventSlim clientReady;

        static String[,] inventoryData;


        public Form1()
        {

        }

        static void Main(string[] args)
        {
            //decrypt the token for github purposes
            String token = decryptToken(encryptedToken);
            grabInventoryData();


            setUpClient(token);
            setUpClientReady();
            while (checkClientReady())
            {
                //do nothing until client is ready
            }
            sendMessage(token, sheet, "Online");

            while (true)
            {
                client.OnMessageReceived += (message) =>
                {
                    sendMessage(token, sheet, message.text);
                    Console.WriteLine(message.text);
                };
                setUpClient(token);
                setUpClientReady();
                while (checkClientReady())
                {
                    //do nothing until client is ready
                }
            }
        }
        private static void sendMessage(String token, String slackChannel, String item)
        {
            new Thread(() =>
            {
                Thread.CurrentThread.IsBackground = true;
                /* run your code here */
                findItemData(item);

                client.GetChannelList((clr) => { Console.WriteLine("Got Channels"); });
                client.PostMessage((mr) => Console.WriteLine("Posted Inventory Data"), slackChannel, item);
            }).Start();
            
        }

        private static String[,] convertIListToArray(IList<IList<object>> values)
        {
            String[,] dataArray = new String[values.Count, values[0].Count];

            int r = 0, c = 0;
            foreach (var row in values)
            {
                foreach (String dataEntry in row)
                {
                    dataArray[r, c] = dataEntry;
                    c++;
                }
                r++;
                c = 0;
            }
            return dataArray;
        }
        private static String decryptToken(String token)
        {
            return token.Replace(" ", String.Empty);
        }

        private static bool setUpClient(String token)
        {
            try
            {
                client = new SlackSocketClient(token);
                return true;
            }
            catch(Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                return false;
            }
        }

        private static bool setUpClientReady()
        {
            try
            {
                clientReady = new ManualResetEventSlim(false);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                return false;
            }   
        }

        private static bool checkClientReady()
        {
            client.Connect((connected) => {
                // This is called once the client has emitted the RTM start command
                clientReady.Set();
            }, () => {
                // This is called once the RTM client has connected to the end point
            });
            clientReady.Wait();
            return !clientReady.IsSet;
        }
        private static void grabInventoryData()
        {
            GoogleCredential credential;
            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream)
                    .CreateScoped(Scopes);
            }

            // Create Google Sheets API service.
            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            var range = $"{sheet}!A:B";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                service.Spreadsheets.Values.Get(SpreadsheetId, range);

            var response = request.Execute();
            IList<IList<object>> values = response.Values;
            inventoryData = convertIListToArray(values);
        }

        private static String[] findItemData(String item)
        {
            String[] dataArray = new String[inventoryData.GetLength(1)];
            String currentItem;

            for (int r = 1; r < inventoryData.GetLength(0); r++)
            {
                currentItem = inventoryData[r, 0];
                
            }
            return dataArray;
        }
    }
}
