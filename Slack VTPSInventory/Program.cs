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

        public Form1()
        {

        }

        static void Main(string[] args)
        {
            //decrypt the token for github purposes
            String token = decryptToken(encryptedToken);


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
            String[,] dataArray;
            if (values != null && values.Count > 0)
            {
                dataArray = convertIListToArray(values);


                //Work within this big ass statement
                String line = "";

                for (int r = 0; r < dataArray.GetLength(0); r++)
                {
                    for(int sdfkjhsfjhsdf = 0; sdfkjhsfjhsdf < dataArray.GetLength(1); sdfkjhsfjhsdf++)
                    {
                        line = line + dataArray[r, sdfkjhsfjhsdf] + "\t";
                    }

                    line = "\n";

                }

            }

            sendMessage(token, "general", "AA1");

            Console.ReadKey();
        


            //client.PostMessage()
            
        }
        private static String receiveMessage()
        {
            return "String";
        }
        private static void sendMessage(String token, String slackChannel, String itemLocation)
        {
            ManualResetEventSlim clientReady = new ManualResetEventSlim(false);
            SlackSocketClient client = new SlackSocketClient(token);
            client.Connect((connected) => { 
                // This is called once the client has emitted the RTM start command
                clientReady.Set();
            }, () => {
                // This is called once the RTM client has connected to the end point
            });
            client.OnMessageReceived += (message) =>
            {
                // Handle each message as you receive them
            };
            //     clientReady.Wait();
            client.GetChannelList((clr) => { Console.WriteLine("got channels"); });
            client.PostMessage((mr) => Console.WriteLine("sent message to general!"), slackChannel, itemLocation);
            Console.ReadKey();
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
    }
    
}