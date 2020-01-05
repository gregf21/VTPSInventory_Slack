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
using System.Net;


namespace VTPSInventory
{
    public partial class Form1
    {
        static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static readonly string ApplicationName = "InventoryLocations";
        static readonly string SpreadsheetId = "14LTmOA1xTyFOZ0eHHOENEZzabqh5bnTXVgC6Zs_QZj4";
        static readonly string sheet = "InventoryLocations";
        static SheetsService service;
        static readonly String botName = "<@URP33QQCC>";
        static readonly String encryptedToken = "xoxb - 428502883536 - 873105840420 - OZDHxvmy5psq1cQzWHmRvBpy";
        public static String decryptedToken = null;

        static SlackSocketClient client;
        static ManualResetEventSlim clientReady;

        static String[,] inventoryData;

        static String previousItem;


        public Form1()
        {

        }

        static void Main(string[] args)
        {
            //decrypt the token for github purposes
            String decryptedToken = decryptToken(encryptedToken);
            String token = decryptedToken;
            grabInventoryData();
            Console.WriteLine("Got Inventory Data");
            previousItem = "";
            setUpClient(token);
            setUpClientReady();
            while (checkClientReady())
            {
                //do nothing until client is ready
            }
            setUpClient(token);
            setUpClientReady();
            while (checkClientReady())
            {
                //do nothing until client is ready
            }
            Console.WriteLine("Client is Ready");
            StartServer();
            Console.WriteLine("Press Enter to Exit...");
            Console.ReadLine();
        }
        public static void StartServer()
        {
            var httpListener = new HttpListener();
            string hostName = Dns.GetHostName();
            var simpleServer = new SimpleServer(httpListener, "http://" + Dns.GetHostByName(hostName).AddressList[0].ToString() + ":1234/test/", ProcessYourResponse);
            simpleServer.Start();
            Console.WriteLine("Server Started :: IP Address is " + Dns.GetHostByName(hostName).AddressList[0].ToString());
        }
        private static string[] ParseInput(string input)
        {
            //0: User
            //1: Message Text
            //2: Slack Channel
            string[] message = new string[3];
            int userIndex = input.LastIndexOf(@"""user"":");
            while (input[userIndex] != ':')
            {
                userIndex++;
            }
            while (input[userIndex] != '"')
            {
                userIndex++;
            }
            userIndex++;
            int indexend = userIndex;
            while (input[indexend] != '"')
            {
                message[0] += input[indexend];
                indexend++;
            }

            int textIndex = input.LastIndexOf(@"""text"":");
            while (input[textIndex] != ':')
            {
                textIndex++;
            }
            while (input[textIndex] != '"')
            {
                textIndex++;
            }
            textIndex++;
            int textIndexEnd = textIndex;
            while (input[textIndexEnd] != '"')
            {
                message[1] += input[textIndexEnd];
                textIndexEnd++;
            }
            message[1] = message[1].Remove(0, 1);

            int channelIndex = input.LastIndexOf(@"""channel"":");
            while (input[channelIndex] != ':')
            {
                channelIndex++;
            }
            while (input[channelIndex] != '"')
            {
                channelIndex++;
            }
            channelIndex++;
            int channelIndexEnd = channelIndex;
            while (input[channelIndexEnd] != '"')
            {
                message[2] += input[channelIndexEnd];
                channelIndexEnd++;
            }
            return message;
        }
        public static byte[] ProcessYourResponse(string input)
        {
            byte[] response;
            if (input.LastIndexOf(@"challenge") != -1 && input.LastIndexOf(@"token") != -1)
            {
                int index = input.LastIndexOf(@"challenge");
                while (input[index] != ':')
                {
                    index++;
                }
                while (input[index] != '"')
                {
                    index++;
                }
                index++;
                int indexend = index;
                string challenge = "";
                while (input[indexend] != '"')
                {
                    challenge += input[indexend];
                    indexend++;
                }
                response = Encoding.ASCII.GetBytes("HTTP 200 OK\nContent-type: te / plain\n" + challenge);
                return response;
            }
            else
            {
                response = Encoding.ASCII.GetBytes("Received");
                string[] message = ParseInput(input);
                Console.WriteLine("User: " + message[0]);
                sendMessage(decryptedToken, message[2], message[1], message[0]);
                return response;
            }
        }
        private static void sendMessage(String token, String slackChannel, String item, String user)
        {
            new Thread(() =>
            {
                Thread.CurrentThread.IsBackground = true;
                /* run your code here */
                if (item != "")
                { }
                    String[] tempData = findItemData(item);

            if (tempData[0] != "" && previousItem != tempData[0])
            {
                String inventoryOutput = arrayToString(tempData);
                client.GetChannelList((clr) => { Console.WriteLine("Got Channels"); });
                Console.WriteLine("Search for: " + item);
                client.PostMessage((mr) => Console.WriteLine("Posted Inventory Data to: " + slackChannel + "\nInventory: " + inventoryOutput), slackChannel, "<@" + user + ">\n" + inventoryOutput);
                previousItem = inventoryOutput;
            }
            else if (tempData[0] == "" && tempData[1] == "")
                    {
                        client.GetChannelList((clr) => { Console.WriteLine("Got Channels"); });
                    Console.WriteLine("Search for: " + item);
                    client.PostMessage((mr) => Console.WriteLine("Posted Inventory Data to: " + slackChannel + "\nInventory: " + "Item not found"), slackChannel, "Item not found.");
                    }
            
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
            string[] dataArray = { "", "" };
            for (int i = 0; i < inventoryData.GetLength(0); i++)
            {
                for (int r = 0; r < 2; r++)
                {
                    if (inventoryData[i,0] != null && inventoryData[i, 1] != null)
                    {
                        if ((inventoryData[i, 0].ToLower().Contains(item.ToLower())))
                        {
                            dataArray[0] = inventoryData[i, 0];
                            dataArray[1] = inventoryData[i, 1];
                        }
                    }
                }
            }
            //EASTER EGGS
            if (item.ToLower() == "greg fairbanks")
            {
                dataArray[0] = "My creator!";
                dataArray[1] = "";
            }
            if (item.ToLower() == "help")
            {
                dataArray[0] = "Post a message like this and I will respond with a location:\n@inventorybot [ITEM]";
                dataArray[1] = "";
            }
            if(item.ToLower().Contains("make me a sandwich"))
            {
                dataArray[0] = "Poof. You're a sandwich.";
                dataArray[1] = "";
            }
            if (item.ToLower().Contains("ain't nobody got time for that") || item.ToLower().Contains("aint nobody got time for that"))
            {
                dataArray[0] = "Truth.";
                dataArray[1] = "";
            }
            if (item.ToLower().Contains("do you know morse code"))
            {
                dataArray[0] = "-. ---";
                dataArray[1] = "";
            }
            if (item.ToLower().Contains("here comes dat boi"))
            {
                dataArray[0] = "Oh Shoot.\nWaddup?!";
                dataArray[1] = "";
            }
            if(item.ToLower().Contains("more cowbell"))
            {
                dataArray[0] = "You always need more cowbell.";
                dataArray[1] = "";
            }
            if (item.ToLower().Contains("lets play a game") || item.ToLower().Contains("let's play a game"))
            {
                dataArray[0] = "Let's play hide and seek in the warehouse!";
                dataArray[1] = "";
            }
            if (item.ToLower() == "meow" || item.ToLower() == "meow" || item.ToLower() == "woof" || item.ToLower() == "roar" ||
                item.ToLower() == "oink" || item.ToLower() == "neigh")
            {
                dataArray[0] = item;
                dataArray[1] = "";
            }
            return dataArray;
        }
        /*
        private static String[] findItemData(String item)
        {
            int maxSimilarLetters = 0, maxSimilarOrder = 0, similarLetters, similarOrder;

            String[] dataArray = null;
            String currentItem;

            item = item.ToLower();
            for (int r = 1; r < inventoryData.GetLength(0); r++)
            {
                if (inventoryData[r, 0] != null)
                {
                    currentItem = inventoryData[r, 0].ToLower();
                    similarLetters = 0;
                    similarOrder = 0;

                    if (item == currentItem)
                    {
                        dataArray = convert2DRowTo1DArray(r);
                        break;
                    }

                    for (int i = 0; i < item.Length; i++)
                    {
                        for (int i2 = 0; i2 < currentItem.Length; i2++)
                        {
                            if (item.Substring(i, 1) == currentItem.Substring(i2, 1))
                            {
                                similarLetters++;
                            }
                            if (item.Contains(currentItem.Substring(0, i2 + 1)) && currentItem.Length <= item.Length)
                            {
                                similarOrder = i2 + 1;
                            }
                        }
                    }


                    if (similarLetters > maxSimilarLetters)
                    {
                        maxSimilarLetters = similarLetters;
                        maxSimilarOrder = similarOrder;
                        dataArray = convert2DRowTo1DArray(r);
                    }
                    else if (similarLetters == maxSimilarLetters && similarOrder > maxSimilarOrder)
                    {
                        maxSimilarOrder = similarOrder;
                        dataArray = convert2DRowTo1DArray(r);
                    }
                }
            }
            return dataArray;
        }
        */
        private static String[] convert2DRowTo1DArray(int row)
        {
            String[] tempData = new String[inventoryData.GetLength(1)];
            for(int c = 0; c < inventoryData.GetLength(1); c++)
            {
                tempData[c] = inventoryData[row, c];
            }
            return tempData;

        }

        private static String arrayToString(String[] data)
        {
            String temp = "";
            for(int i = 0; i < data.Length; i++)
            {
                if (data[i] == null)
                    temp += "null\n";
                else
                    temp += data[i] + "\n";
            }

            return temp;
        }
        private static String checkBotCall(String input)
        {
            String item = "";
            if (input.Length > botName.Length + 1)
            {
                if(input.Substring(0, botName.Length) == botName)
                    item = input.Substring(botName.Length + 1);
            }
            return item;
        }
    }
}
